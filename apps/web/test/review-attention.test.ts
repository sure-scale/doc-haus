import { expect, test } from "bun:test"
import { cellKey, hashQuestion, type GridCellData, type GridColumn } from "../src/api/ingest/grid"
import { matches, primaryAction, reviewAttention, type AttentionInput, type SourceVerdict } from "../src/api/review-attention"
import { NOT_ADDRESSED, needsFill } from "../src/api/review-recheck"

const law = { id: "law", question: "What is the governing law?" }
const term = { id: "term", question: "When does it expire?" }
const pay = { id: "pay", question: "What are the payment terms?" }
const rows = ["Supply.docx", "Lease.docx"]
const [supply, lease] = rows

const cited = (doc: string, column: GridColumn, extra: Partial<GridCellData> = {}): GridCellData => ({
  answer: "England and Wales",
  citation: { documentName: doc, docPath: `/matter/${doc}`, section: "10", excerpt: "English law governs.", charStart: 0, charEnd: 20 },
  status: "filled",
  questionHash: hashQuestion(column.question),
  ...extra,
})
const uncited = (column: GridColumn, answer: string, extra: Partial<GridCellData> = {}): GridCellData => ({ answer, status: "filled", questionHash: hashQuestion(column.question), ...extra })

function model(cells: Record<string, GridCellData>, verdicts: Record<string, SourceVerdict> = {}, extra: Partial<AttentionInput> = {}) {
  return reviewAttention({ rows, columns: [law, term], cells, source: (key) => verdicts[key] ?? "verified", ...extra })
}
// ReviewGrid's fillTargets(): every current cell a Fill would compute.
const fillTargets = (input: { rows: string[]; columns: GridColumn[]; cells: Record<string, GridCellData> }) =>
  input.rows.flatMap((doc) => input.columns.filter((col) => needsFill(input.cells[cellKey(doc, col.id)], hashQuestion(col.question)))).length

test("counts and denominators cover only the current documents and questions", () => {
  const result = model({
    [cellKey(supply, law.id)]: cited(supply, law, { status: "reviewed" }),
    [cellKey(supply, term.id)]: cited(supply, term),
    [cellKey(lease, law.id)]: uncited(law, NOT_ADDRESSED),
    // A removed document and a removed column leave cells behind.
    [cellKey("Removed.docx", law.id)]: uncited(law, "Unclear"),
    [cellKey(supply, "gone")]: uncited(law, "Unclear"),
  })
  expect(result.progress).toEqual({ documents: 2, questions: 2, answered: 3, reviewed: 1 })
  expect(result.cells.size).toBe(4)
  expect(result.attention).toBe(1)
  // Not found and No source are separate states, each with its own count.
  expect(result.counts.not_addressed).toBe(1)
  expect(result.counts.no_citation).toBe(0)
  expect(result.counts.unreviewed).toBe(2)
  expect(result.queue.map((entry) => entry.key)).toEqual([cellKey(lease, law.id), cellKey(supply, term.id)])
})

test("each cell shows the highest-priority marker and lists the rest", () => {
  const differs = (doc: string, column: GridColumn) => column.id === law.id ? `${doc} applies a different law.` : undefined
  const result = model({
    [cellKey(supply, law.id)]: cited(supply, law, { questionHash: "old" }),
    [cellKey(supply, term.id)]: uncited(term, NOT_ADDRESSED, { questionHash: "old" }),
    [cellKey(lease, law.id)]: uncited(law, "Net 30 days"),
    [cellKey(lease, term.id)]: cited(lease, term, { questionHash: "old" }),
  }, { [cellKey(supply, law.id)]: "other_document", [cellKey(lease, term.id)]: "not_found" }, { differs })
  const view = (doc: string, col: GridColumn) => {
    const entry = result.cells.get(cellKey(doc, col.id))!
    return [entry.marker?.word, ...entry.further.map((note) => note.word)]
  }
  expect(view(supply, law)).toEqual(["Cites another document", "Question changed", "Differs"])
  expect(view(supply, term)).toEqual(["Question changed", "Not found"])
  expect(view(lease, law)).toEqual(["No source", "Differs"])
  expect(view(lease, term)).toEqual(["Source changed", "Question changed"])
  expect(result.cells.get(cellKey(lease, law.id))!.further).toEqual([{ state: "differs", word: "Differs", reason: "Lease.docx applies a different law." }])
  expect(result.attention).toBe(4)
})

test("Review next takes markers by priority, then unreviewed cells, ties in table order", () => {
  const columns = [law, term, pay]
  const result = reviewAttention({
    rows,
    columns,
    cells: {
      [cellKey(supply, law.id)]: cited(supply, law),
      [cellKey(supply, term.id)]: uncited(term, "Five years"),
      [cellKey(supply, pay.id)]: cited(supply, pay, { status: "reviewed" }),
      [cellKey(lease, law.id)]: cited(lease, law),
      [cellKey(lease, term.id)]: cited(lease, term, { status: "reviewed" }),
    },
    source: (key) => ({ [cellKey(supply, pay.id)]: "missing", [cellKey(lease, law.id)]: "not_found" } as Record<string, SourceVerdict>)[key] ?? "verified",
  })
  expect(result.queue).toEqual([
    { key: cellKey(supply, pay.id), docName: supply, columnId: pay.id, reason: "The cited document is no longer in the matter. Unlock to refresh." },
    { key: cellKey(lease, law.id), docName: lease, columnId: law.id, reason: "The document no longer contains the cited excerpt. Re-run this cell to answer from the current text." },
    { key: cellKey(supply, term.id), docName: supply, columnId: term.id, reason: "The answer has no cited passage in this document." },
    { key: cellKey(supply, law.id), docName: supply, columnId: law.id, reason: "Not reviewed yet." },
  ])
})

test("a reviewed cell whose question changed is flagged, not filled and not unreviewed", () => {
  const cells = Object.fromEntries(rows.flatMap((doc) => [law, term].map((col) => [cellKey(doc, col.id), cited(doc, col, { status: "reviewed" })])))
  cells[cellKey(supply, law.id)] = { ...cells[cellKey(supply, law.id)], questionHash: hashQuestion("What law applies?") }
  const result = model(cells)
  const entry = result.cells.get(cellKey(supply, law.id))!
  expect(entry.marker).toEqual({ state: "question_changed", word: "Question changed", reason: "Question changed after review. Unlock to refresh." })
  expect(result.attention).toBe(1)
  expect(result.counts.unreviewed).toBe(0)
  expect(matches(entry, ["unreviewed"], "")).toBe(false)
  const pending = fillTargets({ rows, columns: [law, term], cells })
  expect(pending).toBe(0)
  expect(primaryAction(result.progress, pending)).toEqual({ kind: "export" })
  expect(result.queue.map((item) => item.key)).toEqual([cellKey(supply, law.id)])
})

test("a source still checking or not checked is shown but is not attention", () => {
  const cells = { [cellKey(supply, law.id)]: cited(supply, law), [cellKey(lease, law.id)]: cited(lease, law, { status: "reviewed" }) }
  for (const checkingSources of [true, false]) {
    const result = model(cells, {}, { source: () => undefined, checkingSources })
    const entry = result.cells.get(cellKey(supply, law.id))!
    expect(entry.marker).toBeUndefined()
    expect(entry.states).toContain("source_unchecked")
    expect(entry.further.map((note) => note.word)).toEqual([checkingSources ? "Checking source" : "Source not checked"])
    expect(result.attention).toBe(0)
    expect(result.queue).toEqual([{ key: cellKey(supply, law.id), docName: supply, columnId: law.id, reason: "Not reviewed yet." }])
  }
})

test("the primary action moves from start to fill to review next to export", () => {
  const action = (columns: GridColumn[], cells: Record<string, GridCellData>) =>
    primaryAction(reviewAttention({ rows, columns, cells, source: () => "verified" }).progress, fillTargets({ rows, columns, cells }))
  const all = (status: GridCellData["status"]) => Object.fromEntries(rows.flatMap((doc) => [law, term].map((col) => [cellKey(doc, col.id), cited(doc, col, { status })])))
  expect(action([], {})).toEqual({ kind: "start" })
  expect(action([law, term], {})).toEqual({ kind: "fill", count: 4 })
  // A negative cell not yet re-checked is still a fill target.
  expect(action([law, term], { ...all("filled"), [cellKey(lease, term.id)]: uncited(term, NOT_ADDRESSED) })).toEqual({ kind: "fill", count: 1 })
  expect(action([law, term], all("filled"))).toEqual({ kind: "reviewNext" })
  expect(action([law, term], all("reviewed"))).toEqual({ kind: "export" })
})

test("without documents there is no primary action, with or without questions", () => {
  // Export, Fill and a starter's questions all need documents; the table asks for them instead.
  for (const columns of [[], [law, term]])
    expect(primaryAction(reviewAttention({ rows: [], columns, cells: {}, source: () => "verified" }).progress, 0)).toBeUndefined()
})

test("a cell shows when it has any chosen state and matches the search text", () => {
  const result = model({ [cellKey(lease, law.id)]: uncited(law, "Net 30 days"), [cellKey(supply, law.id)]: cited(supply, law) }, { [cellKey(supply, law.id)]: "not_found" })
  const entry = result.cells.get(cellKey(lease, law.id))!
  const changed = result.cells.get(cellKey(supply, law.id))!
  expect(matches(entry, ["no_citation"], "LEASE")).toBe(true)
  expect(matches(entry, ["no_citation"], "governing")).toBe(true)
  expect(matches(entry, [], "net 30")).toBe(true)
  expect(matches(entry, ["no_citation"], "supply")).toBe(false)
  expect(matches(entry, ["source_changed"], "")).toBe(false)
  // Several states widen the view: either one shows the cell.
  for (const cell of [entry, changed]) expect(matches(cell, ["source_changed", "no_citation"], "")).toBe(true)
  expect(matches(changed, ["source_changed", "no_citation"], "lease")).toBe(false)
  // An empty table cell has neither state.
  expect(matches(result.cells.get(cellKey(lease, term.id))!, ["source_changed", "no_citation"], "")).toBe(false)
})
