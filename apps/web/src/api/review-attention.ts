import { cellKey, hashQuestion, type GridCellData, type GridColumn, type SourceCheck } from "./ingest/grid"
import { NOT_ADDRESSED } from "./review-recheck"

// What a review table needs from the lawyer, read from the current documents ×
// questions only: each cell's states and its one marker, the Needs attention
// count, progress, the Review next order and the toolbar's primary action.

export type SourceVerdict = SourceCheck | "other_document"
export type MarkerState = "other_document" | "source_changed" | "source_missing" | "question_changed" | "no_citation" | "not_addressed" | "differs"
export type CellState = MarkerState | "unfilled" | "unreviewed" | "source_unchecked"
export type AttentionFilter = MarkerState | "unreviewed"
export type StateNote<S extends CellState = CellState> = { state: S; word: string; reason: string }
export type CellAttention = {
  key: string
  docName: string
  columnId: string
  question: string
  cell?: GridCellData
  states: CellState[]
  marker?: StateNote<MarkerState>
  // Every other marker or source state on the cell, for its tooltip.
  further: StateNote[]
}
export type QueueEntry = { key: string; docName: string; columnId: string; reason: string }
export type ReviewProgress = { documents: number; questions: number; answered: number; reviewed: number }
export type ReviewAttention = {
  cells: Map<string, CellAttention>
  attention: number
  counts: Record<AttentionFilter, number>
  progress: ReviewProgress
  queue: QueueEntry[]
}
export type PrimaryAction = { kind: "start" } | { kind: "fill"; count: number } | { kind: "reviewNext" } | { kind: "export" }

export type AttentionInput = {
  rows: string[]
  columns: GridColumn[]
  cells: Record<string, GridCellData>
  /** The cited cell's verdict; undefined while checking or when it could not be checked. */
  source: (key: string, cell: GridCellData) => SourceVerdict | undefined
  checkingSources?: boolean
  /** A fresh comparison's reason when this cell's answer differs (`GridColumn.comparison`). */
  differs?: (docName: string, column: GridColumn) => string | undefined
}

const RANK: Record<MarkerState, number> = { other_document: 0, source_changed: 1, source_missing: 1, question_changed: 2, no_citation: 3, not_addressed: 3, differs: 4 }
const UNREVIEWED_RANK = 5
const FILTERS: AttentionFilter[] = ["other_document", "source_changed", "source_missing", "question_changed", "no_citation", "not_addressed", "differs", "unreviewed"]
const isMarker = (state: CellState): state is MarkerState => state in RANK
const notAddressed = (answer: string) => answer.trim().replace(/\.$/, "").toLowerCase() === NOT_ADDRESSED.toLowerCase()

function notes(cell: GridCellData | undefined, checking: boolean, differs?: string): Record<CellState, [string, string]> {
  const reviewed = cell?.status === "reviewed"
  return {
    unfilled: ["Not filled", "No answer yet."],
    unreviewed: ["Not reviewed", "Not reviewed yet."],
    other_document: ["Cites another document", `The cited passage is from another document, so it does not support this answer. ${reviewed ? "Unlock to refresh." : "Re-run this cell to answer from this document."}`],
    source_changed: ["Source changed", `The document no longer contains the cited excerpt. ${reviewed ? "Unlock to refresh." : "Re-run this cell to answer from the current text."}`],
    source_missing: ["Source missing", `The cited document is no longer in the matter. ${reviewed ? "Unlock to refresh." : "Re-run this cell to answer from this document."}`],
    source_unchecked: checking ? ["Checking source", "Checking the excerpt against the document."] : ["Source not checked", "The excerpt could not be checked against the document."],
    question_changed: ["Question changed", reviewed ? "The question was edited after this cell was reviewed. Unlock to refresh." : "Question changed since this answer. Re-run this cell to update it."],
    no_citation: ["No source", "The answer has no cited passage in this document."],
    not_addressed: ["Not found", "The document does not appear to address this question."],
    differs: ["Differs", differs || "This answer differs from the other documents."],
  }
}

// States in marker priority order, so the first marker-class state is the marker.
// No source, Not found and Differs describe the answer itself, so reviewing settles them;
// the others say the reviewed answer may no longer hold.
function cellStates(key: string, cell: GridCellData | undefined, questionHash: string, differs: string | undefined, input: AttentionInput): CellState[] {
  if (!cell) return ["unfilled"]
  const open = cell.status === "filled"
  const states: CellState[] = open ? ["unreviewed"] : []
  if (cell.citation) {
    const verdict = input.source(key, cell)
    if (verdict === "other_document") states.push("other_document")
    else if (verdict === "not_found") states.push("source_changed")
    else if (verdict === "missing") states.push("source_missing")
    else if (!verdict) states.push("source_unchecked")
  }
  if (cell.questionHash !== questionHash) states.push("question_changed")
  if (open && !cell.citation) states.push(notAddressed(cell.answer) ? "not_addressed" : "no_citation")
  if (open && differs !== undefined) states.push("differs")
  return states
}

export function reviewAttention(input: AttentionInput): ReviewAttention {
  const cells = new Map<string, CellAttention>()
  const counts = Object.fromEntries(FILTERS.map((filter) => [filter, 0])) as Record<AttentionFilter, number>
  const progress: ReviewProgress = { documents: input.rows.length, questions: input.columns.length, answered: 0, reviewed: 0 }
  const ranked: Array<{ rank: number; entry: QueueEntry }> = []
  let attention = 0
  const hashes = input.columns.map((column) => hashQuestion(column.question))
  for (const docName of input.rows) {
    input.columns.forEach((column, index) => {
      const key = cellKey(docName, column.id)
      const cell = input.cells[key]
      const differs = cell?.status === "filled" ? input.differs?.(docName, column) : undefined
      const states = cellStates(key, cell, hashes[index], differs, input)
      const text = notes(cell, !!input.checkingSources, differs)
      const note = <S extends CellState>(state: S): StateNote<S> => ({ state, word: text[state][0], reason: text[state][1] })
      const markerState = states.find(isMarker)
      const marker = markerState && note(markerState)
      const further = states.filter((state) => state !== markerState && (isMarker(state) || state === "source_unchecked")).map(note)
      cells.set(key, { key, docName, columnId: column.id, question: column.question, cell, states, marker, further })
      for (const state of states) if (state in counts) counts[state as AttentionFilter]++
      if (cell) progress.answered++
      if (cell?.status === "reviewed") progress.reviewed++
      if (marker) attention++
      if (marker || states.includes("unreviewed"))
        ranked.push({ rank: marker ? RANK[marker.state] : UNREVIEWED_RANK, entry: { key, docName, columnId: column.id, reason: marker?.reason ?? text.unreviewed[1] } })
    })
  }
  // Built in row then column order; the sort is stable, so ties keep it.
  const queue = ranked.sort((a, b) => a.rank - b.rank).map(({ entry }) => entry)
  return { cells, attention, counts, progress, queue }
}

/** Each filter's word and meaning for the legend, in the cells' own vocabulary: the reason's first sentence, without a cell's next step. */
export function legend(): Array<StateNote<AttentionFilter>> {
  const plain = notes(undefined, false)
  const meaning = (reason: string) => reason.split(/(?<=\.) /)[0]
  return FILTERS.map((state) => ({ state, word: plain[state][0], reason: meaning(plain[state][1]) }))
}

/**
 * `fillTargets` is the count of cells a Fill would compute (ReviewGrid fillTargets).
 * Without documents nothing can be added, filled, reviewed or exported, so there
 * is no primary action; the table's empty text asks for documents.
 */
export function primaryAction(progress: ReviewProgress, fillTargets: number): PrimaryAction | undefined {
  if (!progress.documents) return undefined
  if (!progress.questions) return { kind: "start" }
  if (fillTargets > 0) return { kind: "fill", count: fillTargets }
  if (progress.answered > progress.reviewed) return { kind: "reviewNext" }
  return { kind: "export" }
}

/**
 * Any of the chosen states (none chosen: every cell) AND case-insensitive text
 * over document name, question and answer.
 */
export function matches(entry: CellAttention, filters: readonly AttentionFilter[], query: string) {
  if (filters.length && !filters.some((filter) => entry.states.includes(filter))) return false
  const text = query.trim().toLowerCase()
  return !text || [entry.docName, entry.question, entry.cell?.answer ?? ""].some((value) => value.toLowerCase().includes(text))
}
