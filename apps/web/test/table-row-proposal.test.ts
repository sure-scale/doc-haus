import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { Window } from "happy-dom"
import { DocumentProposalReview, DownloadActions } from "../src/components/DocumentViewer"
import type { Redline } from "../src/api/ingest/redlines"

const proposed = ["Literal | <b>  two spaces  & text", "", "A full value".repeat(40)]
const reference = ["Reference first", "Reference second", "Reference third"]
const insertion = { kind: "table_row_insert" as const, position: "after" as "before" | "after", table: 2, row: 4, reference_cells: reference, cells: proposed }
const row = (changes: Partial<Redline> = {}): Redline => ({
  id: 73, doc_name: "Synthetic schedule.docx", scope: "table_row", find_text: "compatibility summary",
  old_text: "Legacy reference must not be shown as deleted", new_text: "Joined cells must not be parsed",
  author: "Synthetic Reviewing Lawyer", review_comment: "Review the additional schedule entry.", created_at: 1,
  table_row: insertion,
  ...changes,
})
function render(proposal: Redline, state: { wordLocked?: boolean; busy?: number | "all" | "prepare" } = {}) {
  const page = new Window()
  page.document.body.innerHTML = renderToStaticMarkup(createElement(DocumentProposalReview, {
    redlines: [proposal], wordLocked: state.wordLocked ?? false, busy: state.busy,
    onAccept() {}, onReject() {}, onAcceptAll() {}, onRejectAll() {},
  }))
  const button = (name: string) => [...page.document.querySelectorAll("button")].find(button => button.textContent === name)!
  return { page, body: page.document.body, button, close: () => page.happyDOM.abort() }
}

test("row previews show exact escaped cells, accessible empty/new/reference semantics and the requested insertion order", () => {
  for (const position of ["before", "after"] as const) {
    const original = row({ table_row: { ...insertion, position } }), frozen = JSON.stringify(original)
    const view = render(original)
    try {
      const region = view.body.querySelector('[role="region"]')!, table = region.querySelector("table")!
      expect(region.getAttribute("tabindex")).toBe("0")
      expect(view.page.document.getElementById(region.getAttribute("aria-labelledby")!)!.textContent).toBe(`Insert a row ${position} row 4 in table 2`)
      expect(table.querySelector("caption")!.textContent).toContain(`Insert a row ${position} row 4 in table 2`)
      expect(table.querySelector("caption")!.textContent).toContain("The reference row is unchanged; the new row is proposed.")
      expect([...table.querySelectorAll('thead th[scope="col"]')].map(cell => cell.textContent)).toEqual(["Row", "Column 1", "Column 2", "Column 3"])
      expect([...table.querySelectorAll('tbody th[scope="row"]')].map(cell => cell.textContent)).toEqual(position === "before" ? ["New row", "Reference row"] : ["Reference row", "New row"])
      const inserted = table.querySelector(".table-row-new")!.querySelectorAll("td")
      expect(inserted[0]!.textContent).toBe(proposed[0]); expect(inserted[2]!.textContent).toBe(proposed[2])
      expect(inserted[1]!.querySelector('[aria-hidden="true"]')!.textContent).toBe("—")
      expect(inserted[1]!.querySelector(".sr-only")!.textContent).toBe("Empty cell")
      expect(table.querySelector("b")).toBeNull(); expect(view.body.querySelector("del,ins,.proposal-redline")).toBeNull()
      expect(view.body.textContent).not.toContain(original.old_text); expect(view.body.textContent).not.toContain(original.new_text)
      expect(view.button("Accept").disabled).toBe(false); expect(view.button("Accept all").disabled).toBe(false)
      expect(JSON.stringify(original)).toBe(frozen)
    } finally { view.close() }
  }
})

test("stale or unresolved row references retain frozen cells and position without guessing geometry or enabling either Accept", () => {
  for (const original of [
    row({ stale: true, stale_reason: "The source changed. Prepare a fresh change list." }),
    row({ table_row: { kind: "table_row_insert", position: "before", cells: proposed } }),
  ]) {
    const view = render(original)
    try {
      expect(view.body.textContent).toContain(proposed[0]); expect(view.body.textContent).toContain(proposed[2])
      expect(view.body.textContent).toContain(`Insert a row ${(original.table_row as typeof insertion).position} the reference row`)
      expect(view.body.textContent).toContain("Reference location needs refreshing.")
      expect(view.body.textContent).not.toContain("table 2"); expect(view.body.textContent).not.toContain("row 4")
      expect(view.body.textContent).not.toContain(reference[0]); expect(view.body.querySelector(".proposal-redline,del")).toBeNull()
      if (original.stale_reason) expect(view.body.textContent).toContain(original.stale_reason)
      expect(view.button("Accept").disabled).toBe(true); expect(view.button("Accept all").disabled).toBe(true)
      expect(view.button("Reject").disabled).toBe(false); expect(view.button("Reject all").disabled).toBe(false)
      const download = renderToStaticMarkup(createElement(DownloadActions, { redlines: [original], preview: false, onDownload() {} }))
      const downloadPage = new Window(); downloadPage.document.body.innerHTML = download
      expect(downloadPage.document.querySelector("button")!.textContent).toBe("Download")
      expect(downloadPage.document.querySelector("button")!.disabled).toBe(false)
      downloadPage.happyDOM.abort()
    } finally { view.close() }
  }
})

test("missing, unknown and disguised structural previews never fall through to text diffs or bulk acceptance", () => {
  for (const original of [
    row({ table_row: undefined }),
    row({ table_row: { ...row().table_row!, kind: "table_row_delete" } as unknown as Redline["table_row"] }),
    row({ scope: "phrase" }),
    row({ scope: "phrase", table_row: undefined, target_json: JSON.stringify({ version: 3, kind: "table_row_insert" }), stale: true }),
  ]) {
    const view = render(original)
    try {
      expect(view.body.textContent).toContain("Table row preview unavailable")
      expect(view.body.textContent).not.toContain(original.old_text); expect(view.body.textContent).not.toContain(original.new_text)
      expect(view.body.querySelector(".proposal-redline,del,ins")).toBeNull()
      expect(view.button("Accept").disabled).toBe(true); expect(view.button("Accept all").disabled).toBe(true)
      expect(view.button("Reject").disabled).toBe(false)
    } finally { view.close() }
  }
})

const deletedCells = ["Source | <b>  exact & unchanged", "", "  ", "Last original cell."]
const deletion = (changes: Partial<Redline> = {}): Redline => row({ table_row: { kind: "table_row_delete", table: 3, row: 2, cells: deletedCells }, ...changes })

test("deletion shows the complete exact row, explicit removal and separate empty/whitespace annotations", () => {
  const original = deletion(), frozen = JSON.stringify(original), view = render(original)
  try {
    expect(view.body.querySelector(".table-row-title")!.textContent).toBe("Delete row 2 in table 3")
    const region = view.body.querySelector('[role="region"]')!, table = region.querySelector("table")!
    expect(region.getAttribute("tabindex")).toBe("0")
    expect(table.querySelector("caption")!.textContent).toBe("Delete row 2 in table 3. The complete row shown will be removed.")
    expect(table.querySelector('tbody th[scope="row"]')!.textContent).toBe("Row to delete")
    expect([...table.querySelectorAll(".table-row-cell-text")].map(cell => cell.textContent)).toEqual(deletedCells)
    const displayedCells = table.querySelectorAll("td")
    expect(displayedCells[1]!.querySelector(".sr-only")!.textContent).toBe("Empty cell")
    expect(displayedCells[2]!.querySelector(".table-row-cell-text")!.textContent).toBe("  ")
    expect(displayedCells[2]!.querySelector(".table-row-cell-note")!.textContent).toBe("Whitespace only")
    expect(displayedCells[2]!.querySelector(".table-row-cell-text .table-row-cell-note")).toBeNull()
    expect(table.querySelector("b,.table-row-reference,.table-row-new")).toBeNull()
    expect(view.body.querySelector(".proposal-redline,ins,del")).toBeNull()
    expect(view.body.textContent).toContain(original.review_comment)
    expect(view.body.textContent).not.toContain(original.old_text); expect(view.body.textContent).not.toContain(original.new_text)
    expect(view.button("Accept").disabled).toBe(false); expect(view.button("Accept all").disabled).toBe(false)
    expect(JSON.stringify(original)).toBe(frozen)
  } finally { view.close() }
})

test("the same deletion becomes non-acceptable when stale while retaining frozen cells and hiding cached coordinates", () => {
  const original = deletion(), current = render(original)
  try {
    expect(current.button("Accept").disabled).toBe(false); expect(current.button("Accept all").disabled).toBe(false)
    expect(current.body.textContent).toContain("Delete row 2 in table 3")
  } finally { current.close() }
  for (const stale of [
    { ...original, stale: true, stale_reason: "The source changed. Prepare a fresh change list." },
    deletion({ table_row: { kind: "table_row_delete", cells: deletedCells } }),
  ]) {
    const view = render(stale)
    try {
      expect([...view.body.querySelectorAll(".table-row-cell-text")].map(cell => cell.textContent)).toEqual(deletedCells)
      expect(view.body.querySelector(".table-row-title")!.textContent).toBe("Delete a row")
      expect(view.body.textContent).toContain("Location needs refreshing.")
      expect(view.body.textContent).not.toContain("row 2"); expect(view.body.textContent).not.toContain("table 3")
      if (stale.stale_reason) expect(view.body.textContent).toContain(stale.stale_reason)
      expect(view.button("Accept").disabled).toBe(true); expect(view.button("Accept all").disabled).toBe(true)
      expect(view.button("Reject").disabled).toBe(false); expect(view.button("Reject all").disabled).toBe(false)
      const page = new Window()
      page.document.body.innerHTML = renderToStaticMarkup(createElement(DownloadActions, { redlines: [stale], preview: false, onDownload() {} }))
      expect(page.document.querySelector("button")!.textContent).toBe("Download"); expect(page.document.querySelector("button")!.disabled).toBe(false)
      page.happyDOM.abort()
    } finally { view.close() }
  }
})

test("a disguised deletion marker never becomes a text diff, and insertion fields cannot turn a deletion into another operation", () => {
  for (const original of [
    row({ scope: "phrase", table_row: undefined, target_json: JSON.stringify({ kind: "table_row_delete" }) }),
    row({ scope: "clause", table_row: undefined, target_json: JSON.stringify({ version: "invalid", kind: "table_row_delete" }) }),
    deletion({ table_row: { kind: "table_row_delete", cells: deletedCells, table: 3, row: 2, position: "after" } as unknown as Redline["table_row"] }),
    deletion({ table_row: { kind: "table_row_delete", cells: deletedCells, table: 3, row: 2, reference_cells: reference } as unknown as Redline["table_row"] }),
  ]) {
    const view = render(original)
    try {
      expect(view.body.textContent).toContain("Table row preview unavailable")
      expect(view.body.querySelector(".proposal-redline,.table-row-new,.table-row-reference,ins,del")).toBeNull()
      expect(view.body.textContent).not.toContain(original.old_text); expect(view.body.textContent).not.toContain(original.new_text)
      expect(view.button("Accept").disabled).toBe(true); expect(view.button("Accept all").disabled).toBe(true)
      expect(view.button("Reject").disabled).toBe(false)
    } finally { view.close() }
  }
})

test("ordinary text cards keep their diff and every row action retains existing busy and Word-lease restrictions", () => {
  const ordinary = row({ scope: "phrase", table_row: undefined, old_text: "Payment in thirty days.", new_text: "Payment in forty-five days." })
  const control = render(ordinary)
  try {
    expect(control.body.querySelector(".proposal-redline")).not.toBeNull()
    expect(control.body.querySelector("del")!.textContent).toBe("Deleted: thirty")
    expect(control.body.querySelector("ins")!.textContent).toBe("Inserted: forty-five")
    expect(control.button("Accept").disabled).toBe(false); expect(control.button("Accept all").disabled).toBe(false)
  } finally { control.close() }
  for (const state of [{ wordLocked: true }, { busy: 73 }, { busy: "all" as const }, { busy: "prepare" as const }]) {
    const view = render(row(), state)
    try { expect([...view.body.querySelectorAll("button")].every(button => button.disabled)).toBe(true) }
    finally { view.close() }
  }
})
