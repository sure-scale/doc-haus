import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { Window } from "happy-dom"
import { DocumentProposalReview, DownloadActions } from "../src/components/DocumentViewer"
import type { Redline } from "../src/api/ingest/redlines"

type Spacing = NonNullable<Redline["paragraph_spacing"]>
const text = "The  Supplier <b>keeps</b> these exact words & spaces."
const body: Spacing = { kind: "paragraph_spacing_after", paragraph_text: text, before_direct_after: { kind: "absent" }, after_direct_after_twips: "120", row_context: null, table_context: null }
const proposal = (paragraph_spacing: Spacing | undefined = body, changes: Partial<Redline> = {}): Redline => ({
  id: 81, doc_name: "Synthetic spacing.docx", scope: "paragraph_spacing", paragraph_spacing,
  find_text: "compatibility summary", old_text: "Legacy old summary must not display", new_text: "Legacy new summary must not display",
  author: "Synthetic Reviewing Lawyer", review_comment: "Use the requested direct spacing after this paragraph.", created_at: 1,
  ...changes,
})
function render(value: Redline) {
  const page = new Window()
  page.document.body.innerHTML = renderToStaticMarkup(createElement(DocumentProposalReview, {
    redlines: [value], wordLocked: false, onAccept() {}, onReject() {}, onAcceptAll() {}, onRejectAll() {},
  }))
  const button = (name: string) => [...page.document.querySelectorAll("button")].find(button => button.textContent === name)!
  return { page, root: page.document.body, button, close: () => page.happyDOM.abort() }
}
function acceptance(view: ReturnType<typeof render>, accepted: boolean) {
  expect(view.button("Accept").disabled).toBe(!accepted); expect(view.button("Accept all").disabled).toBe(!accepted)
  expect(view.button("Reject").disabled).toBe(false); expect(view.button("Reject all").disabled).toBe(false)
}

test("direct before/after distinguishes unset and zero, preserves exact paragraph text and formats large twips without numeric rounding", () => {
  for (const [before, displayed] of [
    [{ kind: "absent" }, "Not set"],
    [{ kind: "twips", raw: "0", normalized: "0" }, "0 pt"],
    [{ kind: "twips", raw: "0080", normalized: "80" }, "4 pt"],
    [{ kind: "twips", raw: "9007199254740993", normalized: "9007199254740993" }, "450359962737049.65 pt"],
  ] as const) {
    const original = proposal({ ...body, before_direct_after: before }), snapshot = JSON.stringify(original), view = render(original)
    try {
      expect(view.root.querySelector(".paragraph-spacing-text .paragraph-spacing-literal")!.textContent).toBe(text)
      expect(view.root.querySelector("b,.proposal-redline,ins,del,.table-row-proposal")).toBeNull()
      expect([...view.root.querySelectorAll(".paragraph-spacing-change dd")].map(value => value.textContent)).toEqual([displayed, "6 pt"])
      expect(view.root.querySelector('[role="group"]')!.getAttribute("aria-label")).toBe("Direct spacing after")
      expect(view.root.textContent).toContain("Wording unchanged · direct setting only")
      expect(view.root.querySelector(".paragraph-spacing-unavailable")).toBeNull()
      expect(view.root.textContent).not.toContain(original.old_text); expect(view.root.textContent).not.toContain(original.new_text)
      acceptance(view, true); expect(JSON.stringify(original)).toBe(snapshot)
    } finally { view.close() }
  }
  const zero = render(proposal({ ...body, after_direct_after_twips: "0" }))
  try { expect([...zero.root.querySelectorAll(".paragraph-spacing-change dd")].map(value => value.textContent)).toEqual(["Not set", "0 pt"]); acceptance(zero, true) }
  finally { zero.close() }
})

const contextParagraph = "These literal words mention Table 5, Row 3, Column 2."
const row = ["A  label", contextParagraph, "", "  "]
const table: Spacing = {
  ...body, paragraph_text: contextParagraph, row_context: row,
  table_context: { table: 5, row: 3, column: 2, row_span: 1, column_span: 1, vertical_merge: "none", row_context: row,
    row_cells: row.map((text, index) => ({ column: index + 1, column_span: 1, row_span: 1, vertical_merge: "none", text })) },
}

test("fresh table context is complete and accessible; the same stale target keeps original text without generated coordinates", () => {
  const current = render(proposal(table))
  try {
    expect(current.root.querySelector(".paragraph-spacing-location")!.textContent).toBe("Table 5 · Row 3 · Column 2")
    expect(current.root.querySelector("summary")!.textContent).toBe("Table context")
    const region = current.root.querySelector('.paragraph-spacing-scroll[role="region"]')!
    expect(region.getAttribute("tabindex")).toBe("0")
    expect([...region.querySelectorAll('th[scope="col"]')].map(cell => cell.textContent)).toEqual(row.map(() => "Cell text"))
    expect([...region.querySelectorAll("td .paragraph-spacing-literal")].map(cell => cell.textContent)).toEqual(row)
    expect(region.querySelectorAll(".paragraph-spacing-note")[0]!.textContent).toBe("Empty cell")
    expect(region.querySelectorAll(".paragraph-spacing-note")[1]!.textContent).toBe("Whitespace only")
    acceptance(current, true)
  } finally { current.close() }
  const frozen: Spacing = { kind: table.kind, paragraph_text: table.paragraph_text, before_direct_after: table.before_direct_after, after_direct_after_twips: table.after_direct_after_twips, row_context: table.row_context }
  for (const original of [proposal(table, { stale: true, stale_reason: "The source changed." }), proposal(frozen, { stale: true })]) {
    const snapshot = JSON.stringify(original), view = render(original)
    try {
      expect(view.root.querySelector(".paragraph-spacing-text .paragraph-spacing-literal")!.textContent).toBe(contextParagraph)
      expect(view.root.querySelector(".paragraph-spacing-location")).toBeNull()
      expect(view.root.querySelector("summary")!.textContent).toBe("Original table context")
      expect(view.root.querySelector(".paragraph-spacing-unavailable")!.textContent).toBe("Location needs refreshing.")
      expect([...view.root.querySelectorAll("td .paragraph-spacing-literal")].map(cell => cell.textContent)).toEqual(row)
      const labels = [...view.root.querySelectorAll("[aria-label], [title], caption, summary")].map(element => `${element.getAttribute("aria-label") ?? ""}${element.getAttribute("title") ?? ""}${["CAPTION", "SUMMARY"].includes(element.tagName) ? element.textContent : ""}`).join(" ")
      expect(labels).not.toContain("Table 5"); expect(labels).not.toContain("Row 3"); expect(labels).not.toContain("Column 2")
      acceptance(view, false); expect(JSON.stringify(original)).toBe(snapshot)
      const page = new Window(); page.document.body.innerHTML = renderToStaticMarkup(createElement(DownloadActions, { redlines: [original], preview: false, onDownload() {} }))
      expect(page.document.querySelector("button")!.textContent).toBe("Download"); expect(page.document.querySelector("button")!.disabled).toBe(false)
      page.happyDOM.abort()
    } finally { view.close() }
  }
})

test("missing body context and unknown table geometry cannot masquerade as resolved locations", () => {
  const missingBody: Spacing = { kind: body.kind, paragraph_text: "  ", before_direct_after: body.before_direct_after, after_direct_after_twips: "1", row_context: null }
  for (const original of [proposal(missingBody), proposal({ ...table, table_context: { ...table.table_context!, unavailable_reason: "unsupported_table_geometry" } })]) {
    const view = render(original)
    try {
      expect(view.root.querySelector(".paragraph-spacing-unavailable")!.textContent).toBe("Location needs refreshing.")
      expect(view.root.querySelector(".paragraph-spacing-location")).toBeNull(); acceptance(view, false)
    } finally { view.close() }
  }
  const view = render(proposal(missingBody))
  try {
    expect(view.root.querySelector(".paragraph-spacing-text .paragraph-spacing-literal")!.textContent).toBe("  ")
    expect(view.root.querySelector(".paragraph-spacing-text .paragraph-spacing-note")!.textContent).toBe("Whitespace only")
    expect([...view.root.querySelectorAll(".paragraph-spacing-change dd")].map(value => value.textContent)).toEqual(["Not set", "0.05 pt"])
  } finally { view.close() }
})

test("inconsistent, noncanonical, out-of-range or no-op direct values are unavailable to both acceptance actions", () => {
  for (const invalid of [
    { ...body, before_direct_after: { kind: "twips", raw: "0080", normalized: "81" } },
    { ...body, after_direct_after_twips: "00" },
    { ...body, after_direct_after_twips: "961" },
    { ...body, before_direct_after: { kind: "twips", raw: "00120", normalized: "120" } },
  ] as Spacing[]) {
    const view = render(proposal(invalid))
    try {
      expect(view.root.querySelector(".paragraph-spacing-unavailable")!.textContent).toContain("Spacing preview unavailable")
      expect(view.root.querySelector(".paragraph-spacing-change,.proposal-redline,.table-row-proposal")).toBeNull(); acceptance(view, false)
    } finally { view.close() }
  }
})

test("v4 and spacing-kind markers never become ordinary or row diffs when scope or display data is malformed", () => {
  for (const original of [
    proposal(undefined, { paragraph_spacing: undefined, scope: "phrase", target_json: JSON.stringify({ version: 4, kind: "unknown" }) }),
    proposal(undefined, { paragraph_spacing: undefined, scope: "clause", target_json: JSON.stringify({ kind: "paragraph_spacing_after" }) }),
    proposal(body, { scope: "phrase" }),
    proposal(body, { target_json: JSON.stringify({ version: 3, kind: "table_row_insert" }) }),
  ]) {
    const view = render(original)
    try {
      expect(view.root.querySelector(".paragraph-spacing-unavailable")!.textContent).toContain("Spacing preview unavailable")
      expect(view.root.querySelector(".proposal-redline,.table-row-proposal,ins,del")).toBeNull()
      expect(view.root.textContent).not.toContain(original.old_text); expect(view.root.textContent).not.toContain(original.new_text); acceptance(view, false)
    } finally { view.close() }
  }
})
