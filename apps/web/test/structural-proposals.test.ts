import { describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import type { Redline } from "../src/api/ingest/redlines"
import { paragraphInsertionMarker, reviewableParagraphInsertion } from "../src/components/ParagraphInsertionProposal"
import { paragraphSpacingMarker, reviewableParagraphSpacing } from "../src/components/ParagraphSpacingProposal"
import { ordinaryProposal, ProposalPreview, reviewableProposal } from "../src/components/TableRowProposal"

type Insertion = NonNullable<Redline["paragraph_insert"]>
type Spacing = NonNullable<Redline["paragraph_spacing"]>

const insertionReference = "The  existing paragraph says <a>literal</a> & keeps every word."
const insertionText = `New <b>literal</b> | &  repeated spaces ${"X".repeat(600)}`
const insertionLocation = {
  body_paragraph_number: 8,
  headings: ["Service terms", "Payment"],
  previous_paragraph: "The  preceding paragraph.",
  next_paragraph: "The following paragraph.",
} satisfies Insertion["location"]
const insertionValue: Insertion = {
  kind: "paragraph_insert",
  reference_text: insertionReference,
  position: "after",
  text: insertionText,
  location: insertionLocation,
}
const insertionProposal = (changes: Partial<Redline> = {}): Redline => ({
  id: 91,
  doc_name: "Synthetic insertion.docx",
  scope: "paragraph_insert",
  paragraph_insert: insertionValue,
  old_text: "Legacy old summary must not display",
  new_text: "Legacy joined summary must not display",
  find_text: "Compatibility summary",
  author: "Synthetic Reviewing Lawyer",
  review_comment: "Insert the requested standalone paragraph.",
  created_at: 1,
  ...changes,
})

const spacingText = "The  Supplier <b>keeps</b> these exact words & spaces."
const spacingBody: Spacing = {
  kind: "paragraph_spacing_after",
  paragraph_text: spacingText,
  before_direct_after: { kind: "absent" },
  after_direct_after_twips: "120",
  row_context: null,
  table_context: null,
}
const spacingRow = ["A  label", spacingText, "", "  "]
const spacingTable: Spacing = {
  ...spacingBody,
  row_context: spacingRow,
  table_context: {
    table: 5,
    row: 3,
    column: 2,
    row_span: 1,
    column_span: 1,
    vertical_merge: "none",
    row_context: spacingRow,
    row_cells: spacingRow.map((text, index) => ({
      column: index + 1,
      column_span: 1,
      row_span: 1,
      vertical_merge: "none",
      text,
    })),
  },
}
const spacingProposal = (paragraph_spacing: Spacing | undefined = spacingBody, changes: Partial<Redline> = {}): Redline => ({
  id: 81,
  doc_name: "Synthetic spacing.docx",
  scope: "paragraph_spacing",
  paragraph_spacing,
  find_text: "compatibility summary",
  old_text: "Legacy old summary must not display",
  new_text: "Legacy new summary must not display",
  author: "Synthetic Reviewing Lawyer",
  review_comment: "Use direct spacing after this paragraph.",
  created_at: 1,
  ...changes,
})

const rowCells = ["Literal | <b>  two spaces  & text", "", "A full value".repeat(40)]
const referenceCells = ["Reference first", "Reference second", "Reference third"]
const rowInsert = {
  kind: "table_row_insert" as const,
  position: "after" as "before" | "after",
  table: 2,
  row: 4,
  reference_cells: referenceCells,
  cells: rowCells,
}
const rowProposal = (changes: Partial<Redline> = {}): Redline => ({
  id: 73,
  doc_name: "Synthetic schedule.docx",
  scope: "table_row",
  find_text: "compatibility summary",
  old_text: "Legacy reference must not be shown as deleted",
  new_text: "Joined cells must not be parsed",
  author: "Synthetic Reviewing Lawyer",
  review_comment: "Review the additional schedule entry.",
  created_at: 1,
  table_row: rowInsert,
  ...changes,
})
const rowDelete = (changes: Partial<Redline> = {}): Redline =>
  rowProposal({ table_row: { kind: "table_row_delete", table: 3, row: 2, cells: rowCells }, ...changes })

describe("paragraph insertion proposals", () => {
  test("validates structured bindings, rejects stale or mismatched targets, and never falls through to ordinary diffs", () => {
    const valid = insertionProposal()
    expect(paragraphInsertionMarker(valid)).toBe(true)
    expect(ordinaryProposal(valid)).toBe(false)
    expect(reviewableParagraphInsertion(valid)).toBe(true)
    expect(reviewableProposal(valid)).toBe(true)

    const matchingTarget = insertionProposal({
      target_json: JSON.stringify({ version: 5, ...insertionValue, target_handle: "opaque-source-reference" }),
    })
    expect(reviewableProposal(matchingTarget)).toBe(true)
    expect(reviewableProposal(insertionProposal({ stale: true }))).toBe(false)

    const malformed: Redline[] = [
      insertionProposal({ paragraph_insert: undefined }),
      insertionProposal({ scope: "phrase", paragraph_insert: undefined, target_json: JSON.stringify({ version: 5, kind: "unknown" }) }),
      insertionProposal({ scope: "clause", paragraph_insert: undefined, target_json: JSON.stringify({ kind: "paragraph_insert" }) }),
      insertionProposal({ paragraph_insert: { ...insertionValue, position: "middle" } as unknown as Insertion }),
      insertionProposal({ paragraph_insert: { ...insertionValue, text: "Two\nparagraphs" } }),
      insertionProposal({ paragraph_insert: { ...insertionValue, style: "Heading1" } as unknown as Insertion }),
      insertionProposal({ paragraph_insert: { ...insertionValue, location: undefined } as unknown as Insertion }),
      insertionProposal({ paragraph_insert: { ...insertionValue, location: { ...insertionLocation, body_paragraph_number: 0 } } }),
      insertionProposal({ paragraph_insert: { ...insertionValue, location: { ...insertionLocation, next_paragraph: 3 } } as unknown as Insertion }),
      insertionProposal({ target_json: JSON.stringify({ version: 5, ...insertionValue, text: "Different stored words." }) }),
      insertionProposal({ target_json: JSON.stringify({ version: 5, ...insertionValue, location: { ...insertionLocation, body_paragraph_number: 9 } }) }),
      insertionProposal({ target_json: JSON.stringify({ version: 5, ...insertionValue, location: { ...insertionLocation, headings: ["Other"] } }) }),
    ]
    for (const bad of malformed) {
      expect(paragraphInsertionMarker(bad)).toBe(true)
      expect(ordinaryProposal(bad)).toBe(false)
      expect(reviewableProposal(bad)).toBe(false)
    }
  })

  test("escapes literal HTML tags and never leaks legacy summary fields or opaque handles", () => {
    const proposal = insertionProposal({
      target_json: JSON.stringify({ version: 5, ...insertionValue, target_handle: "opaque-source-reference" }),
    })
    const html = renderToStaticMarkup(createElement(ProposalPreview, { proposal }))
    expect(html).toContain("&lt;a&gt;literal&lt;/a&gt;")
    expect(html).toContain("&lt;b&gt;literal&lt;/b&gt;")
    expect(html).not.toContain("<a>")
    expect(html).not.toContain("<b>")
    expect(html).not.toContain(proposal.old_text)
    expect(html).not.toContain(proposal.new_text)
    expect(html).not.toContain("opaque-source-reference")
  })
})

describe("paragraph spacing proposals", () => {
  test("validates canonical twips, table geometry, stale gating, and v4 marker isolation", () => {
    expect(paragraphSpacingMarker(spacingProposal())).toBe(true)
    expect(ordinaryProposal(spacingProposal())).toBe(false)
    expect(reviewableParagraphSpacing(spacingProposal())).toBe(true)
    expect(reviewableProposal(spacingProposal(spacingTable))).toBe(true)
    expect(reviewableProposal(spacingProposal(spacingTable, { stale: true }))).toBe(false)
    expect(reviewableProposal(spacingProposal({
      ...spacingBody,
      before_direct_after: { kind: "twips", raw: "9007199254740993", normalized: "9007199254740993" },
    }))).toBe(true)

    const missingBody: Spacing = {
      kind: spacingBody.kind,
      paragraph_text: "  ",
      before_direct_after: spacingBody.before_direct_after,
      after_direct_after_twips: "1",
      row_context: null,
    }
    const invalid: Redline[] = [
      spacingProposal(missingBody),
      spacingProposal({ ...spacingTable, table_context: { ...spacingTable.table_context!, unavailable_reason: "unsupported_table_geometry" } }),
      spacingProposal({ ...spacingBody, before_direct_after: { kind: "twips", raw: "0080", normalized: "81" } }),
      spacingProposal({ ...spacingBody, after_direct_after_twips: "00" }),
      spacingProposal({ ...spacingBody, after_direct_after_twips: "961" }),
      spacingProposal({ ...spacingBody, before_direct_after: { kind: "twips", raw: "00120", normalized: "120" } }),
      spacingProposal(undefined, { paragraph_spacing: undefined, scope: "phrase", target_json: JSON.stringify({ version: 4, kind: "unknown" }) }),
      spacingProposal(undefined, { paragraph_spacing: undefined, scope: "clause", target_json: JSON.stringify({ kind: "paragraph_spacing_after" }) }),
      spacingProposal(spacingBody, { scope: "phrase" }),
      spacingProposal(spacingBody, { target_json: JSON.stringify({ version: 3, kind: "table_row_insert" }) }),
    ]
    for (const bad of invalid) {
      expect(paragraphSpacingMarker(bad)).toBe(true)
      expect(ordinaryProposal(bad)).toBe(false)
      expect(reviewableProposal(bad)).toBe(false)
    }
  })

  test("formats large BigInt twips without rounding and escapes literal HTML tags", () => {
    const proposal = spacingProposal({
      ...spacingBody,
      before_direct_after: { kind: "twips", raw: "9007199254740993", normalized: "9007199254740993" },
    })
    const html = renderToStaticMarkup(createElement(ProposalPreview, { proposal }))
    expect(html).toContain("450359962737049.65 pt")
    expect(html).toContain("&lt;b&gt;keeps&lt;/b&gt;")
    expect(html).not.toContain("<b>")
    expect(html).not.toContain(proposal.old_text)
    expect(html).not.toContain(proposal.new_text)
  })
})

describe("table row proposals", () => {
  test("gates reviewability on located fresh rows and blocks disguised or contaminated v3 markers", () => {
    expect(ordinaryProposal(rowProposal())).toBe(false)
    expect(reviewableProposal(rowProposal())).toBe(true)
    expect(reviewableProposal(rowDelete())).toBe(true)

    const ordinary = rowProposal({ scope: "phrase", table_row: undefined, old_text: "thirty", new_text: "forty-five" })
    expect(ordinaryProposal(ordinary)).toBe(true)
    expect(reviewableProposal(ordinary)).toBe(true)

    const nonReviewable: Redline[] = [
      rowProposal({ stale: true }),
      rowProposal({ table_row: { kind: "table_row_insert", position: "before", cells: rowCells } }),
      rowDelete({ stale: true }),
      rowDelete({ table_row: { kind: "table_row_delete", cells: rowCells } }),
      rowProposal({ table_row: undefined }),
      rowProposal({ scope: "phrase" }),
      rowProposal({ scope: "phrase", table_row: undefined, target_json: JSON.stringify({ version: 3, kind: "table_row_insert" }) }),
      rowProposal({ scope: "phrase", table_row: undefined, target_json: JSON.stringify({ kind: "table_row_delete" }) }),
      rowProposal({ scope: "clause", table_row: undefined, target_json: JSON.stringify({ version: "invalid", kind: "table_row_delete" }) }),
      rowDelete({ table_row: { kind: "table_row_delete", cells: rowCells, table: 3, row: 2, position: "after" } as unknown as Redline["table_row"] }),
      rowDelete({ table_row: { kind: "table_row_delete", cells: rowCells, table: 3, row: 2, reference_cells: referenceCells } as unknown as Redline["table_row"] }),
    ]
    for (const bad of nonReviewable) {
      expect(ordinaryProposal(bad)).toBe(false)
      expect(reviewableProposal(bad)).toBe(false)
    }
  })

  test("escapes literal HTML in row cells and hides stale geometry without falling through to legacy summaries", () => {
    const html = renderToStaticMarkup(createElement(ProposalPreview, { proposal: rowProposal() }))
    expect(html).toContain("&lt;b&gt;")
    expect(html).not.toContain("<b>")
    expect(html).not.toContain(rowProposal().old_text)
    expect(html).not.toContain(rowProposal().new_text)

    const staleHtml = renderToStaticMarkup(createElement(ProposalPreview, { proposal: rowProposal({ stale: true }) }))
    expect(staleHtml).not.toContain(referenceCells[0])
    expect(staleHtml).not.toContain("table 2")
  })
})
