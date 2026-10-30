import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { applyDocxLocalization, createDocxLocalization, createDocxRevision, openRevision } from "./docx-amendments"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const cell = (content: string, properties = "") => `<w:tc><w:tcPr>${properties}</w:tcPr>${content}</w:tc>`
const row = (content: string, properties = "") => `<w:tr><w:trPr>${properties}</w:trPr>${content}</w:tr>`
const table = (rows: string, columns: number) => `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(columns)}</w:tblGrid>${rows}</w:tbl>`
function docx(content: string) {
  return zipSync({
    "[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${O}/styles" Target="styles.xml"/></Relationships>`),
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/document.xml": strToU8(`<w:document xmlns:w="${W}"><w:body>${content}${p("Outside table.")}<w:sectPr/></w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
}

function nested() {
  const inner = table(row(cell(p("Nested left")) + cell(p("Repeated phrase."))), 2)
  return docx(table(row(cell(p("Outer header")) + cell(p("Other header")))
    + row(cell(p("Repeated phrase.") + inner + p("Outer tail")) + cell(p("Outer peer"))), 2))
}

test("nested table context uses the nearest SDK table and direct current cell text", async () => {
  const source = nested(), before = new Uint8Array(source)
  const opened = await openRevision(source)
  try {
    const target = opened.paragraphs.find(item => item.match.text === "Nested left")!
    const cells = Object.entries(opened.session.project().anchorIndex)
      .filter(([, a]) => a.kind === "tbl")
      .flatMap(([id]) => opened.session.getTableMetadata(id).metadata!.rows.flatMap(row => row.cells))
      .filter(cell => cell.paragraphAnchors.some(p => p.id === target.match.enclosingAnchor.id))
    expect(cells).toHaveLength(1)
    expect([cells[0]!.rowIndex, cells[0]!.columnIndex]).toEqual([0, 0])
    const inner = opened.revision.paragraphs.find(item => item.text === "Nested left")!.table!
    expect([inner.table, inner.row, inner.column]).toEqual([2, 1, 1])
    expect(inner.row_context).toEqual(["Nested left", "Repeated phrase."])
    const outer = opened.revision.paragraphs.find(item => item.text === "Outer tail")!.table!
    expect([outer.table, outer.row, outer.column]).toEqual([1, 2, 1])
    expect(outer.row_context).toEqual(["Repeated phrase.\nOuter tail", "Outer peer"])
    expect(opened.revision.markdown).not.toContain("row 0")
  } finally { opened.session.close() }
  expect(source).toEqual(before)
})

test("merged cells and omitted grid columns use logical coordinates and explicit spans", async () => {
  const source = docx(table(row(cell(p("Merged"), '<w:gridSpan w:val="2"/>') + cell(p("Right")),
    '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>'), 5))
  const revision = await createDocxRevision(source)
  const right = revision.paragraphs.find(p => p.text === "Right")!.table!
  expect([right.table, right.row, right.column]).toEqual([1, 1, 4])
  expect(right.row_context).toEqual(["Merged", "Right"])
  expect(right.row_cells).toEqual([
    { column: 2, column_span: 2, row_span: 1, vertical_merge: "none", text: "Merged" },
    { column: 4, column_span: 1, row_span: 1, vertical_merge: "none", text: "Right" },
  ])
  const merged = revision.paragraphs.find(p => p.text === "Merged")!.table!
  expect([merged.column, merged.column_span]).toEqual([2, 2])
})

test("vertical merge context distinguishes the lead span and continuation", async () => {
  const source = docx(table(row(cell(p("Upper left"), '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') + cell(p("Upper right")))
    + row(cell(p(""), '<w:gridSpan w:val="2"/><w:vMerge/>') + cell(p("Lower right"))), 3))
  const revision = await createDocxRevision(source)
  const upper = revision.paragraphs.find(p => p.text === "Upper left")!.table!
  expect([upper.row, upper.column, upper.row_span, upper.column_span]).toEqual([1, 1, 2, 2])
  const lower = revision.paragraphs.find(p => p.text === "Lower right")!.table!
  expect([lower.row, lower.column]).toEqual([2, 3])
  expect(lower.row_cells?.[0]).toEqual({ column: 1, column_span: 2, row_span: 0, vertical_merge: "continue", text: "" })
})

test("pending table topology is explicit uncertainty and does not bypass amendment refusal", async () => {
  const source = docx(table(row(cell(p("Before"))) + row(cell(p("Inserted")),
    '<w:ins w:id="900" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"/>'), 1))
  const options = { pendingRevisions: "layer" as const }
  const revision = await createDocxRevision(source, options)
  expect(revision.paragraphs.find(p => p.text === "Before")!.table).toMatchObject({ table: 1, row: null, column: null, unavailable_reason: "pending_table_structure" })
  expect(revision.markdown).toContain("row and column unavailable")
  const packet = await createDocxLocalization(source, [{ id: "edit", before_quote: "Inserted", replacement: "Changed" }], {}, options)
  await expect(applyDocxLocalization(source, packet, [{ amendment_id: "edit", target_handle: packet.candidates[0]!.target_handle }], options))
    .rejects.toThrow("docx_pending_revision:edit")
})

test("corrected context keeps exact nested duplicate selection and source authority", async () => {
  const source = nested(), before = new Uint8Array(source)
  const packet = await createDocxLocalization(source, [{ id: "edit", before_quote: "Repeated phrase.", replacement: "Changed phrase." }])
  const inner = packet.candidates.find(c => c.paragraph.table?.table === 2)!
  expect(inner).toBeDefined()
  const output = await applyDocxLocalization(source, packet, [{ amendment_id: "edit", target_handle: inner.target_handle }])
  const revision = await createDocxRevision(output.bytes)
  expect(revision.paragraphs.find(p => p.table?.table === 1 && p.text === "Repeated phrase.")).toBeDefined()
  expect(revision.paragraphs.find(p => p.table?.table === 2 && p.text === "Changed phrase.")).toBeDefined()
  expect(output.report.package_structure_verified).toBe(true)
  expect(source).toEqual(before)
})

test("inline text revisions keep stable table coordinates and current row text", async () => {
  const changed = '<w:p><w:r><w:t xml:space="preserve">Visible </w:t></w:r><w:del w:id="1" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>removed</w:delText></w:r></w:del><w:ins w:id="2" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>'
  const source = docx(table(row(cell(changed) + cell(p("Peer"))), 2))
  const revision = await createDocxRevision(source, { pendingRevisions: "layer" })
  const view = revision.paragraphs.find(p => p.text === "Visible inserted")!.table!
  expect([view.row, view.column]).toEqual([1, 1])
  expect(view.row_context).toEqual(["Visible inserted", "Peer"])
})

test("unsupported legacy horizontal geometry stays a table with unavailable coordinates", async () => {
  const source = docx(table(row(cell(p("Legacy lead"), '<w:hMerge w:val="restart"/>')
    + cell(p("Legacy continuation"), '<w:hMerge/>')), 2))
  const revision = await createDocxRevision(source)
  expect(revision.paragraphs.find(p => p.text === "Legacy lead")!.table)
    .toMatchObject({ table: 1, row: null, column: null, unavailable_reason: "unsupported_table_geometry" })
  expect(revision.markdown).not.toContain("row 0")
})


test("pending outer table structure makes nested coordinates unavailable", async () => {
  const inner = table(row(cell(p("Nested target"))), 1)
  const source = docx(table(row(cell(p("Outer lead") + inner),
    '<w:ins w:id="901" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"/>'), 1))
  const revision = await createDocxRevision(source, { pendingRevisions: "layer" })
  expect(revision.paragraphs.find(p => p.text === "Nested target")!.table)
    .toMatchObject({ table: 2, row: null, column: null, unavailable_reason: "pending_table_structure" })
})
