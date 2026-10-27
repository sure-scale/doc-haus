import { expect, test } from "bun:test"
import { Document, LevelFormat, Packer, Paragraph } from "docx"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { documentPreviewBlocks } from "./document-preview"
import { extractDocumentText } from "./ingest"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const p = (text: string, properties = "") => `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ""}<w:r><w:t>${text}</w:t></w:r></w:p>`
async function fixture(body: string, styles = "") {
  const parts = unzipSync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Fixture")] }] })))
  parts["word/document.xml"] = strToU8(`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`)
  parts["word/styles.xml"] = strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${styles}</w:styles>`)
  return Buffer.from(zipSync(parts))
}

test("preview headings follow actual styles, inheritance and direct outline; identical body text stays a paragraph", async () => {
  const bytes = await fixture([
    p("Delivery", '<w:pStyle w:val="Heading1"/>'),
    p("Delivery", '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>'),
    p("Inherited heading", '<w:pStyle w:val="Custom"/>'),
    p("Direct level", '<w:outlineLvl w:val="2"/>'),
    p("Explicit body", '<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/>'),
    p("Inherited body", '<w:pStyle w:val="BodyOverride"/>'),
    p("Cycle is ordinary body", '<w:pStyle w:val="CycleA"/>'),
  ].join(""), [
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Custom"><w:basedOn w:val="Heading1"/></w:style>',
    '<w:style w:type="paragraph" w:styleId="BodyOverride"><w:basedOn w:val="Heading1"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="CycleA"><w:basedOn w:val="CycleB"/></w:style>',
    '<w:style w:type="paragraph" w:styleId="CycleB"><w:basedOn w:val="CycleA"/></w:style>',
  ].join(""))
  const text = await extractDocumentText("Reference.docx", bytes)
  expect(await documentPreviewBlocks("Reference.docx", bytes, text)).toEqual([
    { kind: "heading", level: 1, text: "Delivery" },
    { kind: "paragraph", text: "Delivery" },
    { kind: "heading", level: 1, text: "Inherited heading" },
    { kind: "heading", level: 3, text: "Direct level" },
    { kind: "paragraph", text: "Explicit body" },
    { kind: "paragraph", text: "Inherited body" },
    { kind: "paragraph", text: "Cycle is ordinary body" },
  ])
})

test("structured preview preserves canonical pending insertion, deletion and move disclosures as inert text", async () => {
  const bytes = await fixture(p("Reference heading", '<w:outlineLvl w:val="0"/>') + `<w:p>
    <w:r><w:t>Original </w:t></w:r>
    <w:del w:id="1" w:author="Reviewer"><w:r><w:delText>deleted words</w:delText></w:r></w:del>
    <w:ins w:id="2" w:author="Reviewer"><w:r><w:t>inserted words</w:t></w:r></w:ins>
    <w:moveFrom w:id="3" w:author="Reviewer"><w:r><w:delText>moved words</w:delText></w:r></w:moveFrom>
    <w:moveTo w:id="4" w:author="Reviewer"><w:r><w:t>moved words</w:t></w:r></w:moveTo>
    <w:r><w:tab/><w:t>&lt;script&gt;fictional&lt;/script&gt;\u200b</w:t></w:r>
  </w:p>`)
  const text = await extractDocumentText("Tracked.docx", bytes)
  const blocks = await documentPreviewBlocks("Tracked.docx", bytes, text)
  expect(blocks).toBeDefined()
  expect(blocks!.map(block => block.text + "\n\n").join("")).toBe(text)
  expect(text).toContain("pending tracked changes")
  expect(text).toContain("Tracked deletion by Reviewer: deleted words")
  expect(text).toContain("Tracked insertion by Reviewer: inserted words")
  expect(text).toContain("Tracked move from here by Reviewer: moved words")
  expect(text).toContain("Tracked move to here by Reviewer: moved words")
  expect(blocks!.at(-1)!.text).toContain("\t<script>fictional</script>")
  expect(blocks!.at(-1)!.text).not.toContain("\u200b")
})

test("structured preview retains the same resolved list labels and table paragraph order as canonical extraction", async () => {
  const bytes = await Packer.toBuffer(new Document({
    numbering: { config: [{ reference: "clauses", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", start: 1 }] }] },
    sections: [{ children: [new Paragraph({ text: "First provision", numbering: { reference: "clauses", level: 0 } }), new Paragraph({ text: "Second provision", numbering: { reference: "clauses", level: 0 } })] }],
  }))
  const text = await extractDocumentText("List.docx", bytes)
  const blocks = await documentPreviewBlocks("List.docx", bytes, text)
  expect(blocks?.map(block => block.text + "\n\n").join("")).toBe(text)
  expect(blocks?.map(block => block.text)).toEqual(["1. First provision", "2. Second provision"])
  const table = await fixture(`<w:tbl><w:tr><w:tc>${p("Left")}</w:tc><w:tc>${p("Right")}</w:tc></w:tr></w:tbl>`)
  const tableText = await extractDocumentText("Table.docx", table)
  expect((await documentPreviewBlocks("Table.docx", table, tableText))?.map(block => block.text)).toEqual(["Left", "Right"])
}, 30_000)

test("mismatch, partial preview, unsupported XML and non-DOCX files omit blocks without substituting text", async () => {
  const bytes = await fixture(p("Unchanged canonical text"))
  const text = await extractDocumentText("Reference.docx", bytes)
  expect(await documentPreviewBlocks("Reference.docx", bytes, text)).toBeDefined()
  expect(await documentPreviewBlocks("Reference.docx", bytes, text.replace("Unchanged", "Changed"))).toBeUndefined()
  expect(await documentPreviewBlocks("Reference.docx", bytes, text, true)).toBeUndefined()
  expect(await documentPreviewBlocks("Reference.txt", Buffer.from(text), text)).toBeUndefined()
  const parts = unzipSync(bytes)
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]).replace("</w:body>", "<w:p></w:body>"))
  expect(await documentPreviewBlocks("Invalid.docx", Buffer.from(zipSync(parts)), text)).toBeUndefined()
})
