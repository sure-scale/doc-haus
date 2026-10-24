// Run from dochaus/: bun test workflows/lean/notes-style.test.ts
import { afterAll, afterEach, beforeAll, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { buildCombinedFile } from "../../lib/combined-docx"
import { renderMarkdownToDocx } from "../../lib/render"
import { renderNotes } from "./stages"
import { cleanupMatters, matter } from "./test-support"

// A firm style template whose heading styles are numbered through a numbering
// style (w:numStyleLink), as Word writes a list style. The document takes it;
// the drafting notes stay in the house style, so a combined delivery can carry
// the notes' numbering beside the document's.
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const level = (ilvl: number, text: string, style: string) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="${style}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/></w:lvl>`
const heading = (n: number) =>
  `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="${n - 1}"/><w:numId w:val="7"/></w:numPr><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`

function listStyleTemplate() {
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${TYPE}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${TYPE}.numbering+xml"/></Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`),
    "word/document.xml": xml(`<w:document ${NS}><w:body><w:p><w:r><w:t>Firm template</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`),
    "word/styles.xml": xml(`<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${heading(1)}${heading(2)}<w:style w:type="numbering" w:customStyle="1" w:styleId="HeadingList"><w:name w:val="Heading List"/><w:pPr><w:numPr><w:numId w:val="8"/></w:numPr></w:pPr></w:style></w:styles>`),
    "word/numbering.xml": xml(`<w:numbering ${NS}><w:abstractNum w:abstractNumId="3"><w:nsid w:val="11223344"/><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="HeadingList"/>${level(0, "%1.", "Heading1")}${level(1, "%1.%2", "Heading2")}</w:abstractNum><w:abstractNum w:abstractNumId="4"><w:nsid w:val="55667788"/><w:multiLevelType w:val="multilevel"/><w:numStyleLink w:val="HeadingList"/></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num><w:num w:numId="8"><w:abstractNumId w:val="3"/></w:num></w:numbering>`),
  }, { mtime: new Date(1980, 0, 1) })
}

let workspace: string
const saved = process.env.WORKSPACE_ROOT
beforeAll(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "notes-style-"))
  process.env.WORKSPACE_ROOT = workspace
  mkdirSync(path.join(workspace, ".templates"), { recursive: true })
  mkdirSync(path.join(workspace, ".preferences"), { recursive: true })
  writeFileSync(path.join(workspace, ".templates", "Firm style.docx"), listStyleTemplate())
  writeFileSync(path.join(workspace, ".preferences", "preferences.json"), JSON.stringify({ styleTemplate: "Firm style.docx" }))
})
afterEach(cleanupMatters)
afterAll(() => {
  if (saved === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = saved
  rmSync(workspace, { recursive: true, force: true })
})

test("with a firm template numbered through a list style, the notes are in the house style and combine with the document", async () => {
  const directory = matter()
  const body = (await renderMarkdownToDocx("# Supply terms\n\n## Delivery\n\nGoods are delivered on the agreed date.")).bytes
  expect(strFromU8(unzipSync(body)["word/numbering.xml"]!)).toContain('<w:numStyleLink w:val="HeadingList"/>')
  const notes = "# Advice\n\n## Background\n\nThe supply terms follow the request.\n\n## Open points\n\nNone."
  // Rendered in the firm's style, the notes could not be combined.
  expect(buildCombinedFile({ memo: (await renderMarkdownToDocx(notes)).bytes, body })).toMatchObject({ refused: "unsupported_memo_numbering" })
  const memo = await renderNotes(directory, notes)
  expect(strFromU8(unzipSync(memo)["word/numbering.xml"] ?? new Uint8Array())).not.toContain("numStyleLink")
  const combined = buildCombinedFile({ memo, body })
  expect(combined).toHaveProperty("bytes")
  const document = strFromU8(unzipSync((combined as { bytes: Uint8Array }).bytes)["word/document.xml"]!)
  expect(document.indexOf("Background")).toBeLessThan(document.indexOf("Delivery"))
})
