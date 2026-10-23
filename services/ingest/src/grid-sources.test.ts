import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strToU8, zipSync } from "fflate"
import { checkCellSources, sourceCheckRequest } from "./grid"
import { liveText } from "../../../dochaus/lib/extract"

// Invented supply terms. A grid answer cites one sentence; the document then
// changes under it, and the source check reports what the lawyer would find.
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const docx = (paragraphs: string[]) => Buffer.from(zipSync({
  "[Content_Types].xml": strToU8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
  "_rels/.rels": strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
  "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${paragraphs.map(text => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join("")}</w:body></w:document>`),
}))
const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-grid-sources-")))
afterAll(() => rmSync(directory, { recursive: true, force: true }))

test("a cited cell reads verified, re-anchored, not found or missing as its document changes", async () => {
  const docPath = path.join(directory, "Supply terms.docx")
  const excerpt = "This agreement is governed by the law of England and Wales."
  writeFileSync(docPath, docx(["1. Supply. The Supplier shall supply the Goods.", excerpt]))
  const charStart = (await liveText(docPath)).indexOf(excerpt)
  const request = sourceCheckRequest({ cells: [{ key: "Supply terms.docx::c1", documentName: "Supply terms.docx", excerpt, charStart, charEnd: charStart + excerpt.length }] })!
  const registered = (name: string) => (name === "Supply terms.docx" ? { name, docPath } : undefined)
  const check = async () => (await checkCellSources(request, registered))["Supply terms.docx::c1"]

  expect(await check()).toBe("verified")
  // A clause inserted above moves the cited sentence; the text is still there.
  writeFileSync(docPath, docx(["1. Supply. The Supplier shall supply the Goods.", "2. Price. The Buyer shall pay on delivery.", excerpt]))
  expect(await check()).toBe("reanchored")
  // The cited sentence itself was amended.
  writeFileSync(docPath, docx(["1. Supply. The Supplier shall supply the Goods.", "This agreement is governed by the law of Scotland."]))
  expect(await check()).toBe("not_found")
  expect((await checkCellSources(request, () => undefined))["Supply terms.docx::c1"]).toBe("missing")
}, 60_000)
