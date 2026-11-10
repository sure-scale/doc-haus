import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strToU8, zipSync } from "fflate"
import { openDb } from "./db"
import { documentPreviewBlocks } from "./document-preview"
import { extractDocumentText, ingestDocument } from "./ingest"
import { reconstructText } from "./structure"
import { liveText } from "../../../dochaus/lib/extract"
import { verifyCitation } from "../../../dochaus/plugin/legal/citations"

// Every fixture is invented text: no actual party, address or document.

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const docx = (body: string) => Buffer.from(zipSync({
  "[Content_Types].xml": strToU8('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
  "_rels/.rels": strToU8('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
  "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`),
}))
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`

const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-line-breaks-")))
afterAll(() => rmSync(directory, { recursive: true, force: true }))

test("carriage returns share exact extraction, citation offsets and preview text without indexing", async () => {
  const name = "Carriage-return.docx"
  const bytes = docx(`<w:p><w:r><w:t>Example Office</w:t><w:cr/><w:t>1 Sample Road</w:t><w:br/><w:t>Sampletown</w:t></w:r></w:p>`)
  const docPath = path.join(directory, name)
  writeFileSync(docPath, bytes)
  const canonical = await extractDocumentText(name, bytes)
  expect(canonical).toContain("Example Office\n1 Sample Road\nSampletown")
  // liveText appends the canonical source-record terminator; interior offsets are identical.
  expect(await liveText(docPath)).toBe(`${canonical}\n`)
  const excerpt = "Office\n1 Sample Road"
  const charStart = canonical.indexOf(excerpt)
  expect(await verifyCitation({ documentName: name, docPath, section: "1", excerpt, charStart, charEnd: charStart + excerpt.length }))
    .toMatchObject({ verified: true, charStart, charEnd: charStart + excerpt.length })
  const blocks = await documentPreviewBlocks(name, bytes, canonical)
  expect(blocks?.map(block => block.text)).toEqual(["Example Office\n1 Sample Road\nSampletown"])
  expect(readFileSync(docPath)).toEqual(bytes)
})

test("lines split by line breaks are indexed on their own lines, and a citation across them verifies", async () => {
  await ingestDocument(directory, "Letterhead.docx", docx(
    `<w:p><w:r><w:t>Example Office</w:t><w:br/><w:t>1 Sample Road</w:t><w:br/><w:t>Sampletown</w:t></w:r></w:p>`
    + `<w:p>${run("Dear Reader,")}</w:p>`,
  ))
  const docPath = path.join(directory, "Letterhead.docx")
  const db = openDb(directory)
  try {
    // Ingest, the citation verifier and the preview read the same text.
    const canonical = reconstructText(db, docPath)
    expect(canonical).toBe(await liveText(docPath))
    expect(canonical).toContain("Example Office\n1 Sample Road\nSampletown")

    const excerpt = "1 Sample Road\nSampletown"
    const charStart = canonical.indexOf(excerpt)
    const verified = await verifyCitation({ documentName: "Letterhead.docx", docPath, section: "1", excerpt, charStart, charEnd: charStart + excerpt.length })
    expect(verified).toMatchObject({ verified: true, charStart })

    const bytes = readFileSync(docPath)
    const blocks = await documentPreviewBlocks("Letterhead.docx", bytes, await extractDocumentText("Letterhead.docx", bytes))
    expect(blocks?.map(block => block.text)).toEqual(["Example Office\n1 Sample Road\nSampletown", "Dear Reader,"])
  } finally {
    db.close()
  }
}, 120_000)
