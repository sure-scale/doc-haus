import { expect, test } from "bun:test"
import mammoth from "mammoth"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { writeLineBreaks } from "./docx-line-breaks"
import { markTrackedChanges } from "./docx-tracked-text"

// Every fixture is invented text: no actual party, address or document.

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const TYPES = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
const RELS = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

function docx(body: string) {
  return Buffer.from(zipSync({
    "[Content_Types].xml": strToU8(TYPES),
    "_rels/.rels": strToU8(RELS),
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`),
  }))
}
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const text = async (buffer: Buffer) => (await mammoth.extractRawText({ buffer: writeLineBreaks(markTrackedChanges(buffer)) })).value

test("a file without line breaks is returned unchanged", () => {
  const clean = docx(`<w:p>${run("First line only.")}</w:p>`)
  expect(writeLineBreaks(clean)).toBe(clean)
  const pages = docx(`<w:p>${run("Before.")}<w:r><w:br w:type="page"/></w:r>${run("After.")}</w:p>`)
  expect(writeLineBreaks(pages)).toBe(pages)
})

test("each line break in a paragraph reads as a newline", async () => {
  const block = docx(`<w:p><w:r><w:t>Example Office</w:t><w:br/><w:t>1 Sample Road</w:t></w:r><w:r><w:br w:type="textWrapping"/></w:r>${run("Sampletown")}</w:p><w:p>${run("Next paragraph.")}</w:p>`)
  expect(await text(block)).toBe("Example Office\n1 Sample Road\nSampletown\n\nNext paragraph.\n\n")
  const mixed = docx(`<w:p>${run("Before.")}<w:r><w:br w:type="page"/><w:t>After.</w:t><w:br w:type="column"/><w:br/><w:t>Last.</w:t></w:r></w:p>`)
  expect(await text(mixed)).toBe("Before.After.\nLast.\n\n")
})

test("a line break in a pending deletion is not read as current wording", async () => {
  const markup = docx(`<w:p>${run("Kept line")}<w:del w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:br/></w:r></w:del>${run(" joined.")}</w:p>`)
  expect(await text(markup)).toBe("Kept line joined.\n\n")
  const inserted = docx(`<w:p>${run("Kept line")}<w:ins w:id="2" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:br/></w:r></w:ins>${run("Proposed line.")}</w:p>`)
  expect(await text(inserted)).toBe("Kept line\nProposed line.\n\n")
})

test("Unicode namespace prefixes preserve line breaks while foreign tags stay untouched", async () => {
  const source = docx(`<w:p>${run("First")}<w:r><w:br/></w:r>${run("Second")}</w:p>`)
  const parts = unzipSync(source), xml = strFromU8(parts["word/document.xml"]!)
  const unicode = Buffer.from(zipSync({ ...parts, "word/document.xml": strToU8(xml.replaceAll("w:", "δοκ:").replace("xmlns:w=", "xmlns:δοκ=")) }))
  expect(await text(unicode)).toBe(await text(source))
  expect(await text(unicode)).toBe("First\nSecond\n\n")
  const foreign = Buffer.from(zipSync({ ...parts, "word/document.xml": strToU8(xml.replace("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "urn:unrelated")) }))
  expect(writeLineBreaks(foreign)).toBe(foreign)
})
