import { expect, test } from "bun:test"
import mammoth from "mammoth"
import { strToU8, unzipSync, zipSync } from "fflate"
import { writeListLabels } from "./docx-list-labels"

// Every fixture is invented text: no actual party or transaction.

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const PART = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

function docx(body: string, numbering?: string, styles = "") {
  return Buffer.from(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${PART}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${PART}.styles+xml"/>${numbering ? `<Override PartName="/word/numbering.xml" ContentType="${PART}.numbering+xml"/>` : ""}</Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/>${numbering ? `<Relationship Id="rIdN" Type="${REL}/numbering" Target="numbering.xml"/>` : ""}</Relationships>`),
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`),
    "word/styles.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${styles}</w:styles>`),
    ...(numbering ? { "word/numbering.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>${numbering}</w:numbering>`) } : {}),
  }))
}
const level = (ilvl: number, format: string, text: string, font?: string) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/></w:rPr>` : ""}</w:lvl>`
// numId 1: clause numbering 1. / 1.1 / (a). numId 2: Word's default bullets,
// whose second level is the letter "o" in Courier New.
const NUMBERING = `<w:abstractNum w:abstractNumId="0">${level(0, "decimal", "%1.")}${level(1, "decimal", "%1.%2")}${level(2, "lowerLetter", "(%3)")}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="1">${level(0, "bullet", "•", "Symbol")}${level(1, "bullet", "o", "Courier New")}</w:abstractNum>`
  + `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const para = (text: string, list?: [number, number], style?: string) =>
  `<w:p>${list || style ? `<w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ""}${list ? `<w:numPr><w:ilvl w:val="${list[1]}"/><w:numId w:val="${list[0]}"/></w:numPr>` : ""}</w:pPr>` : ""}${text ? run(text) : ""}</w:p>`
const raw = async (buffer: Buffer) => (await mammoth.extractRawText({ buffer })).value
const labelled = async (buffer: Buffer) => raw(await writeListLabels(buffer))

test("numbered paragraphs read with the numbers Word shows; bulleted ones stay as they are", async () => {
  const terms = docx([
    para("Delivery", [1, 0]),
    para("The Supplier shall deliver the cartons.", [1, 1]),
    para("on the delivery date; and", [1, 2]),
    para("at the site.", [1, 2]),
    para("Packed in boxes", [2, 0]),
    para("Labelled by hand", [2, 1]),
    para("Payment", [1, 0]),
    para("The Customer shall pay each invoice.", [1, 1]),
  ].join(""), NUMBERING)
  expect(await labelled(terms)).toBe(
    "1. Delivery\n\n1.1 The Supplier shall deliver the cartons.\n\n(a) on the delivery date; and\n\n(b) at the site.\n\n"
    + "Packed in boxes\n\nLabelled by hand\n\n2. Payment\n\n2.1 The Customer shall pay each invoice.\n\n",
  )
})

test("table cells and style numbering take their labels; empty and text-box paragraphs take none", async () => {
  const clause = `<w:style w:type="paragraph" w:styleId="Clause"><w:name w:val="Clause"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>`
  const textBox = `<w:p>${run("See box.")}<w:r><w:pict><w:txbxContent>${para("Boxed note", [1, 1])}</w:txbxContent></w:pict></w:r></w:p>`
  const schedule = docx([
    para("Scope", [1, 0]),
    para("The services are listed below.", undefined, "Clause"),
    `<w:tbl><w:tr><w:tc>${para("Cleaning of the synthetic premises.", [1, 1])}</w:tc><w:tc>${para("Weekly")}</w:tc></w:tr></w:tbl>`,
    para("", [1, 0]),
    textBox,
    para("Fees", [1, 0]),
  ].join(""), NUMBERING, clause)
  // The empty clause still counts: Word shows "2." on its own line, and Fees is 3.
  expect(await labelled(schedule)).toBe(
    "1. Scope\n\n1.1 The services are listed below.\n\n1.2 Cleaning of the synthetic premises.\n\nWeekly\n\n\n\nSee box.\n\nBoxed note\n\n3. Fees\n\n",
  )
  // Numbered only through a style based on the numbered one, as headings bound to a list are.
  const subclause = `<w:style w:type="paragraph" w:styleId="Subclause"><w:name w:val="Subclause"/><w:basedOn w:val="Clause"/></w:style>`
  const styled = docx(para("The services are listed below.", undefined, "Clause") + para("Cleaning is weekly.", undefined, "Subclause"), NUMBERING, clause + subclause)
  expect(await labelled(styled)).toBe("1.1 The services are listed below.\n\n1.2 Cleaning is weekly.\n\n")
})

test("clauses of the same text keep their own numbers", async () => {
  // The table's grid columns are anchors too, and are not paragraphs.
  const table = `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${para("Rate")}</w:tc><w:tc>${para("Reserved", [1, 1])}</w:tc></w:tr></w:tbl>`
  const reserved = docx([para("Reserved", [1, 0]), para("Reserved", [1, 0]), table, para("Reserved", [1, 1]), para("Fees", [1, 0])].join(""), NUMBERING)
  expect(await labelled(reserved)).toBe("1. Reserved\n\n2. Reserved\n\nRate\n\n2.1 Reserved\n\n2.2 Reserved\n\n3. Fees\n\n")
})

test("a document without numbered lists is passed through as the same bytes", async () => {
  const plain = docx(para("The Supplier shall deliver the cartons.") + para("The Customer shall pay each invoice."))
  expect(await writeListLabels(plain)).toBe(plain)
  const bulleted = docx(para("Packed in boxes", [2, 0]) + para("Labelled by hand", [2, 1]), NUMBERING)
  expect(await writeListLabels(bulleted)).toBe(bulleted)
})

test("a file the numbering reader cannot open extracts exactly as before", async () => {
  const terms = docx(para("Delivery", [1, 0]) + para("The Supplier shall deliver the cartons.", [1, 1]), NUMBERING)
  // No content types part: mammoth reads it, Docxodus does not.
  const parts = unzipSync(terms)
  delete parts["[Content_Types].xml"]
  const untyped = Buffer.from(zipSync(parts))
  expect(await writeListLabels(untyped)).toBe(untyped)
  expect(await raw(untyped)).toBe("Delivery\n\nThe Supplier shall deliver the cartons.\n\n")
  const notZip = Buffer.from("not a word file")
  expect(await writeListLabels(notZip)).toBe(notZip)
})
