import { readFileSync, writeFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { strToU8, zipSync } from "fflate"

// A synthetic original for revision runs that must keep the original's tables:
// a short supply agreement whose Schedule 1 particulars, approval record and
// two signature grids are Word tables, the approval and signature cells left
// empty for completion. Every name, figure and date is invented. The tests use
// the bytes directly; a live check writes the file and uploads it to a matter:
//
//   bun dochaus/test/support/revision-tables-fixture.ts /path/to/carton-supply-agreement.docx
//
// then asks for a revision that touches no table (TABLE_FIXTURE.untouched) or
// one table cell (TABLE_FIXTURE.cellEdit) and checks the four tables in the
// delivered .docx.

export const TABLE_FIXTURE = {
  name: "carton-supply-agreement.docx",
  title: "Carton Supply Agreement",
  untouched: "Revise the supplied Carton Supply Agreement: change the payment period in clause 2 from thirty days to forty-five days. Deliver the revised agreement with a redline against the original.",
  cellEdit: "Revise the supplied Carton Supply Agreement: change the monthly quantity in Schedule 1 from 40,000 cartons to 45,000 cartons. Deliver the revised agreement with a redline against the original.",
  /** The tables in document order, as rows of cell texts ("" for an empty cell). */
  tables: [
    [["Item", "Particulars"], ["Products", "Printed folding cartons, 300 gsm board"], ["Monthly quantity", "40,000 cartons"], ["Unit price", "0.14 per carton"], ["Delivery point", "Unit 7, Riverside Trading Estate"], ["Commencement Date", "1 March 2027"]],
    [["Version", "Approved by", "Office", "Date"], ["1.0", "", "", ""]],
    [["Signature", ""], ["Name", ""], ["Office", ""], ["Date", ""]],
    [["Signature", ""], ["Name", ""], ["Office", ""], ["Date", ""]],
  ],
}

type Block = { style?: "Title" | "Heading1"; text: string } | { table: string[][]; header: boolean }

const [particulars, approval, supplierGrid, buyerGrid] = TABLE_FIXTURE.tables

const BLOCKS: Block[] = [
  { style: "Title", text: TABLE_FIXTURE.title },
  { text: "This agreement is made between Northwind Cartons Limited (the Supplier) and Bluebell Foods Limited (the Buyer)." },
  { style: "Heading1", text: "1. Supply" },
  { text: "The Supplier shall supply the Products in the quantities and at the prices set out in Schedule 1." },
  { style: "Heading1", text: "2. Payment" },
  { text: "The Buyer shall pay each invoice within thirty days of its date." },
  { style: "Heading1", text: "3. Term" },
  { text: "This agreement starts on the Commencement Date and continues for two years." },
  { style: "Heading1", text: "Schedule 1 – Particulars" },
  { table: particulars!, header: true },
  { style: "Heading1", text: "Approval record" },
  { table: approval!, header: true },
  { style: "Heading1", text: "Execution" },
  { text: "Signed for and on behalf of Northwind Cartons Limited:" },
  { table: supplierGrid!, header: false },
  { text: "Signed for and on behalf of Bluebell Foods Limited:" },
  { table: buyerGrid!, header: false },
]

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

function escape(text: string) {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
}

function paragraph(text: string, style?: string, bold = false) {
  const properties = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""
  // An empty cell holds one empty paragraph, as Word writes it.
  if (!text) return `<w:p>${properties}</w:p>`
  return `<w:p>${properties}<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${escape(text)}</w:t></w:r></w:p>`
}

function table(rows: string[][], header: boolean) {
  const columns = rows[0]!.length
  const width = Math.floor(9000 / columns)
  const borders = ["top", "left", "bottom", "right", "insideH", "insideV"].map(edge => `<w:${edge} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join("")
  return `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>` +
    `<w:tblGrid>${`<w:gridCol w:w="${width}"/>`.repeat(columns)}</w:tblGrid>` +
    rows.map((row, at) => `<w:tr>${at === 0 && header ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${row.map(cell =>
      `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${paragraph(cell, undefined, at === 0 && header)}</w:tc>`).join("")}</w:tr>`).join("") +
    "</w:tbl>"
}

function documentXml() {
  const body = BLOCKS.map(block => "table" in block ? table(block.table, block.header) : paragraph(block.text, block.style)).join("")
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>' +
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>' +
  "</w:styles>"

/** The same original as a word processor saves it, with document defaults,
 * a theme, explicit paragraph styles and cell borders: the bytes above saved
 * again by LibreOffice (soffice --headless --convert-to docx), its document
 * properties empty. */
export function revisionTablesWordStyleDocx(): Uint8Array {
  return new Uint8Array(readFileSync(fileURLToPath(new URL("../fixtures/carton-supply-agreement-libreoffice.docx", import.meta.url))))
}

/** The synthetic original's .docx bytes. */
export function revisionTablesOriginalDocx(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'),
    "_rels/.rels": strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    "word/_rels/document.xml.rels": strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
    "word/document.xml": strToU8(documentXml()),
    "word/styles.xml": strToU8(STYLES),
  }, { mtime: new Date(1980, 0, 1) })
}

if (import.meta.main) {
  const target = process.argv[2]
  if (!target) throw new Error("Usage: bun dochaus/test/support/revision-tables-fixture.ts <output .docx path>")
  writeFileSync(target, revisionTablesOriginalDocx())
  console.log(`Wrote ${target}`)
}
