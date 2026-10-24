import { expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { carryOriginalTables, pipeTable, readDocxTables } from "./docx-tables"
import { renderMarkdownToDocx } from "./render"
import { TABLE_FIXTURE, revisionTablesOriginalDocx } from "../test/support/revision-tables-fixture"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const [particulars, approval, supplierGrid, buyerGrid] = TABLE_FIXTURE.tables as [string[][], string[][], string[][], string[][]]

/** A revised body as the drafter writes it: the tables as Markdown tables after their passages. */
function draft(tables: { particulars?: string[][]; approval?: string[][] | null; supplier?: string[][] | null; buyer?: string[][] } = {}) {
  const table = (rows: string[][] | null | undefined, fallback: string[][]) => rows === null ? "" : pipeTable(rows ?? fallback)
  return [
    "# Carton Supply Agreement",
    "## Agreement",
    "This agreement is made between Northwind Cartons Limited (the Supplier) and Bluebell Foods Limited (the Buyer).",
    "### 2. Payment",
    "The Buyer shall pay each invoice within forty-five days of its date.",
    "### Schedule 1 – Particulars",
    table(tables.particulars, particulars),
    ...(tables.approval === null ? [] : ["### Approval record", table(tables.approval, approval)]),
    "### Execution",
    "Signed for and on behalf of Northwind Cartons Limited:",
    table(tables.supplier, supplierGrid),
    "Signed for and on behalf of Bluebell Foods Limited:",
    table(tables.buyer, buyerGrid),
  ].filter(Boolean).join("\n\n") + "\n"
}

/** Each body table's markup, serialized alike whichever package it is read from. */
function tableXml(bytes: Uint8Array) {
  const document = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  return Array.from(document.getElementsByTagNameNS(W, "tbl")).map(node => new XMLSerializer().serializeToString(node).replace(/ xmlns:w="[^"]*"/, ""))
}

async function carried(markdown: string, original = revisionTablesOriginalDocx()) {
  return carryOriginalTables((await renderMarkdownToDocx(markdown)).bytes, original)
}

test("the original's tables read with their rows, empty cells and the passages they follow", () => {
  const tables = readDocxTables(revisionTablesOriginalDocx())
  expect(tables.map(table => table.rows)).toEqual(TABLE_FIXTURE.tables)
  expect(tables.map(table => table.follows)).toEqual(["Schedule 1 – Particulars", "Approval record",
    "Signed for and on behalf of Northwind Cartons Limited:", "Signed for and on behalf of Bluebell Foods Limited:"])
  expect(pipeTable(supplierGrid)).toBe("| Signature |  |\n| --- | --- |\n| Name |  |\n| Office |  |\n| Date |  |")
})

test("tables the revision leaves alone are the original's own table elements, empty signature cells included", async () => {
  const result = await carried(draft())
  expect(result.carries.map(row => row.outcome)).toEqual(["carried", "carried", "carried", "carried"])
  expect(tableXml(result.bytes)).toEqual(tableXml(revisionTablesOriginalDocx()))
  expect(readDocxTables(result.bytes).map(table => table.rows)).toEqual(TABLE_FIXTURE.tables)
  // The renderer's spacing paragraph after a Markdown table goes with it.
  expect(strFromU8(unzipSync(result.bytes)["word/document.xml"]!)).not.toContain('w:before="80" w:after="120"')
})

test("a changed cell is the only change to its table", async () => {
  const edited = particulars.map(row => row.map(cell => cell === "40,000 cartons" ? "45,000 cartons" : cell))
  const result = await carried(draft({ particulars: edited }))
  expect(result.carries[0]).toMatchObject({ outcome: "edited", cells: ["row 3, column 2"], added: 0, removed: [] })
  const [schedule, ...others] = tableXml(result.bytes)
  const [original, ...unchanged] = tableXml(revisionTablesOriginalDocx())
  expect(schedule).toBe(original!.replace("40,000 cartons", "45,000 cartons"))
  expect(others).toEqual(unchanged)
})

test("typographic variants of a cell's text do not count as a change", async () => {
  const retyped = particulars.map(row => row.map(cell => cell.replace("Printed folding cartons, 300 gsm board", "Printed  folding cartons,  300 gsm board")))
  const result = await carried(draft({ particulars: retyped }).replace("Schedule 1 – Particulars", "Schedule 1 - Particulars"))
  expect(result.carries[0]!.outcome).toBe("carried")
})

test("rows added or removed keep the original's other rows and formatting", async () => {
  const added = [...particulars.slice(0, 4), ["Minimum order", "5,000 cartons"], ...particulars.slice(4)]
  const result = await carried(draft({ particulars: added, buyer: buyerGrid.slice(0, 3) }))
  expect(result.carries[0]).toMatchObject({ outcome: "edited", cells: [], added: 1, removed: [] })
  expect(result.carries[3]).toMatchObject({ outcome: "edited", cells: [], added: 0, removed: ["Date"] })
  const tables = readDocxTables(result.bytes)
  expect(tables[0]!.rows).toEqual(added)
  expect(tables[3]!.rows).toEqual(buyerGrid.slice(0, 3))
  // The added row takes the form of the row before it: a body row, in the original's borders.
  const schedule = tableXml(result.bytes)[0]!
  expect(schedule.match(/<w:tblHeader\/>/g)).toHaveLength(1)
  expect(schedule).toContain('w:color="000000"')
})

test("a table the draft leaves out is reported missing, and one with other columns stays as drafted", async () => {
  const restructured = [["Signature", "Name", "Date"], ["", "", ""]]
  const result = await carried(draft({ approval: null, supplier: restructured }))
  expect(result.carries.map(row => row.outcome)).toEqual(["carried", "missing", "drafted", "carried"])
  expect(result.carries[2]).toMatchObject({ reason: "the drafted table has 3 columns where the original has 2" })
  expect(readDocxTables(result.bytes).map(table => table.rows)).toEqual([particulars, restructured, buyerGrid])
})

test("a table that refers to its own package's numbering is not carried", async () => {
  const parts = unzipSync(revisionTablesOriginalDocx())
  const xml = strFromU8(parts["word/document.xml"]!).replace("<w:tc><w:tcPr><w:tcW w:w=\"2250\" w:type=\"dxa\"/></w:tcPr><w:p>",
    "<w:tc><w:tcPr><w:tcW w:w=\"2250\" w:type=\"dxa\"/></w:tcPr><w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>")
  expect(xml).toContain("<w:numPr>")
  const numbered = zipSync({ ...parts, "word/document.xml": strToU8(xml) })
  const result = await carried(draft(), numbered)
  expect(result.carries[1]).toEqual({ table: 2, follows: "Approval record", outcome: "drafted", reason: "the original table holds list numbering, which refers to its own package" })
  expect(result.carries.filter(row => row.outcome === "carried")).toHaveLength(3)
})

test("a carried table's style comes with it when the rendered body lacks it", async () => {
  const parts = unzipSync(revisionTablesOriginalDocx())
  const xml = strFromU8(parts["word/document.xml"]!).replace("<w:tblPr>", '<w:tblPr><w:tblStyle w:val="TableGrid"/>')
  const styles = strFromU8(parts["word/styles.xml"]!).replace("</w:styles>",
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>')
  const styled = zipSync({ ...parts, "word/document.xml": strToU8(xml), "word/styles.xml": strToU8(styles) })
  const result = await carried(draft(), styled)
  expect(result.carries[0]!.outcome).toBe("carried")
  const copied = strFromU8(unzipSync(result.bytes)["word/styles.xml"]!)
  expect(copied).toContain('w:styleId="TableGrid"')
  expect(copied).toContain('w:styleId="TableNormal"')
  expect(tableXml(result.bytes)[0]).toContain('<w:tblStyle w:val="TableGrid"/>')
})
