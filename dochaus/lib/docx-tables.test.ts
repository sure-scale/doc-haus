import { expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { carryOriginalParagraphs, carryOriginalTables, pipeTable, readDocxTables } from "./docx-tables"
import { renderMarkdownToDocx } from "./render"
import { TABLE_FIXTURE, revisionTablesOriginalDocx, revisionTablesWordStyleDocx } from "../test/support/revision-tables-fixture"

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

/** For each body table, the paragraph directly above it: its text, its style and whether it keeps with the next block. */
function leadIns(bytes: Uint8Array) {
  const document = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const paragraphs = Array.from(document.getElementsByTagNameNS(W, "p"))
  const text = (node: Element) => Array.from(node.getElementsByTagNameNS(W, "t")).map(row => row.textContent).join("")
  const style = (node: Element) => node.getElementsByTagNameNS(W, "pStyle")[0]?.getAttributeNS(W, "val") ?? undefined
  const keeps = (node: Element) => Array.from(node.childNodes).some(child => (child as Element).localName === "pPr" &&
    Array.from(child.childNodes).some(property => (property as Element).localName === "keepNext"))
  return {
    tables: Array.from(document.getElementsByTagNameNS(W, "tbl")).map(table => {
      let before = table.previousSibling
      while (before && before.nodeType !== 1) before = before.previousSibling
      const lead = before as Element
      return [text(lead), style(lead), keeps(lead)] as const
    }),
    kept: paragraphs.filter(node => node.parentNode?.localName === "body" && keeps(node)).map(text),
  }
}

test("a carried table's lead-in keeps with it where the original's did, in its own style; other lead-ins, and a table kept as drafted, are left as rendered", async () => {
  // The original keeps a heading-styled line with each table but the third.
  const heading = (text: string) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  const [particularsGrid, record, alpha, beta] = [[["Item", "Value"], ["Price", "10"]], [["Version", ""], ["Approved by", ""]], [["Signature", ""], ["Name", ""]], [["Signature", ""], ["Date", ""]]]
  const original = originalWith(heading("Particulars") + tbl(particularsGrid) + heading("Approval record") + tbl(record) + p("Signed for Alpha Limited") + tbl(alpha) +
    heading("Beta Limited") + tbl(beta) + p("End"))
  // The drafter writes the first as a heading, the record's and Beta's as plain lines, and Beta's grid with another column.
  const rendered = (await renderMarkdownToDocx(markdown("### Particulars", pipeTable(particularsGrid), "Approval record", pipeTable(record),
    "Signed for Alpha Limited", pipeTable(alpha), "Beta Limited", pipeTable([["Signature", "", ""], ["Date", "", ""]]), "End"))).bytes
  const result = carryOriginalTables(rendered, original)
  expect(result.carries.map(row => row.outcome)).toEqual(["carried", "carried", "carried", "drafted"])
  const before = leadIns(rendered)
  expect(before.kept).toEqual([])
  // Only the record's plain line is given the original's keep: the heading
  // keeps by its style, the original did not keep Alpha's line, and Beta's
  // table stays as drafted.
  expect(leadIns(result.bytes)).toEqual({ tables: before.tables.map(([text, style], index) => [text, style, index === 1]), kept: ["Approval record"] })
})

test("another line the drafter puts directly above a carried table is not kept with it", async () => {
  const heading = (text: string) => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  const record = [["Version", ""], ["Approved by", ""]]
  const sentence = "The parties shall sign this record on every change of version."
  const original = originalWith(heading("Approval record") + tbl(record) + p(sentence) + p("End"))
  // The original's sentence after the table is moved above it, under the record's line.
  const rendered = (await renderMarkdownToDocx(markdown("Approval record", sentence, pipeTable(record), "End"))).bytes
  const result = carryOriginalTables(rendered, original)
  expect(result.carries.map(row => row.outcome)).toEqual(["carried"])
  expect(leadIns(result.bytes)).toEqual({ tables: [[sentence, undefined, false]], kept: [] })
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

// Originals built around one or two tables, for the cases the fixture lacks.

function escape(text: string) {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
}

function p(text: string, runProperties = "") {
  return text ? `<w:p><w:r>${runProperties}<w:t xml:space="preserve">${escape(text)}</w:t></w:r></w:p>` : "<w:p/>"
}

/** A cell of plain text, one paragraph per line, or of the given paragraph markup. */
function tc(cell: string) {
  return `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${cell.startsWith("<w:") ? cell : cell.split("\n").map(line => p(line)).join("")}</w:tc>`
}

function tbl(rows: string[][], colour = "000000") {
  return `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="${colour}"/></w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${rows[0]!.map(() => '<w:gridCol w:w="3000"/>').join("")}</w:tblGrid>${rows.map(row => `<w:tr>${row.map(tc).join("")}</w:tr>`).join("")}</w:tbl>`
}

/** The fixture's package with another body, and optionally other styles. */
function originalWith(body: string, styles: (xml: string) => string = xml => xml, extra: Record<string, string> = {}) {
  const parts = unzipSync(revisionTablesOriginalDocx())
  return zipSync({
    ...parts,
    "word/document.xml": strToU8(strFromU8(parts["word/document.xml"]!).replace(/<w:body>[\s\S]*<w:sectPr>/, `<w:body>${body}<w:sectPr>`)),
    "word/styles.xml": strToU8(styles(strFromU8(parts["word/styles.xml"]!))),
    ...Object.fromEntries(Object.entries(extra).map(([name, xml]) => [name, strToU8(xml)])),
  })
}

function markdown(...blocks: string[]) {
  return ["# Doc", ...blocks].join("\n\n") + "\n"
}

function documentXml(bytes: Uint8Array) {
  return strFromU8(unzipSync(bytes)["word/document.xml"]!)
}

test("cells holding Markdown characters read back as the original's, whether or not the drafter keeps their character references", async () => {
  const values = ["[*]", "[*] cartons per [*]", "Fee*", "* excluding VAT", "**Note**", "`code`", "[Name](x)", "~~old~~", "C:\\path\\x", "x &amp; y", "<br>"]
  const rows = [["Item", "Value"], ...values.map((value, at) => [`Row ${at + 1}`, value])]
  const original = originalWith(p("Intro") + tbl(rows) + p("End"))
  expect(readDocxTables(original)[0]!.rows).toEqual(rows)
  expect(pipeTable([["[*]", "Fee*"]])).toBe("| [&#42;] | Fee&#42; |\n| --- | --- |")
  const referenced = await carried(markdown("Intro", pipeTable(rows), "End"), original)
  expect(referenced.carries[0]!.outcome).toBe("carried")
  expect(readDocxTables(referenced.bytes)[0]!.rows).toEqual(rows)
  // A drafter that writes the cells as they read renders "[*]" as "[]"; the original's cells stand.
  const written = [`| ${rows[0]!.join(" | ")} |`, "| --- | --- |", ...rows.slice(1).map(row => `| ${row.join(" | ")} |`)].join("\n")
  const raw = await carried(markdown("Intro", written, "End"), original)
  expect(raw.carries[0]!.outcome).toBe("carried")
  expect(readDocxTables(raw.bytes)[0]!.rows).toEqual(rows)
})

test("a bar in a cell is a character reference, and a table kept as drafted shows it as a bar", async () => {
  const rows = [["Option", "Terms"], ["A | B", "either"]]
  const original = originalWith(p("Intro") + tbl(rows) + p("End"))
  const drafted = [["Option", "Terms", "Notes"], ["A | B", "either", ""]]
  expect(pipeTable(drafted)).toContain("| A &#124; B | either |  |")
  const result = await carried(markdown("Intro", pipeTable(drafted), "End"), original)
  expect(result.carries[0]).toMatchObject({ outcome: "drafted", reason: "the drafted table has 3 columns where the original has 2" })
  expect(readDocxTables(result.bytes)[0]!.rows).toEqual(drafted)
})

test("an edited cell and a table kept as drafted show the Markdown characters their drafted text holds", async () => {
  const values = ["<Name> of the Buyer", "Written as <br>", "See [1](2)", "Cost &amp; fees"]
  const rows = [["Item", "Value"], ...values.map((value, at) => [`Row ${at + 1}`, value])]
  const original = originalWith(p("Intro") + tbl(rows) + p("End"))
  const amended = rows.map(row => row.map(cell => values.includes(cell) ? `${cell}, as amended` : cell))
  const edited = await carried(markdown("Intro", pipeTable(amended), "End"), original)
  expect(edited.carries[0]).toMatchObject({ outcome: "edited", cells: ["row 2, column 2", "row 3, column 2", "row 4, column 2", "row 5, column 2"] })
  expect(readDocxTables(edited.bytes)[0]!.rows).toEqual(amended)
  const widened = amended.map(row => [...row, ""])
  const drafted = await carried(markdown("Intro", pipeTable(widened), "End"), original)
  expect(drafted.carries[0]!.outcome).toBe("drafted")
  expect(readDocxTables(drafted.bytes)[0]!.rows).toEqual(widened)
})

test("a one-row signature block and a one-column notice are carried, with no Markdown left in the body", async () => {
  const notice = [["Important notice"], ["This agreement limits liability."]]
  const block = [["Signed for Alpha Limited\n\nName:\nTitle:", "Signed for Beta Limited\n\nName:\nTitle:"]]
  const original = originalWith(p("Notice") + tbl(notice) + p("Execution") + tbl(block) + p("End"))
  const result = await carried(markdown("Notice", pipeTable(notice), "Execution", pipeTable(block), "End"), original)
  expect(result.carries.map(row => row.outcome)).toEqual(["carried", "carried"])
  expect(readDocxTables(result.bytes).map(table => table.rows)).toEqual([notice, block])
  expect(documentXml(result.bytes)).not.toMatch(/<w:t[^>]*>[^<]*\|/)
})

test("two tables the original separates by an empty paragraph stay two tables", async () => {
  const first = [["A", "1"], ["B", "2"]]
  const second = [["C", "x", "y"], ["D", "z", "w"]]
  const original = originalWith(p("Particulars") + tbl(first) + p("") + tbl(second) + p("End"))
  const result = await carried(markdown("Particulars", pipeTable(first), pipeTable(second), "End"), original)
  expect(result.carries.map(row => row.outcome)).toEqual(["carried", "carried"])
  expect(documentXml(result.bytes)).not.toMatch(/<\/w:tbl>\s*<w:tbl[\s>]/)
  expect(readDocxTables(result.bytes).map(table => table.rows)).toEqual([first, second])
})

test("identical grids are told apart by the passage each follows", async () => {
  const grid = [["Signature", ""], ["Name", ""], ["Office", ""]]
  const original = originalWith(p("Signed for Alpha Limited:") + tbl(grid, "FF0000") + p("Signed for Beta Limited:") + tbl(grid, "0000FF") + p("End"))
  const colours = (bytes: Uint8Array) => tableXml(bytes).map(xml => xml.match(/w:color="(FF0000|0000FF)"/)?.[1] ?? "drafted")
  // Alpha's grid left out: Alpha's is missing, and Beta's own grid stands after Beta's passage.
  const dropped = await carried(markdown("Signed for Beta Limited:", pipeTable(grid), "End"), original)
  expect(dropped.carries.map(row => row.outcome)).toEqual(["missing", "carried"])
  expect(colours(dropped.bytes)).toEqual(["0000FF"])
  // So too when Beta's passage is reworded: it shares more of its words with Beta's than with Alpha's.
  const reworded = await carried(markdown("Signed for and on behalf of Beta Limited:", pipeTable(grid), "End"), original)
  expect(reworded.carries.map(row => row.outcome)).toEqual(["missing", "carried"])
  expect(colours(reworded.bytes)).toEqual(["0000FF"])
  // A new grid between them stays as drafted; each party keeps its own.
  const between = await carried(markdown("Signed for Alpha Limited:", pipeTable(grid), "Signed for Gamma Guarantor Limited:", pipeTable(grid),
    "Signed for Beta Limited:", pipeTable(grid), "End"), original)
  expect(between.carries.map(row => row.outcome)).toEqual(["carried", "carried"])
  expect(colours(between.bytes)).toEqual(["FF0000", "drafted", "0000FF"])
})

test("a changed line in a cell of several paragraphs is rewritten alone, and a cell holding a nested table is not rewritten", async () => {
  const office = '<w:p><w:r><w:t xml:space="preserve">Registered office:</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>1 Mill Lane</w:t></w:r></w:p>'
  const details = p("Northwind Cartons Limited", "<w:rPr><w:b/></w:rPr>") + p("Company number 01234567", "<w:rPr><w:i/></w:rPr>") + office
  const parties = originalWith(p("Parties") + tbl([["Party", "Details"], ["Supplier", details]]) + p("End"))
  const read = readDocxTables(parties)[0]!.rows
  expect(read[1]).toEqual(["Supplier", "Northwind Cartons Limited\nCompany number 01234567\nRegistered office: 1 Mill Lane"])
  const renumbered = await carried(markdown("Parties", pipeTable(read.map(row => row.map(text => text.replace("01234567", "07654321")))), "End"), parties)
  expect(renumbered.carries[0]).toMatchObject({ outcome: "edited", cells: ["row 2, column 2"] })
  expect(tableXml(renumbered.bytes)[0]).toContain(`<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${details.replace("01234567", "07654321")}</w:tc>`)

  const inner = tbl([["Name", ""], ["Title", ""]])
  const layout = tbl([["For Alpha", "For Beta"], [p("Signed by Alpha") + inner + "<w:p/>", p("Signed by Beta") + inner + "<w:p/>"]])
  const execution = originalWith(p("Execution") + layout + p("End"))
  const grid = readDocxTables(execution)[0]!.rows
  const kept = await carried(markdown("Execution", pipeTable(grid), "End"), execution)
  expect(kept.carries[0]!.outcome).toBe("carried")
  expect(documentXml(kept.bytes).match(/<w:tbl>/g)).toHaveLength(3)
  const renamed = grid.map(row => row.map(text => text.replace("Signed by Beta", "Signed by Beta Holdings")))
  const refused = await carried(markdown("Execution", pipeTable(renamed), "End"), execution)
  expect(refused.carries[0]).toMatchObject({ outcome: "drafted", reason: "the drafted table changes a cell holding a nested table or a content control (row 2, column 2)" })
  expect(readDocxTables(refused.bytes)[0]!.rows).toEqual(renamed)
})

test("a table inside a content control is read and carried, and its row-level controls keep their rows", async () => {
  const grid = [["Signature", ""], ["Name", ""], ["Date", ""]]
  const rows = grid.map(row => `<w:tr>${row.map(tc).join("")}</w:tr>`)
  const table = `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    `${rows[0]}${rows[1]}<w:sdt><w:sdtPr><w:id w:val="6"/></w:sdtPr><w:sdtContent>${rows[2]}</w:sdtContent></w:sdt></w:tbl>`
  const original = originalWith(p("Signed for Alpha Limited:") +
    `<w:sdt><w:sdtPr><w:alias w:val="Signature block"/><w:id w:val="5"/></w:sdtPr><w:sdtContent>${table}</w:sdtContent></w:sdt>` + p("End"))
  expect(readDocxTables(original)).toEqual([{ rows: grid, follows: "Signed for Alpha Limited:" }])
  const kept = await carried(markdown("Signed for Alpha Limited:", pipeTable(grid), "End"), original)
  expect(kept.carries[0]!.outcome).toBe("carried")
  expect(tableXml(kept.bytes)[0]).toContain('<w:sdt><w:sdtPr><w:id w:val="6"/></w:sdtPr><w:sdtContent><w:tr>')
  const witnessed = await carried(markdown("Signed for Alpha Limited:", pipeTable([...grid, ["Witness", ""]]), "End"), original)
  expect(witnessed.carries[0]).toMatchObject({ outcome: "edited", added: 1 })
  expect(readDocxTables(witnessed.bytes)[0]!.rows).toEqual([...grid, ["Witness", ""]])
  const undated = await carried(markdown("Signed for Alpha Limited:", pipeTable(grid.slice(0, 2)), "End"), original)
  expect(undated.carries[0]).toMatchObject({ outcome: "edited", removed: ["Date"] })
  expect(readDocxTables(undated.bytes)[0]!.rows).toEqual(grid.slice(0, 2))
  expect(tableXml(undated.bytes)[0]).not.toContain("<w:sdt>")
  // So too a table and rows inside custom XML markup.
  const marked = originalWith(p("Signed for Alpha Limited:") +
    `<w:customXml w:uri="urn:example:signatures" w:element="block">${table.replace(/<w:sdt>[\s\S]*<\/w:sdt>/, `<w:customXml w:element="row">${rows[2]}</w:customXml>`)}</w:customXml>` + p("End"))
  expect(readDocxTables(marked)).toEqual([{ rows: grid, follows: "Signed for Alpha Limited:" }])
  const carriedMarked = await carried(markdown("Signed for Alpha Limited:", pipeTable(grid), "End"), marked)
  expect(carriedMarked.carries[0]!.outcome).toBe("carried")
  expect(readDocxTables(carriedMarked.bytes)[0]!.rows).toEqual(grid)
})

test("a table whose content refers to its own package is kept as drafted, and one that only looks so is carried", async () => {
  const numbered = '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>'
  const cases: [string, string][] = [
    [`<w:p><w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Value</w:t></w:r></w:ins></w:p>`, "the original table holds a ins element, which refers to its own package"],
    [`<w:p><w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Old</w:delText></w:r></w:del><w:r><w:t>Value</w:t></w:r></w:p>`, "the original table holds a del element, which refers to its own package"],
    [`<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Value</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`, "the original table holds a commentRangeStart element, which refers to its own package"],
    [`<w:p><w:hyperlink r:id="rId9"><w:r><w:t>Value</w:t></w:r></w:hyperlink></w:p>`, "the original table links to another part of its package"],
    [`<w:p><w:fldSimple w:instr=" DOCPROPERTY Company "><w:r><w:t>Value</w:t></w:r></w:fldSimple></w:p>`, "the original table holds a field, which refers to its own package"],
    [`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF Price \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Value</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`, "the original table holds a field, which refers to its own package"],
    [`<w:p><w:hyperlink w:anchor="Price"><w:r><w:t>Value</w:t></w:r></w:hyperlink></w:p>`, "the original table links to a bookmark, which is not carried"],
    [`<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>Value</w:t></w:r></w:p>`, "the original table uses a style with list numbering, which refers to its own package"],
  ]
  const outcome = async (cell: string) => {
    const original = originalWith(p("Intro") + tbl([["Item", "Detail"], ["Thing", cell]]) + p("End"), styles => styles.replace("</w:styles>", `${numbered}</w:styles>`))
    expect(readDocxTables(original)[0]!.rows[1]).toEqual(["Thing", "Value"])
    return carried(markdown("Intro", pipeTable([["Item", "Detail"], ["Thing", "Value"]]), "End"), original)
  }
  for (const [cell, reason] of cases) expect((await outcome(cell)).carries[0]).toMatchObject({ outcome: "drafted", reason })
  const off = await outcome(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>Value</w:t></w:r></w:p>`)
  expect(off.carries[0]!.outcome).toBe("carried")
  // A bookmark's id belongs to the original package: the cell is carried without it.
  const marked = await outcome(`<w:p><w:bookmarkStart w:id="0" w:name="Price"/><w:r><w:t>Value</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`)
  expect(marked.carries[0]!.outcome).toBe("carried")
  expect(tableXml(marked.bytes)[0]).not.toContain("bookmark")
  // A character style linked to a numbered paragraph style brings that style without its numbering.
  const linked = '<w:style w:type="character" w:styleId="NoteMark"><w:name w:val="Note Mark"/><w:link w:val="NoteList"/><w:rPr><w:i/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="NoteList"><w:name w:val="Note List"/><w:link w:val="NoteMark"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>'
  const noted = originalWith(p("Intro") + tbl([["Item", "Detail"], ["Thing", p("Value", '<w:rPr><w:rStyle w:val="NoteMark"/></w:rPr>')]]) + p("End"),
    styles => styles.replace("</w:styles>", `${linked}</w:styles>`))
  const linking = await carried(markdown("Intro", pipeTable([["Item", "Detail"], ["Thing", "Value"]]), "End"), noted)
  expect(linking.carries[0]!.outcome).toBe("carried")
  const copied = strFromU8(unzipSync(linking.bytes)["word/styles.xml"]!)
  expect(copied).toContain('w:styleId="NoteList"')
  expect(copied).not.toContain("<w:numPr>")
})

test("a word processor's original keeps its tables and its rows' spacing", async () => {
  const original = revisionTablesWordStyleDocx()
  expect(readDocxTables(original).map(table => table.rows)).toEqual(TABLE_FIXTURE.tables)
  const result = await carried(draft(), original)
  expect(result.carries.map(row => row.outcome)).toEqual(["carried", "carried", "carried", "carried"])
  expect(readDocxTables(result.bytes).map(table => table.rows)).toEqual(TABLE_FIXTURE.tables)
  // Its cell paragraphs are single-spaced with no space after; the rendered package's defaults would add both.
  const paragraphs = tableXml(result.bytes).flatMap(xml => xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [])
  expect(paragraphs.length).toBe(TABLE_FIXTURE.tables.flat(2).length)
  for (const paragraph of paragraphs) expect(paragraph).toContain('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
  const edited = await carried(draft({ particulars: particulars.map(row => row.map(cell => cell === "40,000 cartons" ? "45,000 cartons" : cell)) }), original)
  expect(edited.carries[0]).toMatchObject({ outcome: "edited", cells: ["row 3, column 2"] })
  expect(tableXml(edited.bytes)[0]).toBe(tableXml(result.bytes)[0]!.replace("<w:t>40,000 cartons</w:t>", '<w:t xml:space="preserve">45,000 cartons</w:t>'))
})

test("a carried table keeps the original's faces: its default font and a run's theme font", async () => {
  const defaults = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
  const theme = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Firm"><a:themeElements><a:fontScheme name="Firm">' +
    '<a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'
  const themed = p("10 GBP", '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr>')
  const original = originalWith(p("Intro") + tbl([["Item", "Value"], ["Price", themed]]) + p("End"), styles => styles.replace(/(<w:styles\b[^>]*>)/, `$1${defaults}`),
    { "word/theme/theme1.xml": theme })
  const result = await carried(markdown("Intro", pipeTable([["Item", "Value"], ["Price", "10 GBP"]]), "End"), original)
  expect(result.carries[0]!.outcome).toBe("carried")
  const [table] = tableXml(result.bytes)
  expect(table).toContain('<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Item</w:t>')
  expect(table).toContain('<w:r><w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">10 GBP</w:t>')
  expect(table).not.toContain("Theme")
})

test("a row added after the header takes the form of a body row; a table whose rows all change in number and label stays as drafted", async () => {
  const added = [particulars[0]!, ["Minimum order", "5,000 cartons"], ...particulars.slice(1)]
  const result = await carried(draft({ particulars: added }))
  expect(result.carries[0]).toMatchObject({ outcome: "edited", added: 1 })
  const rows = tableXml(result.bytes)[0]!.match(/<w:tr>[\s\S]*?<\/w:tr>/g)!
  expect(rows[1]).toContain("Minimum order")
  expect(rows[1]).not.toContain("<w:b/>")
  expect(rows[1]).toBe(rows[2]!.replace("Products", "Minimum order").replace("Printed folding cartons, 300 gsm board", "5,000 cartons"))
  // Added after the second of two header rows, a row takes that row's form but does not repeat as a header.
  const headed = tbl([["Item", "Value"], ["(units)", "(GBP)"], ["Price", "10"]]).replace(/<w:tr>(?=[\s\S]*<w:tr>)/g, "<w:tr><w:trPr><w:tblHeader/></w:trPr>")
  const priced = await carried(markdown("Intro", pipeTable([["Item", "Value"], ["(units)", "(GBP)"], ["Fee", "5"], ["Price", "10"]]), "End"), originalWith(p("Intro") + headed + p("End")))
  expect(priced.carries[0]).toMatchObject({ outcome: "edited", added: 1 })
  const headedRows = tableXml(priced.bytes)[0]!.match(/<w:tr>[\s\S]*?<\/w:tr>/g)!
  expect(headedRows.map(row => row.includes("<w:tblHeader/>"))).toEqual([true, true, false, false])
  const replaced = [["Field", "Amount"], ["X", "9"], ["Y", "8"], ["Z", "7"], ["W", "6"], ["V", "5"], ["U", "4"]]
  const rebuilt = await carried(draft({ particulars: replaced }))
  expect(rebuilt.carries[0]).toMatchObject({ outcome: "drafted", reason: "no row of the drafted table begins as a row of the original does" })
  expect(readDocxTables(rebuilt.bytes)[0]!.rows).toEqual(replaced)
})

test("text drafted into a position a short row lacks keeps the table as drafted", async () => {
  const table = tbl([["Item", "Value"], ["Price", "10"]]).replace("</w:tbl>", `<w:tr>${tc("Note")}</w:tr></w:tbl>`)
  const original = originalWith(p("Intro") + table + p("End"))
  expect(readDocxTables(original)[0]!.rows).toEqual([["Item", "Value"], ["Price", "10"], ["Note", ""]])
  const result = await carried(markdown("Intro", pipeTable([["Item", "Value"], ["Price", "10"], ["Note", "Excluding VAT"]]), "End"), original)
  expect(result.carries[0]).toMatchObject({ outcome: "drafted", reason: "the drafted table writes into a position row 3 does not have (column 2)" })
  expect(readDocxTables(result.bytes)[0]!.rows[2]).toEqual(["Note", "Excluding VAT"])
  // A row added after the short row takes its form, which has no place for its second cell.
  const added = await carried(markdown("Intro", pipeTable([["Item", "Value"], ["Price", "10"], ["Note", ""], ["Tax", "20%"]]), "End"), original)
  expect(added.carries[0]).toMatchObject({ outcome: "drafted", reason: "a drafted row is added with text in a column the original's row 3 does not have" })
})

test("hidden text is not read as a cell's text, and stays hidden in a carried cell", async () => {
  const cell = `<w:p><w:r><w:t xml:space="preserve">Price </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>internal note</w:t></w:r></w:p>`
  const original = originalWith(p("Intro") + tbl([["Item", "Value"], [cell, "10"]]) + p("End"))
  expect(readDocxTables(original)[0]!.rows).toEqual([["Item", "Value"], ["Price", "10"]])
  const result = await carried(markdown("Intro", pipeTable([["Item", "Value"], ["Price", "10"]]), "End"), original)
  expect(result.carries[0]!.outcome).toBe("carried")
  expect(tableXml(result.bytes)[0]).toContain("<w:rPr><w:vanish/></w:rPr><w:t>internal note</w:t>")
})

test("a paragraph is carried as the original's only when its words and emphasis read alike, a clause number's bold aside", async () => {
  const typed = (markdown: string) => renderMarkdownToDocx(markdown, undefined, undefined, { numbering: false }).then(result => result.bytes)
  const parts = unzipSync(await typed([
    "# Supply terms",
    "**1.1** The Supplier shall deliver the “Goods”  on the agreed date.",
    "The Buyer shall pay the **Price** on delivery.",
    "Notice must be given in writing.",
    "The fee is [insert amount].",
    "Either party may terminate on notice.",
  ].join("\n\n")))
  // A revision tracked in the original is not carried, though its words read alike.
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace(/<w:r>(<w:t[^>]*>Either party)/, '<w:ins w:id="90" w:author="Firm"><w:r>$1').replace(/(on notice\.<\/w:t><\/w:r>)/, "$1</w:ins>"))
  const rendered = await typed([
    "# Supply terms",
    '1.1 The Supplier shall deliver the "Goods" on the agreed date.',
    "The Buyer shall pay the Price on delivery.",
    "Notice must be given in Writing.",
    "The fee is [insert amount].",
    "Either party may terminate on notice.",
  ].join("\n\n"))
  const { bytes, carry } = carryOriginalParagraphs(rendered, zipSync(parts))
  // The heading and the clause are the original's; emphasis, case, a placeholder and a tracked revision keep the others as drafted.
  expect(carry).toEqual({ paragraphs: 6, carried: 2, breaks: 0, spaced: 0 })
  const document = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const text = (paragraph: Element) => Array.from(paragraph.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join("")
  expect(Array.from(document.getElementsByTagNameNS(W, "p")).map(text).filter(Boolean)).toEqual(["Supply terms",
    "1.1 The Supplier shall deliver the “Goods”  on the agreed date.", "The Buyer shall pay the Price on delivery.",
    "Notice must be given in Writing.", "The fee is [insert amount].", "Either party may terminate on notice."])
  // Each paragraph keeps the drafted paragraph's pointer.
  const pointers = (bytes: Uint8Array) => (strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/_dh_p[0-9a-f]{12}/g) ?? []).join()
  expect(pointers(bytes)).toBe(pointers(rendered))
}, 60_000)

test("a heading drafted in bold is carried as the original's, while strikethrough, italics or part-emphasis keep a paragraph as drafted, and a clause number takes the original's spacing only where it is the only one", async () => {
  const typed = (markdown: string) => renderMarkdownToDocx(markdown, undefined, undefined, { numbering: false }).then(result => result.bytes)
  const original = await typed(["# Terms", "~~1.3 Old clause~~", "*Note to draft*", "1.4 The Buyer shall pay the Price.", "## Schedule", "(a)  Pay", "(b)  Deliver", "The end."].join("\n\n"))
  const rendered = await typed(["# Terms", "1.3 Old clause", "Note to draft", "1.4 The Buyer shall pay the **Price**.", "**Schedule**", "(a) New item", "(a) Pay on time", "(b) Deliver promptly", "The end."].join("\n\n"))
  const { bytes, carry } = carryOriginalParagraphs(rendered, original)
  // The title, the schedule heading and the last line are the original's; the
  // struck, italic and part-bold paragraphs stay as drafted. Of the drafted
  // "(a)" and "(b)", only "(b)" is drafted once, so it alone takes the two spaces.
  expect(carry).toEqual({ paragraphs: 8, carried: 3, breaks: 0, spaced: 1 })
  const document = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const paragraphs = Array.from(document.getElementsByTagNameNS(W, "p")).filter(paragraph => paragraph.getElementsByTagNameNS(W, "t").length)
  const text = (paragraph: Element) => Array.from(paragraph.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join("")
  expect(paragraphs.map(text)).toEqual(["Terms", "1.3 Old clause", "Note to draft", "1.4 The Buyer shall pay the Price.", "Schedule", "(a) New item", "(a) Pay on time", "(b)  Deliver promptly", "The end."])
  const emphasised = (paragraph: Element, name: string) => paragraph.getElementsByTagNameNS(W, name).length > 0
  expect(paragraphs.map(paragraph => [emphasised(paragraph, "b"), emphasised(paragraph, "i"), emphasised(paragraph, "strike")])).toEqual([
    [false, false, false], [false, false, false], [false, false, false], [true, false, false], [false, false, false], [false, false, false], [false, false, false], [false, false, false], [false, false, false]])
  expect(paragraphs[4]!.getElementsByTagNameNS(W, "pStyle")[0]?.getAttributeNS(W, "val")).toBe("Heading2")
}, 60_000)
