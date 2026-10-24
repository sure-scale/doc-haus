import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strToU8, zipSync } from "fflate"
import { openDb } from "./db"
import { extractDocumentText, ingestDocument } from "./ingest"
import { reconstructText } from "./structure"
import { liveText } from "../../../dochaus/lib/extract"
import { verifyQuote } from "../../../dochaus/lib/quote-verify"

// Every fixture is invented text: no actual party, transaction or reviewer.

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const PART = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const lvl = (ilvl: number, format: string, text: string) => `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/></w:lvl>`
const NUMBERING = `<w:abstractNum w:abstractNumId="0">${lvl(0, "decimal", "%1.")}${lvl(1, "decimal", "%1.%2")}${lvl(2, "lowerLetter", "(%3)")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`
const NORMAL = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`
const docx = (body: string, styles = NORMAL) => Buffer.from(zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${PART}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${PART}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${PART}.numbering+xml"/></Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
  "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdN" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`),
  "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`),
  "word/styles.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>${styles}</w:styles>`),
  "word/numbering.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>${NUMBERING}</w:numbering>`),
}))
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const numbered = (ilvl: number, content: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>${content}</w:p>`
const typed = (content: string) => `<w:p>${content}</w:p>`
const CHANGE = 'w:author="Counterparty" w:date="2026-01-01T00:00:00Z"'
const change = (kind: "ins" | "del" | "moveFrom" | "moveTo", id: number, content: string) => `<w:${kind} w:id="${id}" ${CHANGE}>${content}</w:${kind}>`
// A whole clause under a pending change: its paragraph mark is tracked too.
const changedClause = (kind: "ins" | "del" | "moveFrom" | "moveTo", id: number, content: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:rPr><w:${kind} w:id="${id}" ${CHANGE}/></w:rPr></w:pPr>${change(kind, id + 1, content)}</w:p>`
const moveRange = (kind: "moveFrom" | "moveTo", id: number, paragraph: string) =>
  `<w:${kind}RangeStart w:id="${id}" ${CHANGE} w:name="move1"/>${paragraph}<w:${kind}RangeEnd w:id="${id}"/>`

/** The text both lockstep readers give a file: ingest's, and the citation
 * verifier's, which ends in one more newline. */
async function readers(name: string, buffer: Buffer) {
  const file = path.join(directory, name)
  writeFileSync(file, buffer)
  return { ingest: await extractDocumentText(name, buffer), live: await liveText(file) }
}

const CLAUSES: [number, string, string][] = [
  [0, "1.", "Delivery"],
  [1, "1.1", "The Supplier shall deliver the synthetic cartons to the Customer's site."],
  [0, "2.", "Payment"],
  [1, "2.1", "The Customer shall pay each synthetic invoice within thirty days of receipt."],
  [2, "(a)", "by bank transfer; or"],
  [2, "(b)", "by any other agreed means."],
]
const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-list-labels-")))
afterAll(() => rmSync(directory, { recursive: true, force: true }))

test("a natively numbered source indexes like its typed-number twin, and the citation verifier reads the same text", async () => {
  await ingestDocument(directory, "Numbered terms.docx", docx(CLAUSES.map(([ilvl, , text]) => numbered(ilvl, run(text))).join("")))
  await ingestDocument(directory, "Typed terms.docx", docx(CLAUSES.map(([, label, text]) => typed(run(`${label} ${text}`))).join("")))
  const db = openDb(directory)
  try {
    const canonical = reconstructText(db, path.join(directory, "Numbered terms.docx"))
    expect(canonical).toBe(reconstructText(db, path.join(directory, "Typed terms.docx")))
    expect(canonical).toBe(await liveText(path.join(directory, "Numbered terms.docx")))
    const sections = (name: string) => db.query("SELECT label, heading FROM sections WHERE doc_name = ? ORDER BY ordinal").all(name)
    expect(sections("Numbered terms.docx")).toEqual(sections("Typed terms.docx"))
    expect(sections("Numbered terms.docx")).toContainEqual(expect.objectContaining({ label: "2" }))
  } finally {
    db.close()
  }
  const quote = verifyQuote({ directory, document: "Numbered terms.docx", quote: "2.1 The Customer shall pay each synthetic invoice" })
  expect(quote.found).toBe(true)
}, 120_000)

test("a numbered markup keeps its tracked-changes summary first and numbers its proposed clauses", async () => {
  const inserted = `<w:ins w:id="1" w:author="Counterparty" w:date="2026-01-01T00:00:00Z">${run("Interest accrues on late payments.")}</w:ins>`
  const text = await extractDocumentText("Markup.docx", docx(numbered(0, run("Payment")) + numbered(1, run("The Customer shall pay each invoice.")) + numbered(1, inserted)))
  expect(text).toStartWith("[This Word file contains 1 pending tracked change by Counterparty")
  expect(text).toContain("1. Payment\n\n1.1 The Customer shall pay each invoice.\n\n1.2 [Tracked insertion by Counterparty: Interest accrues on late payments.]")
})

test("a markup that deletes or moves a clause numbers the clauses as Word does, and both readers read it alike", async () => {
  const inspect = "The Customer shall inspect the cartons."
  const { ingest, live } = await readers("Moves.docx", docx([
    numbered(0, run("Supply")),
    numbered(1, run("The Supplier shall deliver the cartons.")),
    changedClause("del", 1, `<w:r><w:delText xml:space="preserve">The Supplier shall insure the cartons.</w:delText></w:r>`),
    numbered(1, run("The Customer shall collect the cartons.")),
    moveRange("moveFrom", 3, changedClause("moveFrom", 4, run(inspect))),
    numbered(1, run("The Customer shall store the cartons.")),
    moveRange("moveTo", 6, changedClause("moveTo", 7, run(inspect))),
    changedClause("ins", 8, run("The Customer shall return the pallets.")),
  ].join("")))
  expect(live).toBe(`${ingest}\n`)
  // A clause deleted or moved away is out of the numbering; Word shows its number where it moves to.
  expect(ingest).toEndWith([
    "1. Supply",
    "1.1 The Supplier shall deliver the cartons.",
    "[Tracked deletion by Counterparty: The Supplier shall insure the cartons.]",
    "1.2 The Customer shall collect the cartons.",
    `[Tracked move from here by Counterparty: ${inspect}]`,
    "1.3 The Customer shall store the cartons.",
    `1.4 [Tracked move to here by Counterparty: ${inspect}]`,
    "1.5 [Tracked insertion by Counterparty: The Customer shall return the pallets.]",
    "",
  ].join("\n\n"))
}, 120_000)

test("with a numbered default paragraph style, the tracked-changes summary takes no number", async () => {
  const styles = `<w:style w:type="paragraph" w:default="1" w:styleId="Clause"><w:name w:val="Clause"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>`
  const { ingest, live } = await readers("Default numbered.docx", docx([
    typed(run("Delivery")),
    typed(run("Payment ") + change("ins", 1, run("on time"))),
    typed(run("Term")),
  ].join(""), styles))
  expect(live).toBe(`${ingest}\n`)
  expect(ingest).toStartWith("[This Word file contains 1 pending tracked change by Counterparty")
  expect(ingest).toEndWith("]\n\n1. Delivery\n\n2. Payment [Tracked insertion by Counterparty: on time]\n\n3. Term\n\n")
}, 120_000)
