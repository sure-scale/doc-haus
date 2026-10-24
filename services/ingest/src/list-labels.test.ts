import { afterAll, expect, spyOn, test } from "bun:test"
import { mkdtempSync, readFileSync, realpathSync, rmSync, statSync, unlinkSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strToU8, zipSync } from "fflate"
import mammoth from "mammoth"
import { openDb } from "./db"
import { assessExtraction } from "./extraction-quality"
import { extractDocumentText, ingestDocument, reindexNumberedSources } from "./ingest"
import { normalizeExtractedText } from "./sanitize"
import { reconstructText } from "./structure"
import { liveText } from "../../../dochaus/lib/extract"
import { verifyCitation } from "../../../dochaus/plugin/legal/citations"
import { verifyQuote } from "../../../dochaus/lib/quote-verify"

// Every fixture is invented text: no actual party, transaction or reviewer.

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const PART = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const lvl = (ilvl: number, format: string, text: string) => `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/></w:lvl>`
const NUMBERING = `<w:abstractNum w:abstractNumId="0">${lvl(0, "decimal", "%1.")}${lvl(1, "decimal", "%1.%2")}${lvl(2, "lowerLetter", "(%3)")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`
const docx = (body: string) => Buffer.from(zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${PART}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${PART}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${PART}.numbering+xml"/></Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
  "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdN" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`),
  "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`),
  "word/styles.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
  "word/numbering.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>${NUMBERING}</w:numbering>`),
}))
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const numbered = (ilvl: number, content: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>${content}</w:p>`
const typed = (content: string) => `<w:p>${content}</w:p>`

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

test("a numbered source indexed before its list numbers were read is re-indexed once at startup; an unnumbered one is left alone", async () => {
  const matter = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-list-labels-reindex-")))
  const warn = spyOn(console, "warn").mockImplementation(() => {})
  try {
    const terms = docx(CLAUSES.map(([ilvl, , text]) => numbered(ilvl, run(text))).join(""))
    // As ingest read a numbered file before: the same bytes, its text without the numbers.
    const before = async (name: string, bytes: Buffer, stale = (text: string) => text) => {
      const old = normalizeExtractedText(stale((await mammoth.extractRawText({ buffer: bytes })).value))
      await ingestDocument(matter, name, bytes, { extracted: { text: old.text, findings: old.findings, quality: assessExtraction(old.text, { format: "docx" }) } })
    }
    await before("Numbered terms.docx", terms)
    await before("Removed terms.docx", terms)
    // Indexed from text that differs from today's, as an older reader's may: still not re-indexed, as it has no list.
    await before("Plain terms.docx", docx(CLAUSES.map(([, , text]) => typed(run(text))).join("")), (text) => text.replace("thirty days", "30 days"))
    unlinkSync(path.join(matter, "Removed terms.docx"))
    const read = <T>(query: string, name: string) => {
      const db = openDb(matter)
      try {
        return db.query(query).all(name) as T[]
      } finally {
        db.close()
      }
    }
    const row = (name: string) => read("SELECT id, created_at, rev FROM documents WHERE name = ?", name)
    const citations = () => Promise.all(
      read<{ doc_path: string; section: string; text: string; char_start: number; char_end: number }>("SELECT doc_path, section, text, char_start, char_end FROM chunks WHERE doc_name = ? ORDER BY char_start", "Numbered terms.docx")
        .map((chunk) => ({ multi: chunk.text.includes("\n"), verified: verifyCitation({ documentName: "Numbered terms.docx", docPath: chunk.doc_path, section: chunk.section, excerpt: chunk.text, charStart: chunk.char_start, charEnd: chunk.char_end }) }))
        .map(async ({ multi, verified }) => ({ multi, verified: await verified })),
    )
    expect((await citations()).filter(({ multi, verified }) => multi && !verified)).not.toHaveLength(0)
    const plain = { row: row("Plain terms.docx"), modified: statSync(path.join(matter, "Plain terms.docx")).mtimeMs }

    expect(await reindexNumberedSources(matter)).toBe(1)
    const after = await citations()
    expect(after.some(({ multi }) => multi)).toBe(true)
    expect(after.every(({ verified }) => verified && !verified.reanchored)).toBe(true)
    expect(readFileSync(path.join(matter, "Numbered terms.docx")).equals(terms)).toBe(true)
    expect(row("Plain terms.docx")).toEqual(plain.row)
    expect(statSync(path.join(matter, "Plain terms.docx")).mtimeMs).toBe(plain.modified)
    expect(warn.mock.calls.map(([message]) => message)).toEqual(["numbered document re-index skipped"])

    const reindexed = row("Numbered terms.docx")
    expect(await reindexNumberedSources(matter)).toBe(0)
    expect(row("Numbered terms.docx")).toEqual(reindexed)
    expect(warn).toHaveBeenCalledTimes(1)
  } finally {
    warn.mockRestore()
    rmSync(matter, { recursive: true, force: true })
  }
}, 120_000)
