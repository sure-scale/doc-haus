import { afterAll, beforeAll, describe, expect, test } from "bun:test"
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import mammoth from "mammoth"
import { utils, write, type WorkSheet } from "xlsx"
import { ingestDocument } from "../../services/ingest/src/ingest"
import { verifyQuote } from "../lib/quote-verify"
import { renderMarkdownToDocx } from "../lib/render"
import cite from "../tool/cite"

const directories: string[] = []
afterAll(() => directories.splice(0).forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function matterDir(prefix: string) {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), prefix)))
  directories.push(directory)
  return directory
}

function workbook(sheet: WorkSheet) {
  const book = utils.book_new()
  utils.book_append_sheet(book, sheet, "Sheet1")
  return Buffer.from(write(book, { type: "buffer", bookType: "xlsx" }))
}

// Formulas saved without calculated values; `label` adds a text cell beside them.
function formulas(label?: string): WorkSheet {
  const sheet: WorkSheet = { "!ref": label ? "A1:B2" : "B1:B2", B1: { t: "n", f: "SUM(C1:C9)" }, B2: { t: "n", f: "B1*2" } }
  if (label) sheet.A1 = { t: "s", v: label }
  return sheet
}

const QUOTE = "Synthetic deposit is payable within thirty days"

async function citeIn(name: string, sheet: WorkSheet, quote = QUOTE) {
  const directory = matterDir("dochaus-cite-quality-")
  await ingestDocument(directory, name, workbook(sheet))
  const result = await cite.execute(
    { docPath: name, documentName: name, quote, reason: "Shows the deposit term.", confidence: 5 },
    { directory, sessionID: crypto.randomUUID() } as never,
  )
  return typeof result === "string"
    ? { output: result, citations: undefined, rejected: undefined }
    : { output: result.output, citations: result.metadata?.citations, rejected: result.metadata?.rejectedCitations }
}

describe("cite extraction quality", () => {
  test("cite records a refused quotation for the answer's grounding state", async () => {
    const { output, citations, rejected } = await citeIn("Terms.xlsx", utils.aoa_to_sheet([[QUOTE]]), "Synthetic words absent from the workbook")
    expect(output).toContain("does not appear in Terms.xlsx")
    expect(citations).toBeUndefined()
    expect(rejected).toEqual([{ documentName: "Terms.xlsx", section: "" }])
  })

  test("cite refuses a document with no readable text", async () => {
    const { output, citations } = await citeIn("Unreadable.xlsx", formulas(), "=SUM(C1:C9)")
    expect(output).toContain("has no readable text")
    expect(output).toContain("Do not use this quotation")
    expect(citations).toBeUndefined()
  })

  test("cite verifies a quote in a degraded document with the warning and without 'safe to quote'", async () => {
    const { output, citations } = await citeIn("Partly computed.xlsx", formulas(QUOTE))
    expect(output).toContain(`\n${QUOTE}\n`)
    expect(output).toContain("[⚠ DEGRADED EXTRACTION]")
    expect(output).toContain("may be incomplete or garbled")
    expect(output).not.toContain("safe to quote")
    expect(citations).toEqual([expect.objectContaining({ verified: true, degradedExtraction: ["2 of 3 filled cells are formulas saved without a calculated value"] })])
  })

  test("cite on a document whose extraction is ok is unchanged", async () => {
    const { output, citations } = await citeIn("Terms.xlsx", utils.aoa_to_sheet([[QUOTE], ["Synthetic notices are given in writing."]]))
    expect(output).toContain("safe to quote")
    expect(output).not.toContain("DEGRADED EXTRACTION")
    expect(citations).toEqual([expect.not.objectContaining({ degradedExtraction: expect.anything() })])
  })
})

describe("cite section attribution", () => {
  const render = async (blocks: string[]) => Buffer.from((await renderMarkdownToDocx(blocks.join("\n\n"))).bytes)
  let terms: Buffer

  beforeAll(async () => {
    terms = await render([
      "## 1. Delivery",
      "The Supplier shall deliver the synthetic cartons to the Customer's site.",
      "## 2. Payment",
      "The Customer shall pay each synthetic invoice within thirty days of receipt.",
    ])
  })

  async function citeQuote(directory: string, name: string, quote: string) {
    const result = await cite.execute(
      { docPath: name, documentName: name, quote, reason: "Shows the term.", confidence: 5 },
      { directory, sessionID: crypto.randomUUID() } as never,
    )
    if (typeof result === "string") throw new Error(result)
    return result.metadata.citations[0]!
  }

  test("a verified quotation carries the section of the indexed passage that holds it", async () => {
    const directory = matterDir("dochaus-cite-section-")
    await ingestDocument(directory, "Supply terms.docx", terms)
    const payment = await citeQuote(directory, "Supply terms.docx", "pay each synthetic invoice within thirty days")
    const delivery = await citeQuote(directory, "Supply terms.docx", "deliver the synthetic cartons")
    expect(payment.section).toBe("2")
    expect(delivery.section).toBe("1")
  })

  test("a re-uploaded draft keeps its clause numbers in citations and quote checks", async () => {
    const draft = await render([
      "## 1. Delivery",
      "The Supplier shall deliver the synthetic cartons to the Customer's site.",
      "## 2. Payment",
      "2.1 The Customer shall pay each synthetic invoice within thirty days of receipt.",
    ])
    // The number is Word list numbering, not text in the paragraph's runs.
    expect((await mammoth.extractRawText({ buffer: draft })).value).not.toContain("2.1")
    const directory = matterDir("dochaus-cite-section-")
    await ingestDocument(directory, "Supply draft.docx", draft)
    expect((await citeQuote(directory, "Supply draft.docx", "pay each synthetic invoice within thirty days")).section).toBe("2.1")
    expect(verifyQuote({ directory, quote: "2.1 The Customer shall pay each synthetic invoice" }).found).toBe(true)
  })

  test("a document outside the index keeps the generic label", async () => {
    const directory = matterDir("dochaus-cite-section-")
    writeFileSync(path.join(directory, "Loose note.docx"), terms)
    expect((await citeQuote(directory, "Loose note.docx", "pay each synthetic invoice within thirty days")).section).toBe("cited passage")
  })
})
