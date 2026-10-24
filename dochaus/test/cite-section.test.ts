import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import mammoth from "mammoth"
import { ingestDocument } from "../../services/ingest/src/ingest"
import { verifyQuote } from "../lib/quote-verify"
import { renderMarkdownToDocx } from "../lib/render"
import cite from "../tool/cite"

// A verified quotation carries the indexed section that holds it, so the chat's
// Sources list can group and open it like a search hit. Synthetic text only.

const directories: string[] = []
afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

// Rendered drafts carry Word list numbering, not typed numbers; ingest reads
// the labels Word shows (docx-list-labels.ts).
const render = async (blocks: string[]) => Buffer.from((await renderMarkdownToDocx(blocks.join("\n\n"))).bytes)
const TERMS = await render([
  "## 1. Delivery",
  "The Supplier shall deliver the synthetic cartons to the Customer's site.",
  "## 2. Payment",
  "The Customer shall pay each synthetic invoice within thirty days of receipt.",
])

async function citeQuote(directory: string, name: string, quote: string) {
  const result = await cite.execute(
    { docPath: name, documentName: name, quote, reason: "Shows the term.", confidence: 5 },
    { directory, sessionID: crypto.randomUUID() } as never,
  )
  if (typeof result === "string") throw new Error(result)
  return result.metadata.citations[0]!
}

function matterDir() {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-cite-section-")))
  directories.push(directory)
  return directory
}

test("a verified quotation carries the section of the indexed passage that holds it", async () => {
  const directory = matterDir()
  await ingestDocument(directory, "Supply terms.docx", TERMS)
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
  const directory = matterDir()
  await ingestDocument(directory, "Supply draft.docx", draft)
  expect((await citeQuote(directory, "Supply draft.docx", "pay each synthetic invoice within thirty days")).section).toBe("2.1")
  expect(verifyQuote({ directory, quote: "2.1 The Customer shall pay each synthetic invoice" }).found).toBe(true)
})

test("a document outside the index keeps the generic label", async () => {
  const directory = matterDir()
  writeFileSync(path.join(directory, "Loose note.docx"), TERMS)
  expect((await citeQuote(directory, "Loose note.docx", "pay each synthetic invoice within thirty days")).section).toBe("cited passage")
})
