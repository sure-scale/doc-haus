import { afterAll, expect, test } from "bun:test"
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { ingestDocument } from "../../services/ingest/src/ingest"
import { renderMarkdownToDocx } from "../lib/render"
import cite from "../tool/cite"

// A verified quotation carries the indexed section that holds it, so the chat's
// Sources list can group and open it like a search hit. Synthetic text only.

const directories: string[] = []
afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

// Typed clause numbers, as many source documents carry them: ingest reads a
// source's text without Word's list labels (#197).
const TERMS = Buffer.from((await renderMarkdownToDocx([
  "## 1. Delivery",
  "The Supplier shall deliver the synthetic cartons to the Customer's site.",
  "## 2. Payment",
  "The Customer shall pay each synthetic invoice within thirty days of receipt.",
].join("\n\n"), undefined, undefined, { numbering: false })).bytes)

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

test("a document outside the index keeps the generic label", async () => {
  const directory = matterDir()
  writeFileSync(path.join(directory, "Loose note.docx"), TERMS)
  expect((await citeQuote(directory, "Loose note.docx", "pay each synthetic invoice within thirty days")).section).toBe("cited passage")
})
