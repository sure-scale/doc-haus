import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { pendingRedlinesForDoc } from "../lib/redlines"
import { renderMarkdownToDocx } from "../lib/render"
import { bake } from "../../services/ingest/src/redline"
import redline from "../tool/redline"
import trackedChanges from "../tool/tracked-changes"
import wordIntegration from "../tool/word-integration"

// Word list numbering prints "2.1" before the clause; the read view shows it,
// so a model may copy it into find text and replacement. The number is the
// list's, never text: the edit lands on that clause without printing it twice.
const previous = { root: process.env.WORKSPACE_ROOT, ingest: process.env.INGEST_URL }
const roots: string[] = []
afterEach(() => {
  for (const [key, value] of [["WORKSPACE_ROOT", previous.root], ["INGEST_URL", previous.ingest]] as const)
    if (value === undefined) delete process.env[key]
    else process.env[key] = value
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
})

async function numbered(markdown = "# Supply Agreement\n\n## 1. Delivery\n\n1.1 The Supplier shall deliver the goods.\n\n## 2. Payment\n\n2.1 The Customer shall pay within thirty days.") {
  const root = mkdtempSync(path.join(tmpdir(), "dh-list-label-tools-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  const document = path.join(root, "supply.docx")
  writeFileSync(document, (await renderMarkdownToDocx(markdown)).bytes)
  const context = { sessionID: "labels", messageID: "labels", agent: "redliner", directory: root, worktree: root, abort: new AbortController().signal, metadata() {}, async ask() {} }
  return { root, document, context }
}

test("a tracked change or redline whose find text and replacement copy the clause's list number is proposed without it", async () => {
  const { root, document, context } = await numbered()
  const comment = "Reflects the agreed payment term."
  await trackedChanges.execute({ document: "supply.docx", find: "2.1 The Customer shall pay", replace: "2.1 The Customer shall promptly pay", comment }, context)
  expect(pendingRedlinesForDoc(root, document).map(row => [row.scope, row.find_text, row.new_text]))
    .toEqual([["phrase", "The Customer shall pay", "The Customer shall promptly pay"]])
  await redline.execute({ document: "supply.docx", clause: "2.1 The Customer shall pay within thirty days.", replacement: "2.1 The Customer shall pay within fourteen days.", comment }, context)
  expect(pendingRedlinesForDoc(root, document).map(row => [row.scope, row.find_text, row.new_text]))
    .toEqual([["clause", "The Customer shall pay within thirty days.", "The Customer shall pay within fourteen days."]])
})

test("a plain edit whose find text and replacement copy the clause's list number edits that clause and prints the number once", async () => {
  const { context } = await numbered()
  let saved: Uint8Array | undefined
  const ingest = Bun.serve({ port: 0, async fetch(request) {
    saved = await ((await request.formData()).get("file") as File).bytes()
    return Response.json({ ok: true })
  } })
  process.env.INGEST_URL = ingest.url.href.replace(/\/$/, "")
  try {
    const result = await wordIntegration.execute({ document: "supply.docx", action: "replace", find: "2.1 The Customer shall pay within thirty days.", replace: "2.1 The Customer shall pay within fourteen days." }, context)
    expect(typeof result === "string" ? result : result.output).toContain("Replaced 1 occurrence(s)")
    const session = (await docxodus()).openDocxSession(saved!, {})
    try {
      const clause = session.findByText("fourteen days")!
      expect([clause.autoNumberPrefix, session.getAnchorInfo(clause.id)?.visibleText]).toEqual(["2.1", "2.1 The Customer shall pay within fourteen days."])
    } finally {
      session.close()
    }
  } finally {
    ingest.stop(true)
  }
})

// Clauses 1.1, 2.1 and 3.1 open with the same words; only a copied number
// tells them apart, and the recorded change keeps no number.
const sharedOpening = ["# Supply Agreement", "## 1. Delivery", "1.1 The Supplier shall deliver the Goods to the Site.", "## 2. Installation",
  "2.1 The Supplier shall install the Goods within ten days.", "## 3. Insurance", "3.1 The Supplier shall insure the Goods until Acceptance."].join("\n\n")

async function clauses(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, {})
  try {return session.findAllByText("The Supplier").map(match => session.getAnchorInfo(match.id)?.visibleText)} finally {session.close()}
}

async function accepted(root: string, document: string) {
  const rows = pendingRedlinesForDoc(root, document).map(row => ({ ...row, doc_path: document, doc_name: "supply.docx", old_text: row.find_text, status: "pending" as const, created_at: 0 }))
  return clauses(await bake(await Bun.file(document).bytes(), rows))
}

test("a tracked change or redline whose find text without its copied number also opens other clauses is refused until it names one", async () => {
  const comment = "Reflects the agreed insurance term."
  const tracked = await numbered(sharedOpening)
  const refusals = [
    await trackedChanges.execute({ document: "supply.docx", find: "3.1 The Supplier shall", replace: "3.1 The Supplier must", comment }, tracked.context),
    await redline.execute({ document: "supply.docx", clause: "3.1 The Supplier shall", replacement: "3.1 The Supplier must insure the Goods until Completion.", comment }, tracked.context),
  ]
  for (const refusal of refusals) expect(typeof refusal === "string" ? refusal : refusal.output).toContain("Quote more of the text of 3.1")
  expect(pendingRedlinesForDoc(tracked.root, tracked.document)).toEqual([])
  await trackedChanges.execute({ document: "supply.docx", find: "3.1 The Supplier shall insure", replace: "3.1 The Supplier must insure", comment }, tracked.context)
  expect(await accepted(tracked.root, tracked.document)).toEqual(["1.1 The Supplier shall deliver the Goods to the Site.", "2.1 The Supplier shall install the Goods within ten days.", "3.1 The Supplier must insure the Goods until Acceptance."])
  const rewritten = await numbered(sharedOpening)
  await redline.execute({ document: "supply.docx", clause: "3.1 The Supplier shall insure", replacement: "3.1 The Supplier must insure the Goods until Completion.", comment }, rewritten.context)
  expect(await accepted(rewritten.root, rewritten.document)).toEqual(["1.1 The Supplier shall deliver the Goods to the Site.", "2.1 The Supplier shall install the Goods within ten days.", "3.1 The Supplier must insure the Goods until Completion."])
})

test("a plain edit whose find text copies one clause's number edits only that clause", async () => {
  const { context } = await numbered(sharedOpening)
  let saved: Uint8Array | undefined
  const ingest = Bun.serve({ port: 0, async fetch(request) {
    saved = await ((await request.formData()).get("file") as File).bytes()
    return Response.json({ ok: true })
  } })
  process.env.INGEST_URL = ingest.url.href.replace(/\/$/, "")
  try {
    await wordIntegration.execute({ document: "supply.docx", action: "replace", find: "3.1 The Supplier shall", replace: "3.1 The Supplier must" }, context)
    expect(await clauses(saved!)).toEqual(["1.1 The Supplier shall deliver the Goods to the Site.", "2.1 The Supplier shall install the Goods within ten days.", "3.1 The Supplier must insure the Goods until Acceptance."])
  } finally {
    ingest.stop(true)
  }
})

// Only the number Word prints is not text. A number typed into the text is
// quoted as written; a find copying it with other spacing names nothing.
test("a number typed into a paragraph's text is not taken for Word's list number", async () => {
  const { root, document, context } = await numbered("# Notice\n\nThe parties agree:\n\n3.1\tThe Supplier shall insure the Goods.")
  const result = await trackedChanges.execute({ document: "supply.docx", find: "3.1 The Supplier shall insure", replace: "3.1 The Supplier must insure", comment: "Reflects the agreed insurance term." }, context)
  expect(typeof result === "string" ? result : result.output).toStartWith("Text not found")
  expect(pendingRedlinesForDoc(root, document)).toEqual([])
})
