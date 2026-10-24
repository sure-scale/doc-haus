import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { pendingRedlinesForDoc } from "../lib/redlines"
import { renderMarkdownToDocx } from "../lib/render"
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

async function numbered() {
  const root = mkdtempSync(path.join(tmpdir(), "dh-list-label-tools-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  const document = path.join(root, "supply.docx")
  writeFileSync(document, (await renderMarkdownToDocx("# Supply Agreement\n\n## 1. Delivery\n\n1.1 The Supplier shall deliver the goods.\n\n## 2. Payment\n\n2.1 The Customer shall pay within thirty days.")).bytes)
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
