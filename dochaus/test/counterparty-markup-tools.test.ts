import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { pendingRedlinesForDoc } from "../lib/redlines"
import redline from "../tool/redline"
import trackedChanges from "../tool/tracked-changes"
import wordIntegration from "../tool/word-integration"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "./support/counterparty-markup-fixture"

// #200 scope 4, the touch rule on the chat tools: a paragraph with the other
// side's pending change waits for it to be decided. These tools' proposals are
// baked untracked and their direct edits are untracked, so in such a paragraph
// doc.haus's words would become the other side's.
const previous = process.env.INGEST_URL
const roots: string[] = []
afterEach(() => {
  if (previous === undefined) delete process.env.INGEST_URL
  else process.env.INGEST_URL = previous
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
})

const refusal = (text: string, author: string, verb = "proposed") =>
  `The change to ${JSON.stringify(text)} was not ${verb}: that paragraph has a pending change by ${author}. Accept or reject it under Their changes, then request edits again.`

test("the chat tools refuse a paragraph with their pending change, and record or save nothing", async () => {
  const root = realpathSync(mkdtempSync(path.join(tmpdir(), "dh-markup-tools-")))
  roots.push(root)
  const document = path.join(root, COUNTERPARTY_MARKUP.name)
  writeFileSync(document, counterpartyMarkupDocx())
  const context = { sessionID: "markup", messageID: "markup", agent: "redliner", directory: root, worktree: root, abort: new AbortController().signal, metadata() {}, async ask() {} }
  const comment = "Reflects our client's position."
  const saved: Uint8Array[] = []
  const ingest = Bun.serve({ port: 0, async fetch(request) {
    saved.push(await ((await request.formData()).get("file") as File).bytes())
    return Response.json({ ok: true })
  } })
  process.env.INGEST_URL = ingest.url.href.replace(/\/$/, "")
  try {
    // Beside Morgan's insertion in clause 2, and in Avery's moved clause 3.
    expect(await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "twelve months", replace: "eighteen months", comment }, context))
      .toBe(refusal("twelve months", "Morgan Okafor"))
    expect(await redline.execute({ document: COUNTERPARTY_MARKUP.name, clause: "Notices must be in writing.", replacement: "Notices must be in writing and signed.", comment }, context))
      .toBe(refusal("Notices must be in writing.", "Avery Lindqvist"))
    expect(pendingRedlinesForDoc(root, document)).toEqual([])
    // Inside Avery's insertion in clause 1.
    expect(await wordIntegration.execute({ document: COUNTERPARTY_MARKUP.name, action: "replace", find: "forty-five days of each", replace: "sixty days of each" }, context))
      .toBe(refusal("forty-five days of each", "Avery Lindqvist", "made"))
    expect(saved).toEqual([])
    expect(new Uint8Array(readFileSync(document))).toEqual(counterpartyMarkupDocx())

    // Clause 9 has no pending change, so it is proposed as before.
    await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "the Province", replace: "the State", comment }, context)
    expect(pendingRedlinesForDoc(root, document).map(row => [row.find_text, row.new_text])).toEqual([["the Province", "the State"]])
  } finally { ingest.stop(true) }
})
