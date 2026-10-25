import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { applyDocxLocalization, createDocxLocalization, createDocxRevision } from "../lib/docx-amendments"
import { docxodus } from "../lib/docxodus"
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
  `The change to ${JSON.stringify(text)} was not ${verb}: that paragraph has a pending change by ${author}. Accept or reject it under Tracked changes, then request edits again.`

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

// #299: doc.haus's own pending change (one an earlier Accept wrote tracked into
// their markup) holds no paragraph, as in Request edits: an edit beside it is
// proposed or made, and one that would rewrite it is refused with Prepare's reason.
test("the chat tools edit beside doc.haus's own pending change and refuse to rewrite it", async () => {
  const root = realpathSync(mkdtempSync(path.join(tmpdir(), "dh-markup-tools-ours-")))
  roots.push(root)
  const layer = { pendingRevisions: "layer" } as const
  const source = counterpartyMarkupDocx()
  const packet = await createDocxLocalization(source, [{ id: "law", before_quote: "the Province", replacement: "the State" }], {}, layer)
  const law = packet.candidates.find(row => row.paragraph.scope === "body")!
  const ours = (await applyDocxLocalization(source, packet, [{ amendment_id: "law", target_handle: law.target_handle }], { ...layer, revisionAuthor: "doc.haus" })).bytes
  const document = path.join(root, COUNTERPARTY_MARKUP.name)
  writeFileSync(document, ours)
  const context = { sessionID: "markup", messageID: "markup", agent: "redliner", directory: root, worktree: root, abort: new AbortController().signal, metadata() {}, async ask() {} }
  const comment = "Reflects our client's position."
  const saved: Uint8Array[] = []
  const ingest = Bun.serve({ port: 0, async fetch(request) {
    saved.push(await ((await request.formData()).get("file") as File).bytes())
    return Response.json({ ok: true })
  } })
  process.env.INGEST_URL = ingest.url.href.replace(/\/$/, "")
  const inside = (text: string, verb = "proposed") =>
    `The change to ${JSON.stringify(text)} was not ${verb}: it would rewrite your earlier change there, which is still a tracked change in this file. Make it in Word, or reject that change under Tracked changes and request edits again.`
  const revisions = async (bytes: Uint8Array) => {
    const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
    try { return session.listRevisions().map(({ id, author, type, text }) => ({ id, author, type, text })) } finally { session.close() }
  }
  const proposed = () => pendingRedlinesForDoc(root, document).map(row => [row.find_text, row.new_text])
  try {
    // Beside it, in the same paragraph.
    await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "governed by", replace: "construed under", comment }, context)
    expect(proposed()).toEqual([["governed by", "construed under"]])
    await wordIntegration.execute({ document: COUNTERPARTY_MARKUP.name, action: "replace", find: "governed by", replace: "construed under" }, context)
    expect(saved).toHaveLength(1)
    expect((await createDocxRevision(saved[0]!, layer)).paragraphs[9]!.text).toBe("9. Governing law. This Agreement is construed under the law of the State.")
    expect(await revisions(saved[0]!)).toEqual(await revisions(ours))

    // Inside its insertion of "State", and a rewrite of its whole clause.
    expect(await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "the State", replace: "the Commonwealth", comment }, context))
      .toBe(inside("the State"))
    expect(await redline.execute({ document: COUNTERPARTY_MARKUP.name, clause: "This Agreement is governed by the law of the State.", replacement: "This Agreement is construed under the law of the State.", comment }, context))
      .toBe(inside("This Agreement is governed by the law of the State."))
    expect(await wordIntegration.execute({ document: COUNTERPARTY_MARKUP.name, action: "replace", find: "the State", replace: "the Commonwealth" }, context))
      .toBe(inside("the State", "made"))
    expect(proposed()).toEqual([["governed by", "construed under"]])
    expect(saved).toHaveLength(1)
  } finally { ingest.stop(true) }
})
