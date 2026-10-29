import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { applyDocxLocalization, createDocxLocalization } from "../lib/docx-amendments"
import { pendingRedlinesForDoc, recordRedlineProposal, supersedeMatchingRedlines } from "../lib/redlines"
import { renderMarkdownToDocx } from "../lib/render"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
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

// The document as ingest registers it in the matter (an uploaded markup is a
// source), with the redlines table the tools write.
function register(root: string, role: "source" | "deliverable") {
  mkdirSync(path.join(root, ".dochaus"), { recursive: true })
  const db = new Database(path.join(root, ".dochaus", "legal.db"))
  try {
    db.exec("CREATE TABLE documents (name TEXT, role TEXT)")
    db.query("INSERT INTO documents VALUES (?, ?)").run(COUNTERPARTY_MARKUP.name, role)
  } finally { db.close() }
  supersedeMatchingRedlines(root, path.join(root, COUNTERPARTY_MARKUP.name), () => false)
}

// word-integration's edit is untracked, so it would not show in their markup.
const untracked = "This uploaded markup has tracked changes, so an untracked edit would not show in it. Use Request edits, which makes tracked changes."
// A tool context in which asking the lawyer to approve fails the test: a file
// the tools refuse is refused before anyone is asked.
const unasked = <T extends object>(context: T) => ({ ...context, async ask(): Promise<void> { throw new Error("Approval was requested for a refused edit.") } })

const refusal = (text: string, author: string, verb = "proposed") =>
  `The change to “${text}” was not ${verb}: that paragraph has a pending change by ${author}. Accept or reject it under Tracked changes, then request edits again.`

test("the chat tools refuse a paragraph with their pending change, and record or save nothing", async () => {
  const root = realpathSync(mkdtempSync(path.join(tmpdir(), "dh-markup-tools-")))
  roots.push(root)
  const document = path.join(root, COUNTERPARTY_MARKUP.name)
  writeFileSync(document, counterpartyMarkupDocx())
  register(root, "source")
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
    expect(await wordIntegration.execute({ document: COUNTERPARTY_MARKUP.name, action: "replace", find: "forty-five days of each", replace: "sixty days of each" }, unasked(context)))
      .toBe(untracked)
    expect(saved).toEqual([])
    expect(new Uint8Array(readFileSync(document))).toEqual(counterpartyMarkupDocx())

    // Clause 9 has no pending change, so it is proposed as before.
    await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "the Province", replace: "the State", comment }, context)
    expect(pendingRedlinesForDoc(root, document).map(row => [row.find_text, row.new_text])).toEqual([["the Province", "the State"]])
  } finally { ingest.stop(true) }
})

// #299: doc.haus's own pending change (one an earlier Accept wrote tracked into
// their markup) holds no paragraph, as in Request edits: an edit beside it is
// proposed, and one that would rewrite it is refused with Prepare's reason.
// word-integration, which edits untracked, makes neither.
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
  register(root, "source")
  const context = { sessionID: "markup", messageID: "markup", agent: "redliner", directory: root, worktree: root, abort: new AbortController().signal, metadata() {}, async ask() {} }
  const comment = "Reflects our client's position."
  const saved: Uint8Array[] = []
  const ingest = Bun.serve({ port: 0, async fetch(request) {
    saved.push(await ((await request.formData()).get("file") as File).bytes())
    return Response.json({ ok: true })
  } })
  process.env.INGEST_URL = ingest.url.href.replace(/\/$/, "")
  const inside = (text: string, verb = "proposed") =>
    `The change to “${text}” was not ${verb}: it would rewrite your earlier change there, which is still a tracked change in this file. Make it in Word, or reject that change under Tracked changes and request edits again.`
  const proposed = () => pendingRedlinesForDoc(root, document).map(row => [row.find_text, row.new_text])
  try {
    // Beside it, in the same paragraph.
    await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "governed by", replace: "construed under", comment }, context)
    expect(proposed()).toEqual([["governed by", "construed under"]])
    expect(await wordIntegration.execute({ document: COUNTERPARTY_MARKUP.name, action: "replace", find: "governed by", replace: "construed under" }, unasked(context)))
      .toBe(untracked)

    // Inside its insertion of "State", and a rewrite of its whole clause.
    expect(await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "the State", replace: "the Commonwealth", comment }, context))
      .toBe(inside("the State"))
    expect(await redline.execute({ document: COUNTERPARTY_MARKUP.name, clause: "This Agreement is governed by the law of the State.", replacement: "This Agreement is construed under the law of the State.", comment }, context))
      .toBe(inside("This Agreement is governed by the law of the State."))
    expect(await wordIntegration.execute({ document: COUNTERPARTY_MARKUP.name, action: "replace", find: "the State", replace: "the Commonwealth" }, unasked(context)))
      .toBe(untracked)
    expect(proposed()).toEqual([["governed by", "construed under"]])
    expect(saved).toEqual([])
    expect(new Uint8Array(readFileSync(document))).toEqual(ours)
  } finally { ingest.stop(true) }
})

// #299: anything but an uploaded source is refused while it has tracked
// changes, as Request edits refuses it. These tools' edits are untracked, so
// beside doc.haus's own pending change in a redline they would be invisible.
test("on a deliverable with doc.haus's own pending change the chat tools refuse the file, as Request edits does", async () => {
  const root = realpathSync(mkdtempSync(path.join(tmpdir(), "dh-markup-tools-deliverable-")))
  roots.push(root)
  const clean = (await renderMarkdownToDocx("Governing law. This Agreement is governed by the law of the Province.")).bytes
  const packet = await createDocxLocalization(clean, [{ id: "law", before_quote: "the Province", replacement: "the State" }], {}, { pendingRevisions: "layer" })
  const tracked = (await applyDocxLocalization(clean, packet, [{ amendment_id: "law", target_handle: packet.candidates[0]!.target_handle }],
    { pendingRevisions: "layer", revisionAuthor: "doc.haus" })).bytes
  const document = path.join(root, COUNTERPARTY_MARKUP.name)
  writeFileSync(document, tracked)
  register(root, "deliverable")
  const context = { sessionID: "markup", messageID: "markup", agent: "redliner", directory: root, worktree: root, abort: new AbortController().signal, metadata() {}, async ask() {} }
  const comment = "Reflects our client's position."
  const saved: Uint8Array[] = []
  const ingest = Bun.serve({ port: 0, async fetch(request) {
    saved.push(await ((await request.formData()).get("file") as File).bytes())
    return Response.json({ ok: true })
  } })
  process.env.INGEST_URL = ingest.url.href.replace(/\/$/, "")
  const refused = "This Word file contains tracked changes. Accept or reject them in Word before preparing changes."
  try {
    expect(await trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "governed by", replace: "construed under", comment }, unasked(context))).toBe(refused)
    expect(await redline.execute({ document: COUNTERPARTY_MARKUP.name, clause: "This Agreement is governed by the law of the State.", replacement: "This Agreement is construed under the law of the State.", comment }, unasked(context)))
      .toBe(refused)
    expect(await wordIntegration.execute({ document: COUNTERPARTY_MARKUP.name, action: "replace", find: "governed by", replace: "construed under" }, unasked(context))).toBe(refused)
    expect(pendingRedlinesForDoc(root, document)).toEqual([])
    expect(saved).toEqual([])
    expect(new Uint8Array(readFileSync(document))).toEqual(tracked)
  } finally { ingest.stop(true) }
})


test("tracked changes rechecks all occurrences after approval against a new live field", async () => {
  const root = realpathSync(mkdtempSync(path.join(tmpdir(), "dh-field-permission-")))
  roots.push(root)
  const document = path.join(root, COUNTERPARTY_MARKUP.name)
  const source = (await renderMarkdownToDocx("Notice requires VALUE.\n\nOther VALUE remains.")).bytes
  writeFileSync(document, source)
  register(root, "source")
  const existing = recordRedlineProposal(root, { docPath: document, docName: COUNTERPARTY_MARKUP.name, scope: "phrase", findText: "VALUE", oldText: "VALUE", newText: "OLD", author: "Counsel", reviewComment: "Existing", anchorId: "prior" })
  const before = pendingRedlinesForDoc(root, document)
  let arrived!: () => void, release!: () => void
  const asked = new Promise<void>(resolve => { arrived = resolve })
  const approval = new Promise<void>(resolve => { release = resolve })
  const context = { sessionID: "field", messageID: "field", agent: "redliner", directory: root, worktree: root, abort: new AbortController().signal,
    metadata() {}, async ask() { arrived(); await approval } }
  const request = trackedChanges.execute({ document: COUNTERPARTY_MARKUP.name, find: "VALUE", replace: "UPDATED", comment: "Review.", author: "Counsel" }, context)
  await asked
  const parts = unzipSync(source)
  const xml = strFromU8(parts["word/document.xml"]!)
  const field = '</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF Target </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>VALUE</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">'
  const altered = zipSync({ ...parts, "word/document.xml": strToU8(xml.replace("VALUE", field)) })
  writeFileSync(document, altered)
  release()
  expect(await request).toContain("Word field")
  expect(pendingRedlinesForDoc(root, document)).toEqual(before)
  expect(pendingRedlinesForDoc(root, document)[0]!.id).toBe(existing.id)
  expect(new Uint8Array(readFileSync(document))).toEqual(altered)
})
