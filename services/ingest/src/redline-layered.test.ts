import { afterEach, describe, expect, setSystemTime, test } from "bun:test"
import { createHash } from "node:crypto"
import { realpathSync } from "node:fs"
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, unzipSync, zipSync } from "fflate"
import { createDocxLocalization, openRevision, verifyPreservation } from "../../../dochaus/lib/docx-amendments"
import { readDocxComments } from "../../../dochaus/lib/docx-comments"
import { assertLayeredRoundTrip } from "../../../dochaus/lib/redline-round-trip"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../../../dochaus/test/support/counterparty-markup-fixture"
import { openDb, type RedlineRow } from "./db"
import { docxodus } from "./docxodus"
import { bake, buildRedlined } from "./redline"

// #200 scope 4: on an uploaded counterparty markup, doc.haus's proposals are
// written into their file as native tracked changes beside theirs, and Accept
// bakes them in untracked; their ten revisions and four comments stay pending.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const layer = { pendingRevisions: "layer" } as const
const AUTHOR = "Jordan Example"
const NOTE = "Our client's governing law."
const LAW = "9. Governing law. This Agreement is governed by the law of the State."
const sha256 = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")

const revisions = async (bytes: Uint8Array) => {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try {
    return session.listRevisions().map(({ id, type, author, text, date, anchorId }) => ({ id, type, author, text, date, anchorId }))
  } finally { session.close() }
}
// Clause 9's accepted text, the projection Prepare and Accept read.
const clause9 = async (bytes: Uint8Array) => {
  const opened = await openRevision(bytes, layer)
  try { return opened.revision.paragraphs[9]!.text } finally { opened.session.close() }
}
const xml = (bytes: Uint8Array, part: string) => new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[part]!), "application/xml")
const commentAuthors = (bytes: Uint8Array) => Array.from(xml(bytes, "word/comments.xml").getElementsByTagNameNS(W, "comment")).map(node => node.getAttributeNS(W, "author"))

// One proposal as Prepare stores it on an uploaded source: a whole-revision
// (version 1) binding to the clause 9 quote, under the attorney's name.
async function proposal(source: Uint8Array, find_text: string, new_text: string, id = 1, review_comment = NOTE): Promise<RedlineRow> {
  const packet = await createDocxLocalization(source, [{ id: String(id), before_quote: find_text, replacement: new_text }], {}, layer)
  const [candidate] = packet.candidates.filter(row => row.paragraph.scope === "body")
  return { id, doc_path: "supply-agreement-markup.docx", doc_name: "supply-agreement-markup.docx", scope: "phrase", find_text, old_text: find_text, new_text,
    author: AUTHOR, review_comment, anchor_id: null, created_at: 1, status: "pending",
    target_json: JSON.stringify({ version: 1, source_sha256: packet.revision.source_sha256, target_handle: candidate!.target_handle, amendment_id: String(id) }) }
}

describe("layered redline on an uploaded markup", () => {
  const source = counterpartyMarkupDocx()
  afterEach(() => setSystemTime())

  test("the export keeps their ten revisions and four comments and adds doc.haus's as its own", async () => {
    const row = await proposal(source, "the Province", "the State")
    const exported = await buildRedlined(source, [row], undefined, layer)
    const theirs = await revisions(source)
    expect(theirs).toHaveLength(10)
    const after = await revisions(exported)
    expect(after.filter(revision => revision.author !== AUTHOR)).toEqual(theirs)
    // One deletion and one insertion of whole words, in clause 9, dated when exported.
    const ours = after.filter(revision => revision.author === AUTHOR)
    expect(ours.map(revision => [revision.type, revision.text])).toEqual([["delete", "Province"], ["insert", "State"]])
    expect(new Set(ours.map(revision => revision.anchorId)).size).toBe(1)
    const ids = ["ins", "del", "moveFrom", "moveTo", "rPrChange"].flatMap(name => Array.from(xml(exported, "word/document.xml").getElementsByTagNameNS(W, name)))
      .map(node => node.getAttributeNS(W, "id"))
    expect(new Set(ids).size).toBe(ids.length)
    // Track Changes stays on, so the other side's next edits are tracked.
    expect(strFromU8(unzipSync(exported)["word/settings.xml"]!)).toContain("<w:trackRevisions")

    // Their four comments, then doc.haus's on clause 9.
    expect(commentAuthors(exported)).toEqual([...COUNTERPARTY_MARKUP.comments.map(comment => comment.author), AUTHOR])
    const comments = readDocxComments(exported)
    expect(comments.map(comment => comment.text)).toEqual([...COUNTERPARTY_MARKUP.comments.map(comment => comment.text), NOTE])
    expect(comments.at(-1)!.markers.map(marker => marker.accepted_paragraph_text)).toEqual([LAW, LAW, LAW])
    // With doc.haus's comment taken out of the comments part again, only
    // clause 9 and the Track Changes setting differ from what they sent: every
    // other part compares canonically, the reply and resolved state in
    // commentsExtended included.
    const parts = unzipSync(exported)
    parts["word/comments.xml"] = unzipSync(source)["word/comments.xml"]!
    const tracked = { ...layer, revisionAuthor: AUTHOR }
    const [before, sent] = [await openRevision(source, tracked), await openRevision(zipSync(parts), tracked)]
    try { expect(() => verifyPreservation(before, sent, new Set([9]))).not.toThrow() } finally { [before, sent].forEach(opened => opened.session.close()) }

    // The round trip both ways: rejecting doc.haus's changes gives their file,
    // accepting them gives what Accept stores; the round trip refuses otherwise.
    const baked = await bake(source, [row], undefined, layer)
    await assertLayeredRoundTrip(exported, source, baked)
    await expect(assertLayeredRoundTrip(exported, source, source)).rejects.toThrow("accepting only doc.haus's changes leaves text that differs at body paragraph 9")
    await expect(assertLayeredRoundTrip(baked, source, baked)).rejects.toThrow("rejecting only doc.haus's changes leaves text that differs at body paragraph 9")
  })

  test("Accept bakes the edit in untracked and leaves their ten revisions pending", async () => {
    const baked = await bake(source, [await proposal(source, "the Province", "the State")], undefined, layer)
    expect(await revisions(baked)).toEqual(await revisions(source))
    expect(await clause9(baked)).toBe(LAW)
  })

  test("a clean source exports exactly as before", async () => {
    const clean = (await renderMarkdownToDocx("Current review period is five days.\n\nThe Province governs.")).bytes
    const row = await proposal(clean, "five days", "seven days")
    setSystemTime(new Date("2026-10-25T12:00:00Z"))
    expect(sha256(await buildRedlined(clean, [row], undefined, layer))).toBe(sha256(await buildRedlined(clean, [row])))
  })
})

test("over HTTP an uploaded markup exports layered, and Accept keeps their changes listed", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-redline-layered-"))
  const config = path.join(temporary, "config")
  await mkdir(path.join(temporary, "workspace"))
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  const engine = Bun.serve({ port: 0, fetch: () => new Response("No engine in this redline test", { status: 503 }) })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const log = path.join(temporary, "ingest.log")
  const service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: path.join(temporary, "workspace"), DOCHAUS_DIR: config,
      INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}` },
    stdout: Bun.file(log), stderr: Bun.file(log),
  })
  const request = (route: string, options?: RequestInit) => fetch(`http://127.0.0.1:${port}${route}`, { ...options, signal: AbortSignal.timeout(60_000) })
  try {
    const deadline = Date.now() + 20_000
    while (!await request("/matters").then(response => response.ok).catch(() => false)) {
      if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(log, "utf8")}`)
      await Bun.sleep(50)
    }
    const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Layered markup matter" }) })
    const matter = await created.json() as { id: string; dir: string }
    const name = COUNTERPARTY_MARKUP.name
    const form = new FormData()
    form.append("file", new File([new Uint8Array(counterpartyMarkupDocx())], name))
    expect((await request(`/matters/${matter.id}/documents`, { method: "POST", body: form })).status).toBe(200)
    const file = realpathSync(path.join(matter.dir, name))
    const uploaded = await Bun.file(file).bytes()
    const row = await proposal(uploaded, "the Province", "the State")
    const db = openDb(matter.dir)
    try {
      row.id = Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,review_comment,anchor_id,target_json,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        [file, name, row.scope, row.find_text, row.old_text, row.new_text, row.author, row.review_comment, null, row.target_json!, Date.now()]).lastInsertRowid)
    } finally { db.close() }
    const listed = await (await request(`/matters/${matter.id}/redlines?name=${name}`)).json() as Array<{ id: number; stale?: true }>
    expect(listed.map(proposal => [proposal.id, proposal.stale])).toEqual([[row.id, undefined]])

    const redlined = await request(`/matters/${matter.id}/documents/redlined?name=${name}`)
    expect(redlined.status).toBe(200)
    // Nothing of theirs was accepted, so there is nothing to disclose.
    expect(redlined.headers.get("X-DocHaus-Revision-Notice")).toBeNull()
    const exported = new Uint8Array(await redlined.arrayBuffer())
    expect((await revisions(exported)).map(revision => revision.author)).toEqual([...(await revisions(uploaded)).map(revision => revision.author), AUTHOR, AUTHOR])
    expect(commentAuthors(exported)).toEqual([...COUNTERPARTY_MARKUP.comments.map(comment => comment.author), AUTHOR])

    expect((await request(`/matters/${matter.id}/redlines/${row.id}/accept`, { method: "POST" })).status).toBe(200)
    const accepted = await Bun.file(file).bytes()
    expect(await revisions(accepted)).toEqual(await revisions(uploaded))
    expect(await clause9(accepted)).toBe(LAW)
    const markup = await (await request(`/matters/${matter.id}/documents/markup?name=${name}`)).json() as { revisions: unknown[] }
    expect(markup.revisions).toHaveLength(10)
  } finally {
    service.kill()
    await service.exited
    engine.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 90_000)
