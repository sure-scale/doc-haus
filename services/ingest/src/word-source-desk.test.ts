import { afterAll, beforeAll, expect, test } from "bun:test"
import { Hono } from "hono"
import { createHash } from "node:crypto"
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import type { DeskComparison, DeskSourcePreview, DeskSources, DeskVersions } from "../../../dochaus/lib/word-source-desk-contract"
import { openDb } from "./db"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { readCurrentVersion } from "./document-versions"
import { wordCors, wordDocumentRoutes, wordTokenPath } from "./word-documents"
import { wordHttpsHandler } from "./word-https"
import { installWordSourceDesk } from "./word-source-desk"
import { serializeMatter } from "./matter-session"
import { claimWordEditLease, clearWordEditLease } from "../../../dochaus/lib/word-edit-lease"
import type { WordGrant } from "./word-grants"

const origin = "https://localhost:47612"
const keys = ["DOCHAUS_WORD_ORIGIN", "DOCHAUS_WORD_EXTERNAL_HTTPS", "DOCHAUS_RUNTIME_DIR"] as const
const previous = keys.map(key => process.env[key])
let directory: string
let app: Hono
const token = "a".repeat(43)
const digest = createHash("sha256").update(token).digest("hex")
const bytes = (text: string) => Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] }))
const request = (action: string, body: object = {}, headers: Record<string, string> = {}) => app.request(`/word/desk/${action}`, {
  method: "POST", headers: { Origin: origin, Authorization: `Bearer ${token}`, "x-dochaus-word-version": WORD_ADDIN_VERSION, "content-type": "application/json", ...headers },
  body: JSON.stringify(body),
})
const uploadHeaders = { Origin: origin, Authorization: `Bearer ${token}`, "x-dochaus-word-version": WORD_ADDIN_VERSION }
const uploadForm = (name: string, bytes: Uint8Array, extra?: [string, string]) => {
  const form = new FormData()
  form.set("file", new File([new Uint8Array(bytes)], name))
  if (extra) form.set(...extra)
  return form
}

beforeAll(async () => {
  directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-source-desk-")))
  process.env.DOCHAUS_WORD_ORIGIN = origin
  process.env.DOCHAUS_WORD_EXTERNAL_HTTPS = "1"
  process.env.DOCHAUS_RUNTIME_DIR = path.join(directory, "runtime")
  await writeDeliverable(directory, "Draft.docx", await bytes("The delivery date is Monday."))
  await writeDeliverable(directory, "Source.docx", await bytes("Neutral source confirms Monday delivery."))
  await writeDeliverable(directory, "Other.docx", await bytes("A different saved document."))
  const kept = readCurrentVersion(directory, "Draft.docx")!
  const db = openDb(directory)
  try {
    db.run("UPDATE documents SET role='source' WHERE name='Source.docx'")
    db.run("INSERT INTO word_grants (token_hash,document_id,origin,issued_version,created_at) VALUES (?,?,?,?,?)", [digest, kept.document.id, origin, kept.current.id, Date.now()])
  } finally { db.close() }
  const matter = { id: "synthetic-matter", title: "Synthetic Desk matter", dir: directory }
  app = new Hono()
  app.use("/word/*", wordCors)
  app.route("/", wordDocumentRoutes({ matter: () => matter, matters: () => [matter], locks: () => [], write: async () => { throw new Error("Desk reads must never write") } }))
}, 30_000)

afterAll(() => {
  keys.forEach((key, index) => { if (previous[index] === undefined) delete process.env[key]; else process.env[key] = previous[index] })
  if (directory) rmSync(directory, { recursive: true, force: true })
})

test("scoped source previews pin exact saved bytes; changed sources fail without changing ownership or saves", async () => {
  const initialBytes = readFileSync(path.join(directory, "Draft.docx"))
  const listing = await request("sources")
  expect(listing.status).toBe(200)
  const sources = await listing.json() as DeskSources
  expect(sources).toMatchObject({ matterName: "Synthetic Desk matter", documentName: "Draft.docx", truncated: false })
  const selected = sources.sources.find(row => row.name === "Source.docx")!
  expect(selected).toMatchObject({ role: "source", version: { number: 1 } })
  expect(selected.version?.createdAt).toBe(readCurrentVersion(directory, "Source.docx")!.current.created_at)
  const preview = await request("source", { sourceId: selected.id, expectedHash: selected.hash })
  expect(preview.status).toBe(200)
  expect(await preview.json() as DeskSourcePreview).toMatchObject({ ...selected, text: expect.stringContaining("Neutral source confirms Monday delivery."), blocks: [{ kind: "paragraph", text: "Neutral source confirms Monday delivery." }], provenance: "saved-version", truncated: false })
  expect((await request("source", { sourceId: "../../other", expectedHash: selected.hash })).status).toBe(400)
  expect((await request("source", { sourceId: "9999999999", expectedHash: selected.hash })).status).toBe(404)
  writeFileSync(path.join(directory, "Source.docx"), await bytes("The neutral source now confirms Tuesday."))
  const changed = await request("source", { sourceId: selected.id, expectedHash: selected.hash })
  expect(changed.status).toBe(409)
  expect(await changed.json()).toMatchObject({ code: "desk_source_changed" })
  // A file changed outside doc.haus has no confirmed current snapshot to show.
  const after = await (await request("sources")).json() as DeskSources
  expect(after.truncated).toBe(true)
  expect(after.sources.some(row => row.name === "Source.docx")).toBe(false)
  const db = openDb(directory)
  try {
    expect(db.query("SELECT COUNT(*) AS n FROM word_sessions").get()).toEqual({ n: 0 })
    expect(db.query("SELECT COUNT(*) AS n FROM word_edit_leases").get()).toEqual({ n: 0 })
  } finally { db.close() }
  expect(readFileSync(path.join(directory, "Draft.docx"))).toEqual(initialBytes)
}, 30_000)

test("linked Documents uploads keep both beside an actively edited document without a chat session", async () => {
  const before = readCurrentVersion(directory, "Draft.docx")!
  const original = readFileSync(path.join(directory, "Draft.docx"))
  const incoming = await bytes("Incoming neutral evidence.")
  const db = openDb(directory)
  claimWordEditLease(db, { document: before.document.id, tokenHash: digest, session: "synthetic-upload-pane" })
  try {
    for (const expected of ["Draft (2).docx", "Draft (3).docx"]) {
      const response = await app.request("/word/desk/upload", { method: "POST", headers: uploadHeaders, body: uploadForm("Draft.docx", incoming) })
      expect(response.status).toBe(200)
      expect(response.headers.get("cache-control")).toBe("no-store")
      expect(await response.json()).toEqual({ name: expected })
      expect(readFileSync(path.join(directory, expected))).toEqual(Buffer.from(incoming))
      expect(db.query("SELECT role FROM documents WHERE name=?").get(expected)).toEqual({ role: "source" })
      expect(readCurrentVersion(directory, expected)?.current.number).toBe(1)
    }
    expect(readFileSync(path.join(directory, "Draft.docx"))).toEqual(original)
    expect(readCurrentVersion(directory, "Draft.docx")?.current.id).toBe(before.current.id)
    expect(db.query("SELECT COUNT(*) AS n FROM word_sessions").get()).toEqual({ n: 0 })
    expect(db.query("SELECT session FROM word_edit_leases").get()).toEqual({ session: "synthetic-upload-pane" })
    const listing = await (await request("sources")).json() as DeskSources
    expect(listing.sources.map(source => source.name)).toContain("Draft (2).docx")
    expect(listing.sources.map(source => source.name)).toContain("Draft (3).docx")
  } finally { clearWordEditLease(db, before.document.id); db.close() }
}, 60_000)

test("Documents upload rejects extra scope/options, invalid multipart and unreadable documents", async () => {
  for (const body of [new FormData(), uploadForm("Empty.txt", new Uint8Array())])
    expect((await app.request("/word/desk/upload", { method: "POST", headers: uploadHeaders, body })).status).toBe(400)
  for (const field of ["matter", "path", "mode", "conversationId"]) {
    const response = await app.request("/word/desk/upload", { method: "POST", headers: uploadHeaders, body: uploadForm("Rejected.txt", Buffer.from("Neutral evidence."), [field, "untrusted"]) })
    expect(response.status).toBe(400)
  }
  const duplicate = uploadForm("Rejected.txt", Buffer.from("Neutral evidence."))
  duplicate.append("file", new File(["More evidence"], "Other.txt"))
  expect((await app.request("/word/desk/upload", { method: "POST", headers: uploadHeaders, body: duplicate })).status).toBe(400)
  expect((await app.request("/word/desk/upload", { method: "POST", headers: uploadHeaders, body: uploadForm("Rejected.exe", Buffer.from("Not supported")) })).status).toBe(400)
  expect((await app.request("/word/desk/upload", { method: "POST", headers: { ...uploadHeaders, "content-type": "multipart/form-data; boundary=bad" }, body: "malformed" })).status).toBe(400)
  const bad = await app.request("/word/desk/upload", { method: "POST", headers: uploadHeaders, body: uploadForm("Unreadable.docx", Buffer.from("Not a Word document")) })
  expect(bad.status).toBe(422)
  expect(await bad.json()).toMatchObject({ code: "desk_upload_failed", error: expect.stringContaining("could not be read") })
  expect((await app.request("/word/desk/upload", { method: "POST", headers: { ...uploadHeaders, Authorization: `Bearer ${"z".repeat(43)}` }, body: uploadForm("Rejected.txt", Buffer.from("Neutral evidence.")) })).status).toBe(401)
  expect(existsSync(path.join(directory, "Rejected.txt"))).toBe(false)
  expect(existsSync(path.join(directory, "Unreadable.docx"))).toBe(false)
})

test("Documents upload revalidates a revoked grant after waiting in the matter queue", async () => {
  const current = readCurrentVersion(directory, "Draft.docx")!
  const db = openDb(directory)
  const grant = db.query("SELECT * FROM word_grants WHERE token_hash=?").get(digest) as WordGrant
  db.close()
  let revoked = false, calls = 0, finished = false
  const queued = new Hono()
  installWordSourceDesk(queued, () => { calls++; return revoked ? undefined : { matter: { id: "synthetic-matter", title: "Synthetic", dir: directory }, grant, target: { name: "Draft.docx", docPath: path.join(directory, "Draft.docx") } } })
  let release!: () => void, entered!: () => void
  const held = new Promise<void>(resolve => { release = resolve })
  const started = new Promise<void>(resolve => { entered = resolve })
  const blocker = serializeMatter(directory, async () => { entered(); await held })
  await started
  try {
    const response = Promise.resolve(queued.request("/word/desk/upload", { method: "POST", body: uploadForm("Revoked.txt", Buffer.from("Neutral evidence.")) })).then(value => { finished = true; return value })
    await Bun.sleep(0)
    expect(calls).toBe(1)
    expect(finished).toBe(false)
    revoked = true; release()
    expect((await response).status).toBe(401)
    expect(calls).toBe(2)
    expect(existsSync(path.join(directory, "Revoked.txt"))).toBe(false)
    expect(readCurrentVersion(directory, "Draft.docx")?.current.id).toBe(current.current.id)
  } finally { release(); await blocker }
})

test("Documents upload bounds chunked multipart before parsing and enforces its file limit", async () => {
  let cancelled = false
  const stream = new ReadableStream<Uint8Array>({ start(controller) { controller.enqueue(new Uint8Array(17 * 1024 * 1024 + 1)) }, cancel() { cancelled = true } })
  const oversized = await app.request("/word/desk/upload", { method: "POST", headers: { ...uploadHeaders, "content-type": "multipart/form-data; boundary=synthetic" }, body: stream, duplex: "half" } as RequestInit)
  expect(oversized.status).toBe(413)
  expect(cancelled).toBe(true)
  const file = await app.request("/word/desk/upload", { method: "POST", headers: uploadHeaders, body: uploadForm("Large.txt", new Uint8Array(16 * 1024 * 1024 + 1)) })
  expect(file.status).toBe(413)
  expect(existsSync(path.join(directory, "Large.txt"))).toBe(false)
})

test("retained-version comparison stays within the linked document and refuses redaction, stale clients and revoked grants", async () => {
  const base = readCurrentVersion(directory, "Draft.docx")!
  await writeDeliverableRevision(directory, "Draft.docx", await bytes("The delivery date is Tuesday."))
  const current = readCurrentVersion(directory, "Draft.docx")!
  const history = await (await request("versions")).json() as DeskVersions
  expect(history.matterName).toBe("Synthetic Desk matter")
  expect(history.versions.find(row => row.id === current.current.id)?.bytes).toBe(current.current.bytes)
  expect(history.document).toEqual({ id: base.document.id, name: "Draft.docx" })
  expect(history.versions.map(row => row.id)).toEqual([current.current.id, base.current.id])
  const response = await request("compare", { baseVersion: base.current.id, currentVersion: current.current.id })
  expect(response.status).toBe(200)
  const compared = await response.json() as DeskComparison
  expect(compared).toMatchObject({ context: "saved-versions", base: { number: 1 }, current: { number: 2 }, truncated: false })
  expect(compared.changes.length).toBeGreaterThan(0)
  const other = readCurrentVersion(directory, "Other.docx")!
  expect((await request("compare", { baseVersion: base.current.id, currentVersion: other.current.id })).status).toBe(404)
  const db = openDb(directory)
  try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [current.current.id]) } finally { db.close() }
  expect((await request("compare", { baseVersion: base.current.id, currentVersion: current.current.id })).status).toBe(410)
  expect((await request("versions", {}, { Origin: "https://unrelated.invalid" })).status).toBe(403)
  expect((await request("versions", {}, { "x-dochaus-word-version": "0.0.0.0" })).status).toBe(426)
  const revoked = openDb(directory)
  try { revoked.run("UPDATE word_grants SET revoked_at=? WHERE token_hash=?", [Date.now(), digest]) } finally { revoked.close() }
  for (const action of ["sources", "source", "versions", "compare", "open", "upload"]) expect((await request(action)).status).toBe(401)
}, 30_000)

test("HTTPS bridge only forwards the exact scoped desk command paths", async () => {
  const forwarded: string[] = []
  const handler = wordHttpsHandler(request => { forwarded.push(new URL(request.url).pathname); return new Response("scoped") },
    { origin, external: true, host: "127.0.0.1", port: 47612 })
  for (const action of ["sources", "source", "versions", "compare", "open", "upload"]) {
    const route = `/word/desk/${action}`
    expect(wordTokenPath(route)).toBe(true)
    expect((await handler(new Request(origin + route, { method: "POST", body: "{}" }))).status).toBe(200)
    expect((await handler(new Request(origin + route))).status).toBe(405)
  }
  for (const route of ["/word/desk/source/extra", "/word/desk/upload/extra", "/word/desk/other", "/matters/synthetic/documents/content"]) {
    expect(wordTokenPath(route)).toBe(false)
    expect((await handler(new Request(origin + route, { method: "POST" }))).status).toBe(404)
  }
  expect(forwarded).toHaveLength(6)
})
