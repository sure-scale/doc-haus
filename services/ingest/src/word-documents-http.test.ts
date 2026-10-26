import { afterAll, beforeAll, expect, test } from "bun:test"
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, InsertedTextRun, Packer, Paragraph } from "docx"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { openDb } from "./db"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"
import { linkedWordDocx } from "../../../dochaus/lib/word-linked-docx"
import { WORD_LINK_SETTING, type WordDocumentLink, type WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"

const paneOrigin = "https://localhost:47612"
const browserOrigin = "http://localhost:5190"
let temporary = "", base = "", matter = { id: "", dir: "" }
let service: ReturnType<typeof Bun.spawn> | undefined
let engine: ReturnType<typeof Bun.serve> | undefined
const request = (route: string, options?: RequestInit) => fetch(base + route, { ...options, signal: AbortSignal.timeout(60_000) })
const route = (suffix: string, options?: RequestInit) => request(`/matters/${matter.id}${suffix}`, options)
const query = (name: string) => new URLSearchParams({ name }).toString()
const plain = (text: string, actor = "Synthetic reviewer") => Packer.toBuffer(new Document({ lastModifiedBy: actor, sections: [{ children: [new Paragraph(text)] }] }))
function file(bytes: Uint8Array, name: string) {
  const form = new FormData()
  form.append("file", new File([new Uint8Array(bytes)], name))
  return form
}
async function create(name: string, text = "Neutral initial text.", source = false) {
  const response = await route(source ? "/documents" : "/deliverables", { method: "POST", body: file(await plain(text), name) })
  expect(response.status).toBe(200)
}
function readLink(bytes: Uint8Array) {
  for (const [name, part] of Object.entries(unzipSync(bytes))) if (name.includes("webextension") && name.endsWith(".xml")) {
    const doc = new DOMParser().parseFromString(strFromU8(part), "application/xml")
    const node = Array.from(doc.getElementsByTagNameNS("http://schemas.microsoft.com/office/webextensions/webextension/2010/11", "property"))
      .find(node => node.getAttribute("name") === WORD_LINK_SETTING)
    if (node) return JSON.parse(node.getAttribute("value")!) as WordDocumentLink
  }
  throw new Error("No Word link")
}
async function handout(name: string) {
  const response = await route(`/documents/word?${query(name)}`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ origin: browserOrigin }) })
  expect(response.status).toBe(200)
  const bytes = new Uint8Array(await response.arrayBuffer())
  return { bytes, link: readLink(bytes) }
}
const tokenHeaders = (link: WordDocumentLink, extra: Record<string, string> = {}) => ({ Origin: paneOrigin, Authorization: `Bearer ${link.token}`, ...extra })
const save = (link: WordDocumentLink, bytes: Uint8Array, baseVersion = link.baseVersion, extra: Record<string, string> = {}) => request("/word/saves", {
  method: "POST", headers: tokenHeaders(link, { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "x-dochaus-word-base-version": baseVersion, ...extra }), body: new Uint8Array(bytes),
})
const ping = (link: WordDocumentLink, session: string, done = false) => request(done ? "/word/done" : "/word/heartbeat", {
  method: "POST", headers: tokenHeaders(link, { "x-dochaus-word-session": session }),
})
const status = async () => (await (await route("/documents/word")).json() as { documents: { name: string; open: boolean; currentVersion: string; overlap: boolean }[] }).documents
const history = async (name: string) => await (await route(`/documents/versions?${query(name)}`)).json() as {
  document: { id: string }; versions: { id: string; number: number; origin: string; parent: string | null; actor: string; removed: string | null }[]
}

async function start() {
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const log = path.join(temporary, "ingest.log")
  service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, DOCHAUS_DATA_ROOT: temporary, DOCHAUS_RUNTIME_DIR: path.join(temporary, "runtime"), WORKSPACE_ROOT: path.join(temporary, "workspace"), DOCHAUS_DIR: path.join(temporary, "config"), DOCHAUS_WORD_ORIGIN: paneOrigin,
      DOCHAUS_WORD_EXTERNAL_HTTPS: "1", INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine!.port}` },
    stdout: Bun.file(log), stderr: Bun.file(log),
  })
  base = `http://127.0.0.1:${port}`
  const deadline = Date.now() + 20_000
  while (!await request("/matters").then(response => response.ok).catch(() => false)) {
    if (Date.now() > deadline) throw new Error(`Ingest startup failed: ${await readFile(log, "utf8")}`)
    await Bun.sleep(50)
  }
}
beforeAll(async () => {
  temporary = await mkdtemp(path.join(tmpdir(), "dochaus-word-http-"))
  await mkdir(path.join(temporary, "workspace"))
  await mkdir(path.join(temporary, "config"))
  await writeFile(path.join(temporary, "config", "opencode.json"), "{}\n")
  engine = Bun.serve({ port: 0, fetch: () => new Response("No engine required", { status: 503 }) })
  await start()
  const response = await request("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Synthetic Word matter" }) })
  matter = await response.json() as typeof matter
}, 30_000)
afterAll(async () => {
  service?.kill(); await service?.exited
  engine?.stop(true)
  if (temporary) await rm(temporary, { recursive: true, force: true })
})

test("handout links the current exact document; hash-only grants survive restart and are scoped to that identity", async () => {
  await create("Identity.docx")
  await create("Other.docx")
  const { bytes, link } = await handout("Identity.docx")
  expect(link).toMatchObject({ version: 1, origin: browserOrigin, apiOrigin: paneOrigin, matter: matter.id, documentName: "Identity.docx" })
  const original = unzipSync(new Uint8Array(await readFile(path.join(matter.dir, "Identity.docx"))))
  expect(unzipSync(bytes)["word/document.xml"]).toEqual(original["word/document.xml"])
  const db = openDb(matter.dir)
  try {
    const grants = db.query("SELECT token_hash, document_id FROM word_grants").all() as { token_hash: string; document_id: string }[]
    expect(grants.some(row => row.token_hash === link.token)).toBe(false)
    expect(grants.find(row => row.document_id === link.document)?.token_hash).toMatch(/^[a-f0-9]{64}$/)
  } finally { db.close() }
  expect((await request(`/matters/${matter.id}/documents/content?${query("Identity.docx")}`, { headers: tokenHeaders(link) })).status).toBe(403)
  expect((await route(`/documents?${query("Identity.docx")}`, { method: "DELETE", headers: tokenHeaders(link) })).status).toBe(403)
  const other = await handout("Other.docx")
  const invalidParent = await save(link, await plain("Wrong parent attempt."), other.link.baseVersion)
  expect([invalidParent.status, (await invalidParent.json() as { code: string }).code]).toEqual([409, "word_base_invalid"])
  expect((await ping(link, "session-one-0001")).status).toBe(200)
  service?.kill(); await service?.exited
  await start()
  expect((await ping(link, "session-one-0001")).status).toBe(200)
  const second = await handout("Identity.docx")
  expect((await route(`/documents/word?${query("Identity.docx")}`, { method: "DELETE" })).status).toBe(200)
  for (const token of [link, second.link]) expect((await ping(token, "session-one-0001")).status).toBe(401)
  expect((await status()).some(document => document.name === "Identity.docx")).toBe(false)
})

test("token endpoints require the exact configured pane origin on success, errors and preflight", async () => {
  const { link } = await handout("Other.docx")
  for (const origin of [undefined, "http://localhost:4612", "https://localhost:47613", "https://unrelated.invalid"]) {
    const response = await request("/word/heartbeat", { method: "POST", headers: { Authorization: `Bearer ${link.token}`, ...(origin ? { Origin: origin } : {}), "x-dochaus-word-session": "session-origin-001" } })
    expect(response.status).toBe(403)
    expect(response.headers.get("Access-Control-Allow-Origin")).toBeNull()
  }
  const preflight = await request("/word/saves", { method: "OPTIONS", headers: { Origin: paneOrigin, "Access-Control-Request-Method": "POST", "Access-Control-Request-Headers": "authorization,x-dochaus-word-base-version" } })
  expect(preflight.status).toBe(204)
  expect(preflight.headers.get("Access-Control-Allow-Origin")).toBe(paneOrigin)
  const invalid = await request("/word/saves", { method: "POST", headers: { Origin: paneOrigin, Authorization: "Bearer invalid" } })
  expect(invalid.status).toBe(401)
  expect(invalid.headers.get("Access-Control-Allow-Origin")).toBe(paneOrigin)
  const malformed = await save(link, new Uint8Array([1, 2, 3]))
  expect([malformed.status, (await malformed.json() as { code: string }).code]).toEqual([400, "invalid_docx"])
  expect(malformed.headers.get("Access-Control-Allow-Origin")).toBe(paneOrigin)
})

test("Word branches keep both versions, caller parents, actor and heartbeat overlap; retries do not revert later work", async () => {
  await create("Branches.docx")
  const { link } = await handout("Branches.docx")
  await ping(link, "session-first-0001")
  await ping(link, "session-second-0002")
  expect((await status()).find(document => document.name === "Branches.docx")).toMatchObject({ open: true, overlap: true })
  await ping(link, "session-second-0002", true)
  const firstBytes = linkedWordDocx(await plain("First Word edit.", "Synthetic lawyer one"), link)
  const first = await save(link, firstBytes)
  expect(first.status).toBe(200)
  const firstSave = await first.json() as WordSaveResponse
  expect(firstSave).toMatchObject({ version: { number: 2, parent: link.baseVersion, created: true }, overlap: false })
  expect(await (await save(link, firstBytes)).json()).toMatchObject({ version: { id: firstSave.version.id, created: false }, overlap: false })
  expect(readWorkingDocx(matter.dir, "Branches.docx")).toBeDefined()
  const second = await save(link, await plain("Parallel Word edit.", "Synthetic lawyer two"))
  expect(second.status).toBe(200)
  const secondSave = await second.json() as WordSaveResponse
  expect(secondSave).toMatchObject({ version: { number: 3, parent: link.baseVersion, created: true }, overlap: true })
  const retried = await save(link, firstBytes)
  expect(await retried.json()).toMatchObject({ version: { id: firstSave.version.id, created: false }, overlap: true })
  const versions = await history("Branches.docx")
  expect(versions.versions.map(row => [row.number, row.origin, row.actor, row.parent])).toEqual([
    [3, "word", "Synthetic lawyer two", link.baseVersion], [2, "word", "Synthetic lawyer one", link.baseVersion], [1, "generated", "unknown", null],
  ])
  const current = new Uint8Array(await (await route(`/documents/content?${query("Branches.docx")}`)).arrayBuffer())
  expect(strFromU8(unzipSync(current)["word/document.xml"])).toContain("Parallel Word edit.")
  expect(Object.values(unzipSync(current)).map(bytes => strFromU8(bytes)).join("\n")).not.toContain(link.token)
  const row = (await status()).find(document => document.name === "Branches.docx")
  expect(row).toMatchObject({ currentVersion: secondSave.version.id, overlap: true })
  await ping(link, "session-first-0001", true)
  expect((await status()).find(document => document.name === "Branches.docx")?.open).toBe(false)
})

test("deleting and recreating the filename does not revive a token; stale copies cannot restore redacted text", async () => {
  await create("Lifecycle.docx", "Original restricted phrase.")
  const old = await handout("Lifecycle.docx")
  expect((await route(`/documents?${query("Lifecycle.docx")}`, { method: "DELETE" })).status).toBe(200)
  await create("Lifecycle.docx", "New document identity.")
  const current = await handout("Lifecycle.docx")
  expect(current.link.document).not.toBe(old.link.document)
  expect((await save(old.link, await plain("Attempt from a deleted document."))).status).toBe(401)
  const redacted = await plain("[REDACTED]")
  const redaction = await route("/documents/revisions", { method: "POST", headers: { "x-dochaus-origin": "redaction" }, body: file(redacted, "Lifecycle.docx") })
  expect(redaction.status).toBe(200)
  const stale = await save(current.link, await plain("New document identity."))
  expect(stale.status).toBe(401)
  const fresh = await handout("Lifecycle.docx")
  expect((await save(fresh.link, await plain("Preserves [REDACTED], with a neutral addition."))).status).toBe(200)
  expect((await history("Lifecycle.docx")).versions.some(row => row.removed === "redaction")).toBe(true)
})

test("unchanged uploads retain and persist overlap from simultaneous sessions and a stale base", async () => {
  await create("Unchanged.docx")
  const { link } = await handout("Unchanged.docx")
  await ping(link, "unchanged-session-first")
  await ping(link, "unchanged-session-second")
  const bytes = new Uint8Array(await readFile(path.join(matter.dir, "Unchanged.docx")))
  const noChange = await save(link, bytes)
  expect(await noChange.json()).toMatchObject({ version: { id: link.baseVersion, created: false }, overlap: true })
  await ping(link, "unchanged-session-first", true)
  await ping(link, "unchanged-session-second", true)
  expect((await status()).find(document => document.name === "Unchanged.docx")).toMatchObject({ open: false, overlap: true })
  const revision = await route("/documents/revisions", { method: "POST", body: file(await plain("An ordinary newer revision."), "Unchanged.docx") })
  expect(revision.status).toBe(200)
  const id = (await revision.json() as { version: { id: string } }).version.id
  const stale = await save(link, new Uint8Array(await readFile(path.join(matter.dir, "Unchanged.docx"))))
  expect(await stale.json()).toMatchObject({ version: { id, created: false }, overlap: true })
  expect((await history("Unchanged.docx")).versions).toHaveLength(2)
})

test("a source save reindexes canonical text and ordinary current/history/redlined downloads strip linkage", async () => {
  await create("Source.docx", "The synthetic source begins here.", true)
  const { link } = await handout("Source.docx")
  const saved = await save(link, linkedWordDocx(await plain("The revised synthetic source is now searchable."), link))
  expect(saved.status).toBe(200)
  const body = await saved.json() as WordSaveResponse
  const db = openDb(matter.dir)
  try {
    const rows = db.query("SELECT text FROM chunks WHERE doc_name = ?").all("Source.docx") as { text: string }[]
    expect(rows.map(row => row.text).join("\n")).toContain("revised synthetic source")
    expect(rows.map(row => row.text).join("\n")).not.toContain("source begins here")
    expect(db.query("SELECT role FROM documents WHERE name = ?").get("Source.docx")).toEqual({ role: "source" })
  } finally { db.close() }
  for (const url of [`/documents/content?${query("Source.docx")}`, `/documents/redlined?${query("Source.docx")}`, `/documents/versions/${body.version.id}/content`]) {
    const response = await route(url)
    expect(response.status).toBe(200)
    const parts = unzipSync(new Uint8Array(await response.arrayBuffer()))
    expect(Object.values(parts).map(bytes => strFromU8(bytes)).join("\n")).not.toContain(link.token)
    expect(Object.keys(parts).some(name => name.includes("webextensions/dochaus"))).toBe(false)
  }
}, 90_000)

test("tracked Word saves are kept while preserving the existing guard for later document mutation", async () => {
  await create("Tracked.docx")
  const { link } = await handout("Tracked.docx")
  const clean = await save(link, await plain("Clean accepted Word state."))
  expect(clean.status).toBe(200)
  const baseVersion = (await clean.json() as WordSaveResponse).version.id
  const accepted = readWorkingDocx(matter.dir, "Tracked.docx")
  const tracked = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph({ children: [new InsertedTextRun({ text: "Pending tracked insertion.", id: 1, author: "Synthetic reviewer", date: "2026-10-26T10:00:00Z" })] })] }] }))
  const response = await save(link, tracked, baseVersion)
  expect(response.status).toBe(200)
  expect(readWorkingDocx(matter.dir, "Tracked.docx")).toEqual(accepted)
  const guarded = await route("/documents/revisions", { method: "POST", body: file(await plain("Unbound replacement."), "Tracked.docx") })
  expect(guarded.status).toBe(409)
})

test("plain DOCX responses strip linkage from externally uploaded files while other archive downloads remain intact", async () => {
  const synthetic: WordDocumentLink = { version: 1, origin: browserOrigin, apiOrigin: paneOrigin, matter: matter.id, matterName: "Synthetic matter",
    document: "external-synthetic-id", documentName: "External.docx", baseVersion: "external-synthetic-base", token: "external-synthetic-capability" }
  const bytes = linkedWordDocx(await plain("An externally uploaded Word copy."), synthetic)
  const upload = await route("/deliverables", { method: "POST", body: file(bytes, "External.docx") })
  expect(upload.status).toBe(200)
  const captured = (await upload.json() as { version: { id: string } }).version
  // The original really carries the linkage; these are causal response-strip checks.
  expect(Object.values(unzipSync(new Uint8Array(await readFile(path.join(matter.dir, "External.docx"))))).map(bytes => strFromU8(bytes)).join("\n")).toContain(synthetic.token)
  for (const url of [`/documents/content?${query("External.docx")}`, `/documents/redlined?${query("External.docx")}`, `/documents/versions/${captured.id}/content`]) {
    const response = await route(url)
    expect(response.status).toBe(200)
    expect(Object.values(unzipSync(new Uint8Array(await response.arrayBuffer()))).map(bytes => strFromU8(bytes)).join("\n")).not.toContain(synthetic.token)
  }
  const archive = zipSync({ "sample.txt": strToU8("Neutral archive contents.") })
  expect((await route("/documents", { method: "POST", body: file(archive, "Sample.zip") })).status).toBe(200)
  const downloaded = await route(`/documents/content?${query("Sample.zip")}`)
  expect(downloaded.status).toBe(200)
  expect(new Uint8Array(await downloaded.arrayBuffer())).toEqual(archive)
})
