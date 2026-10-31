import { afterAll, beforeAll, expect, test } from "bun:test"
import { mkdtemp, mkdir, readFile, readdir, rm, stat, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, InsertedTextRun, Packer, Paragraph } from "docx"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { openDb } from "./db"
import type { DeskSources } from "../../../dochaus/lib/word-source-desk-contract"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"
import { linkedWordDocx, readWordLink } from "../../../dochaus/lib/word-linked-docx"
import { WORD_ADDIN_VERSION, WORD_LINK_SETTING, type WordDocumentLink, type WordSaveResponse, type WordConnectionState, type WordLeaseAction } from "../../../dochaus/lib/word-addin-contract"

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
const tokenHeaders = (link: WordDocumentLink, extra: Record<string, string> = {}) => ({ Origin: paneOrigin, Authorization: `Bearer ${link.token}`, "x-dochaus-word-version": WORD_ADDIN_VERSION, ...extra })
const owners = new Map<string, { session: string; generation: string }>()
async function lease(link: WordDocumentLink, action: WordLeaseAction, session = "session-test-0001", baseVersion = link.baseVersion,
  expected: { expectedCurrent?: string; expectedGeneration?: string } = {}) {
  const response = await request("/word/lease", { method: "POST", headers: tokenHeaders(link, { "content-type": "application/json",
    "x-dochaus-word-session": session, ...(owners.has(link.token) ? { "x-dochaus-word-lease": owners.get(link.token)!.generation } : {}) }),
    body: JSON.stringify({ action, baseVersion, ...expected }) })
  return response
}
async function claim(link: WordDocumentLink, session = "session-test-0001", baseVersion = link.baseVersion) {
  const response = await lease(link, "claim", session, baseVersion)
  expect(response.status).toBe(200)
  const state = await response.json() as WordConnectionState
  expect(state.lease.state).toBe("owned")
  owners.set(link.token, { session, generation: state.lease.generation! })
  return state
}
const save = (link: WordDocumentLink, bytes: Uint8Array, baseVersion = link.baseVersion, extra: Record<string, string> = {}) => request("/word/saves", {
  method: "POST", headers: tokenHeaders(link, { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "x-dochaus-word-base-version": baseVersion,
    "x-dochaus-word-session": owners.get(link.token)?.session ?? "session-test-0001", "x-dochaus-word-lease": owners.get(link.token)?.generation ?? "", ...extra }), body: new Uint8Array(bytes),
})
const ping = (link: WordDocumentLink, session: string, done = false, baseVersion = link.baseVersion, generation = owners.get(link.token)?.generation ?? "") => request(done ? "/word/done" : "/word/heartbeat", {
  method: "POST", headers: tokenHeaders(link, { "x-dochaus-word-session": session, "x-dochaus-word-lease": generation, "x-dochaus-word-base-version": baseVersion }),
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
  await claim(link, "session-one-0001")
  const invalidParent = await save(link, await plain("Wrong parent attempt."), other.link.baseVersion)
  expect([invalidParent.status, (await invalidParent.json() as { code: string }).code]).toEqual([409, "word_base_invalid"])
  expect((await ping(link, "session-one-0001")).status).toBe(200)
  service?.kill(); await service?.exited
  await start()
  expect((await ping(link, "session-one-0001")).status).toBe(200)
  const second = await handout("Identity.docx")
  expect(second.link.token === link.token).toBe(true)
  expect(Buffer.compare(second.bytes, bytes) === 0).toBe(true)
  expect((await route(`/documents/word?${query("Identity.docx")}`, { method: "DELETE" })).status).toBe(200)
  for (const token of [link, second.link]) expect((await ping(token, "session-one-0001")).status).toBe(401)
  expect((await status()).some(document => document.name === "Identity.docx")).toBe(false)
})

test("native handoff is one-use, preserves linkage and reuses its protected live grant across opens", async () => {
  await create("Handoff.docx")
  const issue = async (origin = browserOrigin) => {
    const response = await route(`/documents/word/open?${query("Handoff.docx")}`, { method: "POST",
      headers: { "Content-Type": "application/json" }, body: JSON.stringify({ origin }) })
    expect(response.status).toBe(200)
    const value = await response.json() as { uri: string; id: string }
    const uri = new URL(value.uri), source = new URL(uri.searchParams.get("source")!)
    expect(uri.protocol === "doc.haus:" && uri.host === "word" && uri.pathname === "/open").toBe(true)
    expect(source.origin === paneOrigin && /^\/word\/open\/[A-Za-z0-9_-]{43}$/.test(source.pathname)).toBe(true)
    expect(/^[a-f0-9-]{36}$/.test(value.id)).toBe(true)
    expect(response.headers.get("Cache-Control")).toBe("no-store")
    return Object.assign(source, { id: value.id })
  }
  const state = async (id: string) => (await (await route(`/documents/word/open/status?id=${id}`)).json() as { state: string }).state
  const result = (source: URL, value: object, origin?: string) => request(`${source.pathname}/result`, { method: "POST",
    headers: { "Content-Type": "application/json", ...(origin ? { Origin: origin } : {}) }, body: JSON.stringify(value) })
  const first = await issue()
  expect(await state(first.id)).toBe("pending")
  expect((await result(first, { state: "opened" })).status).toBe(410)
  const refused = await request(first.pathname, { headers: { Origin: browserOrigin } })
  expect(refused.status).toBe(403)
  expect((await request(first.pathname, { method: "HEAD" })).status).toBe(403)
  const redemption = await request(first.pathname)
  expect(redemption.status).toBe(200)
  expect(redemption.headers.get("Content-Type")).toContain("wordprocessingml.document")
  expect(redemption.headers.get("Cache-Control")).toBe("no-store")
  const firstBytes = new Uint8Array(await redemption.arrayBuffer()), firstLink = readLink(firstBytes)
  expect(firstLink.apiOrigin).toBe(paneOrigin)
  expect(await state(first.id)).toBe("redeemed")
  expect((await result(first, { state: "opened", path: "not-accepted" })).status).toBe(400)
  expect((await result(first, { state: "opened" }, browserOrigin)).status).toBe(403)
  expect((await result(first, { state: "opened" })).status).toBe(200)
  expect(await state(first.id)).toBe("completed")
  const otherMatter = await (await request("/matters", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ title: "Synthetic other matter" }) })).json() as { id: string }
  expect((await request(`/matters/${otherMatter.id}/documents/word/open/status?id=${first.id}`)).status).toBe(404)
  expect((await request(first.pathname)).status).toBe(410)
  const second = await issue()
  const secondBytes = new Uint8Array(await (await request(second.pathname)).arrayBuffer())
  expect(Buffer.compare(firstBytes, secondBytes) === 0).toBe(true)
  expect(readLink(secondBytes).token === firstLink.token).toBe(true)
  const otherBrowser = await issue("https://browser.example")
  const otherLink = readLink(new Uint8Array(await (await request(otherBrowser.pathname)).arrayBuffer()))
  expect(otherLink.token !== firstLink.token).toBe(true)
  expect((await result(otherBrowser, { state: "failed" })).status).toBe(200)
  expect(await state(otherBrowser.id)).toBe("failed")
  const cache = path.join(matter.dir, ".dochaus/word-handouts", firstLink.document)
  const files = await readdir(cache)
  expect(files.length > 0 && files.every(name => /^[a-f0-9]{64}\.[a-f0-9]{64}\.docx$/.test(name))).toBe(true)
  if (process.platform !== "win32") {
    expect((await stat(cache)).mode & 0o777).toBe(0o700)
    expect((await stat(path.join(cache, files[0]!))).mode & 0o777).toBe(0o600)
  }
  const revoked = await issue()
  expect((await route(`/documents/word?${query("Handoff.docx")}`, { method: "DELETE" })).status).toBe(200)
  expect((await request(revoked.pathname)).status).toBe(410)
  const renewed = await issue()
  const renewedLink = readLink(new Uint8Array(await (await request(renewed.pathname)).arrayBuffer()))
  expect(renewedLink.token !== firstLink.token).toBe(true)
  const changed = await issue()
  expect((await route("/documents/revisions", { method: "POST", body: file(await plain("A new current version."), "Handoff.docx") })).status).toBe(200)
  expect((await request(changed.pathname)).status).toBe(410)
  const fresh = await handout("Handoff.docx")
  expect(fresh.link.baseVersion !== renewedLink.baseVersion).toBe(true)
  expect(fresh.link.token !== renewedLink.token).toBe(true)
  expect((await readdir(cache)).length).toBe(1)
})

const latestCopy = (link: WordDocumentLink, extra: Record<string, string> = {}, input: object = { origin: browserOrigin }) => request("/word/latest", {
  method: "POST", headers: tokenHeaders(link, { "Content-Type": "application/json", "x-dochaus-word-base-version": link.baseVersion,
    "x-dochaus-word-session": owners.get(link.token)?.session ?? "session-test-0001", ...extra }), body: JSON.stringify(input),
})
const latestSource = async (response: Response) => {
  expect(response.status).toBe(200)
  expect(response.headers.get("Cache-Control")).toBe("no-store")
  const value = await response.json() as { uri: string; id: string }
  expect(value.id).toMatch(/^[a-f0-9-]{36}$/)
  return new URL(new URL(value.uri).searchParams.get("source")!)
}

test("inline latest reads only this capability's current document without versions or lease changes", async () => {
  await create("Inline latest.docx")
  const old = await handout("Inline latest.docx")
  const owner = await claim(old.link)
  const saved = await save(old.link, await plain("Latest inline wording."))
  const current = (await saved.json() as WordSaveResponse).version
  const response = await latestCopy(old.link, {}, { origin: browserOrigin, mode: "inline", name: "Other.docx", matter: "foreign" })
  expect(response.status).toBe(200)
  expect(response.headers.get("Cache-Control")).toBe("no-store")
  const value = await response.json() as { version: { id: string; number: number }; bytes: string; sha256: string }
  expect(value.version).toEqual({ id: current.id, number: current.number })
  const bytes = Buffer.from(value.bytes, "base64")
  expect(value.sha256).toBe(new Bun.CryptoHasher("sha256").update(bytes).digest("hex"))
  expect(strFromU8(unzipSync(bytes)["word/document.xml"]!)).toContain("Latest inline wording.")
  expect(Object.values(unzipSync(bytes)).map(part => strFromU8(part)).join("")).not.toContain(old.link.token)
  expect((await history("Inline latest.docx")).versions).toHaveLength(2)
  expect(await (await lease(old.link, "inspect")).json()).toMatchObject({ version: { state: "stale" }, lease: { state: "owned", generation: owner.lease.generation } })
  expect((await latestCopy(old.link, { "x-dochaus-word-base-version": "foreign" }, { origin: browserOrigin, mode: "inline" })).status).toBe(409)
  await writeFile(path.join(matter.dir, "Inline latest.docx"), await plain("Unrecorded outside edit."))
  expect((await latestCopy(old.link, {}, { origin: browserOrigin, mode: "inline" })).status).toBe(409)
})

test("latest copy opens a separate current handout and hands over only its own stale lease on native read", async () => {
  await create("Latest.docx")
  const old = await handout("Latest.docx")
  const owner = await claim(old.link)
  const saved = await save(old.link, await plain("The newer saved wording."))
  expect(saved.status).toBe(200)
  const current = (await saved.json() as WordSaveResponse).version
  const before = new Uint8Array(await readFile(path.join(matter.dir, "Latest.docx")))
  const source = await latestSource(await latestCopy(old.link, {}, { origin: browserOrigin, name: "Other.docx", matter: "not-this-matter" }))
  expect(source.origin).toBe(paneOrigin)
  expect(await (await lease(old.link, "inspect")).json()).toMatchObject({ version: { state: "stale" }, lease: { state: "owned", generation: owner.lease.generation } })
  expect((await history("Latest.docx")).versions).toHaveLength(2)
  const response = await request(source.pathname)
  expect(response.status).toBe(200)
  const freshBytes = new Uint8Array(await response.arrayBuffer()), fresh = readLink(freshBytes)
  expect(fresh).toMatchObject({ origin: browserOrigin, document: old.link.document, documentName: "Latest.docx", baseVersion: current.id })
  expect(fresh.token === old.link.token).toBe(false)
  expect(strFromU8(unzipSync(freshBytes)["word/document.xml"]!)).toContain("The newer saved wording.")
  expect(old.link.baseVersion === current.id).toBe(false)
  expect(await readFile(path.join(matter.dir, "Latest.docx"))).toEqual(Buffer.from(before))
  expect(await (await lease(old.link, "inspect")).json()).toMatchObject({ version: { state: "stale" }, lease: { state: "available" } })
  expect((await save(old.link, await plain("Old local work must never replace latest."))).status).toBe(423)
  expect((await request(source.pathname)).status).toBe(410)
  const freshOwner = await claim(fresh, "session-latest-0001")
  const currentSource = await latestSource(await latestCopy(fresh))
  expect((await request(currentSource.pathname)).status).toBe(200)
  expect(await (await lease(fresh, "inspect", "session-latest-0001")).json()).toMatchObject({ version: { state: "current" }, lease: { state: "owned", generation: freshOwner.lease.generation } })
  expect((await ping(fresh, "session-latest-0001", true)).status).toBe(200)
  expect((await history("Latest.docx")).versions).toHaveLength(2)
})

test("latest copy ticket cannot release a replacement editor or serve a later changed document", async () => {
  await create("Latest race.docx")
  const old = await handout("Latest race.docx")
  const initial = await claim(old.link)
  const saved = await save(old.link, await plain("Current wording before handoff."))
  const current = (await saved.json() as WordSaveResponse).version
  const source = await latestSource(await latestCopy(old.link))
  const fresh = await handout("Latest race.docx")
  const takeover = await lease(fresh.link, "takeover", "session-replacement-0001", current.id,
    { expectedCurrent: current.id, expectedGeneration: initial.lease.generation! })
  expect(takeover.status).toBe(200)
  const replacement = await takeover.json() as WordConnectionState
  owners.set(fresh.link.token, { session: "session-replacement-0001", generation: replacement.lease.generation! })
  expect((await request(source.pathname)).status).toBe(200)
  expect(await (await lease(fresh.link, "inspect", "session-replacement-0001")).json()).toMatchObject({ lease: { state: "owned", generation: replacement.lease.generation } })
  const changed = await latestSource(await latestCopy(old.link))
  expect((await save(fresh.link, await plain("A later version after preparing the link."))).status).toBe(200)
  expect((await request(changed.pathname)).status).toBe(410)
})

test("latest copy refuses unrecorded changes and invalid or revoked document capabilities", async () => {
  await create("Latest guard.docx")
  const { link } = await handout("Latest guard.docx")
  expect((await latestCopy(link, { Origin: "https://unrelated.invalid" })).status).toBe(403)
  expect((await latestCopy(link, { Authorization: "Bearer invalid" })).status).toBe(401)
  expect((await latestCopy(link, {}, { origin: "file:///tmp" })).status).toBe(400)
  expect((await latestCopy(link, { "x-dochaus-word-base-version": "foreign-version" })).status).toBe(409)
  const outside = await plain("Outside edit must be resolved in doc.haus.")
  await writeFile(path.join(matter.dir, "Latest guard.docx"), outside)
  const refused = await latestCopy(link)
  expect(refused.status).toBe(409)
  expect(await refused.json()).toMatchObject({ code: "word_document_changed", version: { state: "unrecorded-change" } })
  expect((await history("Latest guard.docx")).versions).toHaveLength(1)
  expect(new Uint8Array(await readFile(path.join(matter.dir, "Latest guard.docx")))).toEqual(new Uint8Array(outside))
  expect((await route(`/documents/word?${query("Latest guard.docx")}`, { method: "DELETE" })).status).toBe(200)
  expect((await latestCopy(link)).status).toBe(401)
})

const deskSources = async (link: WordDocumentLink) => (await (await request("/word/desk/sources", { method: "POST", headers: tokenHeaders(link) })).json() as DeskSources).sources
const deskOpen = (link: WordDocumentLink, input: object) => request("/word/desk/open", { method: "POST", headers: tokenHeaders(link, { "Content-Type": "application/json" }), body: JSON.stringify(input) })

test("desk native opening preserves both Word sessions and opens only the selected saved matter document", async () => {
  await create("Desk caller.docx")
  await create("Desk target.docx", "A separate matter document for desktop Word.")
  const caller = await handout("Desk caller.docx"), target = await handout("Desk target.docx")
  await claim(caller.link, "session-desk-caller")
  await claim(target.link, "session-desk-target")
  const selected = (await deskSources(caller.link)).find(row => row.name === "Desk target.docx")!
  const snapshot = () => {
    const db = openDb(matter.dir)
    try { return { leases: db.query("SELECT * FROM word_edit_leases ORDER BY document_id").all(), sessions: db.query("SELECT * FROM word_sessions ORDER BY token_hash, session").all(), versions: db.query("SELECT id, sha256 FROM document_versions ORDER BY id").all() } }
    finally { db.close() }
  }
  const before = snapshot(), callerBytes = await readFile(path.join(matter.dir, "Desk caller.docx")), targetBytes = await readFile(path.join(matter.dir, "Desk target.docx"))
  const input = { sourceId: selected.id, expectedHash: selected.hash, origin: browserOrigin }
  const first = await latestSource(await deskOpen(caller.link, input))
  expect(snapshot()).toEqual(before)
  const native = await request(first.pathname)
  expect(native.status).toBe(200)
  const opened = new Uint8Array(await native.arrayBuffer()), linked = readLink(opened)
  expect(linked).toMatchObject({ document: target.link.document, documentName: "Desk target.docx", baseVersion: target.link.baseVersion, origin: browserOrigin, apiOrigin: paneOrigin })
  expect(linked.token).toBe(target.link.token)
  expect(strFromU8(unzipSync(opened)["word/document.xml"]!)).toContain("A separate matter document for desktop Word.")
  expect((await request(first.pathname)).status).toBe(410)
  const repeated = await latestSource(await deskOpen(caller.link, input))
  expect(new Uint8Array(await (await request(repeated.pathname)).arrayBuffer())).toEqual(opened)
  expect(snapshot()).toEqual(before)
  expect(await readFile(path.join(matter.dir, "Desk caller.docx"))).toEqual(callerBytes)
  expect(await readFile(path.join(matter.dir, "Desk target.docx"))).toEqual(targetBytes)
  // Selecting this pane's current document also leaves its lease untouched.
  const own = (await deskSources(caller.link)).find(row => row.name === "Desk caller.docx")!
  const ownSource = await latestSource(await deskOpen(caller.link, { sourceId: own.id, expectedHash: own.hash, origin: browserOrigin }))
  expect((await request(ownSource.pathname)).status).toBe(200)
  expect(snapshot()).toEqual(before)
})

test("desk native opening refuses other matters, changed files and revoked originating grants", async () => {
  await create("Desk guarded caller.docx")
  await create("Desk guarded target.docx", "Confirmed target wording.")
  const caller = await handout("Desk guarded caller.docx")
  let selected = (await deskSources(caller.link)).find(row => row.name === "Desk guarded target.docx")!
  const input = () => ({ sourceId: selected.id, expectedHash: selected.hash, origin: browserOrigin })
  expect((await deskOpen(caller.link, { ...input(), sourceId: "../../outside.docx" })).status).toBe(400)
  expect((await deskOpen(caller.link, { ...input(), name: "Desk target.docx" })).status).toBe(400)
  expect((await deskOpen(caller.link, { ...input(), origin: "file:///tmp" })).status).toBe(400)
  const otherMatter = await (await request("/matters", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ title: "Synthetic foreign desk matter" }) })).json() as { id: string; dir: string }
  expect((await request(`/matters/${otherMatter.id}/deliverables`, { method: "POST", body: file(await plain("Foreign matter content must not be exposed."), "Foreign.docx") })).status).toBe(200)
  const db = openDb(otherMatter.dir)
  let foreign: { id: number }
  try { foreign = db.query("SELECT id FROM documents WHERE name='Foreign.docx'").get() as { id: number } } finally { db.close() }
  const { sha256Hex } = await import("./document-versions")
  const foreignHash = sha256Hex(await readFile(path.join(otherMatter.dir, "Foreign.docx")))
  expect((await deskOpen(caller.link, { sourceId: String(foreign.id), expectedHash: foreignHash, origin: browserOrigin })).status).toBe(409)
  expect((await deskOpen(caller.link, { ...input(), matter: otherMatter.id })).status).toBe(400)
  const prepared = await latestSource(await deskOpen(caller.link, input()))
  expect((await route("/documents/revisions", { method: "POST", body: file(await plain("A later recorded target version."), "Desk guarded target.docx") })).status).toBe(200)
  expect((await request(prepared.pathname)).status).toBe(410)
  expect((await deskOpen(caller.link, input())).status).toBe(409)
  selected = (await deskSources(caller.link)).find(row => row.name === "Desk guarded target.docx")!
  const revoked = await latestSource(await deskOpen(caller.link, input()))
  expect((await route(`/documents/word?${query("Desk guarded caller.docx")}`, { method: "DELETE" })).status).toBe(200)
  expect((await deskOpen(caller.link, input())).status).toBe(401)
  // The target's fresh grant stays valid; revoking the requester still retires
  // this handoff, so the ticket cannot outlive its original authority.
  expect((await request(revoked.pathname)).status).toBe(410)
  const fresh = await handout("Desk guarded caller.docx")
  const before = await history("Desk guarded target.docx")
  await writeFile(path.join(matter.dir, "Desk guarded target.docx"), await plain("An unrecorded external replacement."))
  expect((await deskOpen(fresh.link, input())).status).toBe(409)
  expect((await history("Desk guarded target.docx")).versions).toEqual(before.versions)
  // Restoring the recorded bytes makes the same selection available again;
  // ambiguous registry aliases and later redaction must still retire tickets.
  const recorded = await route(`/documents/versions/${before.versions[0]!.id}/content`)
  await writeFile(path.join(matter.dir, "Desk guarded target.docx"), new Uint8Array(await recorded.arrayBuffer()))
  const aliasTicket = await latestSource(await deskOpen(fresh.link, input()))
  const aliases = openDb(matter.dir)
  try { aliases.run("INSERT INTO documents(doc_path, name, created_at, role) VALUES (?, ?, ?, 'deliverable')", [`${matter.dir}/./Desk guarded target.docx`, "Desk guarded target.docx", Date.now()]) }
  finally { aliases.close() }
  expect((await deskOpen(fresh.link, input())).status).toBe(409)
  expect((await request(aliasTicket.pathname)).status).toBe(410)
  const restore = openDb(matter.dir)
  try { restore.run("DELETE FROM documents WHERE doc_path=?", [`${matter.dir}/./Desk guarded target.docx`]) } finally { restore.close() }
  const redacted = await latestSource(await deskOpen(fresh.link, input()))
  expect((await route("/documents/revisions", { method: "POST", headers: { "x-dochaus-origin": "redaction" }, body: file(await plain("[REDACTED]"), "Desk guarded target.docx") })).status).toBe(200)
  expect((await request(redacted.pathname)).status).toBe(410)
  expect((await deskOpen(fresh.link, input())).status).toBe(409)
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
  const invalid = await request("/word/saves", { method: "POST", headers: { Origin: paneOrigin, Authorization: "Bearer invalid", "x-dochaus-word-version": WORD_ADDIN_VERSION } })
  expect(invalid.status).toBe(401)
  expect(invalid.headers.get("Access-Control-Allow-Origin")).toBe(paneOrigin)
  const malformed = await save(link, new Uint8Array([1, 2, 3]))
  expect([malformed.status, (await malformed.json() as { code: string }).code]).toEqual([400, "invalid_docx"])
  expect(malformed.headers.get("Access-Control-Allow-Origin")).toBe(paneOrigin)
})

test("one owner saves current ancestry; stale writes, lost-ACK retries and takeovers cannot revert work", async () => {
  await create("Branches.docx")
  const { link } = await handout("Branches.docx")
  const initial = await claim(link, "session-first-0001")
  const other = await lease(link, "claim", "session-second-0002")
  expect(other.status).toBe(423)
  expect(await other.json()).toMatchObject({ version: { state: "current" }, lease: { state: "other", generation: initial.lease.generation } })
  expect((await status()).find(document => document.name === "Branches.docx")).toMatchObject({ open: true, overlap: false })
  const firstBytes = linkedWordDocx(await plain("First Word edit.", "Synthetic lawyer one"), link)
  const first = await save(link, firstBytes)
  expect(first.status).toBe(200)
  const firstSave = await first.json() as WordSaveResponse
  expect(firstSave).toMatchObject({ version: { number: 2, parent: link.baseVersion, created: true }, overlap: false,
    connection: { version: { base: { id: expect.any(String) }, state: "current" }, lease: { state: "owned" } } })
  // An exact current export confirms a lost ACK with its original caller base.
  expect(await (await save(link, firstBytes)).json()).toMatchObject({ version: { id: firstSave.version.id, created: false } })
  expect(readWorkingDocx(matter.dir, "Branches.docx")).toBeDefined()
  const stale = await save(link, await plain("Stale edit must stay local."))
  expect(stale.status).toBe(409)
  expect(await stale.json()).toMatchObject({ code: "word_base_stale", version: { state: "stale", current: { id: firstSave.version.id } } })
  const staleDone = await ping(link, "session-first-0001", true)
  expect(staleDone.status).toBe(409)
  expect(await staleDone.json()).toMatchObject({ version: { state: "stale" }, lease: { state: "owned", generation: initial.lease.generation } })
  const second = await save(link, await plain("Current next edit.", "Synthetic lawyer two"), firstSave.version.id)
  expect(second.status).toBe(200)
  const secondSave = await second.json() as WordSaveResponse
  expect(secondSave).toMatchObject({ version: { number: 3, parent: firstSave.version.id, created: true }, overlap: false })
  expect((await save(link, firstBytes)).status).toBe(409)
  expect((await history("Branches.docx")).versions.map(row => [row.number, row.origin, row.actor, row.parent])).toEqual([
    [3, "word", "Synthetic lawyer two", firstSave.version.id], [2, "word", "Synthetic lawyer one", link.baseVersion], [1, "generated", "unknown", null],
  ])
  const staleTakeover = await lease(link, "takeover", "session-second-0002", link.baseVersion, {
    expectedCurrent: secondSave.version.id, expectedGeneration: initial.lease.generation })
  expect(staleTakeover.status).toBe(409)
  const wrongCAS = await lease(link, "takeover", "session-second-0002", secondSave.version.id, {
    expectedCurrent: secondSave.version.id, expectedGeneration: "wrong-generation" })
  expect(wrongCAS.status).toBe(409)
  const taken = await lease(link, "takeover", "session-second-0002", secondSave.version.id, {
    expectedCurrent: secondSave.version.id, expectedGeneration: initial.lease.generation })
  expect(taken.status).toBe(200)
  const state = await taken.json() as WordConnectionState
  expect(state.lease.generation).not.toBe(initial.lease.generation)
  for (const done of [false, true]) {
    const fenced = await ping(link, "session-first-0001", done, secondSave.version.id)
    expect(fenced.status).toBe(423)
    expect(await fenced.json()).toMatchObject({ version: { state: "current" }, lease: { state: "other", generation: state.lease.generation } })
  }
  const fencedSave = await save(link, await plain("Fenced write."), secondSave.version.id)
  expect(fencedSave.status).toBe(423)
  expect(await fencedSave.json()).toMatchObject({ version: { state: "current" }, lease: { state: "other", generation: state.lease.generation } })
  owners.set(link.token, { session: "session-second-0002", generation: state.lease.generation! })
  expect((await ping(link, "session-second-0002", true, secondSave.version.id)).status).toBe(200)
  // Done is repeatable only until a replacement generation exists.
  expect((await ping(link, "session-second-0002", true, secondSave.version.id)).status).toBe(200)
  const current = new Uint8Array(await (await route(`/documents/content?${query("Branches.docx")}`)).arrayBuffer())
  expect(strFromU8(unzipSync(current)["word/document.xml"])).toContain("Current next edit.")
  expect(Object.values(unzipSync(current)).map(bytes => strFromU8(bytes)).join("\n")).not.toContain(link.token)
  expect((await status()).find(document => document.name === "Branches.docx")).toMatchObject({ currentVersion: secondSave.version.id, open: false, overlap: false })
})

test("deleting and recreating the filename does not revive a token; stale copies cannot restore redacted text", async () => {
  await create("Lifecycle.docx", "Original restricted phrase.")
  const old = await handout("Lifecycle.docx")
  expect((await route(`/documents?${query("Lifecycle.docx")}`, { method: "DELETE" })).status).toBe(200)
  await create("Lifecycle.docx", "New document identity.")
  const current = await handout("Lifecycle.docx")
  expect(current.link.document).not.toBe(old.link.document)
  expect((await save(old.link, await plain("Attempt from a deleted document."))).status).toBe(401)
  expect((await latestCopy(old.link)).status).toBe(401)
  const redacted = await plain("[REDACTED]")
  const redaction = await route("/documents/revisions", { method: "POST", headers: { "x-dochaus-origin": "redaction" }, body: file(redacted, "Lifecycle.docx") })
  expect(redaction.status).toBe(200)
  const stale = await save(current.link, await plain("New document identity."))
  expect(stale.status).toBe(401)
  expect((await latestCopy(current.link)).status).toBe(401)
  const fresh = await handout("Lifecycle.docx")
  await claim(fresh.link)
  expect((await save(fresh.link, await plain("Preserves [REDACTED], with a neutral addition."))).status).toBe(200)
  expect((await history("Lifecycle.docx")).versions.some(row => row.removed === "redaction")).toBe(true)
})

test("expiry fences old saves and browser recovery requires displayed current and generation", async () => {
  await create("Recovery.docx")
  const { link } = await handout("Recovery.docx")
  const acquired = await claim(link)
  const bytes = new Uint8Array(await readFile(path.join(matter.dir, "Recovery.docx")))
  expect(await (await save(link, bytes)).json()).toMatchObject({ version: { id: link.baseVersion, created: false }, overlap: false })
  const db = openDb(matter.dir)
  try { db.run("UPDATE word_edit_leases SET expires_at = 1 WHERE document_id = ?", [link.document]) } finally { db.close() }
  const expiredHeartbeat = await ping(link, "session-test-0001")
  expect(expiredHeartbeat.status).toBe(423)
  expect(await expiredHeartbeat.json()).toMatchObject({ version: { state: "current" }, lease: { state: "available", generation: acquired.lease.generation } })
  const expiredSave = await save(link, await plain("Expired write."))
  expect(expiredSave.status).toBe(423)
  expect(await expiredSave.json()).toMatchObject({ version: { state: "current" }, lease: { state: "available", generation: acquired.lease.generation } })
  const inspected = await lease(link, "inspect")
  expect(await inspected.json()).toMatchObject({ version: { state: "current" }, lease: { state: "available", generation: acquired.lease.generation } })
  const recover = (expectedCurrent: string, expectedGeneration: string | null, action = "release") => route(`/documents/word/recover?${query("Recovery.docx")}`, {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ action, expectedCurrent, expectedGeneration }) })
  expect((await recover("stale-version", acquired.lease.generation!)).status).toBe(409)
  expect((await recover(link.baseVersion, "wrong-generation")).status).toBe(409)
  expect((await recover(link.baseVersion, acquired.lease.generation!, "takeover")).status).toBe(200)
  expect((await lease(link, "claim")).status).toBe(401)
  const fresh = await handout("Recovery.docx")
  await claim(fresh.link)
  expect((await save(fresh.link, await plain("Recovered current edit."))).status).toBe(200)
  expect((await history("Recovery.docx")).versions).toHaveLength(2)
})

test("a source save reindexes canonical text and ordinary current/history/redlined downloads strip linkage", async () => {
  await create("Source.docx", "The synthetic source begins here.", true)
  const { link } = await handout("Source.docx")
  await claim(link)
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
  await claim(link)
  const clean = await save(link, await plain("Clean accepted Word state."))
  expect(clean.status).toBe(200)
  const baseVersion = (await clean.json() as WordSaveResponse).version.id
  const accepted = readWorkingDocx(matter.dir, "Tracked.docx")
  const tracked = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph({ children: [new InsertedTextRun({ text: "Pending tracked insertion.", id: 1, author: "Synthetic reviewer", date: "2026-10-26T10:00:00Z" })] })] }] }))
  const response = await save(link, tracked, baseVersion)
  expect(response.status).toBe(200)
  const trackedAck = await response.json() as WordSaveResponse
  await ping(link, "session-test-0001", true, trackedAck.version.id)
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

test("native operation admission fences takeover, recovery, revoke and every save until its exact owner settles", async () => {
  await create("Native fence.docx")
  const { link, bytes } = await handout("Native fence.docx")
  const acquired = await claim(link), generation = acquired.lease.generation!
  const operation = (input: object, session = "session-test-0001", fence = generation) => request("/word/native-operation", {
    method: "POST", headers: tokenHeaders(link, { "content-type": "application/json", "x-dochaus-word-session": session,
      "x-dochaus-word-lease": fence, "x-dochaus-word-base-version": link.baseVersion }), body: JSON.stringify(input),
  })
  const begin = { action: "begin", sequence: 0, operationId: "synthetic-operation-1", intent: "refresh", expectedCurrent: link.baseVersion }
  expect((await operation({...begin, sequence: -1})).status).toBe(400)
  expect((await operation({...begin, unexpected: true})).status).toBe(400)
  expect((await operation(begin, "different-session")).status).toBe(423)
  expect((await request("/word/native-operation", {method:"POST", headers:tokenHeaders(link, {"content-type":"application/json", "x-dochaus-word-session":"session-test-0001", "x-dochaus-word-lease":generation, "x-dochaus-word-base-version":link.baseVersion}), body:"x".repeat(4097)})).status).toBe(400)
  expect(await (await operation(begin)).json()).toMatchObject({operation: { sequence: 1, operationId: "synthetic-operation-1", state: "pending" }})
  expect(await (await operation(begin)).json()).toMatchObject({operation: { sequence: 1, state: "pending" }})
  const inspected = await (await lease(link, "inspect")).json() as WordConnectionState
  expect(inspected.nativeOperation).toEqual({ sequence: 1, state: "pending" })
  const issued = await route(`/documents/word/open?${query("Native fence.docx")}`, {method:"POST", headers:{"content-type":"application/json"}, body:JSON.stringify({origin:browserOrigin})})
  expect(issued.status).toBe(200)
  const handoff = await issued.json() as { uri:string; id:string }
  const openPath = new URL(new URL(handoff.uri).searchParams.get("source")!).pathname
  const readState = () => { const registry = openDb(matter.dir); try { return { lease: registry.query("SELECT * FROM word_edit_leases WHERE document_id=?").get(link.document), sessions: registry.query("SELECT * FROM word_sessions ORDER BY session").all() } } finally { registry.close() } }
  const beforeOpen = readState(), refusedOpen = await request(openPath)
  expect(refusedOpen.status).toBe(423)
  expect((await refusedOpen.json() as {error:string}).error).toContain("Open in Word again")
  expect(readState()).toEqual(beforeOpen)
  expect((await (await route(`/documents/word/open/status?id=${handoff.id}`)).json() as {state:string}).state).toBe("failed")
  expect((await request(openPath)).status).toBe(410)
  expect((await save(link, bytes)).status).toBe(423)
  expect((await ping(link, "session-test-0001", true)).status).toBe(423)
  expect((await lease(link, "release")).status).toBe(423)
  expect((await lease(link, "takeover", "different-session", link.baseVersion, { expectedCurrent: link.baseVersion, expectedGeneration: generation })).status).toBe(423)
  const recover = () => route(`/documents/word/recover?${query("Native fence.docx")}`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ action: "takeover", expectedCurrent: link.baseVersion, expectedGeneration: generation }) })
  expect((await recover()).status).toBe(423)
  expect((await route(`/documents/word?${query("Native fence.docx")}`, { method: "DELETE" })).status).toBe(423)
  expect((await lease(link, "inspect")).status).toBe(200)
  expect((await operation({action: "settle", sequence: 1, operationId: "other", outcome: "confirmed"})).status).toBe(423)
  const db = openDb(matter.dir)
  try { db.run("UPDATE word_edit_leases SET expires_at=1 WHERE document_id=?", [link.document]) } finally { db.close() }
  expect((await lease(link, "claim", "different-session")).status).toBe(423)
  expect((await ping(link, "session-test-0001")).status).toBe(200)
  const settle = {action: "settle", sequence: 1, operationId: "synthetic-operation-1", outcome: "confirmed"}
  expect((await operation(settle)).status).toBe(200)
  expect((await operation(settle)).status).toBe(200)
  expect((await operation({...settle, outcome: "not-dispatched"})).status).toBe(423)
  expect((await save(link, bytes)).status).toBe(200)
})

test("a revoked grant cannot erase a pending native operation or let a replacement edit", async () => {
  await create("Revoked native fence.docx")
  const { link } = await handout("Revoked native fence.docx")
  const acquired = await claim(link)
  const headers = tokenHeaders(link, {"content-type":"application/json", "x-dochaus-word-session":"session-test-0001", "x-dochaus-word-lease": acquired.lease.generation!, "x-dochaus-word-base-version":link.baseVersion})
  expect((await request("/word/native-operation", { method:"POST", headers, body:JSON.stringify({ action:"begin", sequence:0, operationId:"synthetic-revoked", intent:"refresh", expectedCurrent:link.baseVersion }) })).status).toBe(200)
  const db = openDb(matter.dir)
  try { db.run("UPDATE word_grants SET revoked_at=? WHERE document_id=?", [Date.now(), link.document]); db.run("UPDATE word_edit_leases SET expires_at=1 WHERE document_id=?", [link.document]) } finally { db.close() }
  expect((await request("/word/native-operation", { method:"POST", headers, body:JSON.stringify({ action:"settle", sequence:1, operationId:"synthetic-revoked", outcome:"confirmed" }) })).status).toBe(401)
  const fresh = await handout("Revoked native fence.docx")
  expect((await lease(fresh.link, "claim", "different-session")).status).toBe(423)
  const recover = await route(`/documents/word/recover?${query("Revoked native fence.docx")}`, { method:"POST", headers:{"content-type":"application/json"}, body:JSON.stringify({action:"takeover", expectedCurrent:link.baseVersion, expectedGeneration:acquired.lease.generation}) })
  expect(recover.status).toBe(423)
  expect((await lease(fresh.link, "inspect", "different-session")).status).toBe(200)
})

test("refresh admission binds stale source, exact latest version and original live owner", async () => {
  await create("Native refresh fence.docx")
  const { link } = await handout("Native refresh fence.docx"), acquired = await claim(link)
  const saved = await save(link, await plain("New current version."))
  expect(saved.status).toBe(200)
  const latest = (await saved.json() as WordSaveResponse).version.id
  const headers = tokenHeaders(link, {"content-type":"application/json", "x-dochaus-word-session":"session-test-0001", "x-dochaus-word-lease":acquired.lease.generation!, "x-dochaus-word-base-version":link.baseVersion})
  const begin = (input: object, hdr: Record<string, string> = headers) => request("/word/native-operation", {method:"POST", headers:hdr, body:JSON.stringify(input)})
  const refresh = { action:"begin", sequence:0, operationId:"refresh-operation", intent:"refresh", expectedCurrent:latest }
  expect((await begin({...refresh, expectedCurrent:"changed-version"})).status).toBe(409)
  expect((await begin({...refresh, intent:"amendment", expectedCurrent:undefined})).status).toBe(400)
  expect((await begin(refresh, {...headers,"x-dochaus-word-session":"different-session"})).status).toBe(423)
  const filename = path.join(matter.dir, "Native refresh fence.docx"), before = await readFile(filename)
  await writeFile(filename, await plain("Unrecorded replacement."))
  expect((await begin(refresh)).status).toBe(409)
  await writeFile(filename, before)
  const db = openDb(matter.dir)
  try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [link.baseVersion]) } finally { db.close() }
  expect((await begin(refresh)).status).toBe(401)
  const restore = openDb(matter.dir)
  try { restore.run("UPDATE document_versions SET removed=NULL WHERE id=?", [link.baseVersion]) } finally { restore.close() }
  const beforeAdmission = openDb(matter.dir)
  try { expect(beforeAdmission.query("SELECT native_sequence,native_state FROM word_edit_leases WHERE document_id=?").get(link.document)).toEqual({native_sequence:0,native_state:null}) } finally { beforeAdmission.close() }
  const admitted = await begin(refresh)
  expect(admitted.status).toBe(200)
  expect(await admitted.json()).toMatchObject({operation:{sequence:1,operationId:"refresh-operation",state:"pending"},connection:{version:{base:{id:link.baseVersion},current:{id:latest},state:"stale"},lease:{state:"owned",generation:acquired.lease.generation}}})
  expect((await begin(refresh)).status).toBe(200)
  expect((await begin({action:"settle",sequence:1,operationId:"refresh-operation",outcome:"not-dispatched"})).status).toBe(200)
})

test("pending Word recovery copies only the confirmed saved version and leaves the original fenced", async () => {
  await create("Isolated recovery.docx")
  const { link } = await handout("Isolated recovery.docx"), owner = await claim(link)
  const headers = tokenHeaders(link, {"content-type":"application/json", "x-dochaus-word-session":"session-test-0001", "x-dochaus-word-lease":owner.lease.generation!, "x-dochaus-word-base-version":link.baseVersion})
  expect((await request("/word/native-operation", { method:"POST", headers, body:JSON.stringify({action:"begin",sequence:0,operationId:"recovery-op",intent:"refresh",expectedCurrent:link.baseVersion}) })).status).toBe(200)
  const sourcePath = path.join(matter.dir, "Isolated recovery.docx"), saved = await readFile(sourcePath)
  const db = openDb(matter.dir)
  try { db.run("UPDATE word_grants SET revoked_at=? WHERE document_id=?", [Date.now(), link.document]) } finally { db.close() }
  // Simulate an unrelated, unrecorded external working-file change; the copy
  // must use the retained version, not this mutable file.
  const external = await plain("Unrecorded external local source change.")
  await writeFile(sourcePath, external)
  const list = await (await route("/documents/word")).json() as { documents: { name:string; documentId?:string; nativeOperation?:{sequence:number;state:string} }[] }
  expect(list.documents.find(row => row.name === "Isolated recovery.docx")).toMatchObject({documentId:link.document,nativeOperation:{sequence:1,state:"pending"}})
  expect(JSON.stringify(list)).not.toContain(link.token)
  expect(JSON.stringify(list)).not.toContain(owner.lease.generation!)
  expect(JSON.stringify(list)).not.toContain("recovery-op")
  const snapshot = () => { const registry=openDb(matter.dir); try { return {
    document: registry.query("SELECT * FROM versioned_documents WHERE id=?").get(link.document),
    versions: registry.query("SELECT * FROM document_versions WHERE document_id=? ORDER BY number").all(link.document),
    lease: registry.query("SELECT * FROM word_edit_leases WHERE document_id=?").get(link.document),
    grants: registry.query("SELECT * FROM word_grants WHERE document_id=? ORDER BY token_hash").all(link.document),
    sessions: registry.query("SELECT s.* FROM word_sessions s JOIN word_grants g ON g.token_hash=s.token_hash WHERE g.document_id=? ORDER BY s.session").all(link.document),
    roles: registry.query("SELECT name,role,doc_path FROM documents WHERE name=?").all("Isolated recovery.docx"),
    redlines: registry.query("SELECT * FROM redlines WHERE doc_name=?").all("Isolated recovery.docx"),
  } } finally { registry.close() } }
  const before = snapshot()
  const payload = {source:"Isolated recovery.docx",name:"Recovery copy.docx",document:link.document,version:link.baseVersion,sequence:1}
  const copy = (body:object) => route("/documents/word/recovery-copy", {method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify(body)})
  expect((await copy({...payload,sequence:2})).status).toBe(409)
  expect((await copy({...payload,version:"other"})).status).toBe(409)
  expect((await copy({...payload,name:"../unsafe.docx"})).status).toBe(409)
  expect((await copy({...payload,name:"Isolated recovery.docx"})).status).toBe(409)
  expect(snapshot()).toEqual(before)
  const originalVersion = before.versions[0] as {document_id:string;sha256:string;ext:string}
  const keptFile = path.join(matter.dir, ".dochaus", "revisions", originalVersion.document_id, `${originalVersion.sha256}${originalVersion.ext}`)
  await writeFile(keptFile, "corrupt snapshot")
  expect((await copy(payload)).status).toBe(409)
  expect(snapshot()).toEqual(before)
  await writeFile(keptFile, saved)
  const registry = openDb(matter.dir)
  try { registry.run("CREATE TRIGGER refuse_recovery_copy BEFORE INSERT ON documents WHEN NEW.name='Blocked recovery.docx' BEGIN SELECT RAISE(ABORT, 'synthetic'); END") } finally { registry.close() }
  expect((await copy({...payload,name:"Blocked recovery.docx"})).status).toBe(503)
  expect(await Bun.file(path.join(matter.dir,"Blocked recovery.docx")).exists()).toBe(false)
  expect(snapshot()).toEqual(before)
  const cleanup = openDb(matter.dir)
  try { cleanup.run("DROP TRIGGER refuse_recovery_copy") } finally { cleanup.close() }
  const responses = await Promise.all([copy(payload), copy(payload)])
  expect(responses.map(response => response.status).sort()).toEqual([200,409])
  const result = responses.find(response => response.status === 200)!
  expect(await result.json()).toMatchObject({name:"Recovery copy.docx",sourceVersion:link.baseVersion})
  expect(await readFile(path.join(matter.dir,"Recovery copy.docx"))).toEqual(saved)
  expect(Buffer.compare(await readFile(sourcePath), external)).toBe(0)
  expect(snapshot()).toEqual(before)
  expect((await copy(payload)).status).toBe(409)
  const recovered = await handout("Recovery copy.docx")
  expect(recovered.link.document).not.toBe(link.document)
  expect(recovered.link.token).not.toBe(link.token)
  await claim(recovered.link, "copy-session-0001")
  expect((await save(recovered.link, recovered.bytes)).status).toBe(200)
  expect(snapshot()).toEqual(before)
})


test("saved recovery strips an inherited Word grant while retaining document content", async () => {
  await create("Link seed.docx")
  const seed = await handout("Link seed.docx")
  const linked = linkedWordDocx(await plain("Saved with an earlier Word link."), seed.link)
  expect((await route("/deliverables", {method:"POST", body:file(linked,"Embedded recovery.docx")})).status).toBe(200)
  const original = await readFile(path.join(matter.dir,"Embedded recovery.docx"))
  expect(readWordLink(original)).toBeTruthy()
  const {link} = await handout("Embedded recovery.docx"), owner = await claim(link)
  const headers = tokenHeaders(link,{"content-type":"application/json","x-dochaus-word-session":"session-test-0001","x-dochaus-word-lease":owner.lease.generation!,"x-dochaus-word-base-version":link.baseVersion})
  expect((await request("/word/native-operation", {method:"POST",headers,body:JSON.stringify({action:"begin",intent:"refresh",sequence:0,operationId:"embedded-recovery",expectedCurrent:link.baseVersion})})).status).toBe(200)
  const copied = await route("/documents/word/recovery-copy", {method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({source:"Embedded recovery.docx",name:"Unlinked recovery.docx",document:link.document,version:link.baseVersion,sequence:1})})
  expect(copied.status).toBe(200)
  const result = await readFile(path.join(matter.dir,"Unlinked recovery.docx"))
  expect(readWordLink(result)).toBeNull()
  expect(unzipSync(result)["word/document.xml"]).toEqual(unzipSync(original)["word/document.xml"])
  expect(await readFile(path.join(matter.dir,"Embedded recovery.docx"))).toEqual(original)
  const handoutCopy = await handout("Unlinked recovery.docx")
  expect(handoutCopy.link.document).not.toBe(link.document)
  expect(handoutCopy.link.token).not.toBe(seed.link.token)
})
