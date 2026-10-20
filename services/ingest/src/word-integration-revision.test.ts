import { afterAll, beforeAll, expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { Document, Packer, Paragraph } from "docx"
import wordIntegration from "../../../dochaus/tool/word-integration"
import searchDocument from "../../../dochaus/tool/search-document"
import { paragraphTargets } from "../../../dochaus/lib/drafting-amendment"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import { memoNotesFile, readLeanDocuments, recordLeanDocuments } from "../../../dochaus/workflows/lean/deliverables"
import { writeLeanPlan } from "../../../dochaus/workflows/lean/plan"
import { initLeanState, readLeanState, writeLeanBytes, writeLeanState, writeLeanText } from "../../../dochaus/workflows/lean/state"
import { scriptedLeanModel } from "../../../dochaus/workflows/lean/test-support"
import { getDocumentRole, listSourceDocuments, openDb } from "./db"
import { leanDraftingStatus } from "./lean-drafting"
import { registerLeanDeliverable } from "./lean-deliverables"

// The word-integration tool's replace saves through ingest's revision route, so
// an edited source is re-indexed, a deliverable keeps its role and pointers, a
// saved combined file is rebuilt, and a refused save leaves the file as it was.
// A real ingest process with real local embeddings; no generation engine.

const temporary = mkdtempSync(path.join(tmpdir(), "dochaus-word-integration-"))
let ingest: ReturnType<typeof Bun.spawn>
let engine: ReturnType<typeof Bun.serve>
let base = ""
const priorIngestUrl = process.env.INGEST_URL

beforeAll(async () => {
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  await Bun.write(path.join(config, "opencode.json"), "{}")
  await Bun.write(path.join(workspace, ".keep"), "")
  engine = Bun.serve({ port: 0, fetch: () => new Response("No engine in this test", { status: 503 }) })
  const port = closedPort()
  const log = path.join(temporary, "ingest.log")
  ingest = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config, INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "" },
    stdout: Bun.file(log), stderr: Bun.file(log),
  })
  base = `http://127.0.0.1:${port}`
  const deadline = Date.now() + 20_000
  while (!await request("/matters").then(response => response.ok).catch(() => false)) {
    if (ingest.exitCode !== null || Date.now() > deadline) throw new Error(`ingest did not start:\n${await Bun.file(log).text()}`)
    await Bun.sleep(50)
  }
  process.env.INGEST_URL = base
})

afterAll(async () => {
  if (priorIngestUrl === undefined) delete process.env.INGEST_URL
  else process.env.INGEST_URL = priorIngestUrl
  ingest?.kill()
  await ingest?.exited
  engine?.stop(true)
  rmSync(temporary, { recursive: true, force: true })
})

function closedPort() {
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port!
  probe.stop(true)
  return port
}

function request(route: string, init?: RequestInit) {
  return fetch(base + route, { ...init, signal: AbortSignal.timeout(60_000) })
}

async function newMatter(title: string) {
  const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title }) })
  expect(created.status).toBe(200)
  const matter = await created.json() as { id: string; dir: string }
  return { ...matter, route: `/matters/${matter.id}` }
}

async function register(route: string, kind: "documents" | "deliverables", name: string, bytes: Uint8Array) {
  const form = new FormData()
  form.append("file", new File([new Uint8Array(bytes)], name))
  const response = await request(`${route}/${kind}`, { method: "POST", body: form })
  expect(response.status).toBe(200)
}

async function docx(...paragraphs: string[]) {
  return new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: paragraphs.map(text => new Paragraph(text)) }] })))
}

function context(directory: string, sessionID = "word-integration-test") {
  return { callID: "call", sessionID, messageID: "message", agent: "test", directory, worktree: directory, abort: new AbortController().signal, metadata() {}, async ask() {} }
}

const search = async (directory: string, query: string, sessionID: string) => {
  const result = await searchDocument.execute({ query }, context(directory, sessionID))
  return typeof result === "string" ? result : result.output
}

const replace = (directory: string, document: string, find: string, replacement: string) =>
  wordIntegration.execute({ document, action: "replace", find, replace: replacement }, context(directory))

const documentXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/document.xml"]!)

test("an edited source is re-indexed, so matter search returns the new wording and not the old", async () => {
  const matter = await newMatter("Word edit source")
  await register(matter.route, "documents", "Lease.docx", await docx("1. Upkeep", "The tenant repaints the gazebo vermilion every spring."))
  const before = await search(matter.dir, "gazebo paint colour", "search-before")
  expect(before).toContain("vermilion")

  const result = await replace(matter.dir, "Lease.docx", "vermilion", "cerulean")
  expect(result).toMatchObject({ title: "Edited Lease.docx", metadata: { edits: 1 } })
  expect(documentXml(readFileSync(path.join(matter.dir, "Lease.docx")))).toContain("cerulean")

  const after = await search(matter.dir, "gazebo paint colour", "search-after")
  expect(after).toContain("cerulean")
  expect(after).not.toContain("vermilion")
  const db = openDb(matter.dir)
  try {
    expect(getDocumentRole(db, path.join(matter.dir, "Lease.docx"))).toBe("source")
  } finally { db.close() }
}, 120_000)

test("a refused or unreachable save reports the edit as not saved and leaves the file unchanged", async () => {
  const matter = await newMatter("Word edit refused")
  await register(matter.route, "documents", "Notice.docx", await docx("Notices go by registered post."))
  const registered = readFileSync(path.join(matter.dir, "Notice.docx"))

  // The ingest service cannot be reached.
  process.env.INGEST_URL = `http://127.0.0.1:${closedPort()}`
  try {
    const unreachable = await replace(matter.dir, "Notice.docx", "registered post", "courier")
    expect(unreachable).toBeString()
    expect(unreachable as string).toStartWith("Edit not saved")
    expect(unreachable as string).toEndWith("Notice.docx is unchanged.")
  } finally { process.env.INGEST_URL = base }
  expect(readFileSync(path.join(matter.dir, "Notice.docx"))).toEqual(registered)

  // A file put in the matter folder by hand was never registered, so the route refuses it.
  const loose = await docx("Loose wording stays as written.")
  writeFileSync(path.join(matter.dir, "Loose.docx"), loose)
  const refused = await replace(matter.dir, "Loose.docx", "Loose wording", "Changed wording")
  expect(refused).toBeString()
  expect(refused as string).toStartWith("Edit not saved")
  expect(refused as string).toContain("404")
  expect(refused as string).toContain("not registered")
  expect(refused as string).toEndWith("Loose.docx is unchanged.")
  expect(new Uint8Array(readFileSync(path.join(matter.dir, "Loose.docx")))).toEqual(loose)
}, 120_000)

test("an edited deliverable stays a deliverable outside source retrieval and keeps its paragraph pointers", async () => {
  const matter = await newMatter("Word edit deliverable")
  await register(matter.route, "deliverables", "Draft.docx", await docx("1. Supply", "The supplier ships the widgets within five days."))
  const file = path.join(matter.dir, "Draft.docx")
  const pointers = (await paragraphTargets(readFileSync(file))).map(target => target.pid)
  expect(pointers.length).toBeGreaterThan(0)
  expect(pointers.every(Boolean)).toBe(true)

  const result = await replace(matter.dir, "Draft.docx", "five days", "seven days")
  expect(result).toMatchObject({ title: "Edited Draft.docx", metadata: { edits: 1 } })
  const edited = readFileSync(file)
  expect(documentXml(edited)).toContain("seven days")
  expect((await paragraphTargets(edited)).map(target => target.pid)).toEqual(pointers)
  const db = openDb(matter.dir)
  try {
    expect(getDocumentRole(db, file)).toBe("deliverable")
    expect(listSourceDocuments(db)).toEqual([])
    expect(db.query("SELECT COUNT(*) AS n FROM chunks").get()).toEqual({ n: 0 })
  } finally { db.close() }
}, 120_000)

test("editing a part of a saved combined file rebuilds the combined file with the new wording", async () => {
  const matter = await newMatter("Word edit combined")
  // A completed lean run that delivered one combined document with its memo.
  const documents = [{ title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }], file: "supply-pack.docx", deliverable: "combined" as const }]
  initLeanState(matter.dir, { request: "Draft a short supply note from the supply terms.", ownerSessionID: "ses_owner", model: null })
  writeLeanPlan(matter.dir, documents, undefined, undefined, { name: "cover memo", file: memoNotesFile(documents, "cover memo") })
  const body = (await renderMarkdownToDocx("# Supply note\n\n## Delivery\n\nBody marker: the supplier delivers the goods.")).bytes
  writeLeanBytes(matter.dir, "d1/body-applied.docx", body)
  writeLeanBytes(matter.dir, "d1/redline.docx", body)
  const memo = "# Cover memo\n\nMemo marker: the start date needs confirmation."
  writeLeanBytes(matter.dir, "drafting-notes.docx", (await renderMarkdownToDocx(memo)).bytes)
  writeLeanText(matter.dir, "drafting-notes.md", `${memo}\n`)
  await recordLeanDocuments(matter.dir, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, true)
  writeLeanState(matter.dir, { ...readLeanState(matter.dir)!, status: "completed" })
  const part = readLeanDocuments(matter.dir)!.documents.find(row => row.kind === "body")!.deliverable!
  const saved = leanDraftingStatus(matter.dir)?.combined_output?.file
  expect(saved).toBe("supply-pack.docx")
  expect(documentXml(readFileSync(path.join(matter.dir, saved!)))).not.toContain("by rail")

  const result = await replace(matter.dir, part, "delivers the goods", "delivers the goods by rail")
  expect(result).toMatchObject({ title: `Edited ${part}`, metadata: { edits: 1 } })
  expect(documentXml(readFileSync(path.join(matter.dir, part)))).toContain("by rail")
  expect(documentXml(readFileSync(path.join(matter.dir, saved!)))).toContain("by rail")
}, 120_000)
