import { afterAll, beforeAll, expect, test } from "bun:test"
import { Hono } from "hono"
import { createHash } from "node:crypto"
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
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
  const preview = await request("source", { sourceId: selected.id, expectedHash: selected.hash })
  expect(preview.status).toBe(200)
  expect(await preview.json() as DeskSourcePreview).toMatchObject({ ...selected, text: expect.stringContaining("Neutral source confirms Monday delivery."), provenance: "saved-version", truncated: false })
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

test("retained-version comparison stays within the linked document and refuses redaction, stale clients and revoked grants", async () => {
  const base = readCurrentVersion(directory, "Draft.docx")!
  await writeDeliverableRevision(directory, "Draft.docx", await bytes("The delivery date is Tuesday."))
  const current = readCurrentVersion(directory, "Draft.docx")!
  const history = await (await request("versions")).json() as DeskVersions
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
  for (const action of ["sources", "source", "versions", "compare"]) expect((await request(action)).status).toBe(401)
}, 30_000)

test("HTTPS bridge only forwards the exact scoped desk command paths", async () => {
  const forwarded: string[] = []
  const handler = wordHttpsHandler(request => { forwarded.push(new URL(request.url).pathname); return new Response("scoped") },
    { origin, external: true, host: "127.0.0.1", port: 47612 })
  for (const action of ["sources", "source", "versions", "compare"]) {
    const route = `/word/desk/${action}`
    expect(wordTokenPath(route)).toBe(true)
    expect((await handler(new Request(origin + route, { method: "POST", body: "{}" }))).status).toBe(200)
    expect((await handler(new Request(origin + route))).status).toBe(405)
  }
  for (const route of ["/word/desk/source/extra", "/word/desk/other", "/matters/synthetic/documents/content"]) {
    expect(wordTokenPath(route)).toBe(false)
    expect((await handler(new Request(origin + route, { method: "POST" }))).status).toBe(404)
  }
  expect(forwarded).toHaveLength(4)
})
