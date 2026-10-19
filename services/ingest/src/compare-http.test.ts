import { afterAll, beforeAll, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { existsSync } from "node:fs"
import { mkdtemp, mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, InsertedTextRun, Packer, Paragraph, TextRun } from "docx"
import { strFromU8, unzipSync } from "fflate"
import { recordRedlineProposal } from "../../../dochaus/lib/redlines"

// GET /matters/:id/compare against an isolated ingest service. Fixtures are
// neutral synthetic paragraphs built in code. Comparing needs no generation
// engine; any engine request fails the suite.

const SENT = "Sent draft.docx"
const RETURNED = "Returned draft.docx"
const TRACKED = "Tracked return.docx"
const DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
const SCOPE = "The invented supplier delivers the sample goods."
const PAYMENT = "Payment is due within thirty days of invoice."
const NOTICES = "Notices are given in writing."
const MEETING = "The parties meet once each quarter."

type Hunk = { id: string; kind: string; location: { scope: string; paragraph: number; side: string }; before?: string; after?: string; untracked: boolean }
type Comparison = {
  left: { name: string; label: string; trackedChanges: { count: number; authors: string[] }; pendingProposals: number }
  right: { name: string; label: string; trackedChanges: { count: number; authors: string[] }; pendingProposals: number }
  options: { ignoreFormatting: boolean; detectMoves: boolean; headersFooters: boolean }
  attribution: string
  changes: Hunk[]
  untracked: Hunk[]
  blacklineName: string
}

let temporary = ""
let workspace = ""
let service: ReturnType<typeof Bun.spawn> | undefined
let engine: ReturnType<typeof Bun.serve> | undefined
const engineRequests: string[] = []
let base = ""
let matter = { id: "", dir: "" }

const request = (route: string, options?: RequestInit) => fetch(base + route, { ...options, signal: AbortSignal.timeout(60_000) })
const compare = (query: Record<string, string>) => request(`/matters/${matter.id}/compare?${new URLSearchParams(query)}`)
const failure = async (response: Response) => ({ status: response.status, body: await response.json() as { error: string; message: string } })

const docx = async (paragraphs: Paragraph[]) => new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: paragraphs }] })))
const plain = (...texts: string[]) => docx(texts.map(text => new Paragraph(text)))

beforeAll(async () => {
  temporary = await mkdtemp(path.join(tmpdir(), "dochaus-compare-http-"))
  workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  await mkdir(workspace)
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  engine = Bun.serve({ port: 0, fetch(request) {
    engineRequests.push(request.url)
    return new Response("No engine in this compare test", { status: 503 })
  } })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const serviceLog = path.join(temporary, "ingest.log")
  service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config,
      INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}` },
    stdout: Bun.file(serviceLog), stderr: Bun.file(serviceLog),
  })
  base = `http://127.0.0.1:${port}`
  const deadline = Date.now() + 20_000
  while (true) {
    if (await request("/matters").then(response => response.ok).catch(() => false)) break
    if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(serviceLog, "utf8")}`)
    await Bun.sleep(50)
  }
  const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify({ title: "Version compare" }) })
  expect(created.status).toBe(200)
  matter = await created.json() as { id: string; dir: string }
  const upload = async (name: string, bytes: Uint8Array<ArrayBuffer>) => {
    const form = new FormData()
    form.append("file", new File([bytes], name))
    expect((await request(`/matters/${matter.id}/documents`, { method: "POST", body: form })).status).toBe(200)
  }
  // Right deletes the payment paragraph and adds a meeting paragraph, with no tracked changes.
  await upload(SENT, await plain("1. Scope", SCOPE, PAYMENT, "2. Notices", NOTICES))
  await upload(RETURNED, await plain("1. Scope", SCOPE, "2. Notices", NOTICES, MEETING))
  // A return that tracks one insertion by a fictional reviewer and makes one edit untracked.
  await upload(TRACKED, await docx([
    new Paragraph("1. Scope"), new Paragraph(SCOPE), new Paragraph("Payment is due within sixty days of invoice."),
    new Paragraph("2. Notices"),
    new Paragraph({ children: [new TextRun(NOTICES), new InsertedTextRun({ text: " Email counts as writing.", id: 1, author: "Pat Reviewer", date: "2026-10-15T09:00:00Z" })] }),
  ]))
  await upload("Notice.pdf", pdf("The invented notice is given in writing."))
}, 90_000)

afterAll(async () => {
  service?.kill()
  await service?.exited
  engine?.stop(true)
  if (temporary) await rm(temporary, { recursive: true, force: true })
  expect(engineRequests).toEqual([])
})

test("a compare names both documents or refuses with missing_name", async () => {
  const queries: Record<string, string>[] = [{ right: RETURNED }, { left: SENT }, { left: "", right: RETURNED }, { left: "  ", right: RETURNED }]
  for (const query of queries) {
    expect(await failure(await compare(query))).toEqual({ status: 400, body: { error: "missing_name", message: expect.any(String) } })
  }
})

test("an unknown document, an unsupported name or a path outside the matter is not_found", async () => {
  // A real Word file beside the matter folder must stay out of reach.
  await writeFile(path.join(workspace, "outside.docx"), await plain("Outside the matter."))
  for (const left of ["Missing.docx", "../outside.docx", `../${path.basename(matter.dir)}/${SENT}`, "/etc/hosts.docx", "notes.txt"]) {
    expect(await failure(await compare({ left, right: RETURNED }))).toEqual({ status: 404, body: { error: "not_found", message: expect.any(String) } })
    expect(await failure(await compare({ left: SENT, right: left }))).toEqual({ status: 404, body: { error: "not_found", message: expect.any(String) } })
  }
  expect(await failure(await request(`/matters/no-such-matter/compare?left=${encodeURIComponent(SENT)}&right=${encodeURIComponent(RETURNED)}`)))
    .toEqual({ status: 404, body: { error: "not_found", message: expect.any(String) } })
  expect(existsSync(path.join(workspace, "no-such-matter"))).toBe(false)
})

test("a PDF on either side is not_docx and points to Convert to DOCX", async () => {
  for (const query of [{ left: "Notice.pdf", right: RETURNED }, { left: SENT, right: "Notice.pdf" }]) {
    const refused = await failure(await compare(query))
    expect(refused).toEqual({ status: 400, body: { error: "not_docx", message: expect.stringContaining("Convert to DOCX") } })
    expect(refused.body.message).toContain("Notice.pdf")
  }
})

test("the same document on both sides is same_document, whatever the spelling of its name", async () => {
  expect(await failure(await compare({ left: SENT, right: SENT }))).toEqual({ status: 400, body: { error: "same_document", message: expect.any(String) } })
  // Names match case-insensitively where the filesystem does.
  const other = await failure(await compare({ left: SENT, right: SENT.toUpperCase() }))
  if (process.platform === "darwin" || process.platform === "win32")
    expect(other).toEqual({ status: 400, body: { error: "same_document", message: expect.any(String) } })
  else expect(other).toEqual({ status: 404, body: { error: "not_found", message: expect.any(String) } })
})

test("format=json reports the paragraph changes, the untracked edits and both inputs", async () => {
  const response = await compare({ left: SENT, right: RETURNED, format: "json" })
  expect(response.status).toBe(200)
  const result = await response.json() as Comparison
  expect(result.left).toEqual({ name: SENT, label: "Sent draft", trackedChanges: { count: 0, authors: [] }, pendingProposals: 0 })
  expect(result.right).toEqual({ name: RETURNED, label: "Returned draft", trackedChanges: { count: 0, authors: [] }, pendingProposals: 0 })
  expect(result.options).toEqual({ ignoreFormatting: true, detectMoves: true, headersFooters: true })
  expect(result.blacklineName).toBe("Blackline - Sent draft vs Returned draft.docx")
  expect(["authors", "label"]).toContain(result.attribution)
  const essence = (hunks: Hunk[]) => hunks.map(({ kind, before, after }) => ({ kind, before, after }))
  expect(essence(result.changes)).toEqual([
    { kind: "deleted", before: PAYMENT, after: undefined },
    { kind: "inserted", before: undefined, after: MEETING },
  ])
  expect(result.changes.map(hunk => hunk.id)).toEqual(["C1", "C2"])
  expect(result.changes.every(hunk => hunk.untracked)).toBe(true)
  // With no tracked changes in right, every difference was made without tracking.
  expect(result.untracked.map(({ id: _id, ...hunk }) => hunk)).toEqual(result.changes.map(({ id: _id, ...hunk }) => hunk))
  expect(result.untracked.map(hunk => hunk.id)).toEqual(["U1", "U2"])
})

test("a return with tracked changes discloses them and lists only the untracked edit", async () => {
  const result = await (await compare({ left: SENT, right: TRACKED, format: "json" })).json() as Comparison
  expect(result.right.trackedChanges).toEqual({ count: 1, authors: ["Pat Reviewer"] })
  expect(result.untracked.map(({ kind, before, after }) => ({ kind, before, after }))).toEqual([
    { kind: "modified", before: PAYMENT, after: "Payment is due within sixty days of invoice." },
  ])
  expect(result.changes.map(hunk => hunk.after)).toContain(`${NOTICES} Email counts as writing.`)
})

test("format=docx (the default) returns the Word blackline inline, or as an attachment with download=1", async () => {
  const queries: Record<string, string>[] = [{ left: SENT, right: RETURNED }, { left: SENT, right: RETURNED, format: "docx" }]
  for (const query of queries) {
    const response = await compare(query)
    expect(response.status).toBe(200)
    expect(response.headers.get("content-type")).toBe(DOCX_TYPE)
    expect(response.headers.get("content-disposition")).toStartWith('inline; filename="Blackline - Sent draft vs Returned draft.docx"')
    const document = strFromU8(unzipSync(new Uint8Array(await response.arrayBuffer()))["word/document.xml"]!)
    expect(document).toContain("<w:ins ")
    expect(document).toContain("<w:del ")
    expect(document).toContain(MEETING)
  }
  const download = await compare({ left: SENT, right: RETURNED, download: "1" })
  expect(download.status).toBe(200)
  expect(download.headers.get("content-disposition")).toBe(
    `attachment; filename="Blackline - Sent draft vs Returned draft.docx"; filename*=UTF-8''${encodeURIComponent("Blackline - Sent draft vs Returned draft.docx")}`)
})

test("labels and turned-off options are reflected; the blackline keeps the file-stem name", async () => {
  const query = { left: SENT, right: RETURNED, leftLabel: "Our draft", rightLabel: "Counterparty", ignoreFormatting: "0", moves: "0", headers: "0" }
  const result = await (await compare({ ...query, format: "json" })).json() as Comparison
  expect(result.options).toEqual({ ignoreFormatting: false, detectMoves: false, headersFooters: false })
  expect([result.left.label, result.right.label]).toEqual(["Our draft", "Counterparty"])
  expect(result.blacklineName).toBe("Blackline - Sent draft vs Returned draft.docx")
  const document = strFromU8(unzipSync(new Uint8Array(await (await compare(query)).arrayBuffer()))["word/document.xml"]!)
  expect(document).toContain("Counterparty")
  expect(document).not.toContain('w:author="Returned draft"')
})

test("a compare writes nothing to the matter and reports pending doc.haus proposals without applying them", async () => {
  // Every file and its bytes, except SQLite's shared-memory index, which any
  // reader of the registry rewrites; the database and its log must not change.
  const snapshot = async () => {
    const files = await readdir(matter.dir, { recursive: true, withFileTypes: true })
    const rows = await Promise.all(files.filter(entry => entry.isFile()).map(async entry => {
      const relative = path.relative(matter.dir, path.join(entry.parentPath, entry.name))
      if (relative.endsWith("-shm")) return relative
      return `${relative} ${createHash("sha256").update(await readFile(path.join(matter.dir, relative))).digest("hex")}`
    }))
    return rows.toSorted()
  }
  recordRedlineProposal(matter.dir, {
    docPath: path.join(matter.dir, SENT), docName: SENT, scope: "phrase", findText: "thirty days", oldText: "thirty days",
    newText: "forty-five days", author: "Test reviewer", reviewComment: "Review this proposed wording.", anchorId: "payment",
  })
  const before = await snapshot()
  // A pair no earlier test compared, so any stored output would be a new file.
  const result = await (await compare({ left: TRACKED, right: SENT, format: "json" })).json() as Comparison
  expect([result.left.pendingProposals, result.right.pendingProposals]).toEqual([0, 1])
  // The stored bytes are compared, not the proposal.
  expect(result.changes.map(hunk => hunk.after)).toContain(PAYMENT)
  expect(JSON.stringify(result)).not.toContain("forty-five")
  expect((await compare({ left: TRACKED, right: SENT })).status).toBe(200)
  expect((await compare({ left: TRACKED, right: SENT, download: "1" })).status).toBe(200)
  expect(await snapshot()).toEqual(before)
})

test("a file that is not a readable Word document fails as compare_failed, not a server error", async () => {
  await writeFile(path.join(matter.dir, "Broken.docx"), "not a word document")
  expect(await failure(await compare({ left: SENT, right: "Broken.docx", format: "json" })))
    .toEqual({ status: 422, body: { error: "compare_failed", message: expect.any(String) } })
})

// A one-page PDF with one line of text.
function pdf(wording: string) {
  const stream = `BT /F1 12 Tf 40 700 Td (${wording}) Tj ET`
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ]
  let text = "%PDF-1.4\n"
  const offsets = objects.map((object, index) => {
    const offset = text.length
    text += `${index + 1} 0 obj\n${object}\nendobj\n`
    return offset
  })
  const xref = text.length
  text += `xref\n0 6\n0000000000 65535 f \n${offsets.map(offset => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return new Uint8Array(Buffer.from(text))
}
