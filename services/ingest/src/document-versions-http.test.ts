import { afterAll, beforeAll, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { access, mkdtemp, mkdir, readdir, readFile, rm, writeFile } from "node:fs/promises"
import { Database } from "bun:sqlite"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, InsertedTextRun, Packer, Paragraph, TextRun } from "docx"
import { recordRedlineProposal } from "../../../dochaus/lib/redlines"

// Document versions (#268) through the real ingest routes of an isolated
// service: each existing writer adds exactly one version with its origin, the
// matter keeps one working file per document, no write loses bytes, restore
// copies a version forward and an identical upload stores nothing. Fixtures are
// neutral synthetic paragraphs built in code; no generation engine is used.

const AGREEMENT = "Services agreement.docx"
const LETTER = "Advice letter.docx"
const NOTICE_PDF = "Notice.pdf"
const NOTICE_DOCX = "Notice.docx"
const SCOPE = "The invented supplier delivers the sample goods."
const PAYMENT = "Payment is due within thirty days of invoice."
const NOTICES = "Notices are given in writing."

type Captured = { document_id: string; id: string; number: number; parent: string | null; origin: string; created: boolean }
type Version = { id: string; number: number; sha256: string; ext: string; bytes: number; parent: string | null; derived_from: string | null
  origin: string; actor: string; run: string | null; label: string | null; created_at: number; removed: string | null; current: boolean }
type VersionEvent = { id: number; kind: string; version_id: string | null; actor: string; detail: Record<string, unknown> | null; created_at: number }
type History = { name: string; document: { id: string; name: string; current_version: string } | null; versions: Version[]; events: VersionEvent[]; working_version: string | null }
type Refusal = { error: string; code?: string }

let temporary = ""
let service: ReturnType<typeof Bun.spawn> | undefined
let engine: ReturnType<typeof Bun.serve> | undefined
const engineRequests: string[] = []
let base = ""
let matter = { id: "", dir: "" }
let baseline: string[] = []
// Every working-file state the steps below produced, by document.
const states = new Map<string, Set<string>>()

const sha256 = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
const request = (route: string, options?: RequestInit) => fetch(base + route, { ...options, signal: AbortSignal.timeout(60_000) })
const matterRoute = (route: string, options?: RequestInit) => request(`/matters/${matter.id}${route}`, options)
const form = (name: string, bytes: Uint8Array, fields: Record<string, string> = {}) => {
  const body = new FormData()
  body.append("file", new File([new Uint8Array(bytes)], name))
  for (const [key, value] of Object.entries(fields)) body.append(key, value)
  return body
}
const upload = (name: string, bytes: Uint8Array, fields: Record<string, string> = {}, headers: Record<string, string> = {}) =>
  matterRoute("/documents", { method: "POST", body: form(name, bytes, fields), headers })
const revise = (name: string, bytes: Uint8Array, headers: Record<string, string> = {}) =>
  matterRoute("/documents/revisions", { method: "POST", body: form(name, bytes), headers })
const history = async (name: string) => {
  const response = await matterRoute(`/documents/versions?${new URLSearchParams({ name })}`)
  expect(response.status).toBe(200)
  return await response.json() as History
}
const versionBytes = async (id: string) => {
  const response = await matterRoute(`/documents/versions/${id}/content`)
  expect(response.status).toBe(200)
  return new Uint8Array(await response.arrayBuffer())
}
const working = async (name: string) => new Uint8Array(await readFile(path.join(matter.dir, name)))
// Records the working file as a state of its document, so the end checks can
// prove each one is still kept.
const observe = async (name: string) => {
  const bytes = await working(name)
  states.set(name, (states.get(name) ?? new Set()).add(sha256(bytes)))
  return bytes
}
// The matter's Word and PDF files; the source artifacts beside them are not documents.
const documentFiles = async () => (await readdir(matter.dir)).filter((name) => /\.(docx|pdf)$/i.test(name) && !baseline.includes(name)).sort()
const storedFiles = async () => (await readdir(path.join(matter.dir, ".dochaus", "revisions"), { recursive: true }))
  .filter((name) => !path.basename(name).startsWith(".") && path.extname(name)).sort()
// The origin each source ingest stamped in lineage, in order: one per ingest
// run (its first row carries the provenance).
const lineageOrigins = async () => {
  const rows = (await readFile(path.join(matter.dir, ".dochaus", "lineage.jsonl"), "utf8")).split("\n").filter(Boolean)
    .map((line) => JSON.parse(line) as { op: string; run: string; provenance?: { origin?: string } })
  const runs = new Map<string, string | undefined>()
  for (const row of rows) if (row.op === "ingest" && !runs.has(row.run)) runs.set(row.run, row.provenance?.origin)
  return [...runs.values()]
}

const storeOf = (documentId: string) => path.join(matter.dir, ".dochaus", "revisions", documentId)
const exists = (file: string) => access(file).then(() => true, () => false)
const status = async (response: Response) => [response.status, (await response.json() as Refusal).code] as const
const remove = (name: string) => matterRoute(`/documents?${new URLSearchParams({ name })}`, { method: "DELETE" })
const compareJSON = (query: Record<string, string>) => matterRoute(`/compare?${new URLSearchParams({ ...query, format: "json" })}`)

const docx = async (paragraphs: Paragraph[]) => new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: paragraphs }] })))
const plain = (...texts: string[]) => docx(texts.map((text) => new Paragraph(text)))

beforeAll(async () => {
  temporary = await mkdtemp(path.join(tmpdir(), "dochaus-versions-http-"))
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  await mkdir(workspace)
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  engine = Bun.serve({ port: 0, fetch(request) {
    engineRequests.push(request.url)
    return new Response("No engine in this versions test", { status: 503 })
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
    if (await request("/matters").then((response) => response.ok).catch(() => false)) break
    if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(serviceLog, "utf8")}`)
    await Bun.sleep(50)
  }
  const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" },
    body: JSON.stringify({ title: "Neutral versions matter" }) })
  expect(created.status).toBe(200)
  matter = await created.json() as { id: string; dir: string }
  baseline = (await readdir(matter.dir)).filter((name) => !name.startsWith("."))
}, 90_000)

afterAll(async () => {
  service?.kill()
  await service?.exited
  engine?.stop(true)
  if (temporary) await rm(temporary, { recursive: true, force: true })
  expect(engineRequests).toEqual([])
})

const sent = { agreement: new Uint8Array(), replaced: new Uint8Array() }
const ids = { agreementV1: "", letterV1: "", letterV2: "", letterV3: "", noticePdfV1: "" }

test("upload, generate, Word revision, replace, accept and convert each add one version with the writer's origin", async () => {
  // Upload: version 1 of the agreement.
  sent.agreement = await plain("1. Scope", SCOPE, PAYMENT, "2. Notices", NOTICES)
  const uploaded = await upload(AGREEMENT, sent.agreement, {}, { "x-dochaus-actor": "Lawyer One" })
  expect(uploaded.status).toBe(200)
  const first = (await uploaded.json() as { version: Captured }).version
  expect(first).toMatchObject({ number: 1, parent: null, origin: "upload", created: true })
  ids.agreementV1 = first.id
  expect(await observe(AGREEMENT)).toEqual(sent.agreement)

  // Deliverable write: version 1 of the letter, made by a run.
  const generated = await matterRoute("/deliverables", { method: "POST", body: form(LETTER, await plain("Advice", PAYMENT, NOTICES)),
    headers: { "x-dochaus-actor": "doc.haus", "x-dochaus-run": "ses_neutral_run" } })
  expect(generated.status).toBe(200)
  const letterV1 = (await generated.json() as { version: Captured }).version
  expect(letterV1).toMatchObject({ number: 1, parent: null, origin: "generated", created: true })
  ids.letterV1 = letterV1.id
  await observe(LETTER)

  // The lawyer's Word revision of the letter: version 2.
  const revised = await revise(LETTER, await plain("Advice, as revised", PAYMENT, NOTICES), { "x-dochaus-actor": "Lawyer One" })
  expect(revised.status).toBe(200)
  const letterV2 = (await revised.json() as { version: Captured }).version
  expect(letterV2).toMatchObject({ number: 2, parent: ids.letterV1, origin: "revision", created: true })
  ids.letterV2 = letterV2.id
  await observe(LETTER)

  // Replace in the upload flow saves the agreement's next version and keeps the bytes it replaced.
  sent.replaced = await plain("1. Scope", SCOPE, "Payment is due within forty days of invoice.", "2. Notices", NOTICES)
  const replaced = await upload(AGREEMENT, sent.replaced, { existing: "replace", version_label: "Client return" })
  expect(replaced.status).toBe(200)
  expect((await replaced.json() as { version: Captured }).version).toMatchObject({ number: 2, parent: ids.agreementV1, origin: "replace", created: true })
  expect(await observe(AGREEMENT)).toEqual(sent.replaced)
  expect(await versionBytes(ids.agreementV1)).toEqual(sent.agreement)

  // Accepting a proposal on the letter: version 3, with its accept event.
  const { id: proposal } = recordRedlineProposal(matter.dir, {
    docPath: path.join(matter.dir, LETTER), docName: LETTER, scope: "phrase", findText: "thirty days", oldText: "thirty days",
    newText: "forty-five days", author: "Test reviewer", reviewComment: "Review this proposed wording.", anchorId: "payment",
  })
  const accepted = await matterRoute(`/redlines/${proposal}/accept`, { method: "POST", headers: { "x-dochaus-actor": "Lawyer One" } })
  expect(accepted.status).toBe(200)
  const letterV3 = (await history(LETTER)).versions[0]!
  expect(letterV3).toMatchObject({ number: 3, parent: ids.letterV2, origin: "accept", current: true })
  ids.letterV3 = letterV3.id
  await observe(LETTER)

  // A reject writes nothing: an event against the version the file holds.
  const { id: declined } = recordRedlineProposal(matter.dir, {
    docPath: path.join(matter.dir, LETTER), docName: LETTER, scope: "phrase", findText: "in writing", oldText: "in writing",
    newText: "in writing or by email", author: "Test reviewer", reviewComment: "Review this proposed wording.", anchorId: "notices",
  })
  expect((await matterRoute(`/redlines/${declined}/reject`, { method: "POST" })).status).toBe(200)

  // PDF upload, then conversion: the Word document's version 1 is linked to the PDF's.
  const notice = pdf("The invented notice is given in writing.")
  const pdfUpload = await upload(NOTICE_PDF, notice)
  expect(pdfUpload.status).toBe(200)
  ids.noticePdfV1 = (await pdfUpload.json() as { version: Captured }).version.id
  await observe(NOTICE_PDF)
  const converted = await matterRoute(`/documents/convert?${new URLSearchParams({ name: NOTICE_PDF })}`, { method: "POST" })
  expect(converted.status).toBe(200)
  expect((await converted.json() as { version: Captured }).version).toMatchObject({ number: 1, parent: null, origin: "conversion", created: true })
  await observe(NOTICE_DOCX)

  // One working file per document, under its stable name.
  expect(await documentFiles()).toEqual([LETTER, NOTICE_DOCX, NOTICE_PDF, AGREEMENT].sort())

  const agreement = await history(AGREEMENT)
  expect(agreement.versions.map((row) => [row.number, row.origin, row.actor, row.label, row.current]))
    .toEqual([[2, "replace", "unknown", "Client return", true], [1, "upload", "Lawyer One", null, false]])
  expect(agreement.working_version).toBe(agreement.versions[0]!.id)
  const letter = await history(LETTER)
  expect(letter.versions.map((row) => [row.number, row.origin, row.actor, row.run, row.parent]))
    .toEqual([[3, "accept", "Lawyer One", null, ids.letterV2], [2, "revision", "Lawyer One", null, ids.letterV1], [1, "generated", "doc.haus", "ses_neutral_run", null]])
  expect(letter.events.map((row) => [row.kind, row.version_id, row.detail])).toEqual([
    ["accept", ids.letterV3, { redline_ids: [proposal] }],
    ["reject", ids.letterV3, { redline_ids: [declined] }],
  ])
  const converted_ = await history(NOTICE_DOCX)
  expect(converted_.versions.map((row) => [row.number, row.origin, row.derived_from])).toEqual([[1, "conversion", ids.noticePdfV1]])
  expect((await history(NOTICE_PDF)).versions.map((row) => [row.number, row.origin])).toEqual([[1, "upload"]])

  // Lineage names what made each ingest, not "upload" for all of them.
  expect(await lineageOrigins()).toEqual(["upload", "replace", "upload", "conversion"])
}, 120_000)

test("an agent's edit and a lawyer's save with tracked changes are kept as versions, the latter with a notice", async () => {
  const edited = await revise(AGREEMENT, await plain("1. Scope", SCOPE, "Payment is due within fifty days of invoice.", "2. Notices", NOTICES), {
    "x-dochaus-origin": "agent-edit", "x-dochaus-actor": "agent:drafter", "x-dochaus-run": "ses_agent",
    "x-document-prior-sha256": sha256(await working(AGREEMENT)),
  })
  expect(edited.status).toBe(200)
  expect((await edited.json() as { version: Captured }).version).toMatchObject({ number: 3, origin: "agent-edit" })
  await observe(AGREEMENT)
  expect((await lineageOrigins()).at(-1)).toBe("agent-edit")
  expect((await history(AGREEMENT)).versions[0]).toMatchObject({ number: 3, origin: "agent-edit", actor: "agent:drafter", run: "ses_agent" })

  const tracked = await docx([
    new Paragraph("1. Scope"), new Paragraph(SCOPE), new Paragraph("Payment is due within fifty days of invoice."), new Paragraph("2. Notices"),
    new Paragraph({ children: [new TextRun(NOTICES), new InsertedTextRun({ text: " Email counts as writing.", id: 1, author: "Pat Reviewer", date: "2026-10-15T09:00:00Z" })] }),
  ])
  const saved = await revise(AGREEMENT, tracked)
  expect(saved.status).toBe(200)
  expect(await saved.json()).toMatchObject({
    version: { number: 4, origin: "revision" },
    tracked_changes: { count: 1, authors: ["Pat Reviewer"] },
    notice: "Saved as version 4, with 1 tracked change by Pat Reviewer; resolve them in Word before doc.haus edits it.",
  })
  expect(await observe(AGREEMENT)).toEqual(tracked)

  // A writer may state only a revision origin.
  const refused = await revise(AGREEMENT, await plain("Unrelated"), { "x-dochaus-origin": "restore" })
  expect(refused.status).toBe(400)
  expect((await history(AGREEMENT)).versions).toHaveLength(4)
})

test("restore copies a version forward byte for byte as a new version, and history stays as it was", async () => {
  const before = await history(LETTER)
  const original = await versionBytes(ids.letterV1)
  const restored = await matterRoute(`/documents/versions/${ids.letterV1}/restore`, { method: "POST", headers: { "x-dochaus-actor": "Lawyer One" } })
  expect(restored.status).toBe(200)
  const body = await restored.json() as { version: Captured; restored_from: { id: string; number: number } }
  expect(body.version).toMatchObject({ number: 4, parent: ids.letterV3, origin: "restore", created: true })
  expect(body.restored_from).toEqual({ id: ids.letterV1, number: 1 })
  expect(await observe(LETTER)).toEqual(original)
  expect(await versionBytes(body.version.id)).toEqual(original)

  const after = await history(LETTER)
  expect(after.versions.slice(1)).toEqual(before.versions.map((row) => ({ ...row, current: false })))
  expect(after.versions[0]).toMatchObject({ number: 4, origin: "restore", actor: "Lawyer One", label: "Restored from version 1", current: true })
  expect(after.events.at(-1)).toMatchObject({ kind: "restore", version_id: body.version.id, detail: { restored_from: ids.letterV1, restored_number: 1 } })
  expect(after.working_version).toBe(body.version.id)

  // The working file already holds those bytes: nothing to restore.
  const again = await matterRoute(`/documents/versions/${ids.letterV1}/restore`, { method: "POST" })
  expect(again.status).toBe(409)
  expect((await again.json() as { code: string }).code).toBe("already_current")
  expect((await matterRoute("/documents/versions/not-a-version/restore", { method: "POST" })).status).toBe(404)
  expect((await matterRoute("/documents/versions/not-a-version/content")).status).toBe(404)
})

test("an identical upload is reported as a duplicate and stores nothing", async () => {
  const files = await documentFiles()
  const stored = await storedFiles()
  const counts = async () => Promise.all([AGREEMENT, LETTER, NOTICE_PDF, NOTICE_DOCX].map(async (name) => (await history(name)).versions.length))
  const before = await counts()

  // An earlier version, under its own name, even when replacing.
  const earlier = await upload(AGREEMENT, sent.agreement, { existing: "replace" })
  expect(earlier.status).toBe(409)
  expect(await earlier.json()).toMatchObject({ code: "duplicate", name: AGREEMENT, version: { id: ids.agreementV1, number: 1, current: false } })
  // The current bytes, without a choice or choosing keep both.
  const current = await working(NOTICE_PDF)
  for (const fields of [{}, { existing: "keep-both" }] as Record<string, string>[]) {
    const response = await upload(NOTICE_PDF, current, fields)
    expect(response.status).toBe(409)
    expect(await response.json()).toMatchObject({ code: "duplicate", name: NOTICE_PDF, version: { id: ids.noticePdfV1, number: 1, current: true } })
  }

  expect(await documentFiles()).toEqual(files)
  expect(await storedFiles()).toEqual(stored)
  expect(await counts()).toEqual(before)

  // The same bytes under another name are another document, uploaded as before.
  const renamed = await upload("Another notice.pdf", current)
  expect(renamed.status).toBe(200)
  expect((await renamed.json() as { version: Captured }).version).toMatchObject({ number: 1, origin: "upload" })
})

test("compare takes version ids, with the version before the right side as the default left", async () => {
  type Side = { name: string; label: string; version?: { id: string; number: number } }
  const compare = async (query: Record<string, string>) => matterRoute(`/compare?${new URLSearchParams({ ...query, format: "json" })}`)
  const sides = async (query: Record<string, string>) => {
    const response = await compare(query)
    expect(response.status).toBe(200)
    const body = await response.json() as { left: Side; right: Side; blacklineName: string }
    return [body.left, body.right, body.blacklineName] as const
  }
  const [left, right, blackline] = await sides({ rightVersion: ids.letterV2 })
  expect([left.version, right.version]).toEqual([{ id: ids.letterV1, number: 1 }, { id: ids.letterV2, number: 2 }])
  expect([left.label, right.label, blackline]).toEqual(["Advice letter (version 1)", "Advice letter (version 2)", "Blackline - Advice letter (version 1) vs Advice letter (version 2).docx"])
  // The working file holds version 4 (the restore), so its previous version is 3.
  const [previous, named] = await sides({ right: LETTER, leftVersion: "previous" })
  expect([previous.version?.number, named.name, named.version]).toEqual([3, LETTER, undefined])
  // A version against another document's working file.
  const [other] = await sides({ leftVersion: ids.agreementV1, right: LETTER })
  expect([other.name, other.version?.number]).toEqual([AGREEMENT, 1])

  const refusal = async (query: Record<string, string>) => {
    const response = await compare(query)
    return [response.status, (await response.json() as { error: string }).error]
  }
  expect(await refusal({ rightVersion: ids.agreementV1 })).toEqual([400, "no_previous_version"])
  expect(await refusal({ leftVersion: ids.letterV2, rightVersion: ids.letterV2 })).toEqual([400, "same_version"])
  expect(await refusal({ rightVersion: "not-a-version" })).toEqual([404, "not_found"])
})

test("an outside edit is kept as a version of origin unknown before a replace, even one with the current version's bytes", async () => {
  const name = "Outside edit.docx"
  const original = await plain("Outside edit", SCOPE)
  expect((await upload(name, original)).status).toBe(200)
  await observe(name)
  const outside = await plain("Outside edit", SCOPE, "A clause added outside doc.haus.")
  await writeFile(path.join(matter.dir, name), outside)
  await observe(name)
  // Replacing it with its current version's bytes is not a duplicate: the working file holds other bytes.
  const replaced = await upload(name, original, { existing: "replace" })
  expect(replaced.status).toBe(200)
  expect(await observe(name)).toEqual(original)
  const { versions } = await history(name)
  expect(versions.map((row) => [row.number, row.origin, row.current])).toEqual([[3, "replace", true], [2, "unknown", false], [1, "upload", false]])
  expect(await versionBytes(versions[1]!.id)).toEqual(outside)
  expect(await versionBytes(versions[0]!.id)).toEqual(original)
})

test("a version whose kept bytes are damaged is refused, and uploading those bytes again stores them afresh", async () => {
  const name = "Damaged.docx"
  const original = await plain("Damaged", SCOPE)
  const first = (await (await upload(name, original)).json() as { version: Captured }).version
  expect((await upload(name, await plain("Damaged", PAYMENT), { existing: "replace" })).status).toBe(200)
  const kept = path.join(storeOf(first.document_id), `${sha256(original)}.docx`)
  await writeFile(kept, "damaged")
  expect(await status(await matterRoute(`/documents/versions/${first.id}/content`))).toEqual([409, "version_damaged"])
  expect(await status(await matterRoute(`/documents/versions/${first.id}/restore`, { method: "POST" }))).toEqual([409, "version_damaged"])
  const compared = await compareJSON({ leftVersion: first.id, right: name })
  expect([compared.status, (await compared.json() as Refusal).error]).toEqual([409, "version_damaged"])
  // Not a duplicate of the damaged version: the upload is the next version and stores the bytes again.
  const again = await upload(name, original, { existing: "replace" })
  expect(again.status).toBe(200)
  expect((await again.json() as { version: Captured }).version).toMatchObject({ number: 3, origin: "replace", created: true })
  expect(await versionBytes(first.id)).toEqual(original)
  // The current version's bytes damaged: uploading the working file's bytes repairs them and adds no version.
  await writeFile(kept, "damaged")
  const repaired = await upload(name, original, { existing: "replace" })
  expect(repaired.status).toBe(200)
  expect((await repaired.json() as { version: Captured }).version).toMatchObject({ number: 3, created: false })
  expect(await versionBytes(first.id)).toEqual(original)
  await observe(name)
})

test("a working file deleted outside doc.haus keeps its history, and a restore writes it back", async () => {
  const name = "Missing.docx"
  const original = await plain("Missing", SCOPE)
  const first = (await (await upload(name, original)).json() as { version: Captured }).version
  expect((await upload(name, await plain("Missing", PAYMENT), { existing: "replace" })).status).toBe(200)
  await rm(path.join(matter.dir, name))
  const listed = await history(name)
  expect([listed.versions.map((row) => row.number), listed.working_version]).toEqual([[2, 1], null])
  const restored = await matterRoute(`/documents/versions/${first.id}/restore`, { method: "POST" })
  expect(restored.status).toBe(200)
  expect((await restored.json() as { version: Captured }).version).toMatchObject({ number: 3, origin: "restore" })
  expect(await observe(name)).toEqual(original)
})

test("a redaction deletes the kept bytes of every earlier version, which stay listed as removed by redaction", async () => {
  const secret = "The invented account number is 0000-1111."
  for (const [name, role] of [["Redacted source.docx", "source"], ["Redacted letter.docx", "deliverable"]] as const) {
    const first = await plain("Redaction", secret, SCOPE)
    const second = await plain("Redaction", secret, PAYMENT)
    const created = role === "source" ? await upload(name, first) : await matterRoute("/deliverables", { method: "POST", body: form(name, first) })
    const { document_id: documentId } = (await created.json() as { version: Captured }).version
    expect((role === "source" ? await upload(name, second, { existing: "replace" }) : await revise(name, second)).status).toBe(200)
    // As the redaction tool saves: the file in place, then its revision.
    const redacted = await plain("Redaction", "[REDACTED]", PAYMENT)
    await writeFile(path.join(matter.dir, name), redacted)
    const saved = await revise(name, redacted, { "x-dochaus-origin": "redaction", "x-dochaus-actor": "agent:redactor" })
    expect(saved.status).toBe(200)
    const body = await saved.json() as { version: Captured; purge_failed?: string }
    expect([body.version.number, body.version.origin, body.purge_failed]).toEqual([3, "redaction", undefined])
    const { versions } = await history(name)
    expect(versions.map((row) => [row.number, row.removed])).toEqual([[3, null], [2, "redaction"], [1, "redaction"]])
    // Only the redacted bytes are kept (a deliverable's with the paragraph pointers its revision adds).
    const kept = await working(name)
    expect(await readdir(storeOf(documentId))).toEqual([`${sha256(kept)}.docx`])
    expect(await versionBytes(versions[0]!.id)).toEqual(kept)
    // Listed, but not opened, restored or compared.
    const [, older, oldest] = versions
    expect(await status(await matterRoute(`/documents/versions/${oldest!.id}/content`))).toEqual([410, "removed_by_redaction"])
    expect(await status(await matterRoute(`/documents/versions/${oldest!.id}/restore`, { method: "POST" }))).toEqual([410, "removed_by_redaction"])
    for (const query of [{ rightVersion: versions[0]!.id }, { leftVersion: older!.id, right: name }, { right: name, leftVersion: "previous" }] as Record<string, string>[]) {
      const compared = await compareJSON(query)
      expect([compared.status, (await compared.json() as Refusal).error]).toEqual([410, "removed_by_redaction"])
    }
    // The text is in no kept file of the matter's version store.
    for (const file of await storedFiles())
      expect(Buffer.from(await readFile(path.join(matter.dir, ".dochaus", "revisions", file))).includes(Buffer.from("0000-1111"))).toBe(false)
    await observe(name)
  }
})

test("deleting a document deletes its kept versions, and a later file of that name starts a new document", async () => {
  const name = "Removed.docx"
  const original = await plain("Removed", SCOPE)
  const first = (await (await upload(name, original)).json() as { version: Captured }).version
  expect((await upload(name, await plain("Removed", PAYMENT), { existing: "replace" })).status).toBe(200)
  expect(await exists(storeOf(first.document_id))).toBe(true)
  expect((await remove(name)).status).toBe(200)
  expect(await exists(storeOf(first.document_id))).toBe(false)
  // Its versions are no longer in the matter.
  expect((await matterRoute(`/documents/versions/${first.id}/content`)).status).toBe(404)
  expect((await compareJSON({ leftVersion: first.id, right: LETTER })).status).toBe(404)
  expect(await status(await matterRoute(`/documents/versions/${first.id}/restore`, { method: "POST" }))).toEqual([409, "document_removed"])
  const again = await upload(name, original)
  expect(again.status).toBe(200)
  const next = (await again.json() as { version: Captured }).version
  expect([next.number, next.origin, next.document_id === first.document_id]).toEqual([1, "upload", false])
  expect((await history(name)).versions.map((row) => row.number)).toEqual([1])
  await observe(name)
})

test("an email upload keeps its uploader and label, and one whose attachments are not all in is taken again", async () => {
  const name = "Thread.eml"
  const message = email("Synthetic thread", "Draft attached.", [part("Thread draft.docx", Buffer.from(await plain("Thread draft", SCOPE)))])
  const before = (await lineageOrigins()).length
  const uploaded = await upload(name, message, { version_label: "From the client" }, { "x-dochaus-actor": "Lawyer One" })
  expect(uploaded.status).toBe(200)
  const { versions } = await history(name)
  expect(versions.map((row) => [row.number, row.origin, row.actor, row.label])).toEqual([[1, "upload", "Lawyer One", "From the client"]])
  // The email, its attachment, then doc.haus's own commit of the attachment's outcome.
  expect((await lineageOrigins()).slice(before)).toEqual(["upload", "upload", "system"])
  // A kept email is served as the content route serves its working file.
  const kept = await matterRoute(`/documents/versions/${versions[0]!.id}/content`)
  const current = await matterRoute(`/documents/content?${new URLSearchParams({ name })}`)
  expect([kept.status, kept.headers.get("content-type")]).toEqual([200, current.headers.get("content-type")])
  expect(kept.headers.get("content-type")).not.toContain("rfc822")
  // With its attachments in, the same email again is a duplicate.
  expect(await status(await upload(name, message, { existing: "replace" }))).toEqual([409, "duplicate"])
  // As if the upload had stopped before recording the outcome: uploading it again brings the attachment in.
  const db = new Database(path.join(matter.dir, ".dochaus", "legal.db"))
  try {
    const row = db.query("SELECT attachments FROM documents WHERE name = ?").get(name) as { attachments: string }
    const record = JSON.parse(row.attachments) as { attachments: { outcome: string }[] }
    record.attachments[0]!.outcome = "pending"
    db.run("UPDATE documents SET attachments = ? WHERE name = ?", [JSON.stringify(record), name])
  } finally {
    db.close()
  }
  const again = await upload(name, message, { existing: "replace" })
  expect(again.status).toBe(200)
  expect((await again.json() as { attachments: unknown[] }).attachments).toMatchObject([{ status: "indexed", document: "Thread draft.docx" }])
  expect((await history(name)).versions).toHaveLength(1)
})

test("every working-file state is kept as a version whose bytes match its hash", async () => {
  for (const [name, seen] of states) {
    const { versions } = await history(name)
    expect([...seen].filter((hash) => !versions.some((row) => row.sha256 === hash))).toEqual([])
    // A version a redaction removed has no bytes (above); every other one has its own.
    for (const row of versions) if (!row.removed) expect(sha256(await versionBytes(row.id))).toBe(row.sha256)
  }
})

// A neutral email with base64 Word attachments.
function part(name: string, bytes: Buffer) {
  const type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  return [`Content-Type: ${type}; name="${name}"`, `Content-Disposition: attachment; filename="${name}"`, "Content-Transfer-Encoding: base64", "", bytes.toString("base64")].join("\r\n")
}
function email(subject: string, body: string, parts: string[]) {
  const boundary = "B-versions"
  return new Uint8Array(Buffer.from([
    "From: Sender <sender@example.test>", "To: Recipient <recipient@example.test>", `Subject: ${subject}`,
    "Date: Mon, 01 Jan 2024 00:00:00 +0000", "MIME-Version: 1.0", `Content-Type: multipart/mixed; boundary="${boundary}"`, "",
    `--${boundary}`, "Content-Type: text/plain; charset=utf-8", "", body,
    ...parts.flatMap((raw) => [`--${boundary}`, raw]), `--${boundary}--`, "",
  ].join("\r\n")))
}

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
  text += `xref\n0 6\n0000000000 65535 f \n${offsets.map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return new Uint8Array(Buffer.from(text))
}
