import { expect, spyOn, test } from "bun:test"
import { Hono } from "hono"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { documentExchangeRoutes } from "./document-exchange-routes"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { openDb, upsertDocument } from "./db"
import { readCurrentVersion, readDocumentVersions, readVersion, sha256Hex } from "./document-versions"
import { readDocumentReturnMarker } from "./document-return-matching"
import * as preferences from "./preferences"

const docx = (text: string) => Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] }))
function setup() {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-exchange-http-"))
  let writes = 0
  const app = new Hono().route("/matters", documentExchangeRoutes({ directory: () => directory, saveReturn: async (dir, input) => {
    writes++
    return writeDeliverableRevision(dir, input.name, Buffer.from(input.bytes), { priorSha256: input.prior, version: { origin: "revision", parent: input.parent, actor: input.actor, event: input.event } })
  } }))
  const send = (id: string, body: unknown) => app.request(`/matters/one/documents/versions/${id}/send`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
  const returned = (bytes: Uint8Array, base: string, prior: string, requestId = "received-1", route = "returns", fields: Record<string, string> = {}) => {
    const form = new FormData()
    form.append("file", new File([new Uint8Array(bytes)], "Agreement returned.docx"))
    form.append("baseVersionId", base); form.append("priorSha256", prior); form.append("requestId", requestId); form.append("party", "Other side")
    for (const [key, value] of Object.entries(fields)) form.append(key, value)
    return app.request(`/matters/one/documents/${route}`, { method: "POST", body: form })
  }
  return { directory, app, send, returned, writes: () => writes, close: () => rmSync(directory, { recursive: true, force: true }) }
}

test("explicit send HTTP retains its exact bytes and replay while ordinary inspection writes nothing", async () => {
  const s = setup()
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("Delivery must occur on Monday."))
    const current = readCurrentVersion(s.directory, "Agreement.docx")!.current
    const canonical = readFileSync(path.join(s.directory, "Agreement.docx"))
    const before = readDocumentVersions(s.directory, "Agreement.docx")
    const inspect = await s.returned(await docx("Delivery must occur on Tuesday."), current.id, current.sha256, "inspect-only", "returns/inspect")
    expect(inspect.status).toBe(200)
    expect((await inspect.json()).kind).toBe("needs-confirmation")
    expect(readDocumentVersions(s.directory, "Agreement.docx")).toEqual(before)
    const one = await s.send(current.id, { requestId: "send-1", party: "Other side" })
    expect(one.status).toBe(200)
    const bytes = new Uint8Array(await one.arrayBuffer())
    expect(readDocumentReturnMarker(bytes)?.versionId).toBe(current.id)
    const again = await s.send(current.id, { requestId: "send-1", party: "Other side" })
    expect(again.status).toBe(200)
    expect(new Uint8Array(await again.arrayBuffer())).toEqual(bytes)
    expect(again.headers.get("X-DocHaus-Exchange-Id")).toBe(one.headers.get("X-DocHaus-Exchange-Id"))
    const exact = await s.app.request(`/matters/one/documents/exchanges/${one.headers.get("X-DocHaus-Exchange-Id")}/content`)
    expect(new Uint8Array(await exact.arrayBuffer())).toEqual(bytes)
    expect(readFileSync(path.join(s.directory, "Agreement.docx"))).toEqual(canonical)
    expect(readDocumentVersions(s.directory, "Agreement.docx").versions).toHaveLength(1)
    expect(readDocumentVersions(s.directory, "Agreement.docx").events).toHaveLength(1)
    expect(s.writes()).toBe(0)
    const duplicate = await s.returned(canonical, current.id, current.sha256, "duplicate-return")
    expect(duplicate.status).toBe(409)
    expect((await duplicate.json()).code).toBe("duplicate_version")
  } finally { s.close() }
}, 30_000)

test("a sent blackline retries its computed outbound bytes exactly and keeps the chosen comparison base", async () => {
  const s = setup()
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("Delivery must occur on Monday."))
    const base = readCurrentVersion(s.directory, "Agreement.docx")!.current
    await writeDeliverableRevision(s.directory, "Agreement.docx", await docx("Delivery must occur on Tuesday."))
    const current = readCurrentVersion(s.directory, "Agreement.docx")!.current
    const details = { requestId: "send-blackline", baseVersionId: base.id, party: "Other side" }
    const sent = await s.send(current.id, details), bytes = new Uint8Array(await sent.arrayBuffer())
    expect(sent.status).toBe(200)
    const replay = await s.send(current.id, details)
    expect(replay.status).toBe(200)
    expect(new Uint8Array(await replay.arrayBuffer())).toEqual(bytes)
    const events = readDocumentVersions(s.directory, "Agreement.docx").events
    expect(events).toHaveLength(1)
    expect(events[0]!.detail?.baseVersionId).toBe(base.id)
    expect(events[0]!.detail?.payload).toMatchObject({ sha256: sha256Hex(bytes), bytes: bytes.length })
  } finally { s.close() }
}, 30_000)

test("lost send response retries retained bytes after the firm's marker default changes", async () => {
  const s = setup(), defaults = preferences.readDraftingPreferences()
  const savedPreferences = spyOn(preferences, "readDraftingPreferences").mockReturnValue({ ...defaults, documentMatchingId: true })
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("The parties agree to a saved delivery term."))
    const version = readCurrentVersion(s.directory, "Agreement.docx")!.current
    const body = { requestId: "send-default-marker", party: "Other side" }
    const sent = await s.send(version.id, body), bytes = new Uint8Array(await sent.arrayBuffer())
    expect(sent.status).toBe(200)
    expect(readDocumentReturnMarker(bytes)?.versionId).toBe(version.id)
    savedPreferences.mockReturnValue({ ...defaults, documentMatchingId: false })
    const replay = await s.send(version.id, body)
    expect(replay.status).toBe(200)
    expect(new Uint8Array(await replay.arrayBuffer())).toEqual(bytes)
    expect(replay.headers.get("X-DocHaus-Draft-Number")).toBe(sent.headers.get("X-DocHaus-Draft-Number"))
    expect((await s.send(version.id, { ...body, marker: false })).status).toBe(409)
    expect((await s.send(version.id, { ...body, marker: true })).status).toBe(200)
    expect(readDocumentVersions(s.directory, "Agreement.docx").versions).toHaveLength(1)
    expect(readDocumentVersions(s.directory, "Agreement.docx").events).toHaveLength(1)
  } finally { savedPreferences.mockRestore(); s.close() }
}, 30_000)

test("receipt base correction is explicit, metadata only and idempotent over actual HTTP", async () => {
  const s = setup()
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("Delivery must occur on Monday."))
    const first = readCurrentVersion(s.directory, "Agreement.docx")!.current
    await writeDeliverableRevision(s.directory, "Agreement.docx", await docx("Delivery must occur on Wednesday."))
    const second = readCurrentVersion(s.directory, "Agreement.docx")!.current
    const saved = await s.returned(await docx("Delivery must occur on Tuesday."), first.id, second.sha256)
    const receipt = await saved.json()
    const current = readFileSync(path.join(s.directory, "Agreement.docx"))
    const request = (body: unknown) => s.app.request(`/matters/one/documents/exchanges/${receipt.exchange.id}/base`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
    const body = { baseVersionId: second.id, expectedBaseVersionId: first.id, requestId: "correct-base" }
    const corrected = await request(body)
    expect(corrected.status).toBe(200)
    expect((await corrected.json()).changed).toBe(true)
    expect((await (await request(body)).json()).changed).toBe(false)
    expect(readVersion(s.directory, receipt.version.id)!.version.parent).toBe(second.id)
    expect(readFileSync(path.join(s.directory, "Agreement.docx"))).toEqual(current)
    expect(readDocumentVersions(s.directory, "Agreement.docx").versions).toHaveLength(3)
    expect((await request({ ...body, requestId: "stale-correction" })).status).toBe(409)
    expect((await request(null)).status).toBe(400)
  } finally { s.close() }
}, 30_000)

test("cross-document receipt correction exposes exact linked bytes without moving either document's versions", async () => {
  const s = setup()
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("A known base for the return."))
    const source = readCurrentVersion(s.directory, "Agreement.docx")!
    const incoming = await docx("A revised known base for the return.")
    const saved = await s.returned(incoming, source.current.id, source.current.sha256)
    const receipt = await saved.json()
    await writeDeliverable(s.directory, "Intended.docx", await docx("The intended earlier draft."))
    const intended = readCurrentVersion(s.directory, "Intended.docx")!
    const original = readFileSync(path.join(s.directory, "Agreement.docx")), target = readFileSync(path.join(s.directory, "Intended.docx"))
    const body = { targetDocumentId: intended.document.id, baseVersionId: intended.current.id, expectedDocumentId: source.document.id, expectedBaseVersionId: source.current.id, requestId: "link-intended" }
    const request = () => s.app.request(`/matters/one/documents/exchanges/${receipt.exchange.id}/document`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
    const linked = await request()
    expect(linked.status).toBe(200)
    expect((await linked.json()).changed).toBe(true)
    expect((await (await request()).json()).changed).toBe(false)
    const branch = await s.app.request(`/matters/one/documents/exchanges/${receipt.exchange.id}/branch/content`)
    expect(branch.status).toBe(200)
    expect(Buffer.from(await branch.arrayBuffer())).toEqual(original)
    expect(readFileSync(path.join(s.directory, "Agreement.docx"))).toEqual(original)
    expect(readFileSync(path.join(s.directory, "Intended.docx"))).toEqual(target)
    expect(readDocumentVersions(s.directory, "Intended.docx").versions).toHaveLength(1)
    expect(readDocumentVersions(s.directory, "Intended.docx").linked_returns).toMatchObject([{ versionId: receipt.version.id, sourceName: "Agreement.docx", versionNumber: 2 }])
    const replay = await s.returned(incoming, source.current.id, source.current.sha256)
    expect(replay.status).toBe(200)
    expect((await replay.json()).comparison).toEqual({ leftVersion: intended.current.id, rightVersion: receipt.version.id, left: "Intended.docx", right: "Agreement.docx" })
    const db = openDb(s.directory)
    try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [receipt.version.id]) } finally { db.close() }
    expect((await s.app.request(`/matters/one/documents/exchanges/${receipt.exchange.id}/branch/content`)).status).toBe(410)
  } finally { s.close() }
}, 30_000)

test("email receipt provenance comes from an actual registered attachment and is rechecked on retry", async () => {
  const s = setup()
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("Delivery must occur on Monday."))
    const current = readCurrentVersion(s.directory, "Agreement.docx")!.current
    const bytes = await docx("Delivery must occur on Tuesday.")
    const email = (sender = "reviewer@example.invalid") => `From: Recorded Reviewer <${sender}>\r\nDate: Tue, 10 Nov 2026 12:00:00 +0000\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="synthetic-boundary"\r\n\r\n--synthetic-boundary\r\nContent-Type: text/plain\r\n\r\nPlease see attached.\r\n--synthetic-boundary\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\nContent-Disposition: attachment; filename="Agreement returned.docx"\r\nContent-Transfer-Encoding: base64\r\n\r\n${Buffer.from(bytes).toString("base64")}\r\n--synthetic-boundary--\r\n`
    const emailPath = path.join(s.directory, "Return.eml")
    writeFileSync(emailPath, email())
    const db = openDb(s.directory)
    try { upsertDocument(db, emailPath, "Return.eml", Date.now(), null) } finally { db.close() }
    const fields = { emailName: "Return.eml" }
    const inspected = await s.returned(bytes, current.id, current.sha256, "email-receipt", "returns/inspect", fields)
    expect(inspected.status).toBe(200)
    expect((await inspected.json()).source).toMatchObject({ documentName: "Return.eml", sender: "Recorded Reviewer <reviewer@example.invalid>" })
    const wrong = await s.returned(await docx("Different attachment bytes."), current.id, current.sha256, "wrong-email", "returns", fields)
    expect(wrong.status).toBe(400)
    expect(s.writes()).toBe(0)
    const saved = await s.returned(bytes, current.id, current.sha256, "email-receipt", "returns", fields)
    expect(saved.status).toBe(200)
    expect((await saved.json()).exchange.source.documentName).toBe("Return.eml")
    writeFileSync(emailPath, email("another@example.invalid"))
    expect((await s.returned(bytes, current.id, current.sha256, "email-receipt", "returns", fields)).status).toBe(409)
    expect(s.writes()).toBe(1)
  } finally { s.close() }
}, 30_000)

test("return HTTP preserves the explicitly selected old base, refuses stale current, and reconciles retries once", async () => {
  const s = setup()
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("Delivery must occur on Monday."))
    const base = readCurrentVersion(s.directory, "Agreement.docx")!.current
    await writeDeliverableRevision(s.directory, "Agreement.docx", await docx("Internal delivery must occur on Wednesday."))
    const current = readCurrentVersion(s.directory, "Agreement.docx")!.current
    const duplicate = await s.returned(readFileSync(path.join(s.directory, ".dochaus", "revisions", base.document_id, `${base.sha256}${base.ext}`)), current.id, current.sha256, "old-duplicate")
    expect(duplicate.status).toBe(409)
    expect((await duplicate.json()).version.id).toBe(base.id)
    const bytes = await docx("Delivery must occur on Tuesday.")
    const stale = await s.returned(bytes, base.id, base.sha256)
    expect(stale.status).toBe(409)
    expect(s.writes()).toBe(0)
    const saved = await s.returned(bytes, base.id, current.sha256)
    expect(saved.status).toBe(200)
    const receipt = await saved.json()
    expect(readVersion(s.directory, receipt.version.id)!.version.parent).toBe(base.id)
    expect(receipt.comparison).toEqual({ leftVersion: base.id, rightVersion: receipt.version.id, left: "Agreement.docx", right: "Agreement.docx" })
    expect(readDocumentVersions(s.directory, "Agreement.docx").events).toHaveLength(1)
    const replay = await s.returned(bytes, base.id, current.sha256)
    expect(replay.status).toBe(200)
    expect((await replay.json()).version.id).toBe(receipt.version.id)
    expect(s.writes()).toBe(1)
    expect(readDocumentVersions(s.directory, "Agreement.docx").versions).toHaveLength(3)
    expect((await s.returned(await docx("A completely different receipt."), base.id, current.sha256)).status).toBe(409)
    const db = openDb(s.directory)
    try { db.run("UPDATE document_versions SET removed='redaction' WHERE id=?", [receipt.version.id]) } finally { db.close() }
    const redacted = await s.returned(bytes, base.id, current.sha256)
    expect(redacted.status).toBe(410)
    expect((await redacted.json()).code).toBe("removed_by_redaction")
    expect(s.writes()).toBe(1)
  } finally { s.close() }
}, 30_000)

test("malformed send JSON and unreadable returned packages receive a useful client error", async () => {
  const s = setup()
  try {
    await writeDeliverable(s.directory, "Agreement.docx", await docx("An exact saved paragraph."))
    const current = readCurrentVersion(s.directory, "Agreement.docx")!.current
    for (const body of ["{", "null", "[]", "false"]) {
      const result = await s.app.request(`/matters/one/documents/versions/${current.id}/send`, { method: "POST", headers: { "content-type": "application/json" }, body })
      expect(result.status).toBe(400)
    }
    const invalid = await s.returned(new TextEncoder().encode("This is not a Word package."), current.id, current.sha256)
    expect(invalid.status).toBe(400)
    expect(s.writes()).toBe(0)
    expect(sha256Hex(readFileSync(path.join(s.directory, "Agreement.docx")))).toBe(current.sha256)
  } finally { s.close() }
}, 30_000)
