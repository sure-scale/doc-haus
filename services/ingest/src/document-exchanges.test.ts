import { afterAll, expect, spyOn, test } from "bun:test"
import { randomUUID } from "node:crypto"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { linkedWordDocx, readWordLink, stripWordLink } from "../../../dochaus/lib/word-linked-docx"
import { applyPageFurniture } from "../../../dochaus/lib/page-furniture"
import { openDb } from "./db"
import { DocumentExchangeError, nextDocumentDraftNumber, prepareDocumentDownload, prepareDocumentReceivedEvent, readDocumentExchangePayload, readDocumentExchangeRequest, readDocumentExchanges, readDocumentReceivedBranch, recordDocumentReceived, registerDocumentSent, relinkDocumentReceivedBase, relinkDocumentReceivedDocument } from "./document-exchanges"
import { writeDraftingPreferences } from "./preferences"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import * as documentMutation from "./document-mutation"
import { captureVersion, detachVersionedDocument, purgeVersionBytes, readDocumentVersions, sha256Hex, versionFile } from "./document-versions"
import { inspectDocumentReturn, readDocumentReturnMarker, stampDocumentReturnMarker } from "./document-return-matching"

const dirs: string[] = []
afterAll(() => dirs.forEach(dir => rmSync(dir, { recursive: true, force: true })))
function matter() { const dir = mkdtempSync(path.join(tmpdir(), "dochaus-exchanges-")); dirs.push(dir); openDb(dir).close(); return dir }
const docx = (...text: string[]) => Packer.toBuffer(new Document({ sections: [{ children: text.map(t => new Paragraph(t)) }] }))
function keep(dir: string, name: string, bytes: Uint8Array, parent?: string, origin: "upload" | "revision" | "redaction" | "generated" = "upload") {
  const file = path.join(dir, name), prior = existsSync(file) ? readFileSync(file) : undefined
  const db = openDb(dir)
  try {
    const result = captureVersion(db, dir, { name, bytes, prior, version: { origin, ...(parent ? { parent } : {}) } })
    writeFileSync(file, bytes)
    return result
  } finally { db.close() }
}
function edit(bytes: Uint8Array, before: string, after: string) {
  const parts = unzipSync(bytes)
  const body = strFromU8(parts["word/document.xml"]!)
  expect(body).toContain(before)
  parts["word/document.xml"] = strToU8(body.replace(before, after))
  return zipSync(parts)
}
function rsids(bytes: Uint8Array, root: string, ids: string[]) {
  const parts = unzipSync(bytes)
  const settings = strFromU8(parts["word/settings.xml"]!)
  parts["word/settings.xml"] = strToU8(settings.replace(/<\/w:settings>/, `<w:rsids><w:rsidRoot w:val="${root}"/>${ids.map(id => `<w:rsid w:val="${id}"/>`).join("")}</w:rsids></w:settings>`))
  return zipSync(parts)
}
function failure(fn: () => unknown, code: string) {
  try { fn(); throw new Error("Expected refusal") }
  catch (error) { expect(error).toBeInstanceOf(DocumentExchangeError); expect((error as DocumentExchangeError).code).toBe(code) }
}

test("phase-one event migration preserves old IDs/details and accepts a send without rewriting versions", async () => {
  const dir = matter(), v = keep(dir, "Agreement.docx", await docx("A neutral supply term."))
  const db = openDb(dir)
  db.run("DROP TABLE document_version_events")
  db.run("CREATE TABLE document_version_events(id INTEGER PRIMARY KEY AUTOINCREMENT,document_id TEXT NOT NULL,version_id TEXT,kind TEXT NOT NULL CHECK(kind IN ('accept','reject','restore')),actor TEXT NOT NULL DEFAULT 'unknown',detail TEXT,created_at INTEGER NOT NULL)")
  db.run("INSERT INTO document_version_events VALUES(41,?,?,'accept','Recorded person','{\"ids\":[\"r1\"]}',123)", [v.document_id, v.id])
  db.close()
  const prepared = prepareDocumentDownload(dir, { versionId: v.id })
  const sent = registerDocumentSent(dir, { prepared, requestId: "send-migration" })
  expect(sent.exchange.id).toBe(42)
  expect(readDocumentVersions(dir, "Agreement.docx").events[0]).toMatchObject({ id: 41, actor: "Recorded person", detail: { ids: ["r1"] }, created_at: 123 })
  expect(readDocumentVersions(dir, "Agreement.docx").versions).toHaveLength(1)
})

test("prepared Word downloads strip write capabilities before hashing and remain byte stable on retry", async () => {
  const dir = matter(), plain = await docx("The delivery period is five days.")
  const v = keep(dir, "Agreement.docx", plain)
  const linked = linkedWordDocx(plain, { version: 1, origin: "https://example.invalid", apiOrigin: "https://example.invalid", matter: randomUUID(), matterName: "Synthetic matter", document: v.document_id,
    documentName: "Agreement.docx", baseVersion: v.id, token: "synthetic-write-capability" })
  const first = prepareDocumentDownload(dir, { versionId: v.id, bytes: linked })
  const retry = prepareDocumentDownload(dir, { versionId: v.id, bytes: linked })
  expect(readWordLink(first.bytes)).toBeNull()
  expect(sha256Hex(first.bytes)).toBe(first.payloadSha256)
  expect(first.payloadSha256).toBe(retry.payloadSha256)
  expect(Buffer.from(stripWordLink(first.bytes))).toEqual(Buffer.from(first.bytes))
  expect(readDocumentReturnMarker(first.bytes)).toEqual({ documentId: v.document_id, versionId: v.id })
  expect(Buffer.from(first.bytes).includes(Buffer.from("synthetic-write-capability"))).toBe(false)
  // Internal preparation records no event or version; the canonical file is untouched.
  expect(readDocumentExchanges(dir, v.document_id)).toEqual([])
  expect(readFileSync(path.join(dir, "Agreement.docx"))).toEqual(Buffer.from(plain))
  expect(readDocumentVersions(dir, "Agreement.docx").versions).toHaveLength(1)
})

test("marker-off preserves a plain file exactly and removes only our custom property from a marked copy", async () => {
  const dir = matter(), plain = await docx("An ordinary unmarked copy."), v = keep(dir, "Agreement.docx", plain)
  expect(Buffer.from(prepareDocumentDownload(dir, { versionId: v.id, marker: false }).bytes)).toEqual(Buffer.from(plain))
  const marked = prepareDocumentDownload(dir, { versionId: v.id }).bytes
  const scrubbed = prepareDocumentDownload(dir, { versionId: v.id, bytes: marked, marker: false }).bytes
  expect(readDocumentReturnMarker(scrubbed)).toBeUndefined()
  expect(strFromU8(unzipSync(scrubbed)["word/document.xml"]!)).toBe(strFromU8(unzipSync(plain)["word/document.xml"]!))
})

test("explicit sends alone count drafts, retain exact outbound views and replay idempotently", async () => {
  const dir = matter(), v1 = keep(dir, "Agreement.docx", await docx("Deliver on Monday."))
  const v2 = keep(dir, "Agreement.docx", await docx("Deliver on Tuesday."), v1.id, "revision")
  const view = await docx("Synthetic comparison view: Monday to Tuesday.")
  const prepared = prepareDocumentDownload(dir, { versionId: v2.id, baseVersionId: v1.id, bytes: view })
  const one = registerDocumentSent(dir, { prepared, party: "Counterparty", actor: "Recorded lawyer", sentAt: 100, requestId: "send-1" })
  const replay = registerDocumentSent(dir, { prepared, party: "Counterparty", actor: "Recorded lawyer", sentAt: 100, requestId: "send-1" })
  expect(one).toEqual(replay)
  expect(one.draftNumber).toBe(1)
  expect(one.exchange).toMatchObject({ versionId: v2.id, baseVersionId: v1.id, actor: "Recorded lawyer", occurredAt: 100 })
  expect(Buffer.from(readDocumentExchangePayload(dir, one.exchange.id).bytes)).toEqual(Buffer.from(prepared.bytes))
  expect(readDocumentVersions(dir, "Agreement.docx").versions).toHaveLength(2)
  prepareDocumentDownload(dir, { versionId: v2.id })
  recordDocumentReceived(dir, { versionId: v2.id, baseVersionId: v1.id, requestId: "receipt-1" })
  expect(registerDocumentSent(dir, { prepared, requestId: "send-2" }).draftNumber).toBe(2)
  expect(readDocumentExchanges(dir, v1.document_id).map(e => e.kind)).toEqual(["sent", "received", "sent"])
})

test("only explicit sends number an existing generated draft line, with a stable retry and untouched working version", async () => {
  const dir = matter(), plain = await docx("The operative wording stays the same.")
  const original = applyPageFurniture(plain, { header: ["Private legend", "Fixture: Legal draft: 1 November 2026"], pageNumber: { format: "page", align: "right" } })
  const v = keep(dir, "Advice.docx", original, undefined, "generated")
  const header = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/header1.xml"]!)
  expect(header(prepareDocumentDownload(dir, { versionId: v.id }).bytes)).toContain("Fixture: Legal draft: 1 November 2026")
  expect(nextDocumentDraftNumber(dir, v.document_id)).toBe(1)
  const first = prepareDocumentDownload(dir, { versionId: v.id, draftNumber: nextDocumentDraftNumber(dir, v.document_id, "number-one") })
  expect(header(first.bytes)).toContain("Fixture: Legal draft 1: 1 November 2026")
  expect(header(first.bytes)).toContain("Private legend")
  registerDocumentSent(dir, { prepared: first, requestId: "number-one" })
  const second = prepareDocumentDownload(dir, { versionId: v.id, draftNumber: nextDocumentDraftNumber(dir, v.document_id, "number-two") })
  registerDocumentSent(dir, { prepared: second, requestId: "number-two" })
  expect(header(second.bytes)).toContain("Fixture: Legal draft 2: 1 November 2026")
  const retry = prepareDocumentDownload(dir, { versionId: v.id, draftNumber: nextDocumentDraftNumber(dir, v.document_id, "number-one") })
  expect(retry.payloadSha256).toBe(first.payloadSha256)
  expect(registerDocumentSent(dir, { prepared: retry, requestId: "number-one" }).draftNumber).toBe(1)
  failure(() => registerDocumentSent(dir, { prepared: first, requestId: "stale-number" }), "sent_count_changed")
  expect(sha256Hex(readFileSync(path.join(dir, "Advice.docx")))).toBe(sha256Hex(original))
  expect(readDocumentVersions(dir, "Advice.docx").versions).toHaveLength(1)
  expect(strFromU8(unzipSync(first.bytes)["word/document.xml"]!)).toBe(strFromU8(unzipSync(original)["word/document.xml"]!))
  expect(strFromU8(unzipSync(first.bytes)["word/footer1.xml"]!)).toBe(strFromU8(unzipSync(original)["word/footer1.xml"]!))
  const uploaded = keep(dir, "Their advice.docx", original)
  expect(header(prepareDocumentDownload(dir, { versionId: uploaded.id, draftNumber: 1 }).bytes)).toContain("Fixture: Legal draft: 1 November 2026")
})

test("a saved firm matching-ID opt-out applies by default while a send can explicitly override it", async () => {
  const dir = matter(), previousRoot = process.env.WORKSPACE_ROOT
  process.env.WORKSPACE_ROOT = dir
  try {
    writeDraftingPreferences({ documentMatchingId: false })
    const v = keep(dir, "Agreement.docx", await docx("A copy with a firm-controlled matching ID."))
    expect(readDocumentReturnMarker(prepareDocumentDownload(dir, { versionId: v.id }).bytes)).toBeUndefined()
    expect(readDocumentReturnMarker(prepareDocumentDownload(dir, { versionId: v.id, marker: true }).bytes)).toEqual({ documentId: v.document_id, versionId: v.id })
    writeDraftingPreferences({ documentMatchingId: true })
    expect(readDocumentReturnMarker(prepareDocumentDownload(dir, { versionId: v.id, marker: false }).bytes)).toBeUndefined()
  } finally {
    if (previousRoot === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previousRoot
  }
})

test("request identifiers cannot silently change bytes, party, date, source, document or event kind", async () => {
  const dir = matter(), v = keep(dir, "Agreement.docx", await docx("The supplier delivers five crates."))
  const prepared = prepareDocumentDownload(dir, { versionId: v.id })
  registerDocumentSent(dir, { prepared, party: "Other side", sentAt: 100, requestId: "same-request" })
  failure(() => registerDocumentSent(dir, { prepared, party: "Another side", requestId: "same-request" }), "request_conflict")
  failure(() => registerDocumentSent(dir, { prepared, party: "Other side", sentAt: 101, requestId: "same-request" }), "request_conflict")
  const changed = prepareDocumentDownload(dir, { versionId: v.id, bytes: edit(prepared.bytes, "five crates", "six crates") })
  failure(() => registerDocumentSent(dir, { prepared: changed, party: "Other side", requestId: "same-request" }), "request_conflict")
  const other = keep(dir, "Other.docx", await docx("An unrelated agreement."))
  failure(() => registerDocumentSent(dir, { prepared: prepareDocumentDownload(dir, { versionId: other.id }), requestId: "same-request" }), "request_conflict")
  failure(() => recordDocumentReceived(dir, { versionId: v.id, requestId: "same-request" }), "request_conflict")
  expect(readDocumentExchanges(dir, v.document_id)).toHaveLength(1)
  recordDocumentReceived(dir, { versionId: v.id, requestId: "receipt-source", source: { attachmentId: "a-1" } })
  failure(() => recordDocumentReceived(dir, { versionId: v.id, requestId: "receipt-source", source: { attachmentId: "a-2" } }), "request_conflict")
})

test("a receipt commits with the real revision writer or rolls back with its file/version, and a lost response can be reconciled", async () => {
  const dir = matter(), first = await docx("Delivery is on Monday."), returned = await docx("Delivery is on Tuesday.")
  await writeDeliverable(dir, "Agreement.docx", first)
  const baseline = sha256Hex(readFileSync(path.join(dir, "Agreement.docx")))
  const v1 = readDocumentVersions(dir, "Agreement.docx").versions[0]!
  const event = prepareDocumentReceivedEvent(dir, { baseVersionId: v1.id, requestId: "atomic-receipt", party: "Other side", source: { attachmentId: "attachment-1" } })
  const stop = spyOn(documentMutation, "recordDocumentMutationCommit").mockImplementation(() => { throw new Error("synthetic failure before registry commit") })
  try {
    await expect(writeDeliverableRevision(dir, "Agreement.docx", returned, { version: { origin: "revision", parent: v1.id, event } })).rejects.toThrow("synthetic failure before registry commit")
  } finally { stop.mockRestore() }
  expect(sha256Hex(readFileSync(path.join(dir, "Agreement.docx")))).toBe(baseline)
  expect(readDocumentVersions(dir, "Agreement.docx").versions).toHaveLength(1)
  expect(readDocumentExchangeRequest(dir, "atomic-receipt")).toBeUndefined()
  await writeDeliverableRevision(dir, "Agreement.docx", returned, { version: { origin: "revision", parent: v1.id, event } })
  // Treat the successful writer response as lost: registry history is the result.
  const receipt = readDocumentExchangeRequest(dir, "atomic-receipt")!
  expect(receipt).toMatchObject({ kind: "received", baseVersionId: v1.id, party: "Other side", source: { attachmentId: "attachment-1" } })
  expect(recordDocumentReceived(dir, { versionId: receipt.versionId, baseVersionId: v1.id, requestId: "atomic-receipt", party: "Other side", source: { attachmentId: "attachment-1" } })).toEqual(receipt)
  expect(readDocumentVersions(dir, "Agreement.docx").versions).toHaveLength(2)
  expect(readDocumentExchanges(dir, receipt.documentId)).toHaveLength(1)
})

test("a return from an older sent source keeps that parent beside a later internal version", async () => {
  const dir = matter(), v1 = keep(dir, "Agreement.docx", await docx("Deliver on Monday."))
  const outgoing = prepareDocumentDownload(dir, { versionId: v1.id })
  registerDocumentSent(dir, { prepared: outgoing, requestId: "send-parent" })
  const internal = keep(dir, "Agreement.docx", await docx("Deliver on Wednesday."), v1.id, "revision")
  const incoming = edit(outgoing.bytes, "Monday", "Tuesday")
  const match = await inspectDocumentReturn(dir, { name: "Their returned draft.docx", bytes: incoming })
  expect(match.kind).toBe("matched")
  expect(match.selected?.versionId).toBe(v1.id)
  const saved = keep(dir, "Agreement.docx", incoming, match.selected!.versionId, "revision")
  const receipt = recordDocumentReceived(dir, { versionId: saved.id, baseVersionId: v1.id, requestId: "receive-parent", party: "Other side", source: { attachmentId: "attachment-1", sender: "reviewer@example.invalid", date: "2026-11-01" }, match: { signal: "marker", reason: match.reason } })
  expect(receipt.baseVersionId).toBe(v1.id)
  expect(saved.parent).toBe(v1.id)
  expect(readDocumentVersions(dir, "Agreement.docx").versions.map(v => v.id)).toContain(internal.id)
  failure(() => recordDocumentReceived(dir, { versionId: saved.id, baseVersionId: internal.id, requestId: "wrong-receipt" }), "received_parent_mismatch")
})

test("an explicit received-base correction changes only registry metadata, with cycle, collision and concurrent-edit refusals", async () => {
  const dir = matter(), first = keep(dir, "Agreement.docx", await docx("The first source version."))
  const second = keep(dir, "Agreement.docx", await docx("The second source version."), first.id, "revision")
  const returned = keep(dir, "Agreement.docx", await docx("The returned wording."), first.id, "revision")
  const receipt = recordDocumentReceived(dir, { versionId: returned.id, baseVersionId: first.id, requestId: "received-for-correction", match: { signal: "lawyer-confirmed", reason: "Originally selected by the lawyer." } })
  const baseline = sha256Hex(readFileSync(path.join(dir, "Agreement.docx")))
  const versions = readDocumentVersions(dir, "Agreement.docx").versions.map(version => [version.id, version.sha256])
  const corrected = relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: second.id, expectedBaseVersionId: first.id, requestId: "correct-base", actor: "Recorded lawyer" })
  expect(corrected.changed).toBe(true)
  expect(sha256Hex(readFileSync(path.join(dir, "Agreement.docx")))).toBe(baseline)
  expect(readDocumentVersions(dir, "Agreement.docx").versions.map(version => [version.id, version.sha256])).toEqual(versions)
  expect(corrected.exchange).toMatchObject({ versionId: returned.id, baseVersionId: second.id, match: { signal: "lawyer-corrected" }, baseCorrections: [{ requestId: "correct-base", fromVersionId: first.id, toVersionId: second.id, actor: "Recorded lawyer", priorMatch: { signal: "lawyer-confirmed" } }] })
  expect(readDocumentVersions(dir, "Agreement.docx").versions[0]?.parent).toBe(second.id)
  expect(relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: second.id, expectedBaseVersionId: first.id, requestId: "correct-base" })).toEqual({ exchange: corrected.exchange, changed: false })
  failure(() => relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: first.id, expectedBaseVersionId: second.id, requestId: "correct-base" }), "request_conflict")
  failure(() => relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: first.id, expectedBaseVersionId: first.id, requestId: "stale-base" }), "received_base_changed")
  failure(() => relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: returned.id, expectedBaseVersionId: second.id, requestId: "self-base" }), "received_base_cycle")
  const other = keep(dir, "Other.docx", await docx("A different document."))
  failure(() => relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: other.id, expectedBaseVersionId: second.id, requestId: "foreign-base" }), "base_document_mismatch")
  const descendant = keep(dir, "Agreement.docx", await docx("A later revision of the return."), returned.id, "revision")
  failure(() => relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: descendant.id, expectedBaseVersionId: second.id, requestId: "descendant-base" }), "received_base_cycle")
  failure(() => registerDocumentSent(dir, { prepared: prepareDocumentDownload(dir, { versionId: second.id }), requestId: "correct-base" }), "request_conflict")
  failure(() => readDocumentExchangeRequest(dir, "correct-base"), "request_conflict")
  expect(sha256Hex(readFileSync(path.join(dir, "Agreement.docx")))).toBe(sha256Hex(readFileSync(versionFile(dir, { document_id: descendant.document_id, sha256: readDocumentVersions(dir, "Agreement.docx").versions[0]!.sha256, ext: ".docx" }))))
  expect(readDocumentVersions(dir, "Agreement.docx").versions.slice(1).map(version => [version.id, version.sha256])).toEqual(versions)
  expect(readDocumentVersions(dir, "Agreement.docx").versions.find(version => version.id === returned.id)?.sha256).toBe(baseline)
  recordDocumentReceived(dir, { versionId: returned.id, baseVersionId: second.id, requestId: "second-receipt" })
  failure(() => relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: first.id, expectedBaseVersionId: second.id, requestId: "multiple-receipts" }), "relink_multiple_receipts")
})

test("cross-document relinking projects the exact received artifact as a branch without replacing either working copy or version history", async () => {
  const dir = matter(), owner = "Original owner.docx", target = "Correct agreement.docx"
  const wrongBase = keep(dir, owner, await docx("An unrelated original document."))
  const returned = keep(dir, owner, await docx("The counterparty's returned agreement."), wrongBase.id, "revision")
  const receipt = recordDocumentReceived(dir, { versionId: returned.id, baseVersionId: wrongBase.id, requestId: "misassigned-receipt", party: "Other side" })
  const base = keep(dir, target, await docx("The agreement sent for review."))
  const current = keep(dir, target, await docx("A later internal agreement."), base.id, "revision")
  const before = [owner, target].map(name => ({ name, hash: sha256Hex(readFileSync(path.join(dir, name))),
    rows: readDocumentVersions(dir, name).versions.map(version => [version.id, version.sha256, version.parent]), current: readDocumentVersions(dir, name).document!.current_version }))
  const result = relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: base.document_id, baseVersionId: base.id,
    expectedDocumentId: wrongBase.document_id, expectedBaseVersionId: wrongBase.id, expectedSourceCurrentVersionId: returned.id, expectedTargetCurrentVersionId: current.id, requestId: "reassign-to-agreement", actor: "Recorded lawyer" })
  expect(result.changed).toBe(true)
  expect(result.exchange).toMatchObject({ documentId: returned.document_id, versionId: returned.id, baseVersionId: wrongBase.id,
    association: { documentId: base.document_id, baseVersionId: base.id }, linkedBranch: { sourceName: owner, versionNumber: 2, targetName: target, baseVersionNumber: 1, current: false, available: true } })
  const projected = readDocumentVersions(dir, target)
  expect(projected.linked_returns).toHaveLength(1)
  expect(projected.linked_returns[0]).toMatchObject({ exchangeId: receipt.id, versionId: returned.id, sourceDocumentId: returned.document_id, targetDocumentId: base.document_id, current: false })
  expect(readDocumentExchanges(dir, base.document_id).map(exchange => exchange.id)).toEqual([receipt.id])
  expect(readDocumentExchanges(dir, wrongBase.document_id).map(exchange => exchange.id)).toEqual([receipt.id])
  expect(sha256Hex(readDocumentReceivedBranch(dir, receipt.id).bytes)).toBe(before[0]!.hash)
  for (const saved of before) {
    expect(sha256Hex(readFileSync(path.join(dir, saved.name)))).toBe(saved.hash)
    expect(readDocumentVersions(dir, saved.name).versions.map(version => [version.id, version.sha256, version.parent])).toEqual(saved.rows)
    expect(readDocumentVersions(dir, saved.name).document!.current_version).toBe(saved.current)
  }
  expect(relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: base.document_id, baseVersionId: base.id,
    expectedDocumentId: wrongBase.document_id, expectedBaseVersionId: wrongBase.id, requestId: "reassign-to-agreement" }).changed).toBe(false)
  failure(() => relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: base.document_id, baseVersionId: current.id,
    expectedDocumentId: base.document_id, expectedBaseVersionId: base.id, requestId: "reassign-to-agreement" }), "request_conflict")
  failure(() => relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: base.document_id, baseVersionId: current.id,
    expectedDocumentId: wrongBase.document_id, expectedBaseVersionId: wrongBase.id, requestId: "stale-association" }), "received_association_changed")
  failure(() => relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: base.document_id, baseVersionId: current.id,
    expectedDocumentId: base.document_id, expectedBaseVersionId: base.id, expectedTargetCurrentVersionId: base.id, requestId: "stale-current" }), "received_association_changed")
  const correctedBase = relinkDocumentReceivedBase(dir, { exchangeId: receipt.id, baseVersionId: current.id, expectedBaseVersionId: base.id, requestId: "correct-linked-base" })
  expect(correctedBase.exchange.association?.baseVersionId).toBe(current.id)
  expect(readDocumentVersions(dir, owner).versions[0]!.parent).toBe(wrongBase.id)
  expect(readDocumentReceivedBranch(dir, receipt.id).branch.baseVersionId).toBe(current.id)
  keep(dir, owner, await docx("The redacted owner now has safe wording."), returned.id, "redaction")
  expect(readDocumentVersions(dir, target).linked_returns[0]).toMatchObject({ available: false, unavailableReason: "source-version-removed" })
  failure(() => readDocumentReceivedBranch(dir, receipt.id), "linked_branch_unavailable")
})

test("a linked branch loses access when its target base is redacted or its physical owner is deleted", async () => {
  const dir = matter(), owner = keep(dir, "Returned file.docx", await docx("The saved return."))
  const base = keep(dir, "Agreement.docx", await docx("The retained source agreement."))
  const receipt = recordDocumentReceived(dir, { versionId: owner.id, requestId: "unmatched-receipt" })
  relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: base.document_id, baseVersionId: base.id, expectedDocumentId: owner.document_id, expectedBaseVersionId: null, requestId: "link-unmatched" })
  const baseRow = readDocumentVersions(dir, "Agreement.docx").versions[0]!
  const retainedBase = versionFile(dir, { document_id: base.document_id, sha256: baseRow.sha256, ext: baseRow.ext })
  const originalBase = readFileSync(retainedBase)
  writeFileSync(retainedBase, "Synthetic damaged retained base.")
  expect(readDocumentVersions(dir, "Agreement.docx").linked_returns[0]).toMatchObject({ available: false, unavailableReason: "bytes-unavailable" })
  failure(() => readDocumentReceivedBranch(dir, receipt.id), "linked_branch_unavailable")
  writeFileSync(retainedBase, originalBase)
  expect(readDocumentVersions(dir, "Agreement.docx").linked_returns[0]?.available).toBe(true)
  keep(dir, "Agreement.docx", await docx("Redacted target wording."), base.id, "redaction")
  expect(readDocumentVersions(dir, "Agreement.docx").linked_returns[0]).toMatchObject({ available: false, unavailableReason: "base-removed" })
  failure(() => readDocumentReceivedBranch(dir, receipt.id), "linked_branch_unavailable")
  const db = openDb(dir)
  try { detachVersionedDocument(db, dir, "Returned file.docx") } finally { db.close() }
  expect(readDocumentVersions(dir, "Agreement.docx").linked_returns[0]).toMatchObject({ available: false, unavailableReason: "source-removed" })
  expect(readDocumentExchanges(dir, base.document_id)[0]?.linkedBranch?.available).toBe(false)
})

test("retention preserves exact send bytes while live, then redaction/deletion purge and refuse them", async () => {
  const dir = matter(), v1 = keep(dir, "Agreement.docx", await docx("Confidential synthetic term."))
  const prepared = prepareDocumentDownload(dir, { versionId: v1.id })
  const sent = registerDocumentSent(dir, { prepared, requestId: "send-purge" })
  const payloadFile = versionFile(dir, { document_id: v1.document_id, sha256: prepared.payloadSha256, ext: prepared.ext })
  purgeVersionBytes(dir, v1.document_id)
  expect(existsSync(payloadFile)).toBe(true)
  keep(dir, "Agreement.docx", await docx("[REDACTED]"), undefined, "redaction")
  failure(() => readDocumentExchangePayload(dir, sent.exchange.id), "removed_by_redaction")
  purgeVersionBytes(dir, v1.document_id)
  expect(existsSync(payloadFile)).toBe(false)
  const v2 = keep(dir, "Other.docx", await docx("Public synthetic clause."))
  const second = registerDocumentSent(dir, { prepared: prepareDocumentDownload(dir, { versionId: v2.id }), requestId: "send-delete" })
  const db = openDb(dir)
  detachVersionedDocument(db, dir, "Other.docx"); db.close()
  purgeVersionBytes(dir, v2.document_id)
  failure(() => readDocumentExchangePayload(dir, second.exchange.id), "version_missing")
  expect(readDocumentExchanges(dir, v2.document_id)).toEqual([])
})

test("renamed exact duplicates are found across the matter without a write; multiple documents stay explicit", async () => {
  const dir = matter(), bytes = await docx("The neutral buyer pays within twenty days."), v = keep(dir, "Agreement.docx", bytes)
  const before = readDocumentVersions(dir, "Agreement.docx")
  const first = await inspectDocumentReturn(dir, { name: "renamed-final.docx", bytes })
  expect(first.kind).toBe("duplicate")
  expect(first.selected?.versionId).toBe(v.id)
  expect(readDocumentVersions(dir, "Agreement.docx")).toEqual(before)
  keep(dir, "Deliberate independent copy.docx", bytes)
  const ambiguous = await inspectDocumentReturn(dir, { name: "again.docx", bytes })
  expect(ambiguous.kind).toBe("duplicate")
  expect(ambiguous.selected).toBeUndefined()
  expect(ambiguous.candidates).toHaveLength(2)
})

test("scrubbed near-identical text and Word session metadata only retrieve candidates, never auto-link", async () => {
  const dir = matter(), original = rsids(await docx("The supplier must deliver five sample crates within thirty working days.", "The buyer pays within twenty days after delivery."), "00aaaa01", ["00aaaa01", "00bbbb02"])
  const v = keep(dir, "Supply agreement.docx", original)
  registerDocumentSent(dir, { prepared: prepareDocumentDownload(dir, { versionId: v.id }), requestId: "send-candidates" })
  const incoming = edit(original, "five sample crates", "six sample crates")
  const result = await inspectDocumentReturn(dir, { name: "Supply agreement v2.docx", bytes: incoming })
  expect(result.kind).toBe("needs-confirmation")
  expect(result.selected).toBeUndefined()
  expect(result.candidates[0]).toMatchObject({ versionId: v.id, sent: true })
  expect(result.candidates[0]!.signals).toContain("rsid")
  expect(result.candidates[0]!.signals).toContain("text")
  // Same boilerplate with one changed obligation remains a candidate, never identity proof.
  const unrelated = await inspectDocumentReturn(dir, { name: "Unrelated agreement.docx", bytes: edit(original, "supplier must deliver", "supplier may cancel delivery of") })
  expect(unrelated.kind).toBe("needs-confirmation")
  expect(unrelated.selected).toBeUndefined()
})

test("a scrubbed return suggests the closest sent base ahead of a more similar unsent current version", async () => {
  const dir = matter(), name = "Services agreement.docx"
  const wording = "The supplier must deliver the goods within ten days after receiving the written purchase order. The buyer must inspect the goods promptly and report visible defects in writing. Payment falls due thirty days after delivery and a valid invoice."
  const v1 = keep(dir, name, await docx(wording.replace("ten days", "five days")))
  registerDocumentSent(dir, { prepared: prepareDocumentDownload(dir, { versionId: v1.id }), requestId: "earlier-send" })
  const v2 = keep(dir, name, await docx(wording), v1.id, "revision")
  const sent = prepareDocumentDownload(dir, { versionId: v2.id, marker: true })
  registerDocumentSent(dir, { prepared: sent, requestId: "closest-send" })
  const v4 = keep(dir, name, await docx(wording.replace("ten days", "twelve days")), v2.id, "revision")
  const incoming = stampDocumentReturnMarker(edit(sent.bytes, "ten days", "twelve days"), null)
  expect(sha256Hex(incoming)).not.toBe(readDocumentVersions(dir, name).versions[0]!.sha256)
  const inspection = await inspectDocumentReturn(dir, { name: "Their scrubbed markup.docx", bytes: incoming })
  expect(inspection.kind).toBe("needs-confirmation")
  expect(inspection.selected).toBeUndefined()
  expect(inspection.candidates[0]?.versionId).toBe(v2.id)
  expect(inspection.candidates.find(candidate => candidate.versionId === v4.id)!.score).toBeGreaterThan(inspection.candidates[0]!.score)
})

test("a marked later copy requires confirmation after its retained physical source is linked to another document", async () => {
  const dir = matter(), owner = "Misidentified return.docx", target = "Correct agreement.docx"
  const wrongBase = keep(dir, owner, await docx("An unrelated original document."))
  const returned = keep(dir, owner, await docx("The supplier must deliver within ten days."), wrongBase.id, "revision")
  const receipt = recordDocumentReceived(dir, { versionId: returned.id, baseVersionId: wrongBase.id, requestId: "marked-misassigned-receipt" })
  const base = keep(dir, target, await docx("The supplier must deliver within five days."))
  const current = keep(dir, target, await docx("A later internal agreement."), base.id, "revision")
  const exported = prepareDocumentDownload(dir, { versionId: returned.id, marker: true }).bytes
  const incoming = edit(exported, "ten days", "twelve days")
  const before = await inspectDocumentReturn(dir, { name: "Later marked copy.docx", bytes: incoming })
  expect(before.kind).toBe("matched")
  expect(before.selected?.versionId).toBe(returned.id)
  const physicalHash = readDocumentVersions(dir, owner).versions[0]!.sha256
  relinkDocumentReceivedDocument(dir, { exchangeId: receipt.id, targetDocumentId: base.document_id, baseVersionId: base.id,
    expectedDocumentId: returned.document_id, expectedBaseVersionId: wrongBase.id, requestId: "correct-marked-association" })
  const after = await inspectDocumentReturn(dir, { name: "Later marked copy.docx", bytes: incoming })
  expect(after.kind).toBe("needs-confirmation")
  expect(after.selected).toBeUndefined()
  expect(after.candidates.map(c => [c.documentId, c.versionId])).toEqual([[base.document_id, base.id], [returned.document_id, returned.id]])
  expect(after.candidates[0]).toMatchObject({ documentName: target, versionNumber: 1, current: false, signals: [] })
  expect(after.candidates[0]!.reason).toContain("base version 1")
  expect(after.candidates[0]!.reason).toContain("not proof")
  expect(after.candidates[1]).toMatchObject({ documentName: owner, versionNumber: 2, signals: ["marker"] })
  expect(after.candidates[1]!.reason).toContain("physical source")
  expect(readDocumentReturnMarker(incoming)).toEqual({ documentId: returned.document_id, versionId: returned.id })
  expect(readDocumentVersions(dir, owner).versions[0]!.sha256).toBe(physicalHash)
  expect(readDocumentVersions(dir, target).document!.current_version).toBe(current.id)
  keep(dir, target, await docx("The target's redacted wording."), current.id, "redaction")
  const unavailable = await inspectDocumentReturn(dir, { name: "Later marked copy.docx", bytes: incoming })
  expect(unavailable.kind).toBe("needs-confirmation")
  expect(unavailable.selected).toBeUndefined()
  expect(unavailable.candidates.map(c => c.versionId)).toEqual([returned.id])
  expect(unavailable.warnings[0]).toContain("target or base is unavailable")
})

test("a stale or foreign marker cannot link a file; an unreadable package discloses incomplete coverage", async () => {
  const dir = matter(), bytes = await docx("An unrelated supply provision.")
  const foreign = stampDocumentReturnMarker(bytes, { documentId: randomUUID(), versionId: randomUUID() })
  const result = await inspectDocumentReturn(dir, { name: "foreign.docx", bytes: foreign })
  expect(result.kind).toBe("needs-confirmation")
  expect(result.selected).toBeUndefined()
  expect(result.warnings).toHaveLength(1)
  const bad = await inspectDocumentReturn(dir, { name: "bad.docx", bytes: Buffer.from("not a Word archive") })
  expect(bad.kind).toBe("needs-confirmation")
  expect(bad.complete).toBe(false)
})
