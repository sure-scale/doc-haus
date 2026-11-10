import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { stripWordLink } from "../../../dochaus/lib/word-linked-docx"
import { withSentDraftNumber } from "../../../dochaus/lib/page-furniture"
import { openDb } from "./db"
import { readDraftingPreferences } from "./preferences"
import { oneLine, readLinkedDocumentBranch, readVersion, readVersionBytes, sha256Hex, storeVersionPayload, type LinkedDocumentBranch, type VersionEventInput } from "./document-versions"
import { stampDocumentReturnMarker, type DocumentExchangeSource, type DocumentReturnMarker } from "./document-return-matching"
export type { DocumentExchangeSource } from "./document-return-matching"

export class DocumentExchangeError extends Error {
  constructor(message: string, public readonly code: string, public readonly status: 400 | 404 | 409 | 410 = 409) { super(message) }
}
export type DocumentExchange = {
  id: number; kind: "sent" | "received"; documentId: string; versionId: string
  party: string | null; actor: string; occurredAt: number; recordedAt: number
  baseVersionId: string | null; requestId: string
  payload: { sha256: string; ext: string; bytes: number } | null
  draftNumber?: number; source?: DocumentExchangeSource; match?: { signal: string; reason: string }
  baseCorrections?: { requestId: string; fromVersionId: string | null; toVersionId: string; actor: string; recordedAt: number; priorMatch?: { signal: string; reason: string } }[]
  association?: { documentId: string; baseVersionId: string }
  associationCorrections?: { requestId: string; from: { documentId: string; baseVersionId: string | null }; to: { documentId: string; baseVersionId: string }; actor: string; recordedAt: number; priorMatch?: { signal: string; reason: string } }[]
  linkedBranch?: LinkedDocumentBranch
}
export type PreparedDocumentDownload = {
  documentId: string; name: string; versionId: string; versionNumber: number
  baseVersionId: string | null; bytes: Uint8Array; payloadSha256: string; ext: string
  marker: DocumentReturnMarker | null
  draftNumber?: number
}
type EventRow = { id: number; document_id: string; version_id: string; kind: "sent" | "received"; actor: string; detail: string; created_at: number }
type Detail = Omit<DocumentExchange, "id" | "kind" | "documentId" | "versionId" | "actor" | "recordedAt" | "linkedBranch">

function available(matterDir: string, versionId: string) {
  const found = readVersion(matterDir, versionId)
  if (!found || found.document.removed_at != null) throw new DocumentExchangeError("That retained document version is unavailable.", "version_missing", 404)
  if (found.version.removed) throw new DocumentExchangeError("That version was removed by redaction.", "removed_by_redaction", 410)
  if (!readVersionBytes(matterDir, found.version)) throw new DocumentExchangeError("The retained version bytes are missing or damaged.", "version_damaged")
  return found
}
function baseOf(matterDir: string, documentId: string, baseVersionId?: string | null) {
  if (!baseVersionId) return null
  const base = available(matterDir, baseVersionId)
  if (base.document.id !== documentId) throw new DocumentExchangeError("The comparison base belongs to a different document.", "base_document_mismatch", 400)
  return base.version.id
}

/** Producing a copy is read-only. An ordinary internal download records no
 * send and advances no draft count. Hash after removing the Word save capability
 * and adding the stable, noncredential return marker: these are the exact bytes
 * the caller must send to the browser, without another transformation. */
export function prepareDocumentDownload(matterDir: string, input: { versionId: string; baseVersionId?: string; bytes?: Uint8Array; marker?: boolean; draftNumber?: number }): PreparedDocumentDownload {
  const found = available(matterDir, input.versionId)
  const ext = found.version.ext
  const baseVersionId = baseOf(matterDir, found.document.id, input.baseVersionId)
  const source = input.bytes ?? readVersionBytes(matterDir, found.version)!
  const marker = ext === ".docx" && (input.marker ?? readDraftingPreferences().documentMatchingId) ? { documentId: found.document.id, versionId: found.version.id } : null
  let bytes: Uint8Array = source
  if (ext === ".docx") {
    try {
      bytes = stripWordLink(source)
      // Existing draft-line ownership follows actual generated ancestry, never
      // a guessed filename or similar header on an arbitrary uploaded agreement.
      let generated = false, version: ReturnType<typeof readVersion> = found
      const seen = new Set<string>()
      while (version && !seen.has(version.version.id)) {
        seen.add(version.version.id)
        if (version.version.origin === "generated") { generated = true; break }
        version = version.version.parent ? readVersion(matterDir, version.version.parent) : undefined
      }
      if (input.draftNumber !== undefined && generated) bytes = withSentDraftNumber(bytes, input.draftNumber)
      bytes = stampDocumentReturnMarker(bytes, marker)
    }
    catch { throw new DocumentExchangeError("The Word copy could not be prepared safely.", "invalid_word_copy", 400) }
  }
  return { documentId: found.document.id, name: found.document.name, versionId: found.version.id, versionNumber: found.version.number,
    baseVersionId, bytes, payloadSha256: sha256Hex(bytes), ext, marker, ...(input.draftNumber !== undefined ? { draftNumber: input.draftNumber } : {}) }
}

const requestIdOf = (id: string) => {
  if (typeof id !== "string" || !/^[A-Za-z0-9._:-]{1,160}$/.test(id)) throw new DocumentExchangeError("Provide a stable request identifier for this send or receipt.", "invalid_request_id", 400)
  return id
}
const timestamp = (value?: number) => {
  if (value === undefined) return Date.now()
  if (!Number.isSafeInteger(value) || value < 0 || value > 8_640_000_000_000_000) throw new DocumentExchangeError("The recorded send or receipt date is invalid.", "invalid_exchange_date", 400)
  return value
}
function sourceOf(source?: DocumentExchangeSource): DocumentExchangeSource | undefined {
  if (!source) return
  const clean = Object.fromEntries((["documentName", "attachmentId", "sender", "date"] as const).flatMap(key => {
    const value = source[key]
    if (value !== undefined && typeof value !== "string") throw new DocumentExchangeError("The receipt source must contain text references.", "invalid_exchange_source", 400)
    return value && oneLine(value) ? [[key, oneLine(value)]] : []
  }))
  return Object.keys(clean).length ? clean : undefined
}
function matchOf(match?: { signal: string; reason: string }) {
  if (!match) return
  if (typeof match.signal !== "string" || typeof match.reason !== "string") throw new DocumentExchangeError("The match must name its evidence and reason.", "invalid_exchange_match", 400)
  return { signal: oneLine(match.signal), reason: oneLine(match.reason) }
}
const view = (row: EventRow): DocumentExchange => ({ ...JSON.parse(row.detail) as Detail, id: row.id, kind: row.kind,
  documentId: row.document_id, versionId: row.version_id, actor: row.actor, recordedAt: row.created_at })
const originalRequestRow = (db: Database, requestId: string) => db.query(
  `SELECT * FROM document_version_events WHERE kind IN ('sent','received') AND json_valid(detail) AND json_extract(detail,'$.requestId')=? ORDER BY id LIMIT 1`,
).get(requestId) as EventRow | null
const correctionRequestRow = (db: Database, requestId: string) => db.query(`SELECT e.id,j.value
  FROM document_version_events e,json_each(e.detail,'$.baseCorrections') j
  WHERE e.kind='received' AND json_valid(e.detail) AND json_extract(j.value,'$.requestId')=?
  UNION ALL SELECT e.id,j.value FROM document_version_events e,json_each(e.detail,'$.associationCorrections') j
  WHERE e.kind='received' AND json_valid(e.detail) AND json_extract(j.value,'$.requestId')=? LIMIT 1`).get(requestId, requestId) as { id: number; value: string } | null
function requestRow(db: Database, requestId: string) {
  if (correctionRequestRow(db, requestId)) throw new DocumentExchangeError("This request identifier already records a received-base correction.", "request_conflict")
  return originalRequestRow(db, requestId)
}

function checkReplay(row: EventRow, kind: DocumentExchange["kind"], documentId: string, versionId: string, detail: Detail) {
  const prior = view(row)
  // An omitted historical date on a retry means the same action; use its first
  // recorded date. Caller-supplied differing dates are compared by the caller.
  const key = (d: Detail) => JSON.stringify([d.party, d.baseVersionId, d.payload, d.source ?? null, d.match ?? null])
  if (row.kind !== kind || row.document_id !== documentId || row.version_id !== versionId || key(prior) !== key(detail))
    throw new DocumentExchangeError("This request identifier already records a different send or receipt.", "request_conflict")
  return prior
}
function revalidate(db: Database, documentId: string, versionId: string, baseVersionId: string | null) {
  const live = db.query(`SELECT v.document_id,v.removed,d.removed_at FROM document_versions v JOIN versioned_documents d ON d.id=v.document_id WHERE v.id=?`).get(versionId) as
    { document_id: string; removed: string | null; removed_at: number | null } | null
  if (!live || live.document_id !== documentId || live.removed_at != null) throw new DocumentExchangeError("That document version is no longer available.", "version_missing", 404)
  if (live.removed) throw new DocumentExchangeError("That version was removed by redaction.", "removed_by_redaction", 410)
  if (baseVersionId) {
    const base = db.query("SELECT document_id,removed FROM document_versions WHERE id=?").get(baseVersionId) as { document_id: string; removed: string | null } | null
    if (!base || base.document_id !== documentId) throw new DocumentExchangeError("The selected base is unavailable for this document.", "base_document_mismatch", 400)
    if (base.removed) throw new DocumentExchangeError("The selected base was removed by redaction.", "removed_by_redaction", 410)
  }
}
function insert(db: Database, documentId: string, versionId: string, kind: DocumentExchange["kind"], actor: string | undefined, detail: Detail) {
  const result = db.query("INSERT INTO document_version_events(document_id,version_id,kind,actor,detail,created_at) VALUES(?,?,?,?,?,?) RETURNING *")
    .get(documentId, versionId, kind, oneLine(actor) || "unknown", JSON.stringify(detail), Date.now()) as EventRow
  return view(result)
}

/** Only this explicit action counts as a send. Repeating its requestId returns
 * the same event/draft number; conflicting reuse is refused across the matter. */
export function registerDocumentSent(matterDir: string, input: { prepared: PreparedDocumentDownload; party?: string; actor?: string; sentAt?: number; requestId: string }) {
  const p = input.prepared, found = available(matterDir, p.versionId)
  if (found.document.id !== p.documentId || found.version.ext !== p.ext || sha256Hex(p.bytes) !== p.payloadSha256)
    throw new DocumentExchangeError("The prepared copy does not match this document or its exact payload hash.", "prepared_copy_changed")
  baseOf(matterDir, p.documentId, p.baseVersionId)
  const detail: Detail = { requestId: requestIdOf(input.requestId), party: oneLine(input.party) || null, occurredAt: timestamp(input.sentAt),
    baseVersionId: p.baseVersionId, payload: { sha256: p.payloadSha256, ext: p.ext, bytes: p.bytes.length } }
  const db = openDb(matterDir)
  try {
    const prior = requestRow(db, detail.requestId)
    if (prior) {
      const exchange = checkReplay(prior, "sent", p.documentId, p.versionId, detail)
      if (input.sentAt !== undefined && exchange.occurredAt !== input.sentAt) throw new DocumentExchangeError("This send request already records a different date.", "request_conflict")
      return { exchange, draftNumber: exchange.draftNumber! }
    }
    // Keep filesystem writes outside the short registry write transaction.
    storeVersionPayload(matterDir, p.documentId, p.payloadSha256, p.ext, p.bytes)
    return db.transaction(() => {
      revalidate(db, p.documentId, p.versionId, p.baseVersionId)
      const raced = requestRow(db, detail.requestId)
      if (raced) {
        const exchange = checkReplay(raced, "sent", p.documentId, p.versionId, detail)
        return { exchange, draftNumber: exchange.draftNumber! }
      }
      const count = (db.query("SELECT COUNT(*) AS n FROM document_version_events WHERE document_id=? AND kind='sent'").get(p.documentId) as { n: number }).n
      if (p.draftNumber !== undefined && p.draftNumber !== count + 1)
        throw new DocumentExchangeError("Another send was recorded while this copy was prepared. Prepare the next numbered copy again.", "sent_count_changed")
      detail.draftNumber = count + 1
      const exchange = insert(db, p.documentId, p.versionId, "sent", input.actor, detail)
      return { exchange, draftNumber: detail.draftNumber }
    }).immediate()
  } finally { db.close() }
}

/** A receipt attaches to the version the upload actually saved. Its parent is
 * selected by the upload writer; recording a receipt never overwrites ancestry,
 * the working file, editing ownership or a document version. */
export function recordDocumentReceived(matterDir: string, input: {
  versionId: string; baseVersionId?: string; party?: string; actor?: string; receivedAt?: number; requestId: string
  source?: DocumentExchangeSource; match?: { signal: string; reason: string }
}): DocumentExchange {
  const found = available(matterDir, input.versionId)
  const baseVersionId = baseOf(matterDir, found.document.id, input.baseVersionId)
  if (baseVersionId && found.version.parent !== baseVersionId)
    throw new DocumentExchangeError("The receipt base must be the parent selected when this version was saved.", "received_parent_mismatch")
  const source = sourceOf(input.source)
  const match = matchOf(input.match)
  const detail: Detail = { requestId: requestIdOf(input.requestId), party: oneLine(input.party) || null, occurredAt: timestamp(input.receivedAt), baseVersionId, payload: null,
    ...(source ? { source } : {}), ...(match ? { match } : {}) }
  const db = openDb(matterDir)
  try {
    return db.transaction(() => {
      revalidate(db, found.document.id, input.versionId, baseVersionId)
      const prior = requestRow(db, detail.requestId)
      if (prior) {
        const exchange = checkReplay(prior, "received", found.document.id, input.versionId, detail)
        if (input.receivedAt !== undefined && exchange.occurredAt !== input.receivedAt) throw new DocumentExchangeError("This receipt request already records a different date.", "request_conflict")
        return exchange
      }
      return insert(db, found.document.id, input.versionId, "received", input.actor, detail)
    }).immediate()
  } finally { db.close() }
}

/** The upload writer attaches this event to its new version in its existing
 * registry transaction. This avoids a saved file with a missing receipt after
 * a process stop between two writes. The caller rechecks request identity under
 * the existing source-corpus lock before writing any returned bytes. */
export function prepareDocumentReceivedEvent(matterDir: string, input: {
  baseVersionId: string; party?: string; actor?: string; receivedAt?: number; requestId: string
  source?: DocumentExchangeSource; match?: { signal: string; reason: string }
}): VersionEventInput {
  const base = available(matterDir, input.baseVersionId)
  const source = sourceOf(input.source), match = matchOf(input.match)
  const detail: Detail = { requestId: requestIdOf(input.requestId), party: oneLine(input.party) || null,
    occurredAt: timestamp(input.receivedAt), baseVersionId: base.version.id, payload: null,
    ...(source ? { source } : {}), ...(match ? { match } : {}) }
  return { kind: "received", actor: oneLine(input.actor) || "unknown", detail }
}

/** An explicit correction changes identity metadata only. The saved/working
 * bytes, version number and every other version stay put. A derived output's
 * inherited source basis follows the corrected base, without pretending that
 * the received content was freshly generated or checked.
 * Restrict to a single receipt so one party's correction cannot silently alter
 * the ancestry asserted by another receipt of the same exact saved version. */
export function relinkDocumentReceivedBase(matterDir: string, input: {
  exchangeId: number; baseVersionId: string; expectedBaseVersionId: string | null; requestId: string; actor?: string
}): { exchange: DocumentExchange; changed: boolean } {
  const linked = readLinkedDocumentBranch(matterDir, input.exchangeId)
  if (linked) return relinkDocumentReceivedDocument(matterDir, { ...input, targetDocumentId: linked.targetDocumentId, expectedDocumentId: linked.targetDocumentId })
  const requestId = requestIdOf(input.requestId)
  if (!Number.isSafeInteger(input.exchangeId) || input.exchangeId < 1) throw new DocumentExchangeError("That receipt is unavailable.", "exchange_missing", 404)
  const db = openDb(matterDir)
  try {
    return db.transaction(() => {
      const row = db.query("SELECT * FROM document_version_events WHERE id=? AND kind='received'").get(input.exchangeId) as EventRow | null
      if (!row) throw new DocumentExchangeError("That receipt is unavailable.", "exchange_missing", 404)
      const exchange = view(row)
      available(matterDir, exchange.versionId)
      baseOf(matterDir, exchange.documentId, input.baseVersionId)
      revalidate(db, exchange.documentId, exchange.versionId, input.baseVersionId)
      const priorRequest = originalRequestRow(db, requestId)
      const correction = correctionRequestRow(db, requestId)
      if (priorRequest || correction && (correction.id !== exchange.id || JSON.parse(correction.value).toVersionId !== input.baseVersionId))
        throw new DocumentExchangeError("This correction request already identifies a different action.", "request_conflict")
      if (correction) return { exchange, changed: false }
      const received = db.query("SELECT COUNT(*) AS n FROM document_version_events WHERE kind='received' AND version_id=?").get(exchange.versionId) as { n: number }
      if (received.n !== 1) throw new DocumentExchangeError("This version has multiple receipts. Their ancestry needs a joint correction.", "relink_multiple_receipts")
      const version = db.query("SELECT parent FROM document_versions WHERE id=?").get(exchange.versionId) as { parent: string | null }
      if (exchange.baseVersionId !== input.expectedBaseVersionId || version.parent !== input.expectedBaseVersionId)
        throw new DocumentExchangeError("The received base changed. Reload its history before correcting it.", "received_base_changed")
      // Following the proposed parent's chain catches self/descendant bases and
      // existing corruption; a parent is never chosen by version number alone.
      const seen = new Set<string>([exchange.versionId])
      let ancestor: string | null = input.baseVersionId
      while (ancestor) {
        if (seen.has(ancestor)) throw new DocumentExchangeError("That base would create a version ancestry cycle.", "received_base_cycle", 400)
        seen.add(ancestor)
        const parent = db.query("SELECT document_id,parent FROM document_versions WHERE id=?").get(ancestor) as { document_id: string; parent: string | null } | null
        if (!parent || parent.document_id !== exchange.documentId) throw new DocumentExchangeError("The selected base ancestry is unavailable for this document.", "base_document_mismatch", 400)
        ancestor = parent.parent
      }
      const changed = exchange.baseVersionId !== input.baseVersionId
      const detail = JSON.parse(row.detail) as Detail
      detail.baseCorrections = [...(detail.baseCorrections ?? []), { requestId, fromVersionId: exchange.baseVersionId, toVersionId: input.baseVersionId,
        actor: oneLine(input.actor) || "unknown", recordedAt: Date.now(), ...(detail.match ? { priorMatch: detail.match } : {}) }]
      detail.baseVersionId = input.baseVersionId
      detail.match = { signal: "lawyer-corrected", reason: "The lawyer explicitly corrected this received file's base version." }
      db.run("UPDATE document_versions SET parent=?,source_versions=(SELECT source_versions FROM document_versions WHERE id=?) WHERE id=?", [input.baseVersionId, input.baseVersionId, exchange.versionId])
      db.run("UPDATE document_version_events SET detail=? WHERE id=?", [JSON.stringify(detail), exchange.id])
      return { exchange: view({ ...row, detail: JSON.stringify(detail) }), changed }
    }).immediate()
  } finally { db.close() }
}

/** Correct logical association by projecting one immutable received artifact
 * as a branch of the chosen document. Physical owner/version, retained bytes
 * and both current working copies are preserved. Moving back to the physical
 * owner restores native ancestry instead of creating a duplicate projection. */
export function relinkDocumentReceivedDocument(matterDir: string, input: {
  exchangeId: number; targetDocumentId: string; baseVersionId: string; expectedDocumentId: string; expectedBaseVersionId: string | null
  requestId: string; actor?: string; expectedSourceCurrentVersionId?: string | null; expectedTargetCurrentVersionId?: string | null
}): { exchange: DocumentExchange; changed: boolean } {
  const requestId = requestIdOf(input.requestId)
  if (!Number.isSafeInteger(input.exchangeId) || input.exchangeId < 1) throw new DocumentExchangeError("That receipt is unavailable.", "exchange_missing", 404)
  const db = openDb(matterDir)
  try {
    const result = db.transaction(() => {
      const row = db.query("SELECT * FROM document_version_events WHERE id=? AND kind='received'").get(input.exchangeId) as EventRow | null
      if (!row) throw new DocumentExchangeError("That receipt is unavailable.", "exchange_missing", 404)
      const exchange = view(row), source = available(matterDir, exchange.versionId)
      if (source.document.id !== exchange.documentId) throw new DocumentExchangeError("The received artifact's owner changed.", "received_identity_changed")
      const base = available(matterDir, input.baseVersionId)
      if (base.document.id !== input.targetDocumentId) throw new DocumentExchangeError("Choose a retained base of the target document.", "base_document_mismatch", 400)
      const priorRequest = originalRequestRow(db, requestId), correction = correctionRequestRow(db, requestId)
      const savedCorrection = correction ? JSON.parse(correction.value) as { to?: { documentId: string; baseVersionId: string } } : undefined
      if (priorRequest || correction && (correction.id !== exchange.id || savedCorrection?.to?.documentId !== input.targetDocumentId || savedCorrection.to.baseVersionId !== input.baseVersionId))
        throw new DocumentExchangeError("This correction request already identifies a different action.", "request_conflict")
      if (correction) return { exchange, changed: false }
      const count = db.query("SELECT COUNT(*) AS n FROM document_version_events WHERE kind='received' AND version_id=?").get(exchange.versionId) as { n: number }
      if (count.n !== 1) throw new DocumentExchangeError("This version has multiple receipts. Their association needs a joint correction.", "relink_multiple_receipts")
      const from = exchange.association ?? { documentId: exchange.documentId, baseVersionId: exchange.baseVersionId }
      if (from.documentId !== input.expectedDocumentId || from.baseVersionId !== input.expectedBaseVersionId ||
        input.expectedSourceCurrentVersionId !== undefined && source.document.current_version !== input.expectedSourceCurrentVersionId ||
        input.expectedTargetCurrentVersionId !== undefined && base.document.current_version !== input.expectedTargetCurrentVersionId)
        throw new DocumentExchangeError("The document association or working version changed. Reload before correcting it.", "received_association_changed")
      revalidate(db, source.document.id, exchange.versionId, null)
      revalidate(db, base.document.id, base.version.id, null)
      const to = { documentId: base.document.id, baseVersionId: base.version.id }
      const changed = from.documentId !== to.documentId || from.baseVersionId !== to.baseVersionId
      const detail = JSON.parse(row.detail) as Detail
      detail.associationCorrections = [...(detail.associationCorrections ?? []), { requestId, from, to, actor: oneLine(input.actor) || "unknown", recordedAt: Date.now(), ...(detail.match ? { priorMatch: detail.match } : {}) }]
      if (to.documentId === source.document.id) {
        const seen = new Set<string>([exchange.versionId])
        let ancestor: string | null = to.baseVersionId
        while (ancestor) {
          if (seen.has(ancestor)) throw new DocumentExchangeError("That base would create a version ancestry cycle.", "received_base_cycle", 400)
          seen.add(ancestor)
          const parent = db.query("SELECT document_id,parent FROM document_versions WHERE id=?").get(ancestor) as { document_id: string; parent: string | null } | null
          if (!parent || parent.document_id !== source.document.id) throw new DocumentExchangeError("The selected base ancestry is unavailable.", "base_document_mismatch", 400)
          ancestor = parent.parent
        }
        delete detail.association
        detail.baseVersionId = to.baseVersionId
        db.run("UPDATE document_versions SET parent=?,source_versions=(SELECT source_versions FROM document_versions WHERE id=?) WHERE id=?", [to.baseVersionId, to.baseVersionId, exchange.versionId])
      } else detail.association = to
      detail.match = { signal: "lawyer-corrected", reason: "The lawyer explicitly corrected this received file's document and base association. Neither working copy was replaced." }
      const serialized = JSON.stringify(detail)
      db.run("UPDATE document_version_events SET detail=? WHERE id=?", [serialized, exchange.id])
      return { exchange: view({ ...row, detail: serialized }), changed }
    }).immediate()
    const linkedBranch = readLinkedDocumentBranch(matterDir, result.exchange.id)
    return { ...result, exchange: { ...result.exchange, ...(linkedBranch ? { linkedBranch } : {}) } }
  } finally { db.close() }
}

export function readDocumentReceivedBranch(matterDir: string, eventId: number): { branch: LinkedDocumentBranch; bytes: Uint8Array } {
  if (!Number.isSafeInteger(eventId) || eventId < 1) throw new DocumentExchangeError("That linked received branch is unavailable.", "exchange_missing", 404)
  const branch = readLinkedDocumentBranch(matterDir, eventId)
  if (!branch) throw new DocumentExchangeError("That receipt has no linked document branch.", "linked_branch_missing", 404)
  if (!branch.available) throw new DocumentExchangeError("The linked received branch is unavailable because its document, base or retained bytes were removed or changed.", "linked_branch_unavailable", branch.unavailableReason?.includes("removed") ? 410 : 409)
  const source = available(matterDir, branch.versionId)
  if (source.document.id !== branch.sourceDocumentId) throw new DocumentExchangeError("The received artifact's owner changed.", "received_identity_changed")
  baseOf(matterDir, branch.targetDocumentId, branch.baseVersionId)
  return { branch, bytes: readVersionBytes(matterDir, source.version)! }
}

function readEvents<T>(matterDir: string, read: (db: Database) => T, empty: T): T {
  const file = path.join(matterDir, ".dochaus", "legal.db")
  if (!existsSync(file)) return empty
  const db = new Database(file, { readonly: true })
  try {
    if (!db.query("SELECT name FROM sqlite_master WHERE type='table' AND name='document_version_events'").get()) return empty
    return read(db)
  } finally { db.close() }
}
export function readDocumentExchanges(matterDir: string, documentId: string): DocumentExchange[] {
  return readEvents(matterDir, db => (db.query(
    `SELECT e.* FROM document_version_events e JOIN versioned_documents d ON d.id=e.document_id
     WHERE e.kind IN ('sent','received') AND ((e.document_id=? AND d.removed_at IS NULL)
       OR (json_valid(e.detail) AND json_extract(e.detail,'$.association.documentId')=?)) ORDER BY e.id`,
  ).all(documentId, documentId) as EventRow[]).map(row => {
    const exchange = view(row), linkedBranch = exchange.association ? readLinkedDocumentBranch(matterDir, row.id) : undefined
    return { ...exchange, ...(linkedBranch ? { linkedBranch } : {}) }
  }), [])
}
export function readDocumentExchangeRequest(matterDir: string, requestId: string): DocumentExchange | undefined {
  const id = requestIdOf(requestId)
  return readEvents(matterDir, db => {
    const row = requestRow(db, id)
    return row ? view(row) : undefined
  }, undefined)
}
export function nextDocumentDraftNumber(matterDir: string, documentId: string, requestId?: string): number {
  if (requestId) {
    const prior = readDocumentExchangeRequest(matterDir, requestId)
    if (prior?.kind === "sent" && prior.documentId === documentId) return prior.draftNumber!
  }
  return readEvents(matterDir, db => (db.query("SELECT COUNT(*) AS n FROM document_version_events WHERE document_id=? AND kind='sent'").get(documentId) as { n: number }).n + 1, 1)
}
export function readDocumentExchangePayload(matterDir: string, eventId: number): { exchange: DocumentExchange; bytes: Uint8Array; name: string } {
  if (!Number.isSafeInteger(eventId) || eventId < 1) throw new DocumentExchangeError("That send record is unavailable.", "exchange_missing", 404)
  const row = readEvents(matterDir, db => db.query("SELECT * FROM document_version_events WHERE id=? AND kind='sent'").get(eventId) as EventRow | null, null)
  if (!row) throw new DocumentExchangeError("That send record is unavailable.", "exchange_missing", 404)
  const exchange = view(row), found = available(matterDir, exchange.versionId)
  if (found.document.id !== exchange.documentId) throw new DocumentExchangeError("That send record is unavailable.", "exchange_missing", 404)
  baseOf(matterDir, exchange.documentId, exchange.baseVersionId)
  if (!exchange.payload || !/^[a-f0-9]{64}$/.test(exchange.payload.sha256) || !/^\.[a-z0-9]{1,10}$/.test(exchange.payload.ext))
    throw new DocumentExchangeError("The exact sent copy is unavailable.", "exchange_payload_missing", 404)
  const bytes = readVersionBytes(matterDir, { document_id: exchange.documentId, ...exchange.payload })
  if (!bytes || bytes.length !== exchange.payload.bytes) throw new DocumentExchangeError("The exact sent copy is missing or damaged.", "exchange_payload_damaged")
  return { exchange, bytes, name: found.document.name }
}
