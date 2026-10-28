import { Database } from "bun:sqlite"
import { randomUUID } from "node:crypto"
import { existsSync } from "node:fs"
import path from "node:path"
import { canonicalArtifactIdentity } from "../capabilities/artifacts/mutation-lock"
import type { WordLeaseState } from "./word-addin-contract"

/** Server-validated authority for one pane. Never construct this from ordinary
 * writer headers; only the authenticated Word route supplies it. */
export type WordLeaseOwner = { document: string; tokenHash: string; session: string; generation: string }
export type WordLeaseWriteOptions = { wordLeaseOwner?: WordLeaseOwner }
export type WordLeaseIdentity = Omit<WordLeaseOwner, "generation">
export type WordEditLease = { document_id: string; token_hash: string; session: string; generation: string; expires_at: number }
export type WordNativeOperation = { sequence: number; operationId: string | null; state: "idle" | "pending" | "confirmed" | "not-dispatched" }
export type WordProposalBinding = { conversationId: string; turnId: string }
type NativeLease = WordEditLease & { native_conversation: string | null; native_turn: string | null; retired: number; native_sequence: number; native_operation: string | null; native_generation: string | null; native_state: "pending" | "confirmed" | "not-dispatched" | null }

// Longer than the pane's bounded pair of 30-second uploads. A stopped pane
// cannot hold a file indefinitely; expiry never silently grants it a new lease.
export const WORD_EDIT_LEASE_TTL = 120_000

export class EditingLeaseError extends Error {
  readonly status = 423
  readonly code = "word_document_locked"
  constructor(message = "This document is being edited in Word. Finish there or take over editing in doc.haus.", readonly expiresAt: number | null = null) {
    super(message)
    this.name = "EditingLeaseError"
  }
}

export function createWordEditLeaseTable(db: Database) {
  db.run(`CREATE TABLE IF NOT EXISTS word_edit_leases (
    document_id TEXT PRIMARY KEY,
    token_hash TEXT NOT NULL,
    session TEXT NOT NULL,
    generation TEXT NOT NULL,
    expires_at INTEGER NOT NULL
  )`)
  db.run(`CREATE TABLE IF NOT EXISTS word_chat_applied (document_id TEXT NOT NULL, conversation_id TEXT NOT NULL, turn_id TEXT NOT NULL, PRIMARY KEY (document_id, conversation_id, turn_id))`)
  // Serialize upgrades only when needed; ordinary reads must not acquire a
  // database write lock on every open.
  const columns = [["retired", "INTEGER NOT NULL DEFAULT 0"], ["native_sequence", "INTEGER NOT NULL DEFAULT 0"],
    ["native_operation", "TEXT"], ["native_generation", "TEXT"], ["native_state", "TEXT"], ["native_conversation", "TEXT"], ["native_turn", "TEXT"]] as const
  const fields = () => new Set((db.query("PRAGMA table_info(word_edit_leases)").all() as { name: string }[]).map(row => row.name))
  const existing = fields()
  if (columns.some(([name]) => !existing.has(name))) db.transaction(() => {
    const current = fields()
    for (const [name, declaration] of columns)
      if (!current.has(name)) db.run(`ALTER TABLE word_edit_leases ADD COLUMN ${name} ${declaration}`)
  }).immediate()
}

function nativeColumns(db: Database) {
  return (db.query("PRAGMA table_info(word_edit_leases)").all() as { name: string }[]).some(row => row.name === "native_state")
}
function nativeLease(db: Database, documentId: string): NativeLease | null {
  return db.query("SELECT * FROM word_edit_leases WHERE document_id = ?").get(documentId) as NativeLease | null
}
function unresolved(db: Database, documentId: string) {
  return nativeColumns(db) && nativeLease(db, documentId)?.native_state === "pending"
}
const pendingMessage = "A change in Word has not been confirmed. Keep the Word file and inspect the change before continuing."
function assertResolved(db: Database, documentId: string) {
  if (unresolved(db, documentId)) throw new EditingLeaseError(pendingMessage)
}

export function readWordNativeOperation(db: Database, documentId: string): WordNativeOperation {
  const row = nativeLease(db, documentId)
  return { sequence: row?.native_sequence ?? 0, operationId: row?.native_operation ?? null, state: row?.native_state ?? "idle" }
}

/** Caller holds the existing artifact lock. The expected sequence is the
 * value observed before admission; duplicate BEGIN returns the same operation. */
export function wordProposalApplied(db: Database, document: string, proposal: WordProposalBinding): boolean {
  return !!db.query("SELECT 1 FROM word_chat_applied WHERE document_id = ? AND conversation_id = ? AND turn_id = ?").get(document, proposal.conversationId, proposal.turnId)
}

export function beginWordNativeOperation(db: Database, owner: WordLeaseOwner, expectedSequence: number, operationId: string, now = Date.now(), proposal?: WordProposalBinding): WordNativeOperation {
  if (!Number.isSafeInteger(expectedSequence) || expectedSequence < 0 || expectedSequence >= Number.MAX_SAFE_INTEGER ||
    !/^[A-Za-z0-9_-]{1,128}$/.test(operationId)) throw new EditingLeaseError("Invalid Word operation identity.")
  if (proposal && (!/^[A-Za-z0-9_-]{1,160}$/.test(proposal.conversationId) || !/^[A-Za-z0-9_-]{1,160}$/.test(proposal.turnId))) throw new EditingLeaseError("Invalid proposal identity.")
  requireWordEditLease(db, owner, now)
  const row = nativeLease(db, owner.document)!
  if (row.native_operation === operationId && row.native_generation === owner.generation && row.native_sequence === expectedSequence + 1 && row.native_state && row.native_conversation === (proposal?.conversationId ?? null) && row.native_turn === (proposal?.turnId ?? null))
    return readWordNativeOperation(db, owner.document)
  if (proposal && wordProposalApplied(db, owner.document, proposal)) throw new EditingLeaseError("This proposal was already applied. Prepare a new proposal.")
  if (row.native_state === "pending" || row.native_sequence !== expectedSequence || row.native_operation === operationId)
    throw new EditingLeaseError(pendingMessage)
  const result = db.run(`UPDATE word_edit_leases SET native_sequence = native_sequence + 1, native_operation = ?, native_generation = ?, native_state = 'pending', native_conversation = ?, native_turn = ?
    WHERE document_id = ? AND token_hash = ? AND session = ? AND generation = ? AND retired = 0
      AND native_sequence = ? AND COALESCE(native_state, '') != 'pending' AND expires_at > ?
      AND EXISTS (SELECT 1 FROM word_grants g JOIN versioned_documents d ON d.id = g.document_id
        WHERE g.token_hash = word_edit_leases.token_hash AND g.document_id = word_edit_leases.document_id AND g.revoked_at IS NULL AND d.removed_at IS NULL)`,
    [operationId, owner.generation, proposal?.conversationId ?? null, proposal?.turnId ?? null, owner.document, owner.tokenHash, owner.session, owner.generation, expectedSequence, now])
  if (result.changes !== 1) throw new EditingLeaseError(pendingMessage)
  return readWordNativeOperation(db, owner.document)
}

/** Caller holds the existing artifact lock through the update and readback.
 * Fence and applied outcome are committed together in an immediate transaction.
 * Only the original owner can settle a fully observed operation. This does
 * not imply a failed or lost Office response has stopped executing. */
export function settleWordNativeOperation(db: Database, owner: WordLeaseOwner, sequence: number, operationId: string, outcome: "confirmed" | "not-dispatched"): WordNativeOperation {
  if (!Number.isSafeInteger(sequence) || sequence < 1 || !/^[A-Za-z0-9_-]{1,128}$/.test(operationId) || !["confirmed", "not-dispatched"].includes(outcome))
    throw new EditingLeaseError("Invalid Word operation identity.")
  return db.transaction(() => {
  const result = db.run(`UPDATE word_edit_leases SET native_state = ? WHERE document_id = ? AND token_hash = ? AND session = ? AND generation = ?
    AND retired = 0 AND native_sequence = ? AND native_operation = ? AND native_generation = ? AND native_state = 'pending'`,
    [outcome, owner.document, owner.tokenHash, owner.session, owner.generation, sequence, operationId, owner.generation])
  const row = nativeLease(db, owner.document)
  if (result.changes !== 1 && (!row || !sameIdentity(row, owner) || row.generation !== owner.generation || row.retired ||
    row.native_sequence !== sequence || row.native_operation !== operationId || row.native_generation !== owner.generation || row.native_state !== outcome)) throw new EditingLeaseError(pendingMessage)
  if (outcome === "confirmed" && row?.native_conversation && row.native_turn) db.run("INSERT OR IGNORE INTO word_chat_applied (document_id, conversation_id, turn_id) VALUES (?, ?, ?)", [owner.document, row.native_conversation, row.native_turn])
  return readWordNativeOperation(db, owner.document)
  }).immediate()
}

export function requirePendingWordProposal(db: Database, owner: WordLeaseOwner, proposal: WordProposalBinding) {
  requireWordEditLease(db, owner)
  const row = nativeLease(db, owner.document)
  if (!row || row.native_state !== "pending" || row.native_generation !== owner.generation || row.native_conversation !== proposal.conversationId || row.native_turn !== proposal.turnId || wordProposalApplied(db, owner.document, proposal)) throw new EditingLeaseError("Review the active proposal before applying it.")
}

export function readWordEditLease(db: Database, documentId: string): WordEditLease | null {
  return (nativeColumns(db) ? db.query("SELECT * FROM word_edit_leases WHERE document_id = ? AND retired = 0") : db.query("SELECT * FROM word_edit_leases WHERE document_id = ?")).get(documentId) as WordEditLease | null
}

function liveLease(db: Database, documentId: string, now: number) {
  return db.query(`SELECT l.* FROM word_edit_leases l JOIN word_grants g ON g.token_hash = l.token_hash
    JOIN versioned_documents d ON d.id = l.document_id WHERE l.document_id = ? AND l.expires_at > ?
    AND g.document_id = l.document_id AND g.revoked_at IS NULL AND d.removed_at IS NULL`).get(documentId, now) as WordEditLease | null
}

function sameIdentity(lease: WordEditLease, identity: WordLeaseIdentity) {
  return lease.document_id === identity.document && lease.token_hash === identity.tokenHash && lease.session === identity.session
}

export function wordLeaseState(db: Database, documentId: string, identity?: WordLeaseIdentity, now = Date.now()): WordLeaseState {
  if (unresolved(db, documentId)) {
    const pending = nativeLease(db, documentId)!, live = liveLease(db, documentId, now)
    return { state: live && identity && sameIdentity(live, identity) ? "owned" : "other",
      generation: pending.generation, expiresAt: live?.expires_at ?? null }
  }
  const lease = liveLease(db, documentId, now)
  const observed = lease ?? readWordEditLease(db, documentId)
  return { state: lease ? identity && sameIdentity(lease, identity) ? "owned" : "other" : "available",
    ...(observed ? { generation: observed.generation } : {}), expiresAt: lease?.expires_at ?? null }
}

/** Caller holds the document's existing artifact lock for this entire change. */
export function claimWordEditLease(db: Database, identity: WordLeaseIdentity, now = Date.now()): WordLeaseOwner {
  assertResolved(db, identity.document)
  const current = liveLease(db, identity.document, now)
  if (current && !sameIdentity(current, identity)) throw new EditingLeaseError(undefined, current.expires_at)
  const generation = current?.generation ?? randomUUID()
  const result = db.run(`INSERT INTO word_edit_leases (document_id, token_hash, session, generation, expires_at) VALUES (?, ?, ?, ?, ?)
    ON CONFLICT(document_id) DO UPDATE SET token_hash=excluded.token_hash, session=excluded.session,
    generation=excluded.generation, expires_at=excluded.expires_at, retired=0
    WHERE COALESCE(word_edit_leases.native_state, '') != 'pending'`,
  [identity.document, identity.tokenHash, identity.session, generation, now + WORD_EDIT_LEASE_TTL])
  if (result.changes !== 1) throw new EditingLeaseError(pendingMessage)
  return { ...identity, generation }
}

export function requireWordEditLease(db: Database, owner: WordLeaseOwner, now = Date.now()) {
  const lease = liveLease(db, owner.document, now)
  if (!lease || !sameIdentity(lease, owner) || lease.generation !== owner.generation)
    throw new EditingLeaseError("This Word editing session no longer owns the document. Keep your local file and check editing status before taking over.", lease?.expires_at ?? null)
  return lease
}

export function renewWordEditLease(db: Database, owner: WordLeaseOwner, now = Date.now()) {
  // A native operation cannot expire out from under its original owner.
  // That owner may renew after TTL if its grant and document are still valid.
  if (unresolved(db, owner.document)) {
    const result = db.run(`UPDATE word_edit_leases SET expires_at = ? WHERE document_id = ? AND token_hash = ? AND session = ? AND generation = ?
      AND retired = 0 AND native_state = 'pending' AND native_generation = ?
      AND EXISTS (SELECT 1 FROM word_grants g JOIN versioned_documents d ON d.id = g.document_id
        WHERE g.token_hash = word_edit_leases.token_hash AND g.document_id = word_edit_leases.document_id AND g.revoked_at IS NULL AND d.removed_at IS NULL)`,
      [now + WORD_EDIT_LEASE_TTL, owner.document, owner.tokenHash, owner.session, owner.generation, owner.generation])
    if (result.changes !== 1) throw new EditingLeaseError(pendingMessage)
    return
  }
  requireWordEditLease(db, owner, now)
  db.run("UPDATE word_edit_leases SET expires_at = ? WHERE document_id = ? AND generation = ?", [now + WORD_EDIT_LEASE_TTL, owner.document, owner.generation])
}

export function releaseWordEditLease(db: Database, owner: WordLeaseOwner) {
  assertResolved(db, owner.document)
  const lease = readWordEditLease(db, owner.document)
  if (!lease || !sameIdentity(lease, owner) || lease.generation !== owner.generation)
    throw new EditingLeaseError("This Word editing session no longer owns the document. Its replacement was left unchanged.")
  // Retain the fence for an idempotent lost Done acknowledgement. A new claim
  // always gets a new generation once expiry is zero.
  const result = db.run("UPDATE word_edit_leases SET expires_at = 0 WHERE document_id = ? AND token_hash = ? AND session = ? AND generation = ? AND COALESCE(native_state, '') != 'pending'", [owner.document, owner.tokenHash, owner.session, owner.generation])
  if (result.changes !== 1) throw new EditingLeaseError(pendingMessage)
}

export function clearWordEditLease(db: Database, documentId: string) {
  assertResolved(db, documentId)
  const result = db.run("UPDATE word_edit_leases SET retired = 1, expires_at = 0 WHERE document_id = ? AND COALESCE(native_state, '') != 'pending'", [documentId])
  if (result.changes !== 1) assertResolved(db, documentId)
}

/** Enforced inside the existing artifact critical section, before bytes or
 * approval state change. A missing owner means an ordinary browser/tool/lean
 * write; a supplied expired owner must never degrade into an ordinary write. */
export function assertWordEditAllowed(directory: string, document: string, owner?: WordLeaseOwner, now = Date.now()) {
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) {
    if (owner) throw new EditingLeaseError("This Word document is no longer registered.")
    return
  }
  const db = new Database(file, { readonly: true })
  try {
    if (!db.query("SELECT 1 FROM sqlite_master WHERE type='table' AND name='word_edit_leases'").get()) {
      if (owner) throw new EditingLeaseError("This Word editing session is no longer available.")
      return
    }
    const key = path.basename(canonicalArtifactIdentity(directory, document))
    const target = db.query("SELECT id FROM versioned_documents WHERE name_key = ? AND removed_at IS NULL").get(key) as { id: string } | null
    if (target) assertResolved(db, target.id)
    if (owner) {
      if (target?.id !== owner.document) throw new EditingLeaseError("This Word document was deleted or replaced. Keep your local file and open its current copy from doc.haus.")
      requireWordEditLease(db, owner, now)
    } else if (target) {
      const lease = liveLease(db, target.id, now)
      if (lease) throw new EditingLeaseError(undefined, lease.expires_at)
    }
  } finally { db.close() }
}

export function assertMatterWordEditsAllowed(directory: string, now = Date.now()) {
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) return
  const db = new Database(file, { readonly: true })
  try {
    if (!db.query("SELECT 1 FROM sqlite_master WHERE type='table' AND name='word_edit_leases'").get()) return
    if (nativeColumns(db) && db.query("SELECT 1 FROM word_edit_leases WHERE native_state = 'pending' LIMIT 1").get())
      throw new EditingLeaseError(pendingMessage)
    const lease = db.query(`SELECT l.expires_at FROM word_edit_leases l JOIN word_grants g ON g.token_hash=l.token_hash
      JOIN versioned_documents d ON d.id=l.document_id WHERE l.expires_at > ? AND g.document_id=l.document_id AND g.revoked_at IS NULL
      AND d.removed_at IS NULL LIMIT 1`).get(now) as { expires_at: number } | null
    if (lease) throw new EditingLeaseError("A document in this matter is being edited in Word. Finish or release that editing session before deleting the matter.", lease.expires_at)
  } finally { db.close() }
}
