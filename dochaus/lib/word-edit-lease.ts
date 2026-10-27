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
}

export function readWordEditLease(db: Database, documentId: string): WordEditLease | null {
  return db.query("SELECT * FROM word_edit_leases WHERE document_id = ?").get(documentId) as WordEditLease | null
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
  const lease = liveLease(db, documentId, now)
  const observed = lease ?? readWordEditLease(db, documentId)
  return { state: lease ? identity && sameIdentity(lease, identity) ? "owned" : "other" : "available",
    ...(observed ? { generation: observed.generation } : {}), expiresAt: lease?.expires_at ?? null }
}

/** Caller holds the document's existing artifact lock for this entire change. */
export function claimWordEditLease(db: Database, identity: WordLeaseIdentity, now = Date.now()): WordLeaseOwner {
  const current = liveLease(db, identity.document, now)
  if (current && !sameIdentity(current, identity)) throw new EditingLeaseError(undefined, current.expires_at)
  const generation = current?.generation ?? randomUUID()
  db.run(`INSERT INTO word_edit_leases (document_id, token_hash, session, generation, expires_at) VALUES (?, ?, ?, ?, ?)
    ON CONFLICT(document_id) DO UPDATE SET token_hash=excluded.token_hash, session=excluded.session,
    generation=excluded.generation, expires_at=excluded.expires_at`,
  [identity.document, identity.tokenHash, identity.session, generation, now + WORD_EDIT_LEASE_TTL])
  return { ...identity, generation }
}

export function requireWordEditLease(db: Database, owner: WordLeaseOwner, now = Date.now()) {
  const lease = liveLease(db, owner.document, now)
  if (!lease || !sameIdentity(lease, owner) || lease.generation !== owner.generation)
    throw new EditingLeaseError("This Word editing session no longer owns the document. Keep your local file and check editing status before taking over.", lease?.expires_at ?? null)
  return lease
}

export function renewWordEditLease(db: Database, owner: WordLeaseOwner, now = Date.now()) {
  requireWordEditLease(db, owner, now)
  db.run("UPDATE word_edit_leases SET expires_at = ? WHERE document_id = ? AND generation = ?", [now + WORD_EDIT_LEASE_TTL, owner.document, owner.generation])
}

export function releaseWordEditLease(db: Database, owner: WordLeaseOwner) {
  const lease = readWordEditLease(db, owner.document)
  if (!lease || !sameIdentity(lease, owner) || lease.generation !== owner.generation)
    throw new EditingLeaseError("This Word editing session no longer owns the document. Its replacement was left unchanged.")
  // Retain the fence for an idempotent lost Done acknowledgement. A new claim
  // always gets a new generation once expiry is zero.
  db.run("UPDATE word_edit_leases SET expires_at = 0 WHERE document_id = ? AND generation = ?", [owner.document, owner.generation])
}

export function clearWordEditLease(db: Database, documentId: string) {
  db.run("DELETE FROM word_edit_leases WHERE document_id = ?", [documentId])
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
    const lease = db.query(`SELECT l.expires_at FROM word_edit_leases l JOIN word_grants g ON g.token_hash=l.token_hash
      JOIN versioned_documents d ON d.id=l.document_id WHERE l.expires_at > ? AND g.document_id=l.document_id AND g.revoked_at IS NULL
      AND d.removed_at IS NULL LIMIT 1`).get(now) as { expires_at: number } | null
    if (lease) throw new EditingLeaseError("A document in this matter is being edited in Word. Finish or release that editing session before deleting the matter.", lease.expires_at)
  } finally { db.close() }
}
