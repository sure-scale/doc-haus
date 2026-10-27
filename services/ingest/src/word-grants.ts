import type { Database } from "bun:sqlite"
import { createWordEditLeaseTable } from "../../../dochaus/lib/word-edit-lease"

/** Linked-document capabilities share the existing per-matter registry, never
 * a second document/version store. They allow the document's guarded writes
 * and the Source Desk's scoped matter reads. Only a SHA-256 digest persists in
 * this registry; protected transport DOCX copies retain their embedded link. */
export function createWordGrantTables(db: Database) {
  createWordEditLeaseTable(db)
  db.run(`CREATE TABLE IF NOT EXISTS word_grants (
    token_hash TEXT PRIMARY KEY,
    document_id TEXT NOT NULL,
    origin TEXT NOT NULL,
    issued_version TEXT NOT NULL,
    created_at INTEGER NOT NULL,
    revoked_at INTEGER,
    overlap INTEGER NOT NULL DEFAULT 0
  )`)
  db.run("CREATE INDEX IF NOT EXISTS word_grants_document ON word_grants(document_id)")
  db.run(`CREATE TABLE IF NOT EXISTS word_sessions (
    token_hash TEXT NOT NULL,
    session TEXT NOT NULL,
    seen_at INTEGER NOT NULL,
    done INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (token_hash, session)
  )`)
}

export type WordGrant = { token_hash: string; document_id: string; origin: string; issued_version: string; revoked_at: number | null; overlap: number }
// Allow the pane's bounded pair of 30-second upload attempts to finish before
// its serialized heartbeat resumes. This is a contact hint, never a file lock.
export const WORD_HEARTBEAT_TTL = 90_000

export function activeWordSessions(db: Database, documentId: string) {
  return (db.query(`SELECT COUNT(DISTINCT s.session) AS n FROM word_sessions s
    JOIN word_grants g ON g.token_hash = s.token_hash WHERE g.document_id = ? AND g.revoked_at IS NULL
    AND s.done = 0 AND s.seen_at > ?`).get(documentId, Date.now() - WORD_HEARTBEAT_TTL) as { n: number }).n
}
