import { Database } from "bun:sqlite"
import { existsSync, mkdirSync, realpathSync } from "node:fs"
import path from "node:path"

// How long an engine write to legal.db waits for another connection's write
// lock. The ingest service holds that lock while it commits an upload or a
// startup backfill; it runs in its own process and finishes while this one
// waits, so a short wait turns an at-once SQLITE_BUSY into a write. bun:sqlite
// waits synchronously and blocks this process meanwhile, so the wait stays
// short and a longer commit still fails as SQLITE_BUSY.
const WRITE_BUSY_TIMEOUT_MS = 2000

/** A read-write connection to a matter's legal.db for the engine's own rows
 * (redlines, redactions, key dates). */
export function openLegalDbForWrite(matterDir: string) {
  const dir = path.join(matterDir, ".dochaus")
  mkdirSync(dir, { recursive: true })
  return prepared(new Database(path.join(dir, "legal.db")), (db) => {
    db.run(`PRAGMA busy_timeout = ${WRITE_BUSY_TIMEOUT_MS}`)
    db.run("PRAGMA journal_mode = WAL")
  })
}

/** `db` after `setup`, or closed when setup throws (a lock wait that ran out),
 * so a failed open never leaks the connection. */
export function prepared(db: Database, setup: (db: Database) => void) {
  try {
    setup(db)
    return db
  } catch (error) {
    db.close()
    throw error
  }
}

/** Whether `document`, a Word file in the matter folder, is registered in the
 * matter's legal.db as an uploaded source (ingest's role "source"). */
export function registeredWordSource(directory: string, document: string) {
  if (path.basename(document) !== document || !/\.docx$/i.test(document)) return false
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) return false
  const relative = path.relative(realpathSync(directory), realpathSync(file))
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) return false
  const db = new Database(file, { readonly: true })
  try {
    // The engine creates legal.db for its own rows before any upload registers a document.
    if (!db.query("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'documents'").get()) return false
    return Boolean(db.query("SELECT 1 FROM documents WHERE name = ? AND role = 'source'").get(document))
  }
  finally { db.close() }
}
