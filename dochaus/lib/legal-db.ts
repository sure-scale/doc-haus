import { Database } from "bun:sqlite"
import { mkdirSync } from "node:fs"
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
  const db = new Database(path.join(dir, "legal.db"))
  db.run(`PRAGMA busy_timeout = ${WRITE_BUSY_TIMEOUT_MS}`)
  db.run("PRAGMA journal_mode = WAL")
  return db
}
