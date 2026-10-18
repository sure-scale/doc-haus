import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { canonicalArtifactIdentity } from "../capabilities/artifacts/mutation-lock"
import { sha256File } from "./lineage"

// The sha256 of the bytes the system itself last wrote to a generated Word file
// (a run's registration of a deliverable, ingest's rebuild of a saved combined
// file), stored in the registry transaction that writes them. Every other write
// of the file removes it. The file is the system's to write over only while it
// holds exactly those bytes: equal bytes without a record are not, and a change
// made outside doc.haus no longer matches.
export function systemDocxKey(directory: string, document: string) {
  return `system-docx:${path.basename(canonicalArtifactIdentity(directory, document))}`
}

/** Records, in the caller's open registry transaction, that the system wrote
 * the bytes with this sha256 to `document`. */
export function recordSystemDocx(db: Database, directory: string, document: string, sha256: string) {
  if (!/^[a-f0-9]{64}$/.test(sha256)) throw new Error("A system Word write is recorded by the sha256 of its bytes.")
  db.run("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
  db.run("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
    [systemDocxKey(directory, document), JSON.stringify({ version: 1, sha256 })])
}

/** Removes the record, in the caller's open registry transaction, for any
 * write of `document` that is not the system's own. */
export function forgetSystemDocx(db: Database, directory: string, document: string) {
  if (!db.query("SELECT name FROM sqlite_master WHERE type='table' AND name='meta'").get()) return
  db.run("DELETE FROM meta WHERE key = ?", [systemDocxKey(directory, document)])
}

export function readSystemDocx(directory: string, document: string): { version: 1; sha256: string } | undefined {
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) return
  const db = new Database(file, { readonly: true })
  try {
    if (!db.query("SELECT name FROM sqlite_master WHERE type='table' AND name='meta'").get()) return
    const row = db.query("SELECT value FROM meta WHERE key=?").get(systemDocxKey(directory, document)) as { value: string } | null
    if (!row) return
    const value: unknown = JSON.parse(row.value)
    if (!value || typeof value !== "object" || !("version" in value) || value.version !== 1 ||
      !("sha256" in value) || typeof value.sha256 !== "string" || !/^[a-f0-9]{64}$/.test(value.sha256))
      throw new Error("The record of the system's last write of this Word file is invalid.")
    return { version: 1, sha256: value.sha256 }
  } finally { db.close() }
}

/** Whether the matter file `name` still holds exactly the bytes the system last
 * wrote to it. A missing file, a link or a record that cannot be read
 * (readSystemDocx throws) is not: such a file is never written over. Reads only. */
export function ownedDeliverable(directory: string, name: string) {
  try {
    const record = readSystemDocx(directory, name)
    // sha256File throws when a link or another non-regular file holds the name.
    return record !== undefined && sha256File(path.resolve(directory, name))?.sha256 === record.sha256
  } catch {
    return false
  }
}
