import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { canonicalArtifactIdentity } from "../capabilities/artifacts/mutation-lock"

// Stored in the existing registry transaction with the accepted document bytes.
// A projection is editing context, never source evidence or drafting approval.
export function workingDocxKey(directory: string, document: string) {
  return `working-docx:${path.basename(canonicalArtifactIdentity(directory, document))}`
}

export function readWorkingDocx(directory: string, document: string) {
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) return
  const db = new Database(file, { readonly: true })
  try {
    if (!db.query("SELECT name FROM sqlite_master WHERE type='table' AND name='meta'").get()) return
    const row = db.query("SELECT value FROM meta WHERE key=?").get(workingDocxKey(directory, document)) as { value: string } | null
    if (!row) return
    const value: unknown = JSON.parse(row.value)
    if (!value || typeof value !== "object" || !("version" in value) || value.version !== 1 ||
      !("source_sha256" in value) || typeof value.source_sha256 !== "string" || !/^[a-f0-9]{64}$/.test(value.source_sha256) ||
      !("projection_sha256" in value) || typeof value.projection_sha256 !== "string" || !/^[a-f0-9]{64}$/.test(value.projection_sha256))
      throw new Error("The accepted working Word revision record is invalid.")
    return value
  } finally { db.close() }
}

export function assertDirectWorkingDocxWriteAllowed(directory: string, document: string) {
  if (readWorkingDocx(directory, document))
    throw new Error("This document has an accepted Word revision. Prepare revision-bound changes or use the document revision API.")
}
