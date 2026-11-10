import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { canonicalArtifactIdentity } from "../capabilities/artifacts/mutation-lock"
import { sha256File } from "./lineage"
import type { SourceDependencyStatus, SourceVersionSnapshot } from "./source-version-dependencies-contract"
export type { SourceDependencyStatus, SourceVersionSnapshot } from "./source-version-dependencies-contract"

export function parseSourceVersionSnapshot(value: unknown): SourceVersionSnapshot | undefined {
  try {
    const raw = typeof value === "string" ? JSON.parse(value) : value
    if (!raw || typeof raw !== "object" || typeof raw.complete !== "boolean" || !Array.isArray(raw.versions)) return
    if (!raw.versions.every((row: unknown) => row && typeof row === "object" &&
      "documentId" in row && typeof row.documentId === "string" && /^[a-f0-9-]{36}$/.test(row.documentId) &&
      "versionId" in row && typeof row.versionId === "string" && /^[a-f0-9-]{36}$/.test(row.versionId) &&
      "sha256" in row && typeof row.sha256 === "string" && /^[a-f0-9]{64}$/.test(row.sha256))) return
    return { complete: raw.complete, versions: raw.versions.map((row: SourceVersionSnapshot["versions"][number]) => ({ documentId: row.documentId, versionId: row.versionId, sha256: row.sha256 })) }
  } catch { return }
}

export function mergeSourceVersionSnapshots(snapshots: (SourceVersionSnapshot | null | undefined)[]): SourceVersionSnapshot {
  const versions = new Map<string, SourceVersionSnapshot["versions"][number]>()
  for (const snapshot of snapshots) for (const row of snapshot?.versions ?? []) versions.set(row.versionId, row)
  return { versions: [...versions.values()].sort((a, b) => a.documentId.localeCompare(b.documentId) || a.versionId.localeCompare(b.versionId)),
    complete: snapshots.length > 0 && snapshots.every(snapshot => snapshot?.complete === true) }
}

function fileSha(directory: string, name: string) {
  try { return name === path.basename(name) ? sha256File(path.join(directory, name))?.sha256 : undefined } catch { return }
}
function retainedSha(directory: string, source: SourceVersionSnapshot["versions"][number], ext: string) {
  if (!/^\.[a-z0-9]{1,10}$/.test(ext)) return
  try { return sha256File(path.join(directory, ".dochaus", "revisions", source.documentId, `${source.sha256}${ext}`))?.sha256 } catch { return }
}
function readDb<T>(directory: string, read: (db: Database) => T, fallback: T): T {
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) return fallback
  const db = new Database(file, { readonly: true })
  try { return read(db) } catch { return fallback } finally { db.close() }
}

/** Read before a stage starts; never create/backfill a version or grant write
 * ownership. It lives in the stage's existing state and is descriptive only. */
export function readMatterSourceVersions(directory: string): SourceVersionSnapshot {
  return readDb(directory, db => {
    const inputs = db.query("SELECT name FROM documents WHERE role='source' UNION SELECT doc_name AS name FROM chunks ORDER BY name").all() as { name: string }[]
    const versions: SourceVersionSnapshot["versions"] = []
    let complete = true
    for (const input of inputs) {
      const key = path.basename(canonicalArtifactIdentity(directory, input.name))
      const row = db.query(`SELECT d.id AS documentId,v.id AS versionId,v.sha256,v.removed
        FROM versioned_documents d JOIN document_versions v ON v.id=d.current_version
        WHERE d.name_key=? AND d.removed_at IS NULL`).get(key) as (SourceVersionSnapshot["versions"][number] & { removed: string | null }) | null
      if (!row || row.removed || fileSha(directory, input.name) !== row.sha256) { complete = false; continue }
      versions.push({ documentId: row.documentId, versionId: row.versionId, sha256: row.sha256 })
    }
    return { versions, complete }
  }, { versions: [], complete: false })
}

/** Resolve a saved source hash, never today's source set. A hash repeated in
 * several retained versions does not identify which version was admitted. */
export function readPinnedSourceVersions(directory: string, pins: { name: string; sha256: string | undefined }[]): SourceVersionSnapshot {
  return readDb(directory, db => {
    const versions = new Map<string, SourceVersionSnapshot["versions"][number]>()
    const seen = new Set<string>()
    let complete = pins.length > 0
    for (const pin of pins) {
      if (typeof pin.name !== "string" || pin.name !== path.basename(pin.name) || !pin.name || typeof pin.sha256 !== "string" || !/^[a-f0-9]{64}$/.test(pin.sha256)) { complete = false; continue }
      const key = path.basename(canonicalArtifactIdentity(directory, pin.name))
      const identity = JSON.stringify([key, pin.sha256])
      if (seen.has(identity)) continue
      seen.add(identity)
      const rows = db.query(`SELECT d.id AS documentId,v.id AS versionId,v.sha256,v.ext,v.removed,d.removed_at
        FROM versioned_documents d JOIN document_versions v ON v.document_id=d.id
        WHERE d.name_key=? AND v.sha256=?`).all(key, pin.sha256) as (SourceVersionSnapshot["versions"][number] & { ext: string; removed: string | null; removed_at: number | null })[]
      if (rows.length !== 1) { complete = false; continue }
      const row = rows[0]!
      versions.set(row.versionId, { documentId: row.documentId, versionId: row.versionId, sha256: row.sha256 })
      if (row.removed || row.removed_at !== null || retainedSha(directory, row, row.ext) !== row.sha256) complete = false
    }
    return { versions: [...versions.values()].sort((left, right) => left.versionId.localeCompare(right.versionId)), complete }
  }, { versions: [], complete: false })
}

/** Derived at read time: a received, revised, deleted, redacted or externally
 * changed source immediately makes all saved dependents visibly stale. */
export function readSourceDependencyStatus(directory: string, snapshot: SourceVersionSnapshot | null | undefined, connection?: Database): SourceDependencyStatus {
  const unknown: SourceDependencyStatus = { status: "unknown", complete: false, sources: [] }
  if (!snapshot) return unknown
  const read = (db: Database): SourceDependencyStatus => {
    const sources = snapshot.versions.map(source => {
      const row = db.query(`SELECT d.name,d.removed_at,v.number,v.removed,v.sha256,v.ext,
        c.id AS current_id,c.number AS current_number,c.sha256 AS current_sha256
        FROM document_versions v JOIN versioned_documents d ON d.id=v.document_id
        LEFT JOIN document_versions c ON c.id=d.current_version WHERE v.id=? AND d.id=?`).get(source.versionId, source.documentId) as {
          name: string; removed_at: number | null; number: number; removed: string | null; sha256: string; ext: string
          current_id: string | null; current_number: number | null; current_sha256: string | null
        } | null
      const linked = row && db.query(`SELECT e.id AS exchangeId,v.id AS versionId,v.number AS versionNumber,s.name AS sourceName,
        v.document_id AS documentId,v.sha256,v.ext FROM document_version_events e
        JOIN document_versions v ON v.id=e.version_id JOIN versioned_documents s ON s.id=e.document_id
        JOIN document_versions b ON b.id=json_extract(e.detail,'$.association.baseVersionId')
        WHERE e.kind='received' AND json_valid(e.detail) AND json_extract(e.detail,'$.association.documentId')=?
        AND b.document_id=? AND b.removed IS NULL AND v.document_id=e.document_id AND v.removed IS NULL AND s.removed_at IS NULL
        ORDER BY e.id DESC LIMIT 1`).get(source.documentId, source.documentId) as {
          exchangeId: number; versionId: string; versionNumber: number; sourceName: string; documentId: string; sha256: string; ext: string
        } | null
      const status = !row ? "unavailable" : row.removed_at !== null || row.removed ? "removed"
        : row.sha256 !== source.sha256 ? "unavailable" : row.current_id !== source.versionId ? "newer-version"
        : fileSha(directory, row.name) !== row.current_sha256 ? "unrecorded-change"
        : retainedSha(directory, source, row.ext) !== source.sha256 ? "unavailable"
        : linked ? retainedSha(directory, linked, linked.ext) === linked.sha256 ? "newer-received-branch" : "unavailable" : "current"
      return { ...source, name: row?.name ?? "Unavailable source", versionNumber: row?.number ?? null,
        currentVersionId: row?.current_id ?? null, currentVersionNumber: row?.current_number ?? null, status,
        ...(linked ? { linkedReturn: { exchangeId: linked.exchangeId, versionId: linked.versionId, sourceName: linked.sourceName, versionNumber: linked.versionNumber } } : {}) } satisfies SourceDependencyStatus["sources"][number]
    })
    return { complete: snapshot.complete, sources, status: sources.some(source => source.status !== "current" && source.status !== "unavailable") ? "stale"
      : !snapshot.complete || sources.some(source => source.status === "unavailable") ? "unknown" : "current" }
  }
  return connection ? read(connection) : readDb(directory, read, unknown)
}
