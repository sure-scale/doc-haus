import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

// A matter document's kept versions (#268, services/ingest/src/document-versions.ts).
// Every version's exact bytes are kept unchanged; restoring one copies it
// forward as the next version, so history is never rewritten.

// What wrote a version, as its writer stated it. A value this build does not
// know (a newer ingest's) still reads as something, never as nothing.
export type VersionOrigin =
  | "upload" | "replace" | "generated" | "revision" | "agent-edit" | "redaction"
  | "accept" | "conversion" | "copy" | "restore" | "system" | "unknown"

export type DocumentVersion = {
  id: string
  number: number
  sha256: string
  ext: string
  bytes: number
  // The version this one replaced as current; versions sharing a parent are
  // parallel, shown side by side and never merged.
  parent: string | null
  // A version of another document this one was made from (a conversion's PDF).
  derived_from: string | null
  origin: VersionOrigin
  // "unknown" unless the writer knew who made it.
  actor: string
  run: string | null
  label: string | null
  created_at: number
  current: boolean
  // Why the version's kept bytes were removed: a redaction purges the earlier
  // versions' bytes and keeps their rows. null (or absent) means kept.
  removed?: "redaction" | (string & {}) | null
}

export type VersionEvent = {
  id: number
  kind: "accept" | "reject" | "restore"
  version_id: string | null
  actor: string
  detail: Record<string, unknown> | null
  created_at: number
}

export type DocumentVersions = {
  name: string
  document: { id: string; name: string; current_version: string | null } | null
  // Newest first.
  versions: DocumentVersion[]
  events: VersionEvent[]
  // The version the working file holds; null when it has changed since
  // doc.haus last wrote it, or no version is kept yet.
  working_version: string | null
}

export async function fetchDocumentVersions(matterId: string, name: string): Promise<DocumentVersions> {
  const res = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/versions?name=${encodeURIComponent(name)}`)
  if (!res.ok) throw new Error(await ingestError(res, `Could not load the versions of ${name}`))
  return res.json()
}

export function versionContentUrl(matterId: string, versionId: string, { download = false } = {}) {
  return `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/versions/${encodeURIComponent(versionId)}/content${download ? "?download=1" : ""}`
}

// A version whose kept bytes cannot be served: missing or damaged (409
// version_damaged), removed by a redaction (410 removed_by_redaction), or gone
// with its document (404). The list keeps showing it.
export class VersionUnavailableError extends Error {
  constructor(readonly status: number, readonly code: string | undefined, message: string) {
    super(message)
  }
}

export async function fetchVersionBytes(matterId: string, versionId: string): Promise<Uint8Array> {
  const res = await fetch(versionContentUrl(matterId, versionId))
  if (res.ok) return new Uint8Array(await res.arrayBuffer())
  const body = (await res.json().catch(() => undefined)) as { error?: unknown; code?: unknown } | undefined
  const message = typeof body?.error === "string" ? body.error : `Could not open this version (${res.status})`
  if (res.status === 404 || res.status === 409 || res.status === 410)
    throw new VersionUnavailableError(res.status, typeof body?.code === "string" ? body.code : undefined, message)
  throw new Error(message)
}

export type RestoreResult = {
  name: string
  version?: { id: string; number: number } | null
  restored_from: { id: string; number: number }
  // A restored Word file with pending tracked changes: the version is kept,
  // and doc.haus edits wait until they are resolved.
  notice?: string
}

export class RestoreRefusedError extends Error {
  constructor(readonly code: string | undefined, message: string) {
    super(message)
  }
}

export async function restoreVersion(matterId: string, versionId: string): Promise<RestoreResult> {
  const res = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/versions/${encodeURIComponent(versionId)}/restore`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: "{}",
  })
  if (res.ok) return res.json()
  const body = (await res.json().catch(() => undefined)) as { error?: unknown; code?: unknown } | undefined
  const message = typeof body?.error === "string" ? body.error : `Could not restore this version (${res.status})`
  throw new RestoreRefusedError(typeof body?.code === "string" ? body.code : res.status === 404 ? "not_found" : undefined, message)
}
