import { Database } from "bun:sqlite"
import type { Context, Hono } from "hono"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import type { DeskComparison, DeskSource, DeskSourcePreview, DeskSources, DeskVersion, DeskVersions } from "../../../dochaus/lib/word-source-desk-contract"
import { compareVersions } from "../../../dochaus/lib/version-compare"
import { readCurrentVersion, readDocumentVersions, readVersion, readVersionBytes, sha256Hex, type VersionRow } from "./document-versions"
import { extractDocumentText } from "./ingest"
import { documentPreviewBlocks } from "./document-preview"
import { readDocumentTarget } from "./registered-document"
import { documentRegistryIdentity } from "./document-identity"
import type { WordGrant } from "./word-grants"

type Auth = { matter: { id: string; title: string; dir: string }; grant: WordGrant; target: { name: string; docPath: string } }
type Registered = { id: number; name: string; doc_path: string; role: "source" | "deliverable" }
const LIMIT = 128_000
const refuse = (c: Context, status: 400 | 401 | 404 | 409 | 410 | 413 | 422, code: string, error: string) => c.json({ code, error }, status)
const unavailable = (c: Context) => refuse(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")

function registry<T>(directory: string, read: (db: Database) => T) {
  const db = new Database(path.join(directory, ".dochaus", "legal.db"), { readonly: true })
  try { return read(db) } finally { db.close() }
}

function source(auth: Auth, row: Registered): DeskSource | undefined {
  // Only registry basenames are resolved. Never accept a caller's path or matter.
  if (!row.name || /[\\/\0]/.test(row.name)) return
  const target = readDocumentTarget(auth.matter.dir, row.name)
  const kept = readCurrentVersion(auth.matter.dir, row.name)
  if (!target || !kept || kept.current.removed || !existsSync(target.docPath)) return
  if (documentRegistryIdentity(row.doc_path) !== target.identity) return
  // An external unsaved/unregistered replacement is not the confirmed version.
  if (sha256Hex(readFileSync(target.docPath)) !== kept.current.sha256) return
  return { id: String(row.id), name: target.name, role: row.role, hash: kept.current.sha256,
    version: { id: kept.current.id, number: kept.current.number, createdAt: kept.current.created_at } }
}

/** Resolve the displayed registry identifier to one confirmed matter file.
 * Native opening additionally requires DOCX; neither caller paths nor another
 * matter's document identifiers can select the target. */
export function readDeskSource(auth: Auth, sourceId: string): DeskSource | undefined {
  const rows = registry(auth.matter.dir, db => db.query("SELECT id, name, doc_path, role FROM documents").all() as Registered[])
  const row = rows.find(row => String(row.id) === sourceId)
  const selected = row ? source(auth, row) : undefined
  if (!selected) return
  const target = readDocumentTarget(auth.matter.dir, selected.name)
  if (!target || rows.filter(row => documentRegistryIdentity(row.doc_path) === target.identity).length !== 1) return
  return selected
}

function view(version: VersionRow, current: string | null): DeskVersion {
  return { id: version.id, number: version.number, hash: version.sha256, parent: version.parent,
    createdAt: version.created_at, bytes: version.bytes, origin: version.origin, label: version.label,
    current: version.id === current, removed: Boolean(version.removed) }
}

async function input(c: Context): Promise<Record<string, unknown> | undefined> {
  // Commands carry identifiers only, never document bytes or a second capability.
  const text = await c.req.text()
  if (text.length > 2048) return
  try { const value: unknown = JSON.parse(text); return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : undefined }
  catch { return }
}

/** Source Desk reads use the existing linked-document grant and never touch the
 * editing lease or word_sessions. All version/preview bytes are retained bytes. */
export function installWordSourceDesk(app: Hono, authenticate: (c: Context) => Auth | undefined) {
  app.post("/word/desk/sources", c => {
    const auth = authenticate(c)
    if (!auth) return unavailable(c)
    c.header("Cache-Control", "no-store")
    const rows = registry(auth.matter.dir, db => db.query("SELECT id, name, doc_path, role FROM documents ORDER BY name, id").all() as Registered[])
    const sources = rows.slice(0, 200).flatMap(row => source(auth, row) ?? [])
    const body: DeskSources = { sources, truncated: rows.length > 200 || sources.length !== rows.length,
      matterName: auth.matter.title, documentName: auth.target.name }
    return c.json(body)
  })

  app.post("/word/desk/source", async c => {
    const auth = authenticate(c)
    if (!auth) return unavailable(c)
    c.header("Cache-Control", "no-store")
    const body = await input(c)
    if (!body || typeof body.sourceId !== "string" || !/^[1-9]\d{0,14}$/.test(body.sourceId) || typeof body.expectedHash !== "string" || !/^[a-f0-9]{64}$/.test(body.expectedHash))
      return refuse(c, 400, "invalid_desk_source", "Choose an identified matter source from the desk.")
    const row = registry(auth.matter.dir, db => db.query("SELECT id, name, doc_path, role FROM documents WHERE id = ?").get(Number(body.sourceId)) as Registered | null)
    if (!row) return refuse(c, 404, "desk_source_missing", "That source is no longer in this matter.")
    const selected = source(auth, row)
    if (!selected || selected.hash !== body.expectedHash || !selected.version)
      return refuse(c, 409, "desk_source_changed", "This source changed or is unavailable. Refresh the source list and choose it again.")
    const kept = readVersion(auth.matter.dir, selected.version.id)
    if (!kept || kept.version.removed || kept.document.removed_at != null)
      return refuse(c, 410, "desk_source_removed", "The selected source version is no longer available.")
    const bytes = readVersionBytes(auth.matter.dir, kept.version)
    if (!bytes) return refuse(c, 409, "desk_source_damaged", "The selected source's retained bytes are missing or damaged.")
    if (bytes.length > 32 * 1024 * 1024) return refuse(c, 413, "desk_source_large", "Open this large source in doc.haus to inspect it.")
    let text: string
    try { text = await extractDocumentText(selected.name, bytes) }
    catch { return refuse(c, 422, "desk_preview_failed", "This source cannot be previewed. Open it in doc.haus.") }
    const blocks = await documentPreviewBlocks(selected.name, bytes, text, text.length > LIMIT)
    // Extraction is asynchronous. Revalidate revocation/redaction and identity
    // before releasing the exact snapshot; never return newly removed bytes.
    const live = authenticate(c)
    if (!live) return unavailable(c)
    const liveRow = registry(auth.matter.dir, db => db.query("SELECT id, name, doc_path, role FROM documents WHERE id = ?").get(row.id) as Registered | null)
    const after = liveRow ? source(live, liveRow) : undefined
    if (after?.hash !== selected.hash || after?.version?.id !== selected.version.id)
      return refuse(c, 409, "desk_source_changed", "This source changed while loading. Refresh the source list and choose it again.")
    const preview: DeskSourcePreview = { ...selected, text: text.slice(0, LIMIT), ...(blocks ? { blocks } : {}), characters: text.length,
      truncated: text.length > LIMIT, provenance: "saved-version" }
    return c.json(preview)
  })

  app.post("/word/desk/versions", c => {
    const auth = authenticate(c)
    if (!auth) return unavailable(c)
    c.header("Cache-Control", "no-store")
    const history = readDocumentVersions(auth.matter.dir, auth.target.name)
    if (!history.document || history.document.id !== auth.grant.document_id)
      return refuse(c, 404, "desk_document_missing", "This document's saved versions are unavailable.")
    const body: DeskVersions = { matterName: auth.matter.title, document: { id: history.document.id, name: history.document.name },
      versions: history.versions.map(version => view({ ...version, document_id: history.document!.id }, history.document!.current_version)) }
    return c.json(body)
  })

  app.post("/word/desk/compare", async c => {
    const auth = authenticate(c)
    if (!auth) return unavailable(c)
    c.header("Cache-Control", "no-store")
    const body = await input(c)
    if (!body || typeof body.baseVersion !== "string" || typeof body.currentVersion !== "string" || body.baseVersion.length > 128 || body.currentVersion.length > 128 || body.baseVersion === body.currentVersion)
      return refuse(c, 400, "invalid_desk_compare", "Choose two different saved versions of this Word document.")
    const left = readVersion(auth.matter.dir, body.baseVersion), right = readVersion(auth.matter.dir, body.currentVersion)
    if (!left || !right || [left, right].some(side => side.document.id !== auth.grant.document_id || side.document.removed_at != null))
      return refuse(c, 404, "desk_version_missing", "Both versions must belong to this Word document.")
    if ([left, right].some(side => side.version.removed))
      return refuse(c, 410, "removed_by_redaction", "A selected version was removed by redaction and cannot be compared.")
    const leftBytes = readVersionBytes(auth.matter.dir, left.version), rightBytes = readVersionBytes(auth.matter.dir, right.version)
    if (!leftBytes || !rightBytes) return refuse(c, 409, "desk_version_damaged", "A selected version's retained bytes are missing or damaged.")
    if (leftBytes.length + rightBytes.length > 32 * 1024 * 1024)
      return refuse(c, 413, "desk_compare_large", "Compare these large saved versions in doc.haus.")
    const label = (side: typeof left) => `${side.document.name} (version ${side.version.number})`.replace(/[\r\n]+/g, " ").slice(0, 120)
    const compared = await compareVersions({ bytes: leftBytes, label: label(left) }, { bytes: rightBytes, label: label(right) },
      { ignoreFormatting: true, detectMoves: true, headersFooters: true }, new Date().toISOString()).catch(() => undefined)
    if (!compared) return refuse(c, 422, "compare_failed", "The saved versions could not be compared. Open Versions in doc.haus.")
    if (!authenticate(c)) return unavailable(c)
    const after = [left, right].map(side => readVersion(auth.matter.dir, side.version.id))
    if (after.some(side => !side || side.version.removed || side.document.removed_at != null))
      return refuse(c, 410, "removed_by_redaction", "A selected version is no longer available.")
    const result: DeskComparison = { base: { ...view(left.version, after[0]!.document.current_version), name: left.document.name, trackedChanges: compared.left },
      current: { ...view(right.version, after[1]!.document.current_version), name: right.document.name, trackedChanges: compared.right },
      changes: compared.changes.slice(0, 500), untracked: compared.untracked.slice(0, 500), proposals: compared.proposals.slice(0, 500),
      attribution: compared.attribution, context: "saved-versions",
      truncated: [compared.changes, compared.untracked, compared.proposals].some(rows => rows.length > 500) }
    return c.json(result)
  })
}
