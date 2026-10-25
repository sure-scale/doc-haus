import { Database } from "bun:sqlite"
import { createHash, randomUUID } from "node:crypto"
import { closeSync, existsSync, fsyncSync, mkdirSync, openSync, readdirSync, readFileSync, renameSync, rmSync, writeSync } from "node:fs"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { canonicalArtifactIdentity } from "../../../dochaus/capabilities/artifacts/mutation-lock"

// Document versions (#268, phase 1). A document is the logical file ("the
// services agreement") with a stable id; a version is the exact bytes a doc.haus
// writer put in its working file. The matter folder keeps one working file per
// document under its stable name, holding the current version. Every version's
// bytes are kept once, unchanged, at .dochaus/revisions/<document>/<sha256><ext>,
// and its row (number, parent, origin, actor, run, time, label) in legal.db.
//
// Versions are captured only inside a writer's existing registry transaction
// (upload, replace, deliverable write, revision save, accept, conversion, copy,
// restore), never by a separate step. Origin comes from the writer; actor is
// "unknown" unless the writer knows it. The bytes a write replaces are kept
// when no version holds them yet (a document from before this store, or a file
// changed outside doc.haus): they become a version of origin "unknown" first,
// so no write loses bytes. History is never rewritten: restore copies an old
// version forward as a new one, and deleting a document through doc.haus
// detaches its history, so a later file of that name starts a new document.
//
// Two writes remove kept bytes, as they did before this store: a redaction
// deletes the stored bytes of the document's earlier versions once it commits
// (their rows stay, marked removed by redaction), and deleting a document
// deletes all of its stored bytes. A removed version is listed but cannot be
// read, compared or restored.

export const VERSION_ORIGINS = [
  "upload", "replace", "generated", "revision", "agent-edit", "redaction",
  "accept", "conversion", "copy", "restore", "system", "unknown",
] as const
export type VersionOrigin = (typeof VERSION_ORIGINS)[number]

/** What the writer knows about a write. */
export type VersionProvenance = { origin: VersionOrigin; actor?: string; run?: string | null; label?: string | null }
export type VersionEventKind = "accept" | "reject" | "restore"
/** Why a kept version's bytes were deleted: a later redaction of its document. */
export type VersionRemoval = "redaction"
export type VersionEventInput = { kind: VersionEventKind; actor?: string; detail?: Record<string, unknown> }
/** A writer's version capture: its provenance, the matter document the new
 * version was derived from (a conversion's PDF, a copy's original) and the
 * event the write records (an accept, a restore). */
export type VersionWrite = VersionProvenance & { derivedFrom?: string; event?: VersionEventInput }

export type VersionRow = {
  id: string
  document_id: string
  number: number
  sha256: string
  ext: string
  bytes: number
  parent: string | null
  derived_from: string | null
  origin: VersionOrigin
  actor: string
  run: string | null
  label: string | null
  created_at: number
  removed: VersionRemoval | null
}
export type VersionedDocumentRow = {
  id: string
  name: string
  name_key: string
  current_version: string | null
  created_at: number
  removed_at: number | null
}
export type CapturedVersion = {
  document_id: string
  id: string
  number: number
  parent: string | null
  origin: VersionOrigin
  // false when the write held the current version's bytes, so none was added
  created: boolean
}

const UNKNOWN_ACTOR = "unknown"

/** Creates the version tables; run by openDb with the other registry DDL. */
export function createVersionTables(db: Database) {
  // name_key: the working file's canonical name (artifact identity basename), so
  // a spelling the filesystem treats as the same file finds the same document.
  // Only one live (not removed) document holds a name.
  db.run(`
    CREATE TABLE IF NOT EXISTS versioned_documents (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      name_key TEXT NOT NULL,
      current_version TEXT,
      created_at INTEGER NOT NULL,
      removed_at INTEGER
    )
  `)
  db.run("CREATE UNIQUE INDEX IF NOT EXISTS versioned_documents_live_name ON versioned_documents(name_key) WHERE removed_at IS NULL")
  // parent: the version this one replaced as current (NULL for the first), so
  // parallel versions can share a parent. derived_from: a version of another
  // document this one was made from (a conversion's PDF, a copy's original).
  db.run(`
    CREATE TABLE IF NOT EXISTS document_versions (
      id TEXT PRIMARY KEY,
      document_id TEXT NOT NULL,
      number INTEGER NOT NULL,
      sha256 TEXT NOT NULL,
      ext TEXT NOT NULL,
      bytes INTEGER NOT NULL,
      parent TEXT,
      derived_from TEXT,
      origin TEXT NOT NULL,
      actor TEXT NOT NULL DEFAULT 'unknown',
      run TEXT,
      label TEXT,
      created_at INTEGER NOT NULL,
      removed TEXT,
      UNIQUE (document_id, number)
    )
  `)
  // removed: why the version's stored bytes were deleted (VersionRemoval).
  if (!(db.query("PRAGMA table_info(document_versions)").all() as { name: string }[]).some((column) => column.name === "removed"))
    db.run("ALTER TABLE document_versions ADD COLUMN removed TEXT")
  db.run("CREATE INDEX IF NOT EXISTS document_versions_sha ON document_versions(sha256)")
  // version_id: the version an accept or restore made, or the version a reject
  // left current. detail: JSON (redline ids, the restored version).
  db.run(`
    CREATE TABLE IF NOT EXISTS document_version_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id TEXT NOT NULL,
      version_id TEXT,
      kind TEXT NOT NULL CHECK (kind IN ('accept', 'reject', 'restore')),
      actor TEXT NOT NULL DEFAULT 'unknown',
      detail TEXT,
      created_at INTEGER NOT NULL
    )
  `)
}

export const sha256Hex = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")

function nameKey(matterDir: string, name: string) {
  return path.basename(canonicalArtifactIdentity(matterDir, path.basename(name)))
}

function liveDocument(db: Database, matterDir: string, name: string) {
  return db.query("SELECT * FROM versioned_documents WHERE name_key = ? AND removed_at IS NULL")
    .get(nameKey(matterDir, name)) as VersionedDocumentRow | null
}

function versionById(db: Database, id: string | null) {
  return id ? db.query("SELECT * FROM document_versions WHERE id = ?").get(id) as VersionRow | null : null
}

// `planned`: the id a new document takes, under which its bytes were staged.
function documentOf(db: Database, matterDir: string, name: string, now: number, planned?: string) {
  const existing = liveDocument(db, matterDir, name)
  if (existing) {
    if (existing.name !== name) db.run("UPDATE versioned_documents SET name = ? WHERE id = ?", [name, existing.id])
    return { ...existing, name }
  }
  const row: VersionedDocumentRow = { id: planned ?? randomUUID(), name, name_key: nameKey(matterDir, name), current_version: null, created_at: now, removed_at: null }
  db.run("INSERT INTO versioned_documents (id, name, name_key, current_version, created_at, removed_at) VALUES (?, ?, ?, NULL, ?, NULL)",
    [row.id, row.name, row.name_key, row.created_at])
  return row
}

export function versionFile(matterDir: string, version: Pick<VersionRow, "document_id" | "sha256" | "ext">) {
  return path.join(matterDir, ".dochaus", "revisions", version.document_id, `${version.sha256}${version.ext}`)
}

// Written once under its content hash: complete bytes go to a private file
// first (wx), are synced, and are then renamed into place, and the directories
// holding the new entries are synced, so the store never holds a partial
// version and a committed row's bytes survive a crash. Bytes already stored are
// not written again; missing or damaged ones are written afresh.
function storeVersionBytes(matterDir: string, documentId: string, sha256: string, ext: string, bytes: Uint8Array) {
  const file = versionFile(matterDir, { document_id: documentId, sha256, ext })
  if (existsSync(file) && sha256Hex(readFileSync(file)) === sha256) return
  const directory = path.dirname(file)
  // Each directory this creates (revisions/, then the document's) is synced into its parent.
  const created = mkdirSync(directory, { recursive: true })
  if (created) for (let parent = path.dirname(directory); ; parent = path.dirname(parent)) {
    syncDirectory(parent)
    if (parent === path.dirname(created) || parent === path.dirname(parent)) break
  }
  const temporary = path.join(directory, `.${sha256}.${process.pid}.${randomUUID()}.tmp`)
  try {
    const descriptor = openSync(temporary, "wx")
    try {
      writeSync(descriptor, bytes)
      fsyncSync(descriptor)
    } finally {
      closeSync(descriptor)
    }
    renameSync(temporary, file)
    syncDirectory(directory)
  } finally {
    rmSync(temporary, { force: true })
  }
}

function syncDirectory(directory: string) {
  const descriptor = openSync(directory, "r")
  try {
    fsyncSync(descriptor)
  } finally {
    closeSync(descriptor)
  }
}

/** The bytes a write of `name` will capture, stored before the writer's
 * registry transaction opens (under its artifact lock), so no byte write or
 * sync holds legal.db's write lock: the new bytes and the working file's, which
 * the write keeps as a version when none holds them (a redaction keeps none).
 * Bytes of a write that then fails stay unreferenced in the store. */
export type VersionStage = { document_id: string; stored: ReadonlySet<string> }
export function stageVersionBytes(matterDir: string, name: string, bytes: Uint8Array, origin: VersionOrigin): VersionStage {
  const live = readStore(matterDir, (db) => liveDocument(db, matterDir, name), null)
  const documentId = live?.id ?? randomUUID()
  const ext = path.extname(name).toLowerCase()
  const file = path.join(matterDir, path.basename(name))
  const stored = new Set<string>()
  for (const entry of [bytes, ...(origin !== "redaction" && existsSync(file) ? [readFileSync(file)] : [])]) {
    const sha256 = sha256Hex(entry)
    if (stored.has(sha256)) continue
    storeVersionBytes(matterDir, documentId, sha256, ext, entry)
    stored.add(sha256)
  }
  return { document_id: documentId, stored }
}

function insertVersion(
  db: Database,
  matterDir: string,
  document: VersionedDocumentRow,
  bytes: Uint8Array,
  sha256: string,
  provenance: VersionProvenance,
  parent: string | null,
  derivedFrom: string | null,
  now: number,
  stage?: VersionStage,
): VersionRow {
  const ext = path.extname(document.name).toLowerCase()
  if (!staged(stage, document.id, sha256)) storeVersionBytes(matterDir, document.id, sha256, ext, bytes)
  const number = ((db.query("SELECT MAX(number) AS n FROM document_versions WHERE document_id = ?").get(document.id) as { n: number | null }).n ?? 0) + 1
  const row: VersionRow = {
    id: randomUUID(), document_id: document.id, number, sha256, ext, bytes: bytes.byteLength, parent, derived_from: derivedFrom,
    origin: provenance.origin, actor: oneLine(provenance.actor) || UNKNOWN_ACTOR, run: oneLine(provenance.run) || null,
    label: oneLine(provenance.label) || null, created_at: now, removed: null,
  }
  db.run(
    "INSERT INTO document_versions (id, document_id, number, sha256, ext, bytes, parent, derived_from, origin, actor, run, label, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [row.id, row.document_id, row.number, row.sha256, row.ext, row.bytes, row.parent, row.derived_from, row.origin, row.actor, row.run, row.label, row.created_at],
  )
  db.run("UPDATE versioned_documents SET current_version = ? WHERE id = ?", [row.id, document.id])
  document.current_version = row.id
  return row
}

function staged(stage: VersionStage | undefined, documentId: string, sha256: string) {
  return stage?.document_id === documentId && stage.stored.has(sha256)
}

/** The current version of the matter document `name`, adding the working
 * file's bytes as a version of origin "unknown" when no version holds them (a
 * document from before this store, or one changed outside doc.haus). Undefined
 * when the file is missing. In the caller's open registry transaction. */
function currentOrFoundVersion(db: Database, matterDir: string, name: string, now: number) {
  const file = path.join(matterDir, path.basename(name))
  if (!existsSync(file)) return
  const bytes = readFileSync(file)
  const sha256 = sha256Hex(bytes)
  const document = documentOf(db, matterDir, name, now)
  const current = versionById(db, document.current_version)
  if (current?.sha256 === sha256) return current
  return insertVersion(db, matterDir, document, bytes, sha256, { origin: "unknown" }, current?.id ?? null, null, now)
}

/** Records the bytes a writer has just put in the working file of `name` as
 * that document's next version, in the writer's open registry transaction.
 * `prior`: the bytes the write replaced (its backup), kept as a version first
 * when no version holds them. A write of the current version's bytes adds none
 * (and stores them again if the kept copy is missing or damaged). `stage`: the
 * bytes stageVersionBytes stored for this write. A redaction keeps no prior
 * bytes and marks every earlier version removed; purgeVersionBytes deletes
 * their bytes once the write commits. */
export function captureVersion(
  db: Database,
  matterDir: string,
  input: { name: string; bytes: Uint8Array; prior?: Uint8Array | null; version: VersionWrite; stage?: VersionStage },
): CapturedVersion {
  const now = Date.now()
  const { stage } = input
  const document = documentOf(db, matterDir, input.name, now, stage?.document_id)
  const sha256 = sha256Hex(input.bytes)
  const redaction = input.version.origin === "redaction"
  let current = versionById(db, document.current_version)
  if (input.prior && !redaction) {
    const prior = sha256Hex(input.prior)
    // A writer that saved the file in place before calling replaced nothing
    // unrecorded: its prior bytes are its own.
    if (prior !== current?.sha256 && prior !== sha256)
      current = insertVersion(db, matterDir, document, input.prior, prior, { origin: "unknown" }, current?.id ?? null, null, now, stage)
  }
  const derivedFrom = input.version.derivedFrom ? currentOrFoundVersion(db, matterDir, input.version.derivedFrom, now)?.id ?? null : null
  const created = current?.sha256 !== sha256
  const version = created
    ? insertVersion(db, matterDir, document, input.bytes, sha256, input.version, current?.id ?? null, derivedFrom, now, stage)
    : current!
  if (!created && !staged(stage, document.id, sha256)) storeVersionBytes(matterDir, document.id, sha256, version.ext, input.bytes)
  // A removed version keeps its row, but not its label: free text that may name what was removed.
  if (redaction)
    db.run("UPDATE document_versions SET removed = 'redaction', label = NULL WHERE document_id = ? AND sha256 != ? AND removed IS NULL", [document.id, sha256])
  if (input.version.event) insertEvent(db, document.id, version.id, input.version.event, now)
  return { document_id: document.id, id: version.id, number: version.number, parent: version.parent, origin: version.origin, created }
}

/** Records an event that writes no bytes (a reject), against the current
 * version when the working file of `name` still holds it (else no version: an
 * event is not a write and captures nothing). In the caller's registry transaction. */
export function recordVersionEvent(db: Database, matterDir: string, name: string, event: VersionEventInput) {
  const now = Date.now()
  const document = documentOf(db, matterDir, name, now)
  const current = versionById(db, document.current_version)
  const file = path.join(matterDir, path.basename(name))
  const holds = current && existsSync(file) && sha256Hex(readFileSync(file)) === current.sha256
  insertEvent(db, document.id, holds ? current.id : null, event, now)
}

function insertEvent(db: Database, documentId: string, versionId: string | null, event: VersionEventInput, now: number) {
  db.run("INSERT INTO document_version_events (document_id, version_id, kind, actor, detail, created_at) VALUES (?, ?, ?, ?, ?, ?)",
    [documentId, versionId, event.kind, oneLine(event.actor) || UNKNOWN_ACTOR, event.detail ? JSON.stringify(event.detail) : null, now])
}

/** Detaches the history of `name` when doc.haus deletes its working file, so a
 * later file of that name starts a new document; its rows stay. Returns the
 * detached document ids, whose bytes purgeVersionBytes deletes once the delete
 * commits. In the delete's registry transaction. */
export function detachVersionedDocument(db: Database, matterDir: string, name: string) {
  const key = nameKey(matterDir, name)
  const ids = (db.query("SELECT id FROM versioned_documents WHERE name_key = ? AND removed_at IS NULL").all(key) as { id: string }[]).map((row) => row.id)
  db.run("UPDATE versioned_documents SET removed_at = ? WHERE name_key = ? AND removed_at IS NULL", [Date.now(), key])
  return ids
}

/** Deletes the stored files of document `documentId` that no kept version
 * holds: all of them once the document is removed, else those of versions a
 * redaction removed (and any a failed write left). After the write that marked
 * them commits, under the document's artifact lock, or at startup. Returns how
 * many entries it deleted; throws when one cannot be deleted. */
export function purgeVersionBytes(matterDir: string, documentId: string) {
  const directory = path.join(matterDir, ".dochaus", "revisions", documentId)
  if (!existsSync(directory)) return 0
  const kept = new Set(readStore(matterDir, (db) => (db.query(
    `SELECT v.sha256 || v.ext AS file FROM document_versions v JOIN versioned_documents d ON d.id = v.document_id
     WHERE v.document_id = ? AND v.removed IS NULL AND d.removed_at IS NULL`,
  ).all(documentId) as { file: string }[]).map((row) => row.file), [] as string[]))
  const entries = readdirSync(directory).filter((entry) => !kept.has(entry))
  if (!kept.size) {
    rmSync(directory, { recursive: true, force: true })
    syncDirectory(path.dirname(directory))
    return entries.length + 1
  }
  for (const entry of entries) rmSync(path.join(directory, entry), { recursive: true, force: true })
  if (entries.length) syncDirectory(directory)
  return entries.length
}

/** At startup, before any write: deletes what a purge left (one that failed,
 * or a stop between a commit and its purge), the stored bytes of every version
 * a redaction removed and the folder of every removed document. Returns how
 * many entries it deleted; throws on the first that cannot be deleted. */
export function sweepRemovedVersionBytes(matterDir: string) {
  const ids = readStore(matterDir, (db) => (db.query(
    `SELECT id FROM versioned_documents WHERE removed_at IS NOT NULL
     UNION SELECT document_id AS id FROM document_versions WHERE removed IS NOT NULL`,
  ).all() as { id: string }[]).map((row) => row.id), [] as string[])
  return ids.reduce((deleted, id) => deleted + purgeVersionBytes(matterDir, id), 0)
}

// Reads open legal.db read-only, so a read never creates or migrates it.
function readStore<T>(matterDir: string, read: (db: Database) => T, empty: T): T {
  const file = path.join(matterDir, ".dochaus", "legal.db")
  if (!existsSync(file)) return empty
  const db = new Database(file, { readonly: true })
  try {
    if (!db.query("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'document_versions'").get()) return empty
    return read(db)
  } finally {
    db.close()
  }
}

export type VersionView = Omit<VersionRow, "document_id"> & { current: boolean }
export type VersionEventView = { id: number; kind: VersionEventKind; version_id: string | null; actor: string; detail: Record<string, unknown> | null; created_at: number }

/** The live document `name` with its versions, newest first, and its events. */
export function readDocumentVersions(matterDir: string, name: string) {
  return readStore(matterDir, (db) => {
    const document = liveDocument(db, matterDir, name)
    if (!document) return { document: null, versions: [] as VersionView[], events: [] as VersionEventView[] }
    const versions = (db.query("SELECT * FROM document_versions WHERE document_id = ? ORDER BY number DESC").all(document.id) as VersionRow[])
      .map(({ document_id: _document, ...row }) => ({ ...row, removed: row.removed ?? null, current: row.id === document.current_version }))
    const events = (db.query("SELECT id, kind, version_id, actor, detail, created_at FROM document_version_events WHERE document_id = ? ORDER BY id").all(document.id) as
      Array<Omit<VersionEventView, "detail"> & { detail: string | null }>)
      .map((row) => ({ ...row, detail: row.detail ? JSON.parse(row.detail) as Record<string, unknown> : null }))
    return { document: { id: document.id, name: document.name, current_version: document.current_version }, versions, events }
  }, { document: null, versions: [] as VersionView[], events: [] as VersionEventView[] })
}

/** One version with its document, by version id. */
export function readVersion(matterDir: string, id: string) {
  return readStore(matterDir, (db) => {
    const version = versionById(db, id)
    if (!version) return
    const document = db.query("SELECT * FROM versioned_documents WHERE id = ?").get(version.document_id) as VersionedDocumentRow
    return { version, document, parent: versionById(db, version.parent) ?? undefined, current: versionById(db, document.current_version) ?? undefined }
  }, undefined)
}

/** The live document `name`'s current version and the one before it. */
export function readCurrentVersion(matterDir: string, name: string) {
  return readStore(matterDir, (db) => {
    const document = liveDocument(db, matterDir, name)
    const current = versionById(db, document?.current_version ?? null)
    if (!document || !current) return
    return { document, current, parent: versionById(db, current.parent) ?? undefined }
  }, undefined)
}

/** The stored bytes of `version`, checked against its hash; undefined when
 * they are missing or damaged. */
export function readVersionBytes(matterDir: string, version: Pick<VersionRow, "document_id" | "sha256" | "ext">) {
  const file = versionFile(matterDir, version)
  if (!existsSync(file)) return
  const bytes = readFileSync(file)
  return sha256Hex(bytes) === version.sha256 ? bytes : undefined
}

/** The kept (not removed) version of the live document `name` holding exactly
 * these bytes, the current one first, with that document's name; undefined
 * when none does. */
export function findVersionBySha(matterDir: string, name: string, sha256: string) {
  return readStore(matterDir, (db) => db.query(
    `SELECT v.id, v.number, v.document_id, v.sha256, v.ext, d.name, v.id = d.current_version AS current
     FROM document_versions v JOIN versioned_documents d ON d.id = v.document_id
     WHERE d.name_key = ? AND d.removed_at IS NULL AND v.sha256 = ? AND v.removed IS NULL
     ORDER BY current DESC, v.number DESC LIMIT 1`,
  ).get(nameKey(matterDir, name), sha256) as
    { id: string; number: number; document_id: string; sha256: string; ext: string; name: string; current: number } | null ?? undefined, undefined)
}

/** The pending tracked changes in a Word file: how many, and by whom, in every
 * XML part (body, headers, footers, notes, comments), as the accepted-revision
 * reader (openRevision) checks them. */
export function trackedChanges(name: string, bytes: Uint8Array) {
  if (path.extname(name).toLowerCase() !== ".docx") return { count: 0, authors: [] as string[] }
  try {
    const parts = Object.values(unzipSync(bytes, { filter: (file) => file.name.endsWith(".xml") }))
    const tags = parts.flatMap((part) => strFromU8(part).match(/<w:(?:ins|del|moveFrom|moveTo)\b[^>]*>/g) ?? [])
    const authors = [...new Set(tags.flatMap((tag) => {
      const author = /\bw:author="([^"]*)"/.exec(tag)?.[1]?.trim()
      return author ? [decodeXml(author)] : []
    }))]
    return { count: tags.length, authors }
  } catch {
    return { count: 0, authors: [] as string[] }
  }
}

/** What a save of a version with pending tracked changes tells the lawyer:
 * the version is kept, and doc.haus edits wait until the changes are resolved,
 * in the viewer (an uploaded source, #200) or in Word. */
export function trackedChangesNotice(number: number, tracked: { count: number; authors: string[] }, where: "viewer" | "word") {
  const changes = `${tracked.count} tracked change${tracked.count === 1 ? "" : "s"}`
  const by = tracked.authors.length ? ` by ${listNames(tracked.authors)}` : ""
  return `Saved as version ${number}, with ${changes}${by}; ${where === "viewer" ? "review them in the viewer" : "resolve them in Word"} before doc.haus edits it.`
}

function listNames(names: string[]) {
  const shown = names.slice(0, 3)
  const rest = names.length - shown.length
  const items = rest > 0 ? [...shown, `${rest} other${rest === 1 ? "" : "s"}`] : shown
  return items.length > 1 ? `${items.slice(0, -1).join(", ")} and ${items.at(-1)}` : items[0]!
}

function decodeXml(text: string) {
  return text.replace(/&(lt|gt|quot|apos|amp);/g, (_, entity: string) => ({ lt: "<", gt: ">", quot: "\"", apos: "'", amp: "&" })[entity]!)
}

/** One line, capped: actors, runs and labels come from requests. */
export function oneLine(text: string | null | undefined) {
  if (!text) return ""
  return Array.from(text.replace(/[\p{Cc}\s]+/gu, " ").trim()).slice(0, 120).join("").trim()
}
