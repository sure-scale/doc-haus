import { Database } from "bun:sqlite"
import { mkdirSync } from "node:fs"
import path from "node:path"
import { documentRegistryIdentity } from "./document-identity"

// Per-matter document registry and retrieval index. Lives inside the matter
// directory so it is scoped to that matter. Source rows own searchable chunks;
// deliverable rows exist only so generated files remain listable and viewable.

// How long the startup connection waits for another connection's write lock.
const STARTUP_BUSY_TIMEOUT_MS = 5000
// How long a source-corpus commit connection waits for one.
const COMMIT_BUSY_TIMEOUT_MS = 500
// How long any other request connection waits for one: long enough for another
// process's brief hold, short enough not to stall this event loop.
const REQUEST_BUSY_TIMEOUT_MS = 200

// startup: the per-matter connection the service opens at boot, before it
// listens, to run the compatibility migrations. Only it waits seconds for a
// write lock and checks the lexical index against the chunks (see ftsInSync
// below).
//
// Every other connection serves a request in this process, where an ingest or
// delete commit can hold its write lock across awaited file writes. bun:sqlite
// waits synchronously, so a request waiting for that lock would block the very
// event loop the holder needs to finish: a long wait could only stall the
// service and then fail. A request connection therefore waits only briefly;
// on a current database opening it takes no write lock at all, so reads go on
// during a commit and a competing write fails with SQLITE_BUSY after that short
// wait. The wait covers another process's short holds: the engine's legal.db
// writes, and the checkpoint its last connection runs as it closes, which
// holds the database while it removes the WAL. At startup no request runs yet,
// so a lock can only belong to another process, which finishes while the
// startup connection waits.
//
// commit: the connection a source upload, a document delete or a deliverable
// write commits through (ingest.ts, document-delete.ts, deliverable.ts). Each
// holds the matter's source-corpus lock, which every commit that keeps the
// write lock across an await also takes, so no such commit runs in this
// process meanwhile and a lock it meets belongs to another process. It waits a
// little longer than other request connections, so an engine write of a few
// milliseconds at that instant delays the commit instead of failing it.
export function openDb(matterDir: string, options: { startup?: boolean; commit?: boolean } = {}): Database {
  const dir = path.join(matterDir, ".dochaus")
  mkdirSync(dir, { recursive: true })
  const db = new Database(path.join(dir, "legal.db"))
  const busyTimeoutMs = options.startup
    ? STARTUP_BUSY_TIMEOUT_MS
    : options.commit
      ? COMMIT_BUSY_TIMEOUT_MS
      : REQUEST_BUSY_TIMEOUT_MS
  db.run(`PRAGMA busy_timeout = ${busyTimeoutMs}`)
  db.run("PRAGMA journal_mode = WAL")
  db.run("CREATE TABLE IF NOT EXISTS answer_drafts (scope TEXT PRIMARY KEY, answers TEXT NOT NULL, revision INTEGER NOT NULL)")
  // injection_report: JSON-encoded prompt-injection findings from ingest-time
  // scanning (see sanitize.ts), NULL when the document came up clean.
  db.run(`
    CREATE TABLE IF NOT EXISTS documents (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL UNIQUE,
      name TEXT NOT NULL,
      created_at INTEGER NOT NULL,
      injection_report TEXT,
      role TEXT NOT NULL DEFAULT 'source' CHECK (role IN ('source', 'deliverable'))
    )
  `)
  // flagged: 1 when the chunk overlaps an ingest-time injection finding, so the
  // search-document tool can mark the passage as adversarial when handing it to
  // the model.
  // structural_path: the document's hierarchy down to this chunk's clause
  // ("Doc › §7 Indemnification › §7.2 Cap"), built deterministically from the
  // clause numbering at ingest (see structure.ts buildSectionPaths). It doubles
  // as a near-zero-cost contextual prefix for the vector channel and is surfaced
  // verbatim in search-document citations.
  // party: the document's named participants as a single line,
  // attached to every chunk so a passage carries its matter/party context even
  // when the names sit pages away in the preamble.
  db.run(`
    CREATE TABLE IF NOT EXISTS chunks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      section TEXT NOT NULL,
      structural_path TEXT NOT NULL DEFAULT '',
      party TEXT NOT NULL DEFAULT '',
      chunk_index INTEGER NOT NULL,
      text TEXT NOT NULL,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL,
      embedding BLOB NOT NULL,
      flagged INTEGER NOT NULL DEFAULT 0
    )
  `)
  // Databases created before the injection-defense columns existed (issue #17)
  // gain them here; their documents read as clean until re-ingested.
  if (!hasColumn(db, "documents", "injection_report")) db.run("ALTER TABLE documents ADD COLUMN injection_report TEXT")
  // Existing matter databases predate the source/deliverable boundary. Their
  // registered documents were all user-ingested sources, so the migration-safe
  // default preserves their current meaning.
  if (!hasColumn(db, "documents", "role"))
    db.run("ALTER TABLE documents ADD COLUMN role TEXT NOT NULL DEFAULT 'source'")
  if (!hasColumn(db, "documents", "label")) db.run("ALTER TABLE documents ADD COLUMN label TEXT")
  if (!hasColumn(db, "chunks", "flagged")) db.run("ALTER TABLE chunks ADD COLUMN flagged INTEGER NOT NULL DEFAULT 0")
  // Structure-aware chunking (issue R2): older databases carry chunks without a
  // structural path or party line. The columns are added empty; the rows fill in
  // on the next re-ingest of each document.
  if (!hasColumn(db, "chunks", "structural_path"))
    db.run("ALTER TABLE chunks ADD COLUMN structural_path TEXT NOT NULL DEFAULT ''")
  if (!hasColumn(db, "chunks", "party")) db.run("ALTER TABLE chunks ADD COLUMN party TEXT NOT NULL DEFAULT ''")
  // Recorded structure (see the sections table below). rev: the first 12 hex
  // digits of the sha256 of a source document's canonical text, the untrimmed
  // text rebuilt from its chunks (structure.ts reconstructText) that chunk
  // offsets and section spans index; deliverables have none. section_ordinal:
  // the section the chunk was cut from. Rows the sections backfill (structure.ts
  // migrateSections) has not yet reached hold NULL.
  if (!hasColumn(db, "documents", "rev")) db.run("ALTER TABLE documents ADD COLUMN rev TEXT")
  if (!hasColumn(db, "chunks", "section_ordinal")) db.run("ALTER TABLE chunks ADD COLUMN section_ordinal INTEGER")
  // Small key-value table for index-level facts that aren't rows — currently
  // which embedding model wrote the chunk vectors (see embed.ts migration).
  db.run("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
  // A source mutation installs filesystem bytes and projections before its
  // SQLite transaction commits. The transaction id is recorded here in that
  // same commit so startup recovery can distinguish a committed mutation from
  // one SQLite rolled back after process termination.
  db.run(`
    CREATE TABLE IF NOT EXISTS source_mutation_commits (
      mutation_id TEXT PRIMARY KEY,
      target_name TEXT NOT NULL,
      committed_at INTEGER NOT NULL
    )
  `)
  db.run(`
    CREATE TABLE IF NOT EXISTS document_mutation_commits (
      mutation_id TEXT PRIMARY KEY,
      target_name TEXT NOT NULL,
      committed_at INTEGER NOT NULL
    )
  `)
  // Chat continuations after an interrupted model turn. Keep the claim beside
  // the matter so two browsers or devices cannot dispatch the same
  // server-issued message ID concurrently. The engine response is reconciled
  // separately; this table is the durable idempotency boundary.
  db.run(`
    CREATE TABLE IF NOT EXISTS continuation_dispatches (
      message_id TEXT PRIMARY KEY,
      session_id TEXT NOT NULL,
      kind TEXT NOT NULL,
      run_id TEXT NOT NULL,
      progress_token TEXT NOT NULL,
      payload_hash TEXT NOT NULL,
      status TEXT NOT NULL CHECK (status IN ('claimed', 'accepted', 'observed', 'uncertain')),
      attempts INTEGER NOT NULL DEFAULT 1,
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL
    )
  `)
  // Lexical channel for hybrid retrieval (issue #67): a BM25-ranked FTS5 index
  // over the same chunks the vector channel scans. external-content mode stores
  // only the index and reads row text back from chunks, so chunk text is never
  // duplicated. The section label and document name are indexed alongside the
  // body so a query naming a clause ("Section 8.3") or a document matches that
  // clause's or document's own chunks directly.
  // Triggers keep the index in sync with every write path (insertChunk,
  // upsertDocument's delete-then-reinsert, deleteDocument). External-content
  // FTS5 requires the special 'delete' insert form to unindex a row. Only the
  // indexed columns re-index a row on update: embedding and section_ordinal
  // writes (the embedding and sections backfills) leave the index alone.
  const insertTrigger = `CREATE TRIGGER IF NOT EXISTS chunks_fts_ai AFTER INSERT ON chunks BEGIN
      INSERT INTO chunks_fts(rowid, text, section, doc_name) VALUES (new.id, new.text, new.section, new.doc_name);
    END`
  const deleteTrigger = `CREATE TRIGGER IF NOT EXISTS chunks_fts_ad AFTER DELETE ON chunks BEGIN
      INSERT INTO chunks_fts(chunks_fts, rowid, text, section, doc_name) VALUES ('delete', old.id, old.text, old.section, old.doc_name);
    END`
  const updateTrigger = `CREATE TRIGGER chunks_fts_au AFTER UPDATE OF text, section, doc_name ON chunks BEGIN
      INSERT INTO chunks_fts(chunks_fts, rowid, text, section, doc_name) VALUES ('delete', old.id, old.text, old.section, old.doc_name);
      INSERT INTO chunks_fts(rowid, text, section, doc_name) VALUES (new.id, new.text, new.section, new.doc_name);
    END`
  // Databases indexed before doc_name was added carry a two-column chunks_fts;
  // FTS5 has no ALTER, so drop and recreate. An index this open creates, for
  // such a database or one from before the index existed, is built from the
  // chunks already stored and committed together with its triggers, so a chunk
  // another process writes the moment it commits is indexed too.
  const ftsSql = (db.query("SELECT sql FROM sqlite_master WHERE name = 'chunks_fts'").get() as { sql: string } | null)
    ?.sql
  if (!ftsSql?.includes("doc_name"))
    db.transaction(() => {
      for (const trigger of ["chunks_fts_ai", "chunks_fts_ad", "chunks_fts_au"]) db.run(`DROP TRIGGER IF EXISTS ${trigger}`)
      if (ftsSql) db.run("DROP TABLE chunks_fts")
      db.run(`
        CREATE VIRTUAL TABLE chunks_fts USING fts5(
          text, section, doc_name,
          content='chunks', content_rowid='id',
          tokenize='unicode61 remove_diacritics 2'
        )
      `)
      db.run("INSERT INTO chunks_fts(chunks_fts) VALUES ('rebuild')")
      for (const trigger of [insertTrigger, deleteTrigger, updateTrigger]) db.run(trigger)
    })()
  // A current index that lost a trigger gets it back here, and the startup
  // check below heals what was written meanwhile. Databases created with the
  // earlier any-column update trigger get the current one in its place; the
  // definition check keeps a current database's schema untouched.
  db.run(insertTrigger)
  db.run(deleteTrigger)
  const currentUpdateTrigger = db.query("SELECT sql FROM sqlite_master WHERE type = 'trigger' AND name = 'chunks_fts_au'").get() as
    | { sql: string }
    | null
  if (currentUpdateTrigger?.sql !== updateTrigger)
    db.transaction(() => {
      db.run("DROP TRIGGER IF EXISTS chunks_fts_au")
      db.run(updateTrigger)
    })()
  // Self-heal drift at startup: chunks written or updated while the triggers
  // did not exist or did not match the index, by an older build or tool. Drift
  // is not always a row-count mismatch: a desynced external-content index can
  // hold the right number of rows with the wrong tokens, and then any
  // trigger-driven 'delete' (e.g. a chunk text UPDATE) throws SQLITE_CORRUPT
  // ("database disk image is malformed") even though the file is intact. FTS5's
  // own 'integrity-check' command compares the index against the chunks table
  // and is the only complete drift probe; 'rebuild' atomically reindexes.
  // The check scans the whole index and, being an INSERT, needs the write lock,
  // so it runs once per matter at startup rather than on every request. Drift
  // cannot arise later: the triggers are part of the schema, so every
  // connection in every process maintains the index in the same transaction as
  // the chunk write, and the index and trigger changes above are atomic. A
  // database the service has not seen gets its index built when this open
  // creates it.
  if (options.startup && !ftsInSync(db)) db.run("INSERT INTO chunks_fts(chunks_fts) VALUES ('rebuild')")
  // Contract-structure tables: deterministic, regex-extracted facts about each
  // document (see structure.ts). Every row is a pointer (offsets) into the
  // document's extracted text plus the verbatim text at that pointer — nothing
  // here is generated, so the lookup tools built on these tables cannot
  // fabricate. A parse miss means a missing row, and the tools say "not found"
  // and fall back to search; it never means a wrong fact.
  db.run(`
    CREATE TABLE IF NOT EXISTS defined_terms (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      term TEXT NOT NULL,
      definition TEXT NOT NULL,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL
    )
  `)
  db.run(`
    CREATE TABLE IF NOT EXISTS section_refs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      ref_kind TEXT NOT NULL,
      ref_label TEXT NOT NULL,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL
    )
  `)
  db.run(`
    CREATE TABLE IF NOT EXISTS parties (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      name TEXT NOT NULL,
      role TEXT,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL
    )
  `)
  // relation: currently only 'amends'. target_name is the verbatim title of the
  // referenced document as this document states it; resolution to an actual
  // matter document happens at lookup time by name match, never by guess.
  db.run(`
    CREATE TABLE IF NOT EXISTS doc_relations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      relation TEXT NOT NULL,
      target_name TEXT NOT NULL,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL
    )
  `)
  // Per-section amendment operations (issue #69). Each row is one operative
  // sentence in an amending document that names a section of its base agreement
  // and the verb (amend/restate/delete/add) applied to it. Spans index into the
  // amender's extracted text — same contract as doc_relations recital spans.
  db.run(`
    CREATE TABLE IF NOT EXISTS section_ops (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      section_label TEXT NOT NULL,
      operation TEXT NOT NULL,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL
    )
  `)
  // A source document's sections in document order, as ingest parsed them
  // (structure.ts documentSections). Spans index the text rebuilt from the
  // chunks (structure.ts reconstructText), in the chunk offsets' coordinates,
  // and tile it: each section starts where the previous one ends, and a
  // section's span is its own text up to the next heading of any depth, not its
  // children's. ordinal counts from 1; parent_ordinal is NULL and depth 1 at the
  // top level. label is the clause number and heading the heading text; both are
  // NULL for the opening text before the first heading. structure_source names
  // the parse that found the heading: 'sectionize' for a heading-shaped line,
  // 'markdown' for a heading the extractor marked with #.
  db.run(`
    CREATE TABLE IF NOT EXISTS sections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      ordinal INTEGER NOT NULL,
      parent_ordinal INTEGER,
      depth INTEGER NOT NULL,
      label TEXT,
      heading TEXT,
      char_start INTEGER NOT NULL,
      char_end INTEGER NOT NULL,
      structure_source TEXT NOT NULL
    )
  `)
  db.run("CREATE UNIQUE INDEX IF NOT EXISTS sections_doc_ordinal ON sections(doc_path, ordinal)")
  // A document's chunks in offset order (reconstructText, recordSections and the
  // per-document deletes); without it each read scans every chunk in the matter,
  // so a whole-matter backfill grows with the square of its size.
  db.run("CREATE INDEX IF NOT EXISTS chunks_doc_offset ON chunks(doc_path, char_start)")
  db.run("CREATE INDEX IF NOT EXISTS defined_terms_term ON defined_terms(term)")
  db.run("CREATE INDEX IF NOT EXISTS section_refs_doc ON section_refs(doc_path)")
  db.run("CREATE INDEX IF NOT EXISTS section_ops_doc ON section_ops(doc_path)")
  // Pending redline proposals. The canonical .docx stays clean (the accepted
  // state); each redline a tool proposes is a row here until a reviewer accepts
  // it (baked into the doc) or rejects it. scope drives how the edit is replayed:
  // 'phrase' is a surgical find/replace, 'clause' rewrites a located paragraph.
  // The redline tools (dochaus/tool/{redline,tracked-changes}.ts) create the same
  // table independently, so keep this DDL in sync with them.
  db.run(`
    CREATE TABLE IF NOT EXISTS redlines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      scope TEXT NOT NULL,
      find_text TEXT NOT NULL,
      old_text TEXT NOT NULL,
      new_text TEXT NOT NULL,
      author TEXT NOT NULL,
      review_comment TEXT NOT NULL DEFAULT '',
      anchor_id TEXT,
      status TEXT NOT NULL DEFAULT 'pending',
      created_at INTEGER NOT NULL
    )
  `)
  if (!hasColumn(db, "redlines", "review_comment"))
    db.run("ALTER TABLE redlines ADD COLUMN review_comment TEXT NOT NULL DEFAULT ''")
  if (!hasColumn(db, "redlines", "target_json"))
    db.run("ALTER TABLE redlines ADD COLUMN target_json TEXT")
  return db
}

// FTS5 reports an index/content mismatch only by throwing SQLITE_CORRUPT from
// its 'integrity-check' command (rank=1 verifies against the content table);
// there is no boolean API, so that error is the result. Any other error, such
// as SQLITE_BUSY while another connection holds the write lock, says nothing
// about the index and is rethrown rather than read as drift.
export function ftsInSync(db: Database) {
  try {
    db.run("INSERT INTO chunks_fts(chunks_fts, rank) VALUES ('integrity-check', 1)")
    return true
  } catch (error) {
    if ((error as { code?: unknown }).code?.toString().startsWith("SQLITE_CORRUPT")) return false
    throw error
  }
}

function hasColumn(db: Database, table: string, column: string) {
  return (db.query(`PRAGMA table_info(${table})`).all() as { name: string }[]).some((c) => c.name === column)
}

export function getMeta(db: Database, key: string): string | null {
  return (db.query("SELECT value FROM meta WHERE key = ?").get(key) as { value: string } | null)?.value ?? null
}

export function setMeta(db: Database, key: string, value: string) {
  db.run("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value", [
    key,
    value,
  ])
}

export type RedlineRow = {
  id: number
  doc_path: string
  doc_name: string
  scope: "phrase" | "clause"
  find_text: string
  old_text: string
  new_text: string
  author: string
  review_comment: string
  anchor_id: string | null
  target_json?: string | null
  // 'superseded': a later proposal on the same paragraph replaced this one before
  // review (set by the redline tools), so it never reaches the pending queue.
  status: "pending" | "accepted" | "rejected" | "superseded"
  created_at: number
}

// Pending redlines for one document, oldest first — the order they are replayed
// when building the redlined view and when accepting in bulk.
export function listPendingRedlines(db: Database, docPath: string): RedlineRow[] {
  const identity = documentRegistryIdentity(docPath)
  return (db.query("SELECT * FROM redlines WHERE status = 'pending' ORDER BY created_at, id").all() as RedlineRow[])
    .filter((row) => documentRegistryIdentity(row.doc_path) === identity)
}

export function getRedline(db: Database, id: number): RedlineRow | null {
  return (db.query("SELECT * FROM redlines WHERE id = ?").get(id) as RedlineRow) ?? null
}

export function setRedlineStatus(db: Database, id: number, status: "accepted" | "rejected") {
  db.run("UPDATE redlines SET status = ? WHERE id = ?", [status, id])
}

// Per-document pending counts, keyed by canonical document identity. Historical
// aliases stay in the database; the docs rail resolves its registry paths alike.
export function pendingRedlineCounts(db: Database): Record<string, number> {
  const rows = db
    .query("SELECT doc_path, COUNT(*) AS n FROM redlines WHERE status = 'pending' GROUP BY doc_path")
    .all() as { doc_path: string; n: number }[]
  return rows.reduce<Record<string, number>>((counts, row) => {
    const identity = documentRegistryIdentity(row.doc_path)
    counts[identity] = (counts[identity] ?? 0) + row.n
    return counts
  }, {})
}

export function listDocuments(db: Database) {
  return db.query("SELECT id, doc_path, name, created_at, role, label FROM documents ORDER BY created_at").all()
}

export function listSourceDocuments(db: Database) {
  return db
    .query("SELECT id, doc_path, name, created_at, role FROM documents WHERE role = 'source' ORDER BY created_at")
    .all()
}

export function getDocumentRole(db: Database, docPath: string): "source" | "deliverable" | null {
  const roles = [...new Set(matchingDocumentRows(db, docPath).map((row) => row.role))]
  if (roles.length > 1)
    throw new Error(`One physical matter document has conflicting source and deliverable registry roles.`)
  return roles[0] ?? null
}

// Drop a document and its chunks from the index. Pairs with removing the source
// .docx so the matter holds no orphaned embeddings.
export function deleteDocument(db: Database, docPath: string) {
  matchingDocumentRows(db, docPath).forEach((row) => {
    deleteDocumentIndexData(db, row.doc_path)
    db.run("DELETE FROM redlines WHERE doc_path = ?", [row.doc_path])
    db.run("DELETE FROM documents WHERE id = ?", [row.id])
  })
}

function deleteDocumentIndexData(db: Database, docPath: string) {
  db.run("DELETE FROM chunks WHERE doc_path = ?", [docPath])
  for (const table of ["defined_terms", "section_refs", "parties", "doc_relations", "section_ops", "sections"])
    db.run(`DELETE FROM ${table} WHERE doc_path = ?`, [docPath])
}

// Re-ingesting a document replaces its rows so the index never holds stale chunks.
export function upsertDocument(
  db: Database,
  docPath: string,
  name: string,
  createdAt: number,
  injectionReport: string | null,
): number {
  if (getDocumentRole(db, docPath) === "deliverable")
    throw new Error(`Cannot replace generated deliverable ${name} with a source document.`)
  const previous = matchingDocumentRows(db, docPath)
  const label = (db.query("SELECT label FROM documents WHERE name = ?").get(name) as { label: string | null } | null)?.label ?? null
  previous.forEach((row) => {
    deleteDocumentIndexData(db, row.doc_path)
    db.run("UPDATE redlines SET doc_path = ?, doc_name = ? WHERE doc_path = ?", [docPath, name, row.doc_path])
    db.run("DELETE FROM documents WHERE id = ?", [row.id])
  })
  const result = db.run("INSERT INTO documents (doc_path, name, created_at, injection_report, role, label) VALUES (?, ?, ?, ?, 'source', ?)", [
    docPath,
    name,
    createdAt,
    injectionReport,
    label,
  ])
  return Number(result.lastInsertRowid)
}

// Register generated output for the Documents rail without admitting it to the
// retrieval index. A deliverable path may only replace an earlier deliverable;
// replacing a user-ingested source would silently destroy the source boundary.
export function upsertDeliverable(db: Database, docPath: string, name: string, createdAt: number): number {
  if (getDocumentRole(db, docPath) === "source")
    throw new Error(`Cannot replace source document ${name} with a generated deliverable.`)
  const label = (db.query("SELECT label FROM documents WHERE name = ?").get(name) as { label: string | null } | null)?.label ?? null
  deleteDocument(db, docPath)
  const result = db.run(
    "INSERT INTO documents (doc_path, name, created_at, injection_report, role, label) VALUES (?, ?, ?, NULL, 'deliverable', ?)",
    [docPath, name, createdAt, label],
  )
  return Number(result.lastInsertRowid)
}

// Refresh an existing deliverable while preserving pending redline rows. Alias
// rows from older versions are collapsed onto the canonical physical path so a
// later source upload cannot observe a different role through casing or a
// symlink spelling.
export function refreshDeliverable(db: Database, docPath: string, name: string, createdAt: number) {
  if (getDocumentRole(db, docPath) !== "deliverable")
    throw new Error(`${name} is not a generated deliverable.`)
  const previous = matchingDocumentRows(db, docPath)
  const label = (db.query("SELECT label FROM documents WHERE name = ?").get(name) as { label: string | null } | null)?.label ?? null
  previous.forEach((row) => {
    deleteDocumentIndexData(db, row.doc_path)
    db.run("UPDATE redlines SET doc_path = ?, doc_name = ? WHERE doc_path = ?", [docPath, name, row.doc_path])
    db.run("DELETE FROM documents WHERE id = ?", [row.id])
  })
  const result = db.run(
    "INSERT INTO documents (doc_path, name, created_at, injection_report, role, label) VALUES (?, ?, ?, NULL, 'deliverable', ?)",
    [docPath, name, createdAt, label],
  )
  return Number(result.lastInsertRowid)
}

// The ingest-time injection findings for one document, for the document text route
// (so read-document can warn the model alongside the full text). NULL means clean.
export function getInjectionReport(db: Database, docPath: string) {
  getDocumentRole(db, docPath)
  const row = matchingDocumentRows(db, docPath)[0]
  return row?.injection_report ? (JSON.parse(row.injection_report) as { findings: unknown[] }) : null
}

type DocumentRegistryRow = {
  id: number
  doc_path: string
  name: string
  role: "source" | "deliverable"
  injection_report: string | null
}

function matchingDocumentRows(db: Database, docPath: string) {
  const identity = documentRegistryIdentity(docPath)
  return (db
    .query("SELECT id, doc_path, name, role, injection_report FROM documents")
    .all() as DocumentRegistryRow[]).filter((row) => documentRegistryIdentity(row.doc_path) === identity)
}

export function insertChunk(
  db: Database,
  chunk: {
    documentId: number
    docPath: string
    docName: string
    section: string
    structuralPath: string
    party: string
    chunkIndex: number
    text: string
    charStart: number
    charEnd: number
    embedding: Float32Array
    flagged: boolean
  },
) {
  db.run(
    "INSERT INTO chunks (document_id, doc_path, doc_name, section, structural_path, party, chunk_index, text, char_start, char_end, embedding, flagged) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [
      chunk.documentId,
      chunk.docPath,
      chunk.docName,
      chunk.section,
      chunk.structuralPath,
      chunk.party,
      chunk.chunkIndex,
      chunk.text,
      chunk.charStart,
      chunk.charEnd,
      Buffer.from(chunk.embedding.buffer, chunk.embedding.byteOffset, chunk.embedding.byteLength),
      chunk.flagged ? 1 : 0,
    ],
  )
}
