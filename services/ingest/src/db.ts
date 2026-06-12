import { Database } from "bun:sqlite"
import { mkdirSync } from "node:fs"
import path from "node:path"

// Per-matter retrieval index. Lives inside the matter directory so it is scoped
// to that matter and rebuildable from the source .docx files at any time.

export function openDb(matterDir: string): Database {
  const dir = path.join(matterDir, ".dochaus")
  mkdirSync(dir, { recursive: true })
  const db = new Database(path.join(dir, "legal.db"))
  db.run("PRAGMA journal_mode = WAL")
  // injection_report: JSON-encoded prompt-injection findings from ingest-time
  // scanning (see sanitize.ts), NULL when the document came up clean.
  db.run(`
    CREATE TABLE IF NOT EXISTS documents (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL UNIQUE,
      name TEXT NOT NULL,
      created_at INTEGER NOT NULL,
      injection_report TEXT
    )
  `)
  // flagged: 1 when the chunk overlaps an ingest-time injection finding, so the
  // search-document tool can mark the passage as adversarial when handing it to
  // the model.
  db.run(`
    CREATE TABLE IF NOT EXISTS chunks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      section TEXT NOT NULL,
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
  if (!hasColumn(db, "chunks", "flagged")) db.run("ALTER TABLE chunks ADD COLUMN flagged INTEGER NOT NULL DEFAULT 0")
  // Lexical channel for hybrid retrieval (issue #67): a BM25-ranked FTS5 index
  // over the same chunks the vector channel scans. external-content mode stores
  // only the index and reads row text back from chunks, so chunk text is never
  // duplicated. The section label is indexed alongside the body so a query
  // naming a clause ("Section 8.3") matches the clause's own chunks directly.
  db.run(`
    CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(
      text, section,
      content='chunks', content_rowid='id',
      tokenize='unicode61 remove_diacritics 2'
    )
  `)
  // Triggers keep the index in sync with every write path (insertChunk,
  // upsertDocument's delete-then-reinsert, deleteDocument). External-content
  // FTS5 requires the special 'delete' insert form to unindex a row.
  db.run(`
    CREATE TRIGGER IF NOT EXISTS chunks_fts_ai AFTER INSERT ON chunks BEGIN
      INSERT INTO chunks_fts(rowid, text, section) VALUES (new.id, new.text, new.section);
    END
  `)
  db.run(`
    CREATE TRIGGER IF NOT EXISTS chunks_fts_ad AFTER DELETE ON chunks BEGIN
      INSERT INTO chunks_fts(chunks_fts, rowid, text, section) VALUES ('delete', old.id, old.text, old.section);
    END
  `)
  db.run(`
    CREATE TRIGGER IF NOT EXISTS chunks_fts_au AFTER UPDATE ON chunks BEGIN
      INSERT INTO chunks_fts(chunks_fts, rowid, text, section) VALUES ('delete', old.id, old.text, old.section);
      INSERT INTO chunks_fts(rowid, text, section) VALUES (new.id, new.text, new.section);
    END
  `)
  // Backfill databases that predate the FTS table, and self-heal any drift (a
  // crash between table creation and indexing, or chunks written while the
  // triggers did not exist yet) — a row-count mismatch is the one observable
  // symptom of every such state, and 'rebuild' atomically reindexes from chunks.
  const chunkCount = (db.query("SELECT COUNT(*) AS n FROM chunks").get() as { n: number }).n
  const ftsCount = (db.query("SELECT COUNT(*) AS n FROM chunks_fts").get() as { n: number }).n
  if (ftsCount !== chunkCount) db.run("INSERT INTO chunks_fts(chunks_fts) VALUES ('rebuild')")
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
      anchor_id TEXT,
      status TEXT NOT NULL DEFAULT 'pending',
      created_at INTEGER NOT NULL
    )
  `)
  return db
}

function hasColumn(db: Database, table: string, column: string) {
  return (db.query(`PRAGMA table_info(${table})`).all() as { name: string }[]).some((c) => c.name === column)
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
  anchor_id: string | null
  // 'superseded': a later proposal on the same paragraph replaced this one before
  // review (set by the redline tools), so it never reaches the pending queue.
  status: "pending" | "accepted" | "rejected" | "superseded"
  created_at: number
}

// Pending redlines for one document, oldest first — the order they are replayed
// when building the redlined view and when accepting in bulk.
export function listPendingRedlines(db: Database, docPath: string): RedlineRow[] {
  return db
    .query("SELECT * FROM redlines WHERE doc_path = ? AND status = 'pending' ORDER BY created_at, id")
    .all(docPath) as RedlineRow[]
}

export function getRedline(db: Database, id: number): RedlineRow | null {
  return (db.query("SELECT * FROM redlines WHERE id = ?").get(id) as RedlineRow) ?? null
}

export function setRedlineStatus(db: Database, id: number, status: "accepted" | "rejected") {
  db.run("UPDATE redlines SET status = ? WHERE id = ?", [status, id])
}

// Per-document pending counts, keyed by absolute doc_path, for the docs-rail badge.
export function pendingRedlineCounts(db: Database): Record<string, number> {
  const rows = db
    .query("SELECT doc_path, COUNT(*) AS n FROM redlines WHERE status = 'pending' GROUP BY doc_path")
    .all() as { doc_path: string; n: number }[]
  return Object.fromEntries(rows.map((r) => [r.doc_path, r.n]))
}

export function listDocuments(db: Database) {
  return db.query("SELECT id, doc_path, name, created_at FROM documents ORDER BY created_at").all()
}

// Drop a document and its chunks from the index. Pairs with removing the source
// .docx so the matter holds no orphaned embeddings.
export function deleteDocument(db: Database, docPath: string) {
  db.run("DELETE FROM chunks WHERE doc_path = ?", [docPath])
  db.run("DELETE FROM documents WHERE doc_path = ?", [docPath])
}

// Re-ingesting a document replaces its rows so the index never holds stale chunks.
export function upsertDocument(
  db: Database,
  docPath: string,
  name: string,
  createdAt: number,
  injectionReport: string | null,
): number {
  db.run("DELETE FROM chunks WHERE doc_path = ?", [docPath])
  db.run("DELETE FROM documents WHERE doc_path = ?", [docPath])
  const result = db.run("INSERT INTO documents (doc_path, name, created_at, injection_report) VALUES (?, ?, ?, ?)", [
    docPath,
    name,
    createdAt,
    injectionReport,
  ])
  return Number(result.lastInsertRowid)
}

// The ingest-time injection findings for one document, for the document text route
// (so read-document can warn the model alongside the full text). NULL means clean.
export function getInjectionReport(db: Database, docPath: string) {
  const row = db.query("SELECT injection_report FROM documents WHERE doc_path = ?").get(docPath) as {
    injection_report: string | null
  } | null
  return row?.injection_report ? (JSON.parse(row.injection_report) as { findings: unknown[] }) : null
}

export function insertChunk(
  db: Database,
  chunk: {
    documentId: number
    docPath: string
    docName: string
    section: string
    chunkIndex: number
    text: string
    charStart: number
    charEnd: number
    embedding: Float32Array
    flagged: boolean
  },
) {
  db.run(
    "INSERT INTO chunks (document_id, doc_path, doc_name, section, chunk_index, text, char_start, char_end, embedding, flagged) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [
      chunk.documentId,
      chunk.docPath,
      chunk.docName,
      chunk.section,
      chunk.chunkIndex,
      chunk.text,
      chunk.charStart,
      chunk.charEnd,
      Buffer.from(chunk.embedding.buffer, chunk.embedding.byteOffset, chunk.embedding.byteLength),
      chunk.flagged ? 1 : 0,
    ],
  )
}
