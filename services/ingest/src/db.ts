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
  db.run(`
    CREATE TABLE IF NOT EXISTS documents (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL UNIQUE,
      name TEXT NOT NULL,
      created_at INTEGER NOT NULL
    )
  `)
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
      embedding BLOB NOT NULL
    )
  `)
  return db
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
export function upsertDocument(db: Database, docPath: string, name: string, createdAt: number): number {
  db.run("DELETE FROM chunks WHERE doc_path = ?", [docPath])
  db.run("DELETE FROM documents WHERE doc_path = ?", [docPath])
  const result = db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [docPath, name, createdAt])
  return Number(result.lastInsertRowid)
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
  },
) {
  db.run(
    "INSERT INTO chunks (document_id, doc_path, doc_name, section, chunk_index, text, char_start, char_end, embedding) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
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
    ],
  )
}
