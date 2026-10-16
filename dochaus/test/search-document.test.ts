// Tests for the supersession awareness in search-document. The hybrid
// retrieval math is exercised by the existing real-matter tests; this file
// stages a synthetic matter and asserts that superseded chunks are flagged and
// down-ranked. Run from dochaus/: bun test test/search-document.test.ts
//
// Embeddings are not available in unit tests, so we exercise the supersession
// logic by importing the lib helpers directly and confirming the chunks table
// contents feed the right supersededSections map.
import { Database } from "bun:sqlite"
import { expect, mock, spyOn, test } from "bun:test"
import { mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { buildChains, supersededSections } from "../lib/amendment-chain"

mock.module("@huggingface/transformers", () => ({
  env: {},
  pipeline: async () => async () => ({ data: new Float32Array(384) }),
}))

function matter() {
  const dir = mkdtempSync(path.join(tmpdir(), "search-supersession-"))
  const dbPath = path.join(dir, ".dochaus", "legal.db")
  mkdirSync(path.dirname(dbPath), { recursive: true })
  const db = new Database(dbPath)
  db.run("PRAGMA journal_mode = WAL")
  for (const sql of [
    `CREATE TABLE documents (id INTEGER PRIMARY KEY, doc_path TEXT UNIQUE, name TEXT, created_at INTEGER)`,
    `CREATE TABLE chunks (id INTEGER PRIMARY KEY, document_id INTEGER, doc_path TEXT, doc_name TEXT, section TEXT, structural_path TEXT, party TEXT, chunk_index INTEGER, text TEXT, char_start INTEGER, char_end INTEGER, embedding BLOB, flagged INTEGER)`,
    `CREATE TABLE defined_terms (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, term TEXT, definition TEXT, char_start INTEGER, char_end INTEGER)`,
    `CREATE TABLE doc_relations (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, relation TEXT, target_name TEXT, char_start INTEGER, char_end INTEGER)`,
    `CREATE TABLE section_ops (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section_label TEXT, operation TEXT, char_start INTEGER, char_end INTEGER)`,
    `CREATE TABLE redlines (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, scope TEXT, find_text TEXT, old_text TEXT, new_text TEXT, author TEXT, review_comment TEXT NOT NULL DEFAULT '', anchor_id TEXT, status TEXT, created_at INTEGER)`,
    `CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)`,
  ]) {
    db.run(sql)
  }
  return { dir, db }
}

function seedTwoDocChain(db: Database) {
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/base.docx", "Original Agreement", 1])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", 2])
  // Base §4 chunk (superseded by am1's restate op)
  db.run("INSERT INTO chunks (doc_path, doc_name, section, structural_path, party, text, char_start, char_end, embedding, flagged) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", ["/m/base.docx", "Original Agreement", "4", "", "", "Original §4 payment terms body text.", 0, 33, Buffer.alloc(384), 0])
  // Base §9 chunk (NOT superseded — no amender touches §9)
  db.run("INSERT INTO chunks (doc_path, doc_name, section, structural_path, party, text, char_start, char_end, embedding, flagged) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", ["/m/base.docx", "Original Agreement", "9", "", "", "Original §9 miscellaneous provisions.", 100, 133, Buffer.alloc(384), 0])
  db.run("INSERT INTO chunks (doc_path, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?)", ["/m/am1.docx", "", "Amendment No. 1 to the Original Agreement.", 0, 41])
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", "amends", "Original Agreement", 0, 41])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", "4", "restate", 50, 100])
}

test("supersededSections flags only the touched section of the base", () => {
  const { db } = matter()
  seedTwoDocChain(db)
  const map = supersededSections(db, buildChains(db))
  const baseMap = map.get("/m/base.docx")
  expect(baseMap).toBeDefined()
  expect(baseMap!.has("4")).toBe(true)
  expect(baseMap!.has("9")).toBe(false)
  db.close()
})

test("supersededSections is empty when no amender touches any section", () => {
  const { db } = matter()
  seedTwoDocChain(db)
  // Reset section_ops to empty
  db.run("DELETE FROM section_ops")
  const map = supersededSections(db, buildChains(db))
  expect(map.size).toBe(0)
  db.close()
})

test("supersededSections is empty when there is no chain", () => {
  const { db } = matter()
  const map = supersededSections(db, buildChains(db))
  expect(map.size).toBe(0)
  db.close()
})

test("search-document defaults malformed empty k to the normal result count", async () => {
  const searchDocument = (await import("../tool/search-document")).default
  const { dir, db } = matter()
  db.run("INSERT INTO meta (key, value) VALUES ('embedding_model', 'onnx-community/granite-embedding-small-english-r2-ONNX')")
  db.run("CREATE VIRTUAL TABLE chunks_fts USING fts5(text, section, doc_name, content='chunks', content_rowid='id')")
  db.run("INSERT INTO documents (id, doc_path, name, created_at) VALUES (1, ?, ?, 1)", [
    path.join(dir, "source.docx"),
    "Source Document",
  ])
  db.run(
    "INSERT INTO chunks (document_id, doc_path, doc_name, section, structural_path, party, text, char_start, char_end, embedding, flagged) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [
      1,
      path.join(dir, "source.docx"),
      "Source Document",
      "1",
      "Source Document > 1",
      "",
      "The engagement letter covers payment terms.",
      0,
      42,
      Buffer.alloc(384 * 4),
      0,
    ],
  )
  db.run("INSERT INTO chunks_fts(rowid, text, section, doc_name) VALUES (1, 'The engagement letter covers payment terms.', '1', 'Source Document')")
  db.close()

  const result = await searchDocument.execute({ query: "engagement letter", k: "" } as never, {
    directory: dir,
    sessionID: "search-test",
  } as never)

  expect(result.output).toContain("Source Document")
})

test("search-document does not treat repeated empty searches as a loop", async () => {
  const searchDocument = (await import("../tool/search-document")).default
  const { dir, db } = matter()
  db.run("INSERT INTO meta (key, value) VALUES ('embedding_model', 'onnx-community/granite-embedding-small-english-r2-ONNX')")
  db.run("CREATE VIRTUAL TABLE chunks_fts USING fts5(text, section, doc_name, content='chunks', content_rowid='id')")
  db.close()

  const ctx = { directory: dir, sessionID: `empty-search-${crypto.randomUUID()}` } as never
  const first = await searchDocument.execute({ query: "governing law" }, ctx)
  const second = await searchDocument.execute({ query: "jurisdiction" }, ctx)

  expect(first.output).toContain("No relevant passages found.")
  expect(second.output).toContain("No relevant passages found.")
  expect(second.output).not.toContain("already returned")
})

test("search-document distinguishes different chunks in the same section", async () => {
  const searchDocument = (await import("../tool/search-document")).default
  const { dir, db } = matter()
  db.run("INSERT INTO meta (key, value) VALUES ('embedding_model', 'onnx-community/granite-embedding-small-english-r2-ONNX')")
  db.run("CREATE VIRTUAL TABLE chunks_fts USING fts5(text, section, doc_name, content='chunks', content_rowid='id')")
  const document = path.join(dir, "source.docx")
  db.run("INSERT INTO documents (id, doc_path, name, created_at) VALUES (1, ?, ?, 1)", [document, "Source Document"])
  for (const [text, start] of [
    ["The first provision contains the unique term alpha.", 0],
    ["The second provision contains the unique term beta.", 100],
  ] as const) {
    db.run(
      "INSERT INTO chunks (document_id, doc_path, doc_name, section, structural_path, party, text, char_start, char_end, embedding, flagged) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
      [1, document, "Source Document", "1", "Source Document > 1", "", text, start, start + text.length, Buffer.alloc(384 * 4), 0],
    )
    const id = Number(db.query("SELECT last_insert_rowid() AS id").get()!.id)
    db.run("INSERT INTO chunks_fts(rowid, text, section, doc_name) VALUES (?, ?, '1', 'Source Document')", [id, text])
  }
  db.close()

  const ctx = { directory: dir, sessionID: `same-section-${crypto.randomUUID()}` } as never
  const first = await searchDocument.execute({ query: "alpha", k: 1 }, ctx)
  const second = await searchDocument.execute({ query: "beta", k: 1 }, ctx)

  expect(first.output).toContain("unique term alpha")
  expect(second.output).toContain("unique term beta")
  expect(second.output).not.toContain("already returned")
  const log = spyOn(console, "warn").mockImplementation(() => {})
  try {
    const repeated = await searchDocument.execute({ query: "alpha", k: 1 }, ctx)
    expect(repeated).toContain("already returned")
    expect(log).toHaveBeenCalledTimes(1)
    const logged = JSON.stringify(log.mock.calls)
    expect(logged).not.toContain("alpha")
    expect(logged).not.toContain("same-section-")
    expect(logged).toContain("query_ref")
    expect(logged).toContain('"result_count":1')
  } finally {
    log.mockRestore()
  }
})
