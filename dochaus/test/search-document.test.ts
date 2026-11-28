import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

const tempDirs: string[] = []
afterEach(() => {
  for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

function matter() {
  const dir = mkdtempSync(path.join(tmpdir(), "search-document-"))
  tempDirs.push(dir)
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
    `CREATE TABLE redlines (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, scope TEXT, find_text TEXT, old_text TEXT, new_text TEXT, author TEXT, review_comment TEXT NOT NULL DEFAULT '', anchor_id TEXT, target_json TEXT, status TEXT, created_at INTEGER)`,
    `CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)`,
    `CREATE VIRTUAL TABLE chunks_fts USING fts5(text, section, doc_name, content='chunks', content_rowid='id')`,
  ]) {
    db.run(sql)
  }
  db.run("INSERT INTO meta (key, value) VALUES ('embedding_model', 'onnx-community/granite-embedding-small-english-r2-ONNX')")
  return { dir, db }
}

// Keep the embedding stub in a subprocess so it never replaces @huggingface/transformers
// in another test file during a shared bun test run.
async function runSearches(directory: string, sessionID: string, calls: Record<string, unknown>[]) {
  const child = Bun.spawn([process.execPath, "--no-env-file", "-e", `
    import { mock, spyOn } from "bun:test"
    mock.module("@huggingface/transformers", () => ({
      env: {},
      pipeline: async () => async () => ({ data: new Float32Array(384) }),
    }))
    const searchDocument = (await import("../tool/search-document")).default
    const input = JSON.parse(await Bun.stdin.text())
    const warnings = []
    const log = spyOn(console, "warn").mockImplementation((...args) => { warnings.push(args) })
    const results = []
    try {
      const ctx = { directory: input.directory, sessionID: input.sessionID }
      for (const args of input.calls) {
        results.push(await searchDocument.execute(args, ctx))
      }
    } finally {
      log.mockRestore()
    }
    console.log(JSON.stringify({ results, warnings }))
  `], {
    cwd: import.meta.dir,
    env: { PATH: process.env.PATH ?? "" },
    stdin: "pipe",
    stdout: "pipe",
    stderr: "pipe",
  })
  child.stdin.write(JSON.stringify({ directory, sessionID, calls }))
  child.stdin.end()
  const [stdout, stderr, code] = await Promise.all([
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
    child.exited,
  ])
  expect({ code, stderr }).toEqual({ code: 0, stderr: "" })
  return JSON.parse(stdout) as { results: Array<string | { output: string }>; warnings: unknown[][] }
}

const textOf = (result: string | { output: string }) => (typeof result === "string" ? result : result.output)

test("search-document defaults malformed empty k to the normal result count", async () => {
  const { dir, db } = matter()
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

  const { results } = await runSearches(dir, "search-test", [{ query: "engagement letter", k: "" }])
  expect(textOf(results[0]!)).toContain("Source Document")
})

test("search-document does not treat repeated empty searches as a loop", async () => {
  const { dir, db } = matter()
  db.close()

  const { results } = await runSearches(dir, `empty-search-${crypto.randomUUID()}`, [
    { query: "governing law" },
    { query: "jurisdiction" },
  ])
  expect(textOf(results[0]!)).toContain("No relevant passages found.")
  expect(textOf(results[1]!)).toContain("No relevant passages found.")
  expect(textOf(results[1]!)).not.toContain("already returned")
})

test("search-document distinguishes different chunks in the same section and logs redacted loop diagnostics", async () => {
  const { dir, db } = matter()
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
    const id = Number((db.query("SELECT last_insert_rowid() AS id").get() as { id: number }).id)
    db.run("INSERT INTO chunks_fts(rowid, text, section, doc_name) VALUES (?, ?, '1', 'Source Document')", [id, text])
  }
  db.close()

  const { results, warnings } = await runSearches(dir, `same-section-${crypto.randomUUID()}`, [
    { query: "alpha", k: 1 },
    { query: "beta", k: 1 },
    { query: "alpha", k: 1 },
  ])

  expect(textOf(results[0]!)).toContain("unique term alpha")
  expect(textOf(results[1]!)).toContain("unique term beta")
  expect(textOf(results[1]!)).not.toContain("already returned")
  expect(textOf(results[2]!)).toContain("already returned")
  expect(warnings).toHaveLength(1)
  const logged = JSON.stringify(warnings)
  expect(logged).not.toContain("alpha")
  expect(logged).not.toContain("same-section-")
  expect(logged).toContain("query_ref")
  expect(logged).toContain('"result_count":1')
})
