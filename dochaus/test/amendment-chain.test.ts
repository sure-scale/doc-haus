// End-to-end test of the amendment-chain tool against a synthetic matter.
// Run from dochaus/: bun test test/amendment-chain.test.ts
import { Database } from "bun:sqlite"
import { expect, test } from "bun:test"
import { mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

const tool = (await import("../tool/amendment-chain")).default

function ctxFor(directory: string) {
  return {
    sessionID: "test",
    messageID: "test",
    agent: "test",
    directory,
    worktree: directory,
    abort: new AbortController().signal,
    metadata() {},
    async ask() {},
  }
}

function matter() {
  const dir = mkdtempSync(path.join(tmpdir(), "amend-tool-"))
  const dbPath = path.join(dir, ".dochaus", "legal.db")
  require("node:fs").mkdirSync(path.dirname(dbPath), { recursive: true })
  const db = new Database(dbPath)
  db.run("PRAGMA journal_mode = WAL")
  for (const sql of [
    `CREATE TABLE documents (id INTEGER PRIMARY KEY, doc_path TEXT UNIQUE, name TEXT, created_at INTEGER)`,
    `CREATE TABLE chunks (id INTEGER PRIMARY KEY, doc_path TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)`,
    `CREATE TABLE defined_terms (id INTEGER PRIMARY KEY, doc_path TEXT, term TEXT)`,
    `CREATE TABLE doc_relations (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, relation TEXT, target_name TEXT, char_start INTEGER, char_end INTEGER)`,
    `CREATE TABLE section_ops (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section_label TEXT, operation TEXT, char_start INTEGER, char_end INTEGER)`,
  ]) {
    db.run(sql)
  }
  return { dir, db }
}

test("reports no amendment relationships on an empty matter", async () => {
  const { dir, db } = matter()
  db.close()
  const result = await tool.execute({}, ctxFor(dir))
  expect(result).toMatch(/No amendment relationships were extracted/)
})

test("reports chain with verbatim recital and per-amender Operates on line", async () => {
  const { dir, db } = matter()
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/base.docx", "Original Agreement", 1])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", 2])
  db.run("INSERT INTO chunks (doc_path, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?)", ["/m/am1.docx", "", "Amendment No. 1 to the Original Agreement dated 2026-01-01.", 0, 59])
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", "amends", "Original Agreement", 0, 59])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", "4", "restate", 100, 200])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", "7", "delete", 300, 400])
  db.close()

  const result = await tool.execute({}, ctxFor(dir))
  expect(result).toMatch(/Chain: Original Agreement is amended/)
  expect(result).toMatch(/Amendment No. 1 amends "Original Agreement"/)
  expect(result).toMatch(/Recital .*"Amendment No. 1 to the Original Agreement/)
  expect(result).toMatch(/Operates on:.*§4 \(restate\).*§7 \(delete\)/)
})

test("omits Operates on line when amender has no section_ops", async () => {
  const { dir, db } = matter()
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/base.docx", "Original Agreement", 1])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", 2])
  db.run("INSERT INTO chunks (doc_path, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?)", ["/m/am1.docx", "", "Amendment No. 1 to the Original Agreement.", 0, 40])
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", "amends", "Original Agreement", 0, 40])
  db.close()

  const result = await tool.execute({}, ctxFor(dir))
  expect(result).toMatch(/Amendment No. 1 amends/)
  expect(result).not.toMatch(/Operates on:/)
})

test("document arg filters to chains that involve that document", async () => {
  const { dir, db } = matter()
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/base.docx", "Original Agreement", 1])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", 2])
  db.run("INSERT INTO chunks (doc_path, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?)", ["/m/am1.docx", "", "Amendment No. 1 to the Original Agreement.", 0, 40])
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", ["/m/am1.docx", "Amendment No. 1", "amends", "Original Agreement", 0, 40])
  db.close()

  const result = await tool.execute({ document: "Original Agreement" }, ctxFor(dir))
  expect(result).toMatch(/Chain: Original Agreement/)

  const miss = await tool.execute({ document: "Nonexistent" }, ctxFor(dir))
  expect(miss).toMatch(/No extracted amendment relationship involves "Nonexistent"/)
})
