// End-to-end tests for the current-section resolver tool.
// Run from dochaus/: bun test test/current-section.test.ts
import { Database } from "bun:sqlite"
import { expect, test } from "bun:test"
import { mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"

const tool = (await import("../tool/current-section")).default

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
  const dir = mkdtempSync(path.join(tmpdir(), "current-section-"))
  const dbPath = path.join(dir, ".dochaus", "legal.db")
  mkdirSync(path.dirname(dbPath), { recursive: true })
  const db = new Database(dbPath)
  db.run("PRAGMA journal_mode = WAL")
  for (const sql of [
    `CREATE TABLE documents (id INTEGER PRIMARY KEY, doc_path TEXT UNIQUE, name TEXT, created_at INTEGER)`,
    `CREATE TABLE chunks (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)`,
    `CREATE TABLE defined_terms (id INTEGER PRIMARY KEY, doc_path TEXT, term TEXT)`,
    `CREATE TABLE doc_relations (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, relation TEXT, target_name TEXT, char_start INTEGER, char_end INTEGER)`,
    `CREATE TABLE section_ops (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section_label TEXT, operation TEXT, char_start INTEGER, char_end INTEGER)`,
  ]) {
    db.run(sql)
  }
  return { dir, db }
}

const BASE = "/m/base.docx"
const AM1 = "/m/am1.docx"
const AM2 = "/m/am2.docx"

function seedChain(db: Database) {
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [BASE, "Original Agreement", 1])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [AM1, "Amendment No. 1", 2])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [AM2, "Amendment No. 2", 3])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [BASE, "Original Agreement", "4", "Original §4 text.", 0, 16])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [BASE, "Original Agreement", "8", "Original §8 text.", 100, 116])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "", "Amendment No. 1 to the Original Agreement.", 0, 41])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "", "Amendment No. 2 to the Original Agreement.", 0, 41])
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "amends", "Original Agreement", 0, 41])
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "amends", "Original Agreement", 0, 41])
}

test("restated section returns latest amender's replacement text + lineage", async () => {
  const { dir, db } = matter()
  seedChain(db)
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "AM1 replacement body.", 200, 220])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "4", "AM2 replacement body.", 200, 220])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "amend", 50, 100])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "4", "restate", 50, 100])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  expect(result).toMatch(/Current text of §4/)
  expect(result).toMatch(/Latest operative source: Amendment No. 2/)
  expect(result).toMatch(/AM2 replacement body\./)
  expect(result).toMatch(/Lineage \(newest first\):/)
  expect(result).toMatch(/Amendment No. 2.*restated/)
  expect(result).toMatch(/Amendment No. 1.*amended.*now superseded/)
  expect(result).toMatch(/Original Agreement §4.*base text.*superseded/)
})

test("operative amender is NOT marked superseded when a newer amender skips the queried section", async () => {
  const { dir, db } = matter()
  seedChain(db)
  // am1 amends §4 (operative for §4). am2 touches only §7 (different section).
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "AM1 replacement body.", 200, 220])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "amend", 50, 100])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "7", "delete", 50, 100])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  // Header names am1 as operative
  expect(result).toMatch(/Latest operative source: Amendment No\. 1/)
  // The same am1 line in the lineage must NOT carry the "(now superseded)" tag.
  // The line shape is "- Amendment No. 1 §4 (chars 50..100): amended — ..."
  expect(result).not.toMatch(/Amendment No\. 1 §4[^]*\(now superseded\)/)
})

test("deleted section returns 'deleted by' with no replacement text", async () => {
  const { dir, db } = matter()
  seedChain(db)
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "8", "delete", 50, 100])
  db.close()

  const result = await tool.execute({ section: "8" }, ctxFor(dir))
  expect(result).toMatch(/Deleted by Amendment No\. 1/)
  expect(result).not.toMatch(/Original §8 text/)
})

test("section not touched by any amender returns base text + note", async () => {
  const { dir, db } = matter()
  seedChain(db)
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "amend", 50, 100])
  db.close()

  const result = await tool.execute({ section: "8" }, ctxFor(dir))
  expect(result).toMatch(/Original §8 text\./)
  expect(result).toMatch(/no amendment modifies §8/)
})

test("amender has op but no replacement chunk falls back honestly", async () => {
  const { dir, db } = matter()
  seedChain(db)
  // AM1 amends §4 but has no chunk labeled "4" of its own.
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "amend", 50, 100])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  expect(result).toMatch(/replacement text not separately indexed/)
  expect(result).toMatch(/read-document/)
})

test("no chain in matter behaves like get-section with a note", async () => {
  const { dir, db } = matter()
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [BASE, "Original Agreement", 1])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [BASE, "Original Agreement", "4", "Original §4 text.", 0, 16])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  expect(result).toMatch(/Original §4 text\./)
  expect(result).toMatch(/no amendment chain in this matter/)
})

test("no chain and document specified returns that document's section text", async () => {
  const { dir, db } = matter()
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [BASE, "Original Agreement", 1])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [BASE, "Original Agreement", "4", "Original §4 text.", 0, 16])
  db.close()

  const result = await tool.execute({ section: "4", document: "Original Agreement" }, ctxFor(dir))
  expect(result).toMatch(/Original §4 text\./)
})
