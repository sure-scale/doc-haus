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

test("restated section returns latest amender's operative wording + lineage", async () => {
  const { dir, db } = matter()
  seedChain(db)
  // Each amender's operative statement sits at its op span (chars 50..100).
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "1", "AM1 operative statement.", 50, 74])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "1", "AM2 operative statement.", 50, 74])
  // The amenders' own clauses labelled "4" are not the base's §4.
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "AM1 own clause 4.", 200, 217])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "4", "AM2 own clause 4.", 200, 217])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "amend", 50, 74])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM2, "Amendment No. 2", "4", "restate", 50, 74])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  expect(result).toMatch(/Current state of §4/)
  expect(result).toMatch(/Latest operative source: Amendment No. 2/)
  expect(result).toContain('Operative wording of Amendment No. 2 (verbatim, chars 50..74)')
  expect(result).toContain('"AM2 operative statement."')
  // The newest entry's quote is not repeated in the lineage.
  expect(result.split("AM2 operative statement.").length - 1).toBe(1)
  expect(result).not.toMatch(/own clause 4/)
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

test("amender op span not recoverable from the index says the current text is not resolved", async () => {
  const { dir, db } = matter()
  seedChain(db)
  // AM1 amends §4, but no indexed chunk covers its op span (chars 50..100).
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "amend", 50, 100])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  expect(result).toMatch(/current text of §4 is not resolved/)
  expect(result).not.toMatch(/Operative wording of/)
  // The indexed text has a gap before the op, so no code-point position is given.
  expect(result).toContain("Next step: read Amendment No. 1 around chars 50..100; its position is unavailable because the indexed text is incomplete.")
  expect(result).not.toMatch(/read-document|code-point offset/)
})

test("amender's own same-numbered clause is never presented as the base section's current text", async () => {
  // Amendment No. 1 restates base §4 in its own §1; its own §4 is Governing Law.
  const { dir, db } = matter()
  const statement = "Section 4 of the Agreement is hereby amended and restated in its entirety to read as follows:"
  const parts = [
    { section: "", text: "AMENDMENT NO. 1 TO MASTER SERVICES AGREEMENT\n" },
    { section: "1", text: `1. Fees. ${statement}\n“4. Fees. Customer shall pay each invoice within forty-five (45) days.”\n` },
    { section: "4", text: "4. Governing Law. This Amendment is governed by the laws of the State of New York.\n" },
  ]
  const full = parts.map((p) => p.text).join("")
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [BASE, "Master Services Agreement", 1])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [AM1, "Amendment No. 1", 2])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [BASE, "Master Services Agreement", "4", "4. Fees. Customer shall pay each invoice within thirty (30) days.", 0, 66])
  let at = 0
  for (const part of parts) {
    db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", part.section, part.text, at, at + part.text.length])
    at += part.text.length
  }
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "amends", "Master Services Agreement", 0, 44])
  const opStart = full.indexOf(statement)
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "restate", opStart, opStart + statement.length])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  const following = "Text following the amendment's statement in Amendment No. 1 (verbatim; may contain the restated wording; not verified as the resulting text of §4):"
  expect(result).not.toMatch(/Current text of §4/)
  expect(result).toMatch(/Latest operative source: Amendment No\. 1 — restated §4/)
  expect(result).toContain(`Operative wording of Amendment No. 1 (verbatim, chars ${opStart}..${opStart + statement.length})`)
  expect(result).toContain(`"${statement}"`)
  expect(result).toMatch(/current text of §4 is not resolved/)
  // The amender's own §4 appears only as part of the positional text after the
  // statement, labelled as unverified, never as the base section's text.
  expect(result).toContain(`${following}\n“4. Fees. Customer shall pay each invoice within forty-five (45) days.”\n4. Governing Law.`)
  expect(result.slice(0, result.indexOf(following))).not.toMatch(/Governing Law/)
  expect(result).toContain(`Next step: read Amendment No. 1 from code-point offset ${opStart} (${full.length - opStart} characters)`)
  expect(result).not.toMatch(/read-document/)
  expect(result.split(statement).length - 1).toBe(1)
})

test("unresolved base: an untouched section lists non-amending documents and says the base was not matched", async () => {
  const { dir, db } = matter()
  const SIDE = "/m/side.docx"
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [AM1, "Amendment No. 1", 2])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [SIDE, "Side Letter", 3])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "4", "4. Governing Law. This Amendment is governed by the laws of the State of New York.", 0, 83])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [SIDE, "Side Letter", "4", "4. Notices. Notices go to the addresses above.", 0, 46])
  // The stated base matches no indexed document, so the chain root is unresolved.
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1", "amends", "Supply Agreement", 0, 10])
  db.close()

  const result = await tool.execute({ section: "4" }, ctxFor(dir))
  expect(result).not.toMatch(/Governing Law/)
  expect(result).not.toMatch(/Text below is from the base document/)
  expect(result).toMatch(/base document "Supply Agreement" was not matched/)
  expect(result).toMatch(/4\. Notices\./)
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

// The amendment states its base as "Lease"; the base file is named differently,
// so the chain root is unresolved.
function seedUnmatchedBase(db: Database) {
  const LEASE = "/m/lease-executed.docx"
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [LEASE, "Lease (executed).docx", 1])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [AM1, "Amendment No. 1.docx", 2])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [LEASE, "Lease (executed).docx", "4", "4. Rent. Base rent text.", 0, 24])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [LEASE, "Lease (executed).docx", "8", "8. Notices. Base notices text.", 100, 130])
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1.docx", "1", "Section 4 of the Lease is hereby amended and restated.", 0, 54])
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1.docx", "amends", "Lease", 0, 21])
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [AM1, "Amendment No. 1.docx", "4", "restate", 0, 54])
}

test("a document not matched to any chain base is never reported as outside every chain while a base is unmatched", async () => {
  const { dir, db } = matter()
  seedUnmatchedBase(db)
  db.close()

  const result = await tool.execute({ section: "4", document: "Lease (executed).docx" }, ctxFor(dir))
  expect(result).toMatch(/4\. Rent\. Base rent text\./)
  expect(result).not.toMatch(/no amendment chain involves/)
  expect(result).toContain('"Lease (executed).docx" was not matched to any amendment chain base')
  expect(result).toContain('stated bases not matched to an indexed document: "Lease"')
})

test("an unmatched base's stated name selects its chain", async () => {
  const { dir, db } = matter()
  seedUnmatchedBase(db)
  db.close()

  const touched = await tool.execute({ section: "4", document: "Lease" }, ctxFor(dir))
  expect(touched).toMatch(/Current state of §4 \(chain root: "Lease" \(unresolved\)\)/)
  expect(touched).toContain('"Section 4 of the Lease is hereby amended and restated."')
  // An untouched section is looked up across non-amending documents, never
  // scoped to the stated name, which is not a document.
  const untouched = await tool.execute({ section: "8", document: "Lease" }, ctxFor(dir))
  expect(untouched).toMatch(/base document "Lease" was not matched/)
  expect(untouched).toMatch(/8\. Notices\. Base notices text\./)
  const missing = await tool.execute({ section: "9", document: "Lease" }, ctxFor(dir))
  expect(missing).toMatch(/^No section "9" in this matter\./)
})
