// Unit tests for the shared amendment-chain helper on a synthetic legal.db.
// Run from dochaus/: bun test lib/amendment-chain.test.ts
import { Database } from "bun:sqlite"
import { expect, test } from "bun:test"
import { mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { buildChains, operativePassage, resolveCurrentSection, supersededSections, verbatimSpan, type SectionOpRow } from "./amendment-chain"

function syntheticDb() {
  const dir = mkdtempSync(path.join(tmpdir(), "amend-helper-"))
  const db = new Database(path.join(dir, "legal.db"))
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
  return { db, dir }
}

type DocSeed = { path: string; name: string; createdAt: number }
type ChunkSeed = { docPath: string; section: string; text: string; charStart: number; charEnd: number }

function seed(db: Database, docs: DocSeed[], chunks: ChunkSeed[], relations: Record<string, unknown>[], ops: Record<string, unknown>[]) {
  for (const d of docs) db.run("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)", [d.path, d.name, d.createdAt])
  for (const c of chunks) db.run("INSERT INTO chunks (doc_path, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?)", [c.docPath, c.section, c.text, c.charStart, c.charEnd])
  for (const r of relations) db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [r.doc_path, r.doc_name, r.relation, r.target_name, r.char_start, r.char_end])
  for (const o of ops) db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [o.doc_path, o.doc_name, o.section_label, o.operation, o.char_start, o.char_end])
}

test("buildChains returns empty when there are no doc_relations", () => {
  const { db } = syntheticDb()
  expect(buildChains(db)).toEqual([])
  db.close()
})

test("buildChains resolves a single amendment to its base", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [],
  )
  const chains = buildChains(db)
  expect(chains).toHaveLength(1)
  expect(chains[0]!.base).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/base.docx", name: "Original Agreement" } })
  expect(chains[0]!.entries).toHaveLength(1)
  expect(chains[0]!.entries[0]!.rel.doc_name).toBe("Amendment No. 1")
  db.close()
})

test("buildChains flattens a transitive chain (Am2 -> Am1 -> Base)", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
      { path: "/m/am2.docx", name: "Amendment No. 2", createdAt: 3 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [],
  )
  const chains = buildChains(db)
  expect(chains).toHaveLength(1)
  expect(chains[0]!.entries.map((e) => e.rel.doc_name)).toEqual(["Amendment No. 1", "Amendment No. 2"])
  db.close()
})

test("buildChains attaches per-amender section_ops to entries", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "restate", char_start: 100, char_end: 200 },
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "7", operation: "delete", char_start: 300, char_end: 400 },
    ],
  )
  const chains = buildChains(db)
  const ops = chains[0]!.entries[0]!.ops
  expect(ops.map((o) => `${o.section_label}:${o.operation}`).sort()).toEqual(["4:restate", "7:delete"])
  db.close()
})

test("supersededSections maps base path to label -> latest amender op", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
      { path: "/m/am2.docx", name: "Amendment No. 2", createdAt: 3 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "amend", char_start: 100, char_end: 200 },
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", section_label: "4", operation: "restate", char_start: 100, char_end: 200 },
    ],
  )
  const chains = buildChains(db)
  const map = supersededSections(db, chains)
  const baseMap = map.get("/m/base.docx")
  expect(baseMap).toBeDefined()
  // Amendment No. 2 is newer, so its restate op wins for label "4".
  expect(baseMap!.get("4")?.op.operation).toBe("restate")
  expect(baseMap!.get("4")?.amender).toBe("Amendment No. 2")
  db.close()
})

test("supersededSections does not flag sections no amender touches", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "amend", char_start: 100, char_end: 200 },
    ],
  )
  const chains = buildChains(db)
  const map = supersededSections(db, chains)
  expect(map.get("/m/base.docx")?.has("5")).toBe(false)
  db.close()
})

test("resolveCurrentSection walks newest-first and returns the latest op", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
      { path: "/m/am2.docx", name: "Amendment No. 2", createdAt: 3 },
    ],
    [
      // am1's operative statement on §4 at its op span (superseded by am2)
      { docPath: "/m/am1.docx", section: "1", text: "OLD OPERATIVE STATEMENT", charStart: 100, charEnd: 123 },
      // am2's operative statement on §4 at its op span (operative)
      { docPath: "/m/am2.docx", section: "1", text: "NEW OPERATIVE STATEMENT", charStart: 100, charEnd: 123 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "amend", char_start: 100, char_end: 123 },
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", section_label: "4", operation: "restate", char_start: 100, char_end: 123 },
    ],
  )
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "4")
  expect(resolved?.operation).toBe("restate")
  expect(resolved?.amender.rel.doc_name).toBe("Amendment No. 2")
  expect(resolved?.operativeText).toBe("NEW OPERATIVE STATEMENT")
  db.close()
})

test("resolveCurrentSection returns delete with its op and no wording when the span is unindexed", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "8", operation: "delete", char_start: 100, char_end: 200 },
    ],
  )
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "8")
  expect(resolved?.operation).toBe("delete")
  expect(resolved?.op.char_start).toBe(100)
  expect(resolved?.operativeText).toBeNull()
  db.close()
})

test("resolveCurrentSection returns null when no amender touches the section", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "amend", char_start: 100, char_end: 200 },
    ],
  )
  const [chain] = buildChains(db)
  expect(resolveCurrentSection(db, chain!, "99")).toBeNull()
  db.close()
})

// A synthetic amendment whose own numbering collides with the base section it
// amends: its §1 restates base §4, while its own §4 is an unrelated clause.
// Chunks and the op span are laid out on one text so offsets match ingest.
function collidingAmendment(db: Database) {
  const statement = "Section 4 of the Agreement is hereby amended and restated in its entirety to read as follows:"
  const subStatement = "Section 4.2(a) of the Agreement is hereby amended by replacing “thirty (30) days” with “forty-five (45) days”"
  const parts = [
    // One astral character (two UTF-16 units, one code point) before the ops,
    // so read-document offsets must be converted to code points.
    { section: "", text: "AMENDMENT NO. 1 TO MASTER SERVICES AGREEMENT 𝔄\n" },
    { section: "1", text: `1. Fees. ${statement}\n“4. Fees. Customer shall pay each invoice within forty-five (45) days.”\n` },
    { section: "2", text: `2. Payment Terms. ${subStatement}\n` },
    { section: "4", text: "4. Governing Law. This Amendment is governed by the laws of the State of New York.\n" },
    { section: "4.2", text: "4.2 Counterparts. This Amendment may be signed in counterparts.\n" },
  ]
  const full = parts.map((p) => p.text).join("")
  let at = 0
  const chunks = parts.map((p) => {
    const chunk = { docPath: "/m/am1.docx", section: p.section, text: p.text, charStart: at, charEnd: at + p.text.length }
    at += p.text.length
    return chunk
  })
  const opStart = full.indexOf(statement)
  const subStart = full.indexOf(subStatement)
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Master Services Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [
      { docPath: "/m/base.docx", section: "4", text: "4. Fees. Customer shall pay each invoice within thirty (30) days.", charStart: 0, charEnd: 66 },
      ...chunks,
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Master Services Agreement", char_start: 0, char_end: 44 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "restate", char_start: opStart, char_end: opStart + statement.length },
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4.2(a)", operation: "amend", char_start: subStart, char_end: subStart + subStatement.length },
    ],
  )
  return { statement, subStatement, opStart, subStart, full }
}

test("resolveCurrentSection returns the amender's operative wording, never its own same-numbered clause", () => {
  const { db } = syntheticDb()
  const { statement, opStart } = collidingAmendment(db)
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "4")
  // The amendment's own §4 (Governing Law) is not the base's §4.
  expect(JSON.stringify(resolved)).not.toContain("Governing Law")
  expect(resolved?.operation).toBe("restate")
  expect(resolved?.amender.rel.doc_name).toBe("Amendment No. 1")
  expect(resolved?.op.char_start).toBe(opStart)
  expect(resolved?.operativeText).toBe(statement)
  db.close()
})

test("resolveCurrentSection quotes a sub-clause op's wording, not the amender's parent-numbered chunk", () => {
  const { db } = syntheticDb()
  const { subStatement } = collidingAmendment(db)
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "4.2(a)")
  expect(resolved?.operativeText).not.toContain("Counterparts")
  expect(resolved?.operation).toBe("amend")
  expect(resolved?.operativeText).toBe(subStatement)
  db.close()
})

test("resolveCurrentSection returns the text after the statement up to the next op, with code-point read arguments", () => {
  const { db } = syntheticDb()
  const { opStart, subStart } = collidingAmendment(db)
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "4")
  // Position-based only: everything between the §4 statement and the next
  // operative statement in the same amending document.
  expect(resolved?.followingText).toBe("“4. Fees. Customer shall pay each invoice within forty-five (45) days.”\n2. Payment Terms.")
  // offset and limit count code points; the preamble's astral character is one.
  expect(resolved?.read).toEqual({ document: "Amendment No. 1", offset: opStart - 1, limit: subStart - opStart })
  db.close()
})

test("resolveCurrentSection runs the following text to the end of the indexed text when no later op bounds it", () => {
  const { db } = syntheticDb()
  const { subStart, subStatement, full } = collidingAmendment(db)
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "4.2(a)")
  expect(resolved?.followingText).toBe(full.slice(subStart + subStatement.length).trim())
  expect(resolved?.read).toEqual({ document: "Amendment No. 1", offset: subStart - 1, limit: full.length - subStart })
  db.close()
})

test("resolveCurrentSection caps the following text at 3000 characters", () => {
  const { db } = syntheticDb()
  const statement = "Section 6 of the Agreement is hereby amended and restated in its entirety to read as follows:"
  const full = `${statement}\n${"Restated body paragraph. ".repeat(200)}`
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [{ docPath: "/m/am1.docx", section: "1", text: full, charStart: 0, charEnd: full.length }],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "6", operation: "restate", char_start: 0, char_end: statement.length },
    ],
  )
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "6")
  expect(resolved?.followingText).toBe(full.slice(statement.length, statement.length + 3000).trim())
  expect(resolved?.read).toEqual({ document: "Amendment No. 1", offset: 0, limit: statement.length + 3000 })
  db.close()
})

test("no code-point position is given when indexed text has a gap before the op", () => {
  // Unstored regions are padded one space per UTF-16 unit, so a code-point
  // count across them can be wrong; the position is withheld instead.
  const { db } = syntheticDb()
  const statement = "Section 5 of the Agreement is hereby deleted in its entirety."
  const tail = `2. Deletion. ${statement}\n`
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [
      { docPath: "/m/am1.docx", section: "", text: "AMENDMENT NO. 1\n", charStart: 0, charEnd: 16 },
      // chars 16..40 are not stored
      { docPath: "/m/am1.docx", section: "2", text: tail, charStart: 40, charEnd: 40 + tail.length },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 15 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "5", operation: "delete", char_start: 53, char_end: 53 + statement.length },
    ],
  )
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "5")
  expect(resolved?.operativeText).toBe(statement)
  expect(resolved?.read).toEqual({ document: "Amendment No. 1", offset: null, limit: null })
  db.close()
})

test("operativePassage rebuilds each amending document once per shared cache", () => {
  const { db } = syntheticDb()
  collidingAmendment(db)
  const [chain] = buildChains(db)
  const ops = chain!.entries[0]!.ops
  let chunkReads = 0
  // operativePassage reads only through db.query; count the chunk reads.
  const counting = {
    query: (sql: string) => {
      if (sql.includes("FROM chunks")) chunkReads++
      return db.query(sql)
    },
  } as unknown as Database
  const cache = new Map()
  const passages = ops.map((op: SectionOpRow) => operativePassage(counting, op, ops, cache))
  expect(passages.map((p) => p.operativeText)).toEqual(ops.map((op) => verbatimSpan(db, op.doc_path, op.char_start, op.char_end)))
  expect(chunkReads).toBe(1)
  db.close()
})

test("verbatimSpan rebuilds overlapping chunks at their offsets and refuses an unanchored span", () => {
  const { db } = syntheticDb()
  const statement = "Section 9 of the Agreement is hereby amended and restated in its entirety to read as follows:"
  const full =
    "AMENDMENT NO. 2 TO SUPPLY AGREEMENT\n" +
    "The parties wish to amend the Agreement as set out below.\n" +
    `1. Term. ${statement}\n“9. Term. This Agreement continues until terminated by either party.”\n`
  const opStart = full.indexOf(statement)
  // Structure-aware chunking carries a trailing overlap: the second chunk
  // starts inside the first one.
  const firstEnd = opStart + 20
  const secondStart = opStart - 30
  seed(
    db,
    [
      { path: "/m/supply.docx", name: "Supply Agreement", createdAt: 1 },
      { path: "/m/am2.docx", name: "Amendment No. 2", createdAt: 2 },
    ],
    [
      { docPath: "/m/am2.docx", section: "1", text: full.slice(0, firstEnd), charStart: 0, charEnd: firstEnd },
      { docPath: "/m/am2.docx", section: "1", text: full.slice(secondStart), charStart: secondStart, charEnd: full.length },
    ],
    [
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", relation: "amends", target_name: "Supply Agreement", char_start: 0, char_end: 35 },
    ],
    [
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", section_label: "9", operation: "restate", char_start: opStart, char_end: opStart + statement.length },
    ],
  )
  expect(verbatimSpan(db, "/m/am2.docx", opStart, opStart + statement.length)).toBe(statement)
  // A span running past the indexed text cannot be anchored.
  expect(verbatimSpan(db, "/m/am2.docx", opStart, full.length + 50)).toBeNull()
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "9")
  expect(resolved?.operativeText).toBe(statement)
  expect(resolved?.followingText).toBe("“9. Term. This Agreement continues until terminated by either party.”")
  db.close()
})

test("ops are taken in document order even when stored grouped by operation", () => {
  // Ingest inserts section_ops grouped by verb (restate rows before delete
  // rows), so storage order is not document order. Here §4 is deleted early
  // and restated later in the same amender: the later restate is operative.
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "restate", char_start: 300, char_end: 400 },
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "delete", char_start: 100, char_end: 200 },
    ],
  )
  const chains = buildChains(db)
  expect(resolveCurrentSection(db, chains[0]!, "4")?.operation).toBe("restate")
  expect(supersededSections(db, chains).get("/m/base.docx")?.get("4")?.op.operation).toBe("restate")
  db.close()
})

test("resolveCurrentSection reports no operative wording when the op span is not recoverable", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    // The amender's only chunk labelled "4" lies outside the op span.
    [{ docPath: "/m/am1.docx", section: "4", text: "UNRELATED AMENDER CLAUSE", charStart: 500, charEnd: 524 }],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "amend", char_start: 100, char_end: 200 },
    ],
  )
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "4")
  expect(JSON.stringify(resolved)).not.toContain("UNRELATED AMENDER CLAUSE")
  expect(resolved?.operation).toBe("amend")
  expect(resolved?.operativeText).toBeNull()
  db.close()
})

test("supersededSections and resolveCurrentSection both pick last op within one amender", () => {
  // One amender with two ops on §4: amend (earlier in document) then delete
  // (later in document). The delete supersedes the amend — both surfaces must
  // agree. Before the fix, supersededSections kept the first (amend) while
  // resolveCurrentSection kept the last (delete), so search-document would
  // warn "(amend, chars 100..200)" while current-section reported §4 deleted.
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/base.docx", name: "Original Agreement", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "amend", char_start: 100, char_end: 200 },
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "delete", char_start: 300, char_end: 400 },
    ],
  )
  const chains = buildChains(db)
  const map = supersededSections(db, chains)
  expect(map.get("/m/base.docx")?.get("4")?.op.operation).toBe("delete")
  expect(map.get("/m/base.docx")?.get("4")?.op.char_start).toBe(300)

  const [chain] = chains
  const resolved = resolveCurrentSection(db, chain!, "4")
  expect(resolved?.operation).toBe("delete")
  expect(resolved?.amender.rel.doc_name).toBe("Amendment No. 1")
  db.close()
})

// Base resolution: a stated target resolves only on an exact match — a recorded
// defined term, or a file name after structural normalisation (case, extension,
// whitespace/underscore/hyphen runs). A near miss stays unmatched; two exact
// matches stay ambiguous. Nothing is picked by partial name overlap.
function resolutionFor(docs: DocSeed[], targetName: string, terms: { docPath: string; term: string }[] = []) {
  const { db } = syntheticDb()
  seed(
    db,
    [...docs, { path: "/m/amender.docx", name: "Amendment No. 1.docx", createdAt: 99 }],
    [],
    [
      { doc_path: "/m/amender.docx", doc_name: "Amendment No. 1.docx", relation: "amends", target_name: targetName, char_start: 0, char_end: 10 },
    ],
    [],
  )
  for (const t of terms) db.run("INSERT INTO defined_terms (doc_path, term) VALUES (?, ?)", [t.docPath, t.term])
  const chains = buildChains(db)
  db.close()
  expect(chains).toHaveLength(1)
  return { chain: chains[0]!, resolution: chains[0]!.entries[0]!.resolution }
}

test("resolveTarget does not match a target that is only a substring of a file name", () => {
  const { chain, resolution } = resolutionFor([{ path: "/m/release.docx", name: "Release.docx", createdAt: 1 }], "Lease")
  expect(resolution).toEqual({ kind: "unmatched" })
  expect(chain.base).toEqual({ kind: "unresolved", targetName: "Lease" })
})

test("resolveTarget does not choose a lone file whose name merely contains the target", () => {
  const { chain, resolution } = resolutionFor(
    [{ path: "/m/master.docx", name: "Master Agreement.docx", createdAt: 1 }],
    "Agreement",
  )
  expect(resolution).toEqual({ kind: "unmatched" })
  expect(chain.base.kind).toBe("unresolved")
})

test("resolveTarget does not match on a numbered title that only shares a prefix", () => {
  const { resolution } = resolutionFor([{ path: "/m/am10.docx", name: "Amendment No. 10.docx", createdAt: 1 }], "Amendment No. 1")
  expect(resolution).toEqual({ kind: "unmatched" })
})

test("resolveTarget resolves an exact file-name match and ignores a near miss beside it", () => {
  const { chain, resolution } = resolutionFor(
    [
      { path: "/m/lease.docx", name: "Lease.docx", createdAt: 1 },
      { path: "/m/release.docx", name: "Release.docx", createdAt: 2 },
    ],
    "Lease",
  )
  expect(resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease.docx" } })
  expect(chain.base).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease.docx" } })
})

test("resolveTarget folds only case, extension and separators in file names", () => {
  const { resolution } = resolutionFor(
    [{ path: "/m/mla.docx", name: "master_lease-agreement.DOCX", createdAt: 1 }],
    "Master Lease  Agreement",
  )
  expect(resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/mla.docx" } })
})

test("resolveTarget resolves a target that states the full file name, but not another extension", () => {
  const docs = [{ path: "/m/source.docx", name: "Source.docx", createdAt: 1 }]
  expect(resolutionFor(docs, "Source.docx").resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/source.docx" } })
  expect(resolutionFor(docs, "Source.pdf").resolution).toEqual({ kind: "unmatched" })
})

test("resolveTarget resolves an exact recorded defined term", () => {
  const { resolution } = resolutionFor(
    [
      { path: "/m/premises.docx", name: "Premises.docx", createdAt: 1 },
      { path: "/m/other.docx", name: "Other.docx", createdAt: 2 },
    ],
    "Head Lease",
    [
      { docPath: "/m/premises.docx", term: "head lease" },
      { docPath: "/m/other.docx", term: "Head Lease Guarantee" },
    ],
  )
  expect(resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/premises.docx" } })
})

test("resolveTarget reports two exact matches as ambiguous", () => {
  const { chain, resolution } = resolutionFor(
    [
      { path: "/m/lease.docx", name: "Lease.docx", createdAt: 1 },
      { path: "/m/lease.pdf", name: "Lease.pdf", createdAt: 2 },
    ],
    "Lease",
  )
  expect(resolution.kind).toBe("ambiguous")
  if (resolution.kind === "ambiguous") expect(resolution.names.sort()).toEqual(["Lease.docx", "Lease.pdf"])
  expect(chain.base.kind).toBe("unresolved")
})

test("supersededSections does not down-rank a document the target only partially names", () => {
  const { db } = syntheticDb()
  seed(
    db,
    [
      { path: "/m/release.docx", name: "Release.docx", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", createdAt: 2 },
    ],
    [],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1.docx", relation: "amends", target_name: "Lease", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1.docx", section_label: "4", operation: "restate", char_start: 100, char_end: 200 },
    ],
  )
  const map = supersededSections(db, buildChains(db))
  expect(map.has("/m/release.docx")).toBe(false)
  db.close()
})

test("resolveTarget strips only a document's real extension, never a dotted word in its name", () => {
  // The indexed path records the extension; a name without it keeps its dots.
  expect(resolutionFor([{ path: "/m/lease.docx", name: "Lease.final", createdAt: 1 }], "Lease").resolution).toEqual({ kind: "unmatched" })
  // No recorded extension: only a supported document extension is stripped.
  expect(resolutionFor([{ path: "/m/lease-final", name: "Lease.final", createdAt: 1 }], "Lease").resolution).toEqual({ kind: "unmatched" })
  expect(resolutionFor([{ path: "/m/Lease.final", name: "Lease.final", createdAt: 1 }], "Lease").resolution).toEqual({ kind: "unmatched" })
  expect(resolutionFor([{ path: "/m/lease", name: "Lease.docx", createdAt: 1 }], "Lease").resolution).toMatchObject({
    kind: "resolved",
    doc: { doc_path: "/m/lease" },
  })
  // A target stating a document extension names the whole file.
  const converted = [{ path: "/m/lease.pdf.docx", name: "Lease.pdf.docx", createdAt: 1 }]
  expect(resolutionFor(converted, "Lease.pdf").resolution).toEqual({ kind: "unmatched" })
  expect(resolutionFor(converted, "Lease.pdf.docx").resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease.pdf.docx" } })
})

test("resolveTarget keeps dots inside names: a version or suffix must be stated in full", () => {
  const final = [{ path: "/m/lease.final.docx", name: "Lease.final.docx", createdAt: 1 }]
  expect(resolutionFor(final, "Lease.final").resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease.final.docx" } })
  expect(resolutionFor(final, "Lease").resolution).toEqual({ kind: "unmatched" })
  const versioned = [{ path: "/m/lease-v12.docx", name: "Lease v1.2.docx", createdAt: 1 }]
  expect(resolutionFor(versioned, "Lease v1.2").resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease-v12.docx" } })
  expect(resolutionFor(versioned, "Lease v1").resolution).toEqual({ kind: "unmatched" })
})

test("resolveTarget folds surrounding and non-breaking spaces but not unicode dashes", () => {
  const lease = [{ path: "/m/lease.docx", name: "Lease Agreement.docx", createdAt: 1 }]
  expect(resolutionFor(lease, "Lease Agreement ").resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease.docx" } })
  expect(resolutionFor(lease, "Lease\u00a0Agreement").resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease.docx" } })
  // A different dash is a different character: unmatched, never guessed.
  const dashed = [{ path: "/m/lease-a.docx", name: "Lease \u2013 A.docx", createdAt: 1 }]
  expect(resolutionFor(dashed, "Lease - A").resolution).toEqual({ kind: "unmatched" })
})

test("resolveTarget trims the target before the defined-term lookup", () => {
  const { resolution } = resolutionFor(
    [{ path: "/m/premises.docx", name: "Premises.docx", createdAt: 1 }],
    "Head Lease ",
    [{ docPath: "/m/premises.docx", term: "Head Lease" }],
  )
  expect(resolution).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/premises.docx" } })
})

// File-name and defined-term matches are not ranked against each other. The
// non-amending preference narrows them, but never by dropping every file-name
// match in favour of a defined term: that is reported as ambiguous.
function chainsFor(docs: DocSeed[], relations: { path: string; name: string; target: string }[], terms: { docPath: string; term: string }[]) {
  const { db } = syntheticDb()
  seed(
    db,
    docs,
    [],
    relations.map((r) => ({ doc_path: r.path, doc_name: r.name, relation: "amends", target_name: r.target, char_start: 0, char_end: 10 })),
    [],
  )
  for (const t of terms) db.run("INSERT INTO defined_terms (doc_path, term) VALUES (?, ?)", [t.docPath, t.term])
  const chains = buildChains(db)
  db.close()
  return chains
}

test("resolveTarget reports an amending exact-name match beside a non-amending defined-term match as ambiguous", () => {
  const chains = chainsFor(
    [
      { path: "/m/original.docx", name: "Original.docx", createdAt: 1 },
      { path: "/m/agreement.docx", name: "Agreement.docx", createdAt: 2 },
      { path: "/m/side.docx", name: "Side.docx", createdAt: 3 },
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", createdAt: 4 },
    ],
    [
      { path: "/m/agreement.docx", name: "Agreement.docx", target: "Original" },
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", target: "Agreement" },
    ],
    [{ docPath: "/m/side.docx", term: "Agreement" }],
  )
  const entry = chains.flatMap((c) => c.entries).find((e) => e.rel.doc_path === "/m/am1.docx")!
  expect(entry.resolution.kind).toBe("ambiguous")
  if (entry.resolution.kind === "ambiguous") expect(entry.resolution.names.sort()).toEqual(["Agreement.docx", "Side.docx"])
})

test("resolveTarget resolves a non-amending exact-name match over an amender that defines the same term", () => {
  // Amendments routinely define the base's name as a term; both preferences
  // point at the base, so it resolves.
  const chains = chainsFor(
    [
      { path: "/m/lease.docx", name: "Lease.docx", createdAt: 1 },
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", createdAt: 2 },
      { path: "/m/am2.docx", name: "Amendment No. 2.docx", createdAt: 3 },
    ],
    [
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", target: "Lease" },
      { path: "/m/am2.docx", name: "Amendment No. 2.docx", target: "Lease" },
    ],
    [
      { docPath: "/m/am1.docx", term: "Lease" },
      { docPath: "/m/am2.docx", term: "Lease" },
    ],
  )
  expect(chains).toHaveLength(1)
  expect(chains[0]!.base).toMatchObject({ kind: "resolved", doc: { doc_path: "/m/lease.docx" } })
  expect(chains[0]!.entries.map((e) => e.resolution.kind)).toEqual(["resolved", "resolved"])
})

test("buildChains groups unmatched targets by their trimmed stated name", () => {
  const chains = chainsFor(
    [
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", createdAt: 1 },
      { path: "/m/am2.docx", name: "Amendment No. 2.docx", createdAt: 2 },
    ],
    [
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", target: "Lease" },
      { path: "/m/am2.docx", name: "Amendment No. 2.docx", target: "Lease " },
    ],
    [],
  )
  expect(chains).toHaveLength(1)
  expect(chains[0]!.base).toEqual({ kind: "unresolved", targetName: "Lease" })
  expect(chains[0]!.entries.map((e) => e.rel.doc_name)).toEqual(["Amendment No. 1.docx", "Amendment No. 2.docx"])
})

test("resolveTarget reports a non-amending exact-name match beside a non-amending definer as ambiguous", () => {
  const chains = chainsFor(
    [
      { path: "/m/lease.docx", name: "Lease.docx", createdAt: 1 },
      { path: "/m/guarantee.docx", name: "Guarantee.docx", createdAt: 2 },
      { path: "/m/am1.docx", name: "Amendment No. 1.docx", createdAt: 3 },
    ],
    [{ path: "/m/am1.docx", name: "Amendment No. 1.docx", target: "Lease" }],
    [{ docPath: "/m/guarantee.docx", term: "Lease" }],
  )
  const resolution = chains[0]!.entries[0]!.resolution
  expect(resolution.kind).toBe("ambiguous")
  if (resolution.kind === "ambiguous") expect(resolution.names.sort()).toEqual(["Guarantee.docx", "Lease.docx"])
})
