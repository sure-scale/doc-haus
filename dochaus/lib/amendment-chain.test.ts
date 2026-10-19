// Unit tests for the shared amendment-chain helper on a synthetic legal.db.
// Run from dochaus/: bun test lib/amendment-chain.test.ts
import { Database } from "bun:sqlite"
import { expect, test } from "bun:test"
import { mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { buildChains, resolveCurrentSection, supersededSections, verbatimSpan } from "./amendment-chain"

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
