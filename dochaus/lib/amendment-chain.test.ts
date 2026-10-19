// Unit tests for the shared amendment-chain helper on a synthetic legal.db.
// Run from dochaus/: bun test lib/amendment-chain.test.ts
import { Database } from "bun:sqlite"
import { expect, test } from "bun:test"
import { mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { buildChains, resolveCurrentSection, supersededSections } from "./amendment-chain"

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
      // am1's replacement text for §4 (superseded by am2)
      { docPath: "/m/am1.docx", section: "4", text: "OLD REPLACEMENT TEXT", charStart: 0, charEnd: 20 },
      // am2's replacement text for §4 (operative)
      { docPath: "/m/am2.docx", section: "4", text: "NEW REPLACEMENT TEXT", charStart: 0, charEnd: 20 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", relation: "amends", target_name: "Original Agreement", char_start: 0, char_end: 10 },
    ],
    [
      { doc_path: "/m/am1.docx", doc_name: "Amendment No. 1", section_label: "4", operation: "amend", char_start: 100, char_end: 200 },
      { doc_path: "/m/am2.docx", doc_name: "Amendment No. 2", section_label: "4", operation: "restate", char_start: 100, char_end: 200 },
    ],
  )
  const [chain] = buildChains(db)
  const resolved = resolveCurrentSection(db, chain!, "4")
  expect(resolved?.operation).toBe("restate")
  expect(resolved?.amender.rel.doc_name).toBe("Amendment No. 2")
  expect(resolved?.replacementText).toBe("NEW REPLACEMENT TEXT")
  db.close()
})

test("resolveCurrentSection returns delete with null replacementText", () => {
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
  expect(resolved?.replacementText).toBeNull()
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
