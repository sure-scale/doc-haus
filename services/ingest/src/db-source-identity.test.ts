import { expect, test } from "bun:test"
import type { Database } from "bun:sqlite"
import { mkdtempSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  deleteDocument,
  ftsInSync,
  getDocumentRole,
  getPreparedChanges,
  insertChunk,
  listPendingRedlines,
  openDb,
  setPreparedChanges,
  upsertDeliverable,
  upsertDocument,
} from "./db"
import { recordSections } from "./structure"

const indexTables = ["chunks", "defined_terms", "section_refs", "parties", "doc_relations", "section_ops", "sections"]
const count = (db: Database, table: string) => (db.query(`SELECT COUNT(*) AS n FROM ${table}`).get() as { n: number }).n
const matches = (db: Database, token: string) =>
  db
    .query("SELECT c.document_id FROM chunks_fts JOIN chunks c ON c.id = chunks_fts.rowid WHERE chunks_fts MATCH ?")
    .all(token)

function withRegistry(run: (db: Database, directory: string, docPath: string) => void) {
  const directory = realpathSync.native(mkdtempSync(path.join(tmpdir(), "dochaus-source-identity-")))
  const docPath = path.join(directory, "Record.md")
  writeFileSync(docPath, "Invented source record.")
  const db = openDb(directory)
  try {
    run(db, directory, docPath)
  } finally {
    db.close()
    rmSync(directory, { recursive: true, force: true })
  }
}

// Explicit local index rows isolate registry identity from extractors/models.
// These are source/index lifecycle tests, not semantic embedding tests.
function indexed(db: Database, id: number, docPath: string, marker: string) {
  insertChunk(db, {
    documentId: id,
    docPath,
    docName: "Record.md",
    section: "1",
    structuralPath: "Record > 1",
    party: "",
    chunkIndex: 0,
    text: `${marker}\n`,
    charStart: 0,
    charEnd: marker.length + 1,
    embedding: new Float32Array([1, 0, 0]),
    flagged: false,
  })
  recordSections(db, docPath, "Record.md")
}
function oldStructure(db: Database, docPath: string) {
  db.run(
    "INSERT INTO defined_terms(doc_path,doc_name,term,definition,char_start,char_end) VALUES(?,?,'Container','Invented container',0,10)",
    [docPath, "Record.md"],
  )
  db.run(
    "INSERT INTO section_refs(doc_path,doc_name,ref_kind,ref_label,char_start,char_end) VALUES(?,?,'section','1',0,10)",
    [docPath, "Record.md"],
  )
  db.run(
    "INSERT INTO parties(doc_path,doc_name,name,role,char_start,char_end) VALUES(?,?,'Invented Logistics LLP','provider',0,10)",
    [docPath, "Record.md"],
  )
  db.run(
    "INSERT INTO doc_relations(doc_path,doc_name,relation,target_name,char_start,char_end) VALUES(?,?,'amends','Earlier invented terms',0,10)",
    [docPath, "Record.md"],
  )
  db.run(
    "INSERT INTO section_ops(doc_path,doc_name,section_label,operation,char_start,char_end) VALUES(?,?,'1','restate',0,10)",
    [docPath, "Record.md"],
  )
}
function redline(db: Database, docPath: string) {
  return Number(
    db.run(
      "INSERT INTO redlines(doc_path,doc_name,scope,find_text,old_text,new_text,author,status,created_at) VALUES(?,?,'phrase','old','old','new','Synthetic reviewer','pending',1)",
      [docPath, "Record.md"],
    ).lastInsertRowid,
  )
}

test("replacing one physical source keeps its id and lawyer metadata while rebuilding every old index row", () =>
  withRegistry((db, _directory, docPath) => {
    const original = upsertDocument(db, docPath, "Record.md", 1, null, {
      parentName: "Enclosing email.eml",
      attachments: "[]",
    })
    indexed(db, original, docPath, "ORIGINALSOURCE947")
    oldStructure(db, docPath)
    db.run("UPDATE documents SET label='draft' WHERE id=?", [original])
    const pending = redline(db, docPath)
    setPreparedChanges(db, docPath, { redline_ids: [pending], open_items: ["Review the original wording."] })
    expect(indexTables.every((table) => count(db, table) > 0)).toBe(true)
    expect(matches(db, "ORIGINALSOURCE947")).toEqual([{ document_id: original }])
    db.transaction(() => {
      const replaced = upsertDocument(
        db,
        docPath,
        "Record.md",
        2,
        '{"findings":[]}',
        {},
        '{"status":"ok","signals":[]}',
      )
      expect(replaced).toBe(original)
      for (const table of indexTables) expect(count(db, table)).toBe(0)
      expect(db.query("SELECT rev FROM documents WHERE id=?").get(original)).toEqual({ rev: null })
      indexed(db, replaced, docPath, "REPLACEMENTSOURCE284")
    })()
    expect(
      db
        .query(
          "SELECT id,role,label,parent_name,attachments,created_at,injection_report,extraction_quality FROM documents",
        )
        .get(),
    ).toEqual({
      id: original,
      role: "source",
      label: "draft",
      parent_name: "Enclosing email.eml",
      attachments: "[]",
      created_at: 2,
      injection_report: '{"findings":[]}',
      extraction_quality: '{"status":"ok","signals":[]}',
    })
    expect(listPendingRedlines(db, docPath).map((row) => row.id)).toEqual([pending])
    expect(getPreparedChanges(db, docPath)).toEqual({
      redline_ids: [pending],
      open_items: ["Review the original wording."],
    })
    expect(matches(db, "ORIGINALSOURCE947")).toEqual([])
    expect(matches(db, "REPLACEMENTSOURCE284")).toEqual([{ document_id: original }])
    expect(ftsInSync(db)).toBe(true)
  }))

test("an explicit deletion followed by a same-name upload creates a new source identity", () =>
  withRegistry((db, _directory, docPath) => {
    const content = "SAMECONTENTSOURCE613\n"
    writeFileSync(docPath, content)
    const original = upsertDocument(db, docPath, "Record.md", 1, null)
    indexed(db, original, docPath, "SAMECONTENTSOURCE613")
    const originalChunk = (db.query("SELECT id FROM chunks WHERE document_id=?").get(original) as { id: number }).id
    const pending = redline(db, docPath)
    setPreparedChanges(db, docPath, { redline_ids: [pending], open_items: ["Earlier source review."] })
    deleteDocument(db, docPath)
    expect(matches(db, "SAMECONTENTSOURCE613")).toEqual([])
    rmSync(docPath)
    writeFileSync(docPath, content)
    expect(readFileSync(docPath, "utf8")).toBe(content)
    expect(getDocumentRole(db, docPath)).toBeNull()
    const replacement = upsertDocument(db, docPath, "Record.md", 2, null)
    indexed(db, replacement, docPath, "SAMECONTENTSOURCE613")
    expect(replacement).toBeGreaterThan(original)
    expect(count(db, "documents")).toBe(1)
    expect(count(db, "redlines")).toBe(0)
    expect(getPreparedChanges(db, docPath)).toBeNull()
    expect(db.query("SELECT id FROM chunks WHERE id=?").get(originalChunk)).toBeNull()
    expect(matches(db, "SAMECONTENTSOURCE613")).toEqual([{ document_id: replacement }])
  }))

test("legacy physical aliases collapse onto the existing source owner and cannot cross the deliverable boundary", () =>
  withRegistry((db, directory, docPath) => {
    const aliasFolder = path.join(directory, "legacy-folder")
    symlinkSync(directory, aliasFolder, "dir")
    const aliasPath = path.join(aliasFolder, "Record.md")
    // The old registry can contain aliases; current upload path validation still
    // rejects user-selected symlinks before this lower-level registry helper.
    const owner = Number(
      db.run(
        "INSERT INTO documents(doc_path,name,created_at,role,label,parent_name,attachments) VALUES(?,'Earlier name.md',1,'source','notes','Enclosing email.eml','[]')",
        [aliasPath],
      ).lastInsertRowid,
    )
    indexed(db, owner, aliasPath, "LEGACYSOURCE368")
    const pending = redline(db, aliasPath)
    expect(upsertDocument(db, docPath, "Record.md", 2, null)).toBe(owner)
    expect(db.query("SELECT id,doc_path,name,label,parent_name,attachments FROM documents").get()).toEqual({
      id: owner,
      doc_path: docPath,
      name: "Record.md",
      label: "notes",
      parent_name: "Enclosing email.eml",
      attachments: "[]",
    })
    const alias = Number(
      db.run("INSERT INTO documents(doc_path,name,created_at,role) VALUES(?,'Historical alias.md',3,'source')", [
        aliasPath,
      ]).lastInsertRowid,
    )
    indexed(db, alias, aliasPath, "DUPLICATEALIAS759")
    expect(upsertDocument(db, docPath, "Record.md", 4, null)).toBe(owner)
    expect(count(db, "documents")).toBe(1)
    expect(count(db, "chunks")).toBe(0)
    expect(upsertDocument(db, docPath, "Record.md", 5, null)).toBe(owner)
    expect(count(db, "documents")).toBe(1)
    expect(listPendingRedlines(db, docPath)).toMatchObject([{ id: pending, doc_path: docPath, doc_name: "Record.md" }])
    expect(() => upsertDeliverable(db, aliasPath, "Record.md", 5)).toThrow("Cannot replace source")
    db.run("UPDATE documents SET role='deliverable' WHERE id=?", [owner])
    expect(() => upsertDocument(db, aliasPath, "Record.md", 6, null)).toThrow("Cannot replace generated deliverable")
    expect(getDocumentRole(db, docPath)).toBe("deliverable")
  }))

test("a failed surrounding source commit restores the source owner, metadata and old FTS rows", () =>
  withRegistry((db, _directory, docPath) => {
    const original = upsertDocument(db, docPath, "Record.md", 1, null, { parentName: "Enclosing email.eml" })
    indexed(db, original, docPath, "KEPTSOURCE492")
    const pending = redline(db, docPath)
    const before = db.query("SELECT * FROM documents").all()
    expect(() =>
      db.transaction(() => {
        upsertDocument(db, docPath, "Record.md", 2, '{"findings":[]}', { parentName: null })
        indexed(db, original, docPath, "ROLLEDBACKSOURCE571")
        throw new Error("Synthetic projection installation failure")
      })(),
    ).toThrow("Synthetic projection installation failure")
    expect(db.query("SELECT * FROM documents").all()).toEqual(before)
    expect(listPendingRedlines(db, docPath).map((row) => row.id)).toEqual([pending])
    expect(matches(db, "KEPTSOURCE492")).toEqual([{ document_id: original }])
    expect(matches(db, "ROLLEDBACKSOURCE571")).toEqual([])
    expect(ftsInSync(db)).toBe(true)
  }))
