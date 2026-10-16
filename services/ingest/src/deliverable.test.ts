import { Database } from "bun:sqlite"
import { describe, expect, test } from "bun:test"
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { getDocumentRole, insertChunk, listDocuments, listSourceDocuments, openDb, upsertDocument } from "./db"
import { DeliverableConflictError, writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { writeSourceArtifacts } from "./source-artifacts"
import { sha256File } from "../../../dochaus/lib/lineage"
import { Document, Packer, Paragraph, Table, TableCell, TableRow } from "docx"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"

async function generated(text: string) {
  return Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph(text)] }] }))))
}

describe("deliverable persistence boundary", () => {
  test("registers generated files for listing without changing source artifacts or retrieval rows", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-deliverable-"))
    try {
      const sourcePath = path.join(dir, "Source.docx")
      const sourceText = "SOURCE MATERIAL\nFacts supplied by the user."
      const db = openDb(dir)
      const documentId = upsertDocument(db, sourcePath, "Source.docx", Date.now(), null)
      insertChunk(db, {
        documentId,
        docPath: sourcePath,
        docName: "Source.docx",
        section: "Preamble",
        structuralPath: "Source.docx > Preamble",
        party: "",
        chunkIndex: 0,
        text: sourceText,
        charStart: 0,
        charEnd: sourceText.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()
      await writeSourceArtifacts(dir)
      const sourceIndexBefore = await Bun.file(path.join(dir, "source-index.json")).text()

      const saved = await writeDeliverable(dir, "Result.docx", await generated("Generated wording."))
      expect(saved.role).toBe("deliverable")
      expect(existsSync(path.join(dir, "Result.docx"))).toBe(true)

      const after = openDb(dir)
      expect(listDocuments(after)).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ name: "Source.docx", role: "source" }),
          expect.objectContaining({ name: "Result.docx", role: "deliverable" }),
        ]),
      )
      expect(listSourceDocuments(after)).toEqual([expect.objectContaining({ name: "Source.docx", role: "source" })])
      expect(
        (after.query("SELECT COUNT(*) AS count FROM chunks WHERE doc_path = ?").get(path.join(dir, "Result.docx")) as {
          count: number
        }).count,
      ).toBe(0)
      after.close()

      expect(await Bun.file(path.join(dir, "source-index.json")).text()).toBe(sourceIndexBefore)
      const redlineDb = openDb(dir)
      redlineDb.run(
        "INSERT INTO redlines (doc_path, doc_name, scope, find_text, old_text, new_text, author, status, created_at) VALUES (?, ?, 'phrase', 'old', 'old', 'new', 'reviewer', 'pending', ?)",
        [path.join(dir, "Result.docx"), "Result.docx", Date.now()],
      )
      redlineDb.close()
      const revisedBytes = await generated("Revised generated wording.")
      await writeDeliverableRevision(dir, "Result.docx", revisedBytes)
      expect(await Bun.file(path.join(dir, "Result.docx")).bytes()).toEqual(new Uint8Array(revisedBytes))
      expect(await Bun.file(path.join(dir, "source-index.json")).text()).toBe(sourceIndexBefore)
      const revised = openDb(dir)
      expect((revised.query("SELECT COUNT(*) AS count FROM redlines WHERE status = 'pending'").get() as { count: number }).count).toBe(1)
      revised.close()
      await writeSourceArtifacts(dir)
      const rebuilt = await Bun.file(path.join(dir, "source-index.json")).json()
      expect(rebuilt.sources.map((source: { document: string }) => source.document)).toEqual(["Source.docx"])
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("enforces create-only and exact-prior replacement inside the deliverable lock", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-deliverable-precondition-"))
    try {
      const initial = await generated("Initial generated wording.")
      const replacement = await generated("Authorized replacement.")
      await writeDeliverable(
        dir,
        "Result.docx",
        initial,
        { precondition: { kind: "create" } },
      )
      await expect(writeDeliverable(
        dir,
        "Result.docx",
        Buffer.from("foreign overwrite"),
        { precondition: { kind: "create" } },
      )).rejects.toThrow("target already exists")
      expect(await Bun.file(path.join(dir, "Result.docx")).bytes()).toEqual(new Uint8Array(initial))
      const prior = sha256File(path.join(dir, "Result.docx"))!
      await expect(writeDeliverable(
        dir,
        "Result.docx",
        Buffer.from("wrong-prior overwrite"),
        { precondition: { kind: "replace", prior_sha256: "0".repeat(64) } },
      )).rejects.toThrow("prior identity changed")
      await writeDeliverable(
        dir,
        "Result.docx",
        replacement,
        { precondition: { kind: "replace", prior_sha256: prior.sha256 } },
      )
      expect(await Bun.file(path.join(dir, "Result.docx")).bytes()).toEqual(new Uint8Array(replacement))
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("publication points a previously unpointed generated package including cells and repeated delivery keeps ids", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-deliverable-pointers-"))
    try {
      const input = await Packer.toBuffer(new Document({ sections: [{ children: [
        new Paragraph("Generated cover note."),
        new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph("Topic")] }), new TableCell({ children: [new Paragraph("Treatment")] })] })] }),
      ] }] }))
      const original = new Uint8Array(input)
      expect(readDocxCommentParagraphs(input).every(row => row.pids.length === 0)).toBe(true)
      await writeDeliverable(dir, "First.docx", input)
      const published = await Bun.file(path.join(dir, "First.docx")).bytes()
      const rows = readDocxCommentParagraphs(published)
      expect(rows.map(row => row.paragraph_text)).toEqual(["Generated cover note.", "Topic", "Treatment"])
      expect(rows.every(row => row.pids.length === 1)).toBe(true)
      expect(new Set(rows.flatMap(row => row.pids)).size).toBe(3)
      expect(new Uint8Array(input)).toEqual(original)
      await writeDeliverable(dir, "Copy.docx", Buffer.from(published))
      expect(await Bun.file(path.join(dir, "Copy.docx")).bytes()).toEqual(published)
      const prior = sha256File(path.join(dir, "First.docx"))!
      await writeDeliverable(dir, "First.docx", Buffer.from(published), { precondition: { kind: "replace", prior_sha256: prior.sha256 } })
      expect(await Bun.file(path.join(dir, "First.docx")).bytes()).toEqual(published)
    } finally { rmSync(dir, { recursive: true, force: true }) }
  })

  test("migrates legacy document rows as sources", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-deliverable-migration-"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      const legacy = new Database(path.join(dir, ".dochaus", "legal.db"))
      legacy.run(`
        CREATE TABLE documents (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          doc_path TEXT NOT NULL UNIQUE,
          name TEXT NOT NULL,
          created_at INTEGER NOT NULL,
          injection_report TEXT
        )
      `)
      const docPath = path.join(dir, "Legacy.docx")
      legacy.run("INSERT INTO documents (doc_path, name, created_at, injection_report) VALUES (?, ?, ?, NULL)", [
        docPath,
        "Legacy.docx",
        Date.now(),
      ])
      legacy.close()

      const migrated = openDb(dir)
      expect(getDocumentRole(migrated, docPath)).toBe("source")
      expect(listSourceDocuments(migrated)).toEqual([expect.objectContaining({ name: "Legacy.docx", role: "source" })])
      migrated.close()
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("refuses to overwrite a registered source through the deliverable path", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-deliverable-conflict-"))
    try {
      const file = path.join(dir, "Source.docx")
      writeFileSync(file, "source bytes")
      const db = openDb(dir)
      upsertDocument(db, file, "Source.docx", Date.now(), null)
      db.close()

      expect(() => writeDeliverable(dir, "Source.docx", Buffer.from("generated bytes"))).toThrow(
        DeliverableConflictError,
      )
      expect(await Bun.file(file).text()).toBe("source bytes")
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
