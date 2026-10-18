import { afterEach, expect, test } from "bun:test"
import { readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { recordLeanDocuments } from "../../../dochaus/workflows/lean/deliverables"
import { runLean } from "../../../dochaus/workflows/lean/runner"
import { leanFile } from "../../../dochaus/workflows/lean/state"
import recordedCompleted from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import { FIXTURE_SOURCE, FIXTURE_SOURCE_TEXT, leanFixtureMatter, scriptedLeanModel } from "../../../dochaus/workflows/lean/test-support"
import { listDocuments, openDb } from "./db"
import { leanDraftingStatus } from "./lean-drafting"
import { registerLeanDeliverable } from "./lean-deliverables"
import { Document, Packer, Paragraph, Table, TableCell, TableRow } from "docx"
import { writeLeanPlan } from "../../../dochaus/workflows/lean/plan"
import { writeLeanBytes } from "../../../dochaus/workflows/lean/state"
import { sha256File } from "../../../dochaus/lib/lineage"
import { readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { readLeanJson } from "../../../dochaus/workflows/lean/state"
import { writeDeliverableRevision } from "./deliverable"
import { acceptedWorkingDocxUpdate } from "./document-amendments"

// Real coordinator, real deliverable registry write and real matter database.
// Only the model is scripted.

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

function ingestMatter() {
  const directory = leanFixtureMatter("lean-deliverables-", dir => {
    const db = openDb(dir)
    const docPath = path.join(dir, FIXTURE_SOURCE)
    const id = db.query("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?) RETURNING id").get(docPath, FIXTURE_SOURCE, 1) as { id: number }
    db.query("INSERT INTO chunks (document_id, doc_path, doc_name, section, chunk_index, text, char_start, char_end, embedding) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)")
      .run(id.id, docPath, FIXTURE_SOURCE, "1", 0, FIXTURE_SOURCE_TEXT, 0, FIXTURE_SOURCE_TEXT.length, new Uint8Array(4))
    db.close()
  })
  directories.push(directory)
  return directory
}

function shape(value: unknown): unknown {
  if (Array.isArray(value)) return value.slice(0, 1).map(shape)
  if (!value || typeof value !== "object") return value === null ? null : typeof value
  return Object.fromEntries(Object.entries(value).toSorted(([a], [b]) => a.localeCompare(b)).map(([key, field]) => [key, shape(field)]))
}

function deliverables(directory: string) {
  const db = openDb(directory)
  const rows = (listDocuments(db) as { name: string; role: string }[]).filter(row => row.role === "deliverable").map(row => row.name)
  db.close()
  return rows.toSorted()
}

test("a completed lean run registers its Word files as matter deliverables and drafting status reports them", async () => {
  const directory = ingestMatter()
  const deps = { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }
  const state = await runLean(directory, deps)
  expect(state.status).toBe("completed")

  expect(deliverables(directory)).toEqual(["Supply note - drafting notes.docx", "Supply note - redline.docx", "Supply note.docx"])
  expect(readFileSync(path.join(directory, "Supply note.docx"))).toEqual(readFileSync(leanFile(directory, "d1/body-applied.docx")))
  const documents = leanDraftingStatus(directory)?.documents
  expect(documents?.status).toBe("completed")
  expect(documents?.documents.map(row => [row.kind, row.path, row.deliverable])).toEqual([
    ["body", ".dochaus/lean/d1/body-applied.docx", "Supply note.docx"],
    ["redline", ".dochaus/lean/d1/redline.docx", "Supply note - redline.docx"],
    ["drafting_notes", ".dochaus/lean/delivered-notes.docx", "Supply note - drafting notes.docx"],
    ["drafting_notes_markdown", ".dochaus/lean/delivered-notes.md", null],
  ])
  expect(leanDraftingStatus(directory)?.plan).toEqual({ documents: [{ id: "d1", title: "Supply note", purpose: "" }] })
  // Keep the older completed-run contract while allowing additive preparation
  // and review stages (this fixture starts without a model).
  expect(shape({ ...leanDraftingStatus(directory), model: recordedCompleted.model })).toMatchObject(shape(recordedCompleted) as Record<string, unknown>)

  // Registering again reuses names that still hold the run's bytes; a file the
  // lawyer replaced is kept and the run's copy takes the next free name.
  writeFileSync(path.join(directory, "Supply note.docx"), "replaced by the lawyer")
  const again = await recordLeanDocuments(directory, deps, true)
  expect(again.documents.map(row => row.deliverable)).toEqual(
    ["Supply note (2).docx", "Supply note - redline.docx", "Supply note - drafting notes.docx", null])
  expect(readFileSync(path.join(directory, "Supply note.docx"), "utf8")).toBe("replaced by the lawyer")
  expect(deliverables(directory)).toContain("Supply note (2).docx")
}, 120_000)

test("a source document's name is never taken over by a deliverable", async () => {
  const directory = ingestMatter()
  writeFileSync(path.join(directory, "Supply note.docx"), "a source")
  const db = openDb(directory)
  db.query("INSERT INTO documents (doc_path, name, created_at) VALUES (?, ?, ?)").run(path.join(directory, "Supply note.docx"), "Supply note.docx", 2)
  db.close()
  // The registry refuses the source's name; the run is told to try another.
  expect(await registerLeanDeliverable(directory, "Supply note.docx", new TextEncoder().encode("a source"), null)).toBeUndefined()
  expect(readFileSync(path.join(directory, "Supply note.docx"), "utf8")).toBe("a source")
})

test("legacy unpointed produced files retain one delivery name and a truthful hash after pointer publication", async () => {
  const directory = ingestMatter()
  writeFileSync(path.join(directory, FIXTURE_SOURCE), await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Imported original.")] }] })))
  const sourceBefore = readFileSync(path.join(directory, FIXTURE_SOURCE))
  writeLeanPlan(directory, [{ title: "Review note", purpose: "Generic record", parts: [{ heading: "Record", scope: "Generic record" }] }])
  const unpointed = await Packer.toBuffer(new Document({ sections: [{ children: [
    new Paragraph("Retain the supplied record."),
    new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph("Topic")] }), new TableCell({ children: [new Paragraph("Treatment")] })] })] }),
  ] }] }))
  expect(readDocxCommentParagraphs(unpointed).every(row => row.pids.length === 0)).toBe(true)
  for (const name of ["d1/body-applied.docx", "d1/redline.docx", "drafting-notes.docx"])
    writeLeanBytes(directory, name, unpointed)
  const deps = { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }
  const first = await recordLeanDocuments(directory, deps, true)
  for (const row of first.documents) {
    expect(row.sha256).toBe(sha256File(path.join(directory, row.path))!.sha256)
    expect(row.sha256).toBe(sha256File(path.join(directory, row.deliverable!))!.sha256)
    expect(readDocxCommentParagraphs(readFileSync(path.join(directory, row.path))).every(paragraph => paragraph.pids.length === 1)).toBe(true)
  }
  const published = readFileSync(path.join(directory, "Review note.docx"))
  const second = await recordLeanDocuments(directory, deps, true)
  expect(second).toEqual(first)
  expect(deliverables(directory)).toEqual(["Review note - drafting notes.docx", "Review note - redline.docx", "Review note.docx"])
  expect(readFileSync(path.join(directory, "Review note.docx"))).toEqual(published)
  writeFileSync(path.join(directory, "Review note.docx"), "A lawyer changed this file.")
  const third = await recordLeanDocuments(directory, deps, true)
  expect(third.documents.find(row => row.kind === "body")!.deliverable).toBe("Review note (2).docx")
  expect(readFileSync(path.join(directory, "Review note.docx"), "utf8")).toBe("A lawyer changed this file.")
  expect(readFileSync(path.join(directory, FIXTURE_SOURCE))).toEqual(sourceBefore)
}, 120_000)

test("a lawyer's accepted revision keeps its name; a later output is neither written over it nor forked to a new name", async () => {
  const directory = ingestMatter()
  const deps = { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }
  expect((await runLean(directory, deps)).status).toBe("completed")
  const file = path.join(directory, "Supply note.docx")
  const delivered = sha256File(file)!.sha256
  const revised = Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Revised by the lawyer.")] }] }))))
  await writeDeliverableRevision(directory, "Supply note.docx", revised, { priorSha256: delivered, updateWorkingRevision: await acceptedWorkingDocxUpdate(directory, file, revised) })

  // The run registers again (as a later pass would): the registry refuses the
  // accepted revision's name and the run does not fall back to "(2)".
  const again = await recordLeanDocuments(directory, deps, true)
  expect(again.documents.map(row => [row.kind, row.deliverable, row.conflict])).toEqual([
    ["body", "Supply note.docx", "accepted_revision"],
    ["redline", "Supply note - redline.docx", undefined],
    ["drafting_notes", "Supply note - drafting notes.docx", undefined],
    ["drafting_notes_markdown", null, undefined],
  ])
  expect(again.documents[0]!.sha256).toBe(delivered)
  expect(readFileSync(file)).toEqual(revised)
  expect(deliverables(directory)).toEqual(["Supply note - drafting notes.docx", "Supply note - redline.docx", "Supply note.docx"])
  // The status keeps the final check's own verdict and discloses the change.
  const gate = readLeanJson(directory, "gate.json") as { status: string; checked_at: string }
  expect(gate.status).toBe("pass")
  expect(leanDraftingStatus(directory)).toMatchObject({ gate, changed_after_check: { since: gate.checked_at, documents: ["Supply note.docx"] } })
}, 120_000)
