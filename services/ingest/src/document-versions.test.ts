import { afterAll, expect, spyOn, test } from "bun:test"
import { Database } from "bun:sqlite"
import * as fs from "node:fs"
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Header, InsertedTextRun, Packer, Paragraph } from "docx"
import { openDb } from "./db"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import * as documentMutation from "./document-mutation"
import { readDocumentVersions, trackedChanges } from "./document-versions"
import { ingestDocument } from "./ingest"

// Document version capture inside the writers (#268), in process. Fixtures are
// neutral synthetic paragraphs built in code.

const directories: string[] = []
afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function matter() {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-document-versions-")))
  directories.push(directory)
  openDb(directory).close()
  return directory
}

const docx = async (...texts: string[]) =>
  Buffer.from(await Packer.toBuffer(new Document({ sections: [{ children: texts.map((text) => new Paragraph(text)) }] })))
const versions = (directory: string, name: string) => readDocumentVersions(directory, name).versions.map((row) => [row.number, row.origin])

test("tracked changes are counted in every part of a Word file, not only its body", async () => {
  const inserted = new InsertedTextRun({ text: "Amended heading", id: 1, author: "Pat Reviewer", date: "2026-10-15T09:00:00Z" })
  const bytes = await Packer.toBuffer(new Document({ sections: [{
    headers: { default: new Header({ children: [new Paragraph({ children: [inserted] })] }) },
    children: [new Paragraph("The invented supplier delivers the sample goods.")],
  }] }))
  expect(trackedChanges("Agreement.docx", new Uint8Array(bytes))).toEqual({ count: 1, authors: ["Pat Reviewer"] })
})

test("version bytes are stored before legal.db's write lock is taken, never while a writer holds it", async () => {
  const directory = matter()
  // A second connection that tries for the write lock without waiting.
  const probe = new Database(path.join(directory, ".dochaus", "legal.db"))
  probe.run("PRAGMA busy_timeout = 0")
  const locked = () => {
    try {
      probe.run("BEGIN IMMEDIATE")
      probe.run("ROLLBACK")
      return false
    } catch {
      return true
    }
  }
  const store = `${path.sep}.dochaus${path.sep}revisions${path.sep}`
  const opened: boolean[] = []
  const { openSync: open } = fs
  const spy = spyOn(fs, "openSync").mockImplementation(((file: fs.PathLike, flags?: fs.OpenMode, mode?: fs.Mode | null) => {
    if (String(file).includes(store)) opened.push(locked())
    return open(file, flags as fs.OpenMode, mode)
  }) as typeof fs.openSync)
  try {
    await writeDeliverable(directory, "Letter.docx", await docx("Advice"))
    await writeDeliverableRevision(directory, "Letter.docx", await docx("Advice, as revised"), { version: { origin: "revision" } })
    await ingestDocument(directory, "Agreement.docx", await docx("The invented supplier delivers the sample goods."))
    // Changed outside doc.haus, then replaced: the outside bytes are kept too.
    writeFileSync(path.join(directory, "Agreement.docx"), await docx("A clause added outside doc.haus."))
    await ingestDocument(directory, "Agreement.docx", await docx("Payment is due within thirty days."), { existing: "replace" })
  } finally {
    spy.mockRestore()
    probe.close()
  }
  expect(opened.length).toBeGreaterThan(0)
  expect(opened.filter(Boolean)).toEqual([])
  expect(versions(directory, "Letter.docx")).toEqual([[2, "revision"], [1, "generated"]])
  expect(versions(directory, "Agreement.docx")).toEqual([[3, "replace"], [2, "unknown"], [1, "upload"]])
}, 60_000)

test("a write that fails after its version is captured leaves the working file and history as they were", async () => {
  const directory = matter()
  const original = await docx("The invented supplier delivers the sample goods.")
  const replacement = await docx("Payment is due within thirty days.")
  await ingestDocument(directory, "Agreement.docx", original)
  await writeDeliverable(directory, "Letter.docx", await docx("Advice"))
  const letter = readFileSync(path.join(directory, "Letter.docx"))
  // The registry commit fails after captureVersion has run in the transaction.
  const commit = spyOn(documentMutation, "recordDocumentMutationCommit").mockImplementation(() => {
    throw new Error("synthetic commit failure")
  })
  try {
    await expect(ingestDocument(directory, "Agreement.docx", replacement, { existing: "replace" })).rejects.toThrow("synthetic commit failure")
    await expect(writeDeliverableRevision(directory, "Letter.docx", await docx("Advice, as revised"))).rejects.toThrow("synthetic commit failure")
  } finally {
    commit.mockRestore()
  }
  expect(readFileSync(path.join(directory, "Agreement.docx"))).toEqual(original)
  expect(readFileSync(path.join(directory, "Letter.docx"))).toEqual(letter)
  expect(versions(directory, "Agreement.docx")).toEqual([[1, "upload"]])
  expect(versions(directory, "Letter.docx")).toEqual([[1, "generated"]])
  // The same write then succeeds as the next version.
  expect((await ingestDocument(directory, "Agreement.docx", replacement, { existing: "replace" })).version).toMatchObject({ number: 2, origin: "replace" })
}, 60_000)
