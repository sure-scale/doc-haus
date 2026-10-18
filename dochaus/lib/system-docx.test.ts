import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { recordRedline } from "./redlines"
import { forgetSystemDocx, ownedDeliverable, readSystemDocx, recordSystemDocx, replaceableDeliverable, systemDocxKey } from "./system-docx"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

const sha256 = (bytes: string) => createHash("sha256").update(bytes).digest("hex")

function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-system-docx-"))
  directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  return directory
}

function registry(directory: string, write: (db: Database) => void) {
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  try { db.transaction(() => write(db)).immediate() } finally { db.close() }
}

test("a file is owned while it holds exactly the bytes the system recorded writing", () => {
  const directory = matter()
  writeFileSync(path.join(directory, "Pack.docx"), "system bytes")
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(false)
  registry(directory, db => recordSystemDocx(db, directory, "Pack.docx", sha256("system bytes")))
  expect(readSystemDocx(directory, "Pack.docx")).toEqual({ version: 1, sha256: sha256("system bytes") })
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(true)
  // The matter path names the same file.
  expect(ownedDeliverable(directory, path.join(directory, "Pack.docx"))).toBe(true)
  // A change made outside doc.haus no longer matches.
  writeFileSync(path.join(directory, "Pack.docx"), "edited bytes")
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(false)
  writeFileSync(path.join(directory, "Pack.docx"), "system bytes")
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(true)
  // Any other write removes the record: the same bytes are then not owned.
  registry(directory, db => forgetSystemDocx(db, directory, "Pack.docx"))
  expect(readSystemDocx(directory, "Pack.docx")).toBeUndefined()
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(false)
})

test("without a registry, a record or a regular file, nothing is owned", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-system-docx-"))
  directories.push(directory)
  writeFileSync(path.join(directory, "Pack.docx"), "bytes")
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(false)
  mkdirSync(path.join(directory, ".dochaus"))
  registry(directory, db => db.run("CREATE TABLE chunks (text TEXT)"))
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(false)
  // Forgetting needs no record, nor a meta table.
  registry(directory, db => forgetSystemDocx(db, directory, "Pack.docx"))
  // A recorded file that is gone, or a link in its place, is not owned.
  registry(directory, db => recordSystemDocx(db, directory, "Gone.docx", sha256("bytes")))
  expect(ownedDeliverable(directory, "Gone.docx")).toBe(false)
  symlinkSync(path.join(directory, "Pack.docx"), path.join(directory, "Gone.docx"))
  expect(ownedDeliverable(directory, "Gone.docx")).toBe(false)
})

test("a record that is not the system write record is refused on read and owns nothing, and only a sha256 is recorded", () => {
  const directory = matter()
  writeFileSync(path.join(directory, "Pack.docx"), "bytes")
  expect(() => registry(directory, db => recordSystemDocx(db, directory, "Pack.docx", "not-a-hash"))).toThrow("sha256")
  registry(directory, db => db.run("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)"))
  for (const value of [JSON.stringify({ version: 2, sha256: sha256("bytes") }), JSON.stringify({ version: 1 }), JSON.stringify({ version: 1, sha256: "A".repeat(64) }), "null", "{"]) {
    registry(directory, db => db.run("INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value", [systemDocxKey(directory, "Pack.docx"), value]))
    expect(() => readSystemDocx(directory, "Pack.docx")).toThrow()
    // Never the system's to write over.
    expect(ownedDeliverable(directory, "Pack.docx")).toBe(false)
  }
})

test("the system may write over its own unchanged write only while no proposal on the file is pending", () => {
  const directory = matter()
  writeFileSync(path.join(directory, "Pack.docx"), "system bytes")
  writeFileSync(path.join(directory, "Other.docx"), "system bytes")
  registry(directory, db => {
    recordSystemDocx(db, directory, "Pack.docx", sha256("system bytes"))
    recordSystemDocx(db, directory, "Other.docx", sha256("system bytes"))
  })
  // A registry without proposals has none pending.
  expect(replaceableDeliverable(directory, "Pack.docx")).toBe(true)
  recordRedline(directory, { docPath: "Pack.docx", docName: "Pack.docx", scope: "phrase", findText: "system", oldText: "system", newText: "lawyer", author: "reviewer", reviewComment: "", anchorId: "p000000000001" })
  expect(ownedDeliverable(directory, "Pack.docx")).toBe(true)
  expect(replaceableDeliverable(directory, "Pack.docx")).toBe(false)
  expect(replaceableDeliverable(directory, path.join(directory, "Pack.docx"))).toBe(false)
  // A proposal on another file does not hold this one.
  expect(replaceableDeliverable(directory, "Other.docx")).toBe(true)
  // Once the lawyer resolves it, the file is the system's again.
  registry(directory, db => db.run("UPDATE redlines SET status = 'rejected'"))
  expect(replaceableDeliverable(directory, "Pack.docx")).toBe(true)
  // A file that is not the system's own is not, proposals or not.
  writeFileSync(path.join(directory, "Pack.docx"), "edited bytes")
  expect(replaceableDeliverable(directory, "Pack.docx")).toBe(false)
  // Nor is one whose proposals cannot be read.
  registry(directory, db => { db.run("DROP TABLE redlines"); db.run("CREATE TABLE redlines (id INTEGER)") })
  expect(replaceableDeliverable(directory, "Other.docx")).toBe(false)
})
