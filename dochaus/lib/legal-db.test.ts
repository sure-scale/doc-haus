import { afterEach, expect, spyOn, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { listKeyDates, recordKeyDate } from "./key-dates"
import { listRedactions, recordRedaction } from "./redactions"
import { openLegalDbForWrite } from "./legal-db"
import { pendingRedlinesForDoc, recordRedline } from "./redlines"

// The engine's legal.db writers against a write lock held by another process,
// as the ingest service holds it while it commits an upload. Every value is an
// invented test fixture.

const directories: string[] = []
afterEach(() => directories.splice(0).forEach((directory) => rmSync(directory, { recursive: true, force: true })))

function matter() {
  const directory = mkdtempSync(path.join(tmpdir(), "legal-db-writers-"))
  directories.push(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.run("PRAGMA journal_mode = WAL")
  db.run("CREATE TABLE ingest_rows (id INTEGER PRIMARY KEY)")
  db.close()
  writeFileSync(path.join(directory, "Synthetic.docx"), "")
  return directory
}

// Another process takes the write lock and releases it after a short commit.
async function lockedBriefly(directory: string) {
  const holder = Bun.spawn(
    [
      process.execPath,
      "-e",
      `const { Database } = require("bun:sqlite")
      const db = new Database(${JSON.stringify(path.join(directory, ".dochaus", "legal.db"))})
      db.run("BEGIN IMMEDIATE")
      db.run("INSERT INTO ingest_rows DEFAULT VALUES")
      console.log("locked")
      setTimeout(() => { db.run("COMMIT"); db.close() }, 300)`,
    ],
    { stdout: "pipe" },
  )
  const reader = holder.stdout.getReader()
  expect(new TextDecoder().decode((await reader.read()).value)).toContain("locked")
  reader.releaseLock()
  return holder
}

test("the engine's legal.db writers wait for another process's short write instead of failing at once", async () => {
  const directory = matter()
  const docPath = path.join(directory, "Synthetic.docx")
  const writes = [
    () =>
      recordRedaction(directory, {
        docPath,
        docName: "Synthetic.docx",
        redactedText: "Invented Name",
        label: "[NAME]",
        reason: "invented personal name",
        author: "test",
        occurrences: 1,
        metadataHits: 0,
      }),
    () =>
      recordKeyDate(directory, {
        docName: "Synthetic.docx",
        kind: "renewal",
        dateIso: "2030-01-01",
        quote: "renews on 1 January 2030",
        reason: "invented renewal date",
        author: "test",
      }),
    () =>
      recordRedline(directory, {
        docPath,
        docName: "Synthetic.docx",
        scope: "phrase",
        findText: "invented",
        oldText: "invented",
        newText: "synthetic",
        author: "test",
        reviewComment: "",
        anchorId: "synthetic-anchor",
      }),
  ]
  for (const write of writes) {
    const holder = await lockedBriefly(directory)
    try {
      write()
    } finally {
      holder.kill()
      await holder.exited
    }
  }
  expect(listRedactions(directory).length).toBe(1)
  expect(listKeyDates(directory).length).toBe(1)
  expect(pendingRedlinesForDoc(directory, docPath).length).toBe(1)
}, 30_000)

// Fails the first statement matching `failing` as a lock timeout would, runs
// `write`, and returns how many connections it left open after the throw.
function leakedOnFailure(failing: RegExp, write: () => unknown) {
  const open = new Set<Database>()
  const run = Database.prototype.run
  const close = Database.prototype.close
  const runSpy = spyOn(Database.prototype, "run").mockImplementation(function (this: Database, ...args: Parameters<Database["run"]>) {
    open.add(this)
    if (failing.test(args[0])) throw new Error("database is locked")
    return run.apply(this, args)
  })
  const closeSpy = spyOn(Database.prototype, "close").mockImplementation(function (this: Database, ...args: Parameters<Database["close"]>) {
    open.delete(this)
    return close.apply(this, args)
  })
  try {
    expect(write).toThrow("database is locked")
  } finally {
    runSpy.mockRestore()
    closeSpy.mockRestore()
  }
  open.forEach((db) => db.close())
  return open.size
}

const redaction = (docPath: string) => ({
  docPath,
  docName: "Synthetic.docx",
  redactedText: "Invented Name",
  label: "[NAME]",
  reason: "invented personal name",
  author: "test",
  occurrences: 1,
  metadataHits: 0,
})
const keyDate = {
  docName: "Synthetic.docx",
  kind: "renewal" as const,
  dateIso: "2030-01-01",
  quote: "renews on 1 January 2030",
  reason: "invented renewal date",
  author: "test",
}
const redline = (docPath: string) => ({
  docPath,
  docName: "Synthetic.docx",
  scope: "phrase" as const,
  findText: "invented",
  oldText: "invented",
  newText: "synthetic",
  author: "test",
  reviewComment: "",
  anchorId: "synthetic-anchor",
})

test("a failed connection setup closes the connection before it throws", () => {
  const directory = matter()
  expect({
    open: leakedOnFailure(/journal_mode/, () => openLegalDbForWrite(directory)),
    keyDate: leakedOnFailure(/journal_mode/, () => recordKeyDate(directory, keyDate)),
    redline: leakedOnFailure(/CREATE TABLE IF NOT EXISTS redlines/, () => recordRedline(directory, redline(path.join(directory, "Synthetic.docx")))),
  }).toEqual({ open: 0, keyDate: 0, redline: 0 })
})

test("a failed redaction or key-date write closes its connection", () => {
  const directory = matter()
  expect({
    redaction: leakedOnFailure(/INSERT INTO redactions/, () => recordRedaction(directory, redaction(path.join(directory, "Synthetic.docx")))),
    keyDate: leakedOnFailure(/INSERT INTO key_dates/, () => recordKeyDate(directory, keyDate)),
  }).toEqual({ redaction: 0, keyDate: 0 })
  expect(listRedactions(directory)).toEqual([])
  expect(listKeyDates(directory)).toEqual([])
})
