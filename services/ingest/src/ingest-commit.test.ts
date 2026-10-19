import { describe, expect, spyOn, test } from "bun:test"
import { mkdtempSync, realpathSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import * as atomicWrite from "./atomic-document-write"
import { getMeta, openDb } from "./db"
import { ingestDocument } from "./ingest"

// Every fixture below is invented test content: no actual person,
// organization, transaction or obligation.

const docx = async (paragraphs: string[]) =>
  Buffer.from(await Packer.toBuffer(new Document({ sections: [{ children: paragraphs.map((text) => new Paragraph(text)) }] })))

const legalDb = (directory: string) => path.join(directory, ".dochaus", "legal.db")

// Another process (as the engine is) takes legal.db's write lock, writes one
// row and commits it `holdMs` later. Resolves once the lock is held.
async function holdWriteLockInAnotherProcess(directory: string, holdMs: number) {
  const holder = Bun.spawn(
    [
      process.execPath,
      "-e",
      `const { Database } = require("bun:sqlite")
      const db = new Database(${JSON.stringify(legalDb(directory))})
      db.run("BEGIN IMMEDIATE")
      db.run("INSERT INTO meta (key, value) VALUES ('synthetic-holder', '1')")
      console.log("locked")
      setTimeout(() => { db.run("COMMIT"); db.close() }, ${holdMs})`,
    ],
    { stdout: "pipe" },
  )
  const reader = holder.stdout.getReader()
  expect(new TextDecoder().decode((await reader.read()).value)).toContain("locked")
  reader.releaseLock()
  return holder
}

describe("an upload's legal.db commit", () => {
  test("waits briefly for another process's write lock instead of failing", async () => {
    const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-ingest-commit-")))
    let holder: ReturnType<typeof Bun.spawn> | undefined
    try {
      // The first upload leaves a current legal.db, so only the second one's
      // write transaction meets the lock.
      await ingestDocument(directory, "First Notice.docx", await docx(["1. Scope", "The invented scope."]))
      const commitDocumentWriteAsync = atomicWrite.commitDocumentWriteAsync
      const commit = spyOn(atomicWrite, "commitDocumentWriteAsync").mockImplementation(async (...args) => {
        holder = await holdWriteLockInAnotherProcess(directory, 100)
        return commitDocumentWriteAsync(...args)
      })
      try {
        await ingestDocument(directory, "Second Notice.docx", await docx(["2. Transport", "The invented box moves."]))
      } finally {
        commit.mockRestore()
      }
      expect(holder).toBeDefined()
      expect(await holder!.exited).toBe(0)
      const db = openDb(directory)
      try {
        expect((db.query("SELECT name FROM documents ORDER BY created_at").all() as { name: string }[]).map((row) => row.name))
          .toEqual(["First Notice.docx", "Second Notice.docx"])
        expect(getMeta(db, "synthetic-holder")).toBe("1")
      } finally {
        db.close()
      }
    } finally {
      holder?.kill()
      rmSync(directory, { recursive: true, force: true })
    }
  }, 60_000)
})
