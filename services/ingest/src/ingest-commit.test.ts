import { afterAll, beforeAll, describe, expect, spyOn, test } from "bun:test"
import { existsSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import * as atomicWrite from "./atomic-document-write"
import { deleteDocument, getMeta, insertChunk, openDb, setMeta, upsertDocument } from "./db"
import { MODEL } from "./embed"
import * as identifierExtraction from "./identifiers"
import { ingestDocument } from "./ingest"
import { sourceArtifactsMatchRegistry, writeSourceArtifacts } from "./source-artifacts"
import { chunkSection, migrateSections, migrateStructure, sectionize } from "./structure"

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

// A matter of `count` invented agreements written straight into the registry
// (zero vectors under the current model id, so nothing is embedded again),
// large enough that extracting its identifiers takes a noticeable time.
function writeCorpus(directory: string, count: number) {
  const db = openDb(directory)
  try {
    db.transaction(() => {
      for (let d = 0; d < count; d++) {
        const lines = [
          "MASTER SERVICES AGREEMENT",
          `This invented agreement is made between Example Holdings ${d} Ltd and Sample Services ${d} LLC.`,
        ]
        for (let clause = 1; clause <= 30; clause++) {
          lines.push(`${clause}. Clause heading ${clause}`)
          for (let paragraph = 1; paragraph <= 4; paragraph++)
            lines.push(`${clause}.${paragraph} ${`The invented party ${d} performs step ${clause}.${paragraph} with due care. `.repeat(6)}`)
        }
        const text = lines.join("\n")
        const name = `Corpus ${String(d).padStart(2, "0")}.docx`
        const docPath = path.join(directory, name)
        writeFileSync(docPath, "synthetic")
        const documentId = upsertDocument(db, docPath, name, d + 1, null)
        let chunkIndex = 0
        for (const section of sectionize(text))
          for (const chunk of chunkSection(section))
            insertChunk(db, {
              documentId,
              docPath,
              docName: name,
              section: section.label,
              structuralPath: name,
              party: "",
              chunkIndex: chunkIndex++,
              text: chunk.text,
              charStart: chunk.charStart,
              charEnd: chunk.charEnd,
              embedding: new Float32Array(384),
              flagged: false,
            })
      }
    })()
    setMeta(db, "embedding_model", MODEL)
    migrateStructure(db)
    migrateSections(db)
  } finally {
    db.close()
  }
}

// Every source artifact file's text, keyed by its path in the matter.
function sourceArtifactFiles(directory: string) {
  const files: Record<string, string> = {}
  for (const name of ["source-index.json", "source-map.md", "source-map.json", "source-ledger.json", "matter-identifiers.json", "matter-identifiers.md"])
    files[name] = readFileSync(path.join(directory, name), "utf8")
  for (const folder of ["sources", "source-chunks"])
    for (const name of readdirSync(path.join(directory, folder)).sort())
      files[`${folder}/${name}`] = readFileSync(path.join(directory, folder, name), "utf8")
  return files
}

// Another process (as the engine is) writes legal.db through the engine's
// write connection every 20 ms until `stop` exists, and logs each write.
async function startEngineWriter(directory: string) {
  const log = path.join(directory, "engine-writes.jsonl")
  const stop = path.join(directory, "engine-writes.stop")
  const writer = Bun.spawn(
    [
      process.execPath,
      "-e",
      `const { appendFileSync, existsSync } = require("node:fs")
      const { openLegalDbForWrite } = require(${JSON.stringify(path.resolve(import.meta.dir, "../../../dochaus/lib/legal-db.ts"))})
      const db = openLegalDbForWrite(${JSON.stringify(directory)})
      const write = () => {
        const at = Date.now()
        let ok = true
        try { db.run("INSERT INTO meta (key, value) VALUES ('synthetic-engine', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value", [String(at)]) }
        catch { ok = false }
        appendFileSync(${JSON.stringify(log)}, JSON.stringify({ at, ms: Date.now() - at, ok }) + "\\n")
      }
      write()
      console.log("writing")
      const timer = setInterval(() => {
        if (existsSync(${JSON.stringify(stop)})) { clearInterval(timer); db.close(); return }
        write()
      }, 20)`,
    ],
    { stdout: "pipe" },
  )
  const reader = writer.stdout.getReader()
  expect(new TextDecoder().decode((await reader.read()).value)).toContain("writing")
  reader.releaseLock()
  return {
    async stop() {
      writeFileSync(stop, "")
      await writer.exited
      return readFileSync(log, "utf8").trim().split("\n").map((line) => JSON.parse(line) as { at: number; ms: number; ok: boolean })
    },
    kill: () => writer.kill(),
  }
}

describe("an upload into a multi-document matter", () => {
  let directory = ""
  beforeAll(() => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-ingest-corpus-")))
    writeCorpus(directory, 12)
  })
  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  test("holds the legal.db write lock for milliseconds, so the engine writes throughout", async () => {
    const writer = await startEngineWriter(directory)
    const extract = spyOn(identifierExtraction, "extractIdentifiers")
    let writes: { at: number; ms: number; ok: boolean }[] = []
    let started = 0
    let finished = 0
    try {
      started = Date.now()
      await ingestDocument(directory, "Uploaded Notice.docx", await docx(["1. Scope", "The invented scope of Example Holdings 1 Ltd."]))
      finished = Date.now()
      // The whole matter's identifiers are extracted once, before the lock.
      expect(extract).toHaveBeenCalledTimes(1)
    } finally {
      extract.mockRestore()
      writes = await writer.stop().catch((error) => (writer.kill(), Promise.reject(error)))
    }
    const during = writes.filter((write) => write.at >= started && write.at <= finished)
    expect(during.filter((write) => !write.ok)).toEqual([])
    expect(Math.max(...during.map((write) => write.ms))).toBeLessThan(500)
    expect(during.length).toBeGreaterThan(10)
    expect(sourceArtifactsMatchRegistry(directory)).toBe(true)
  }, 120_000)

  test("installs the snapshot staged before the lock, byte for byte the one built in the lock", async () => {
    const extract = spyOn(identifierExtraction, "extractIdentifiers")
    try {
      await ingestDocument(directory, "Second Notice.docx", await docx(["2. Transport", "Sample Services 2 LLC moves the invented box."]))
      expect(extract).toHaveBeenCalledTimes(1)
    } finally {
      extract.mockRestore()
    }
    const staged = sourceArtifactFiles(directory)
    expect(Object.keys(staged)).toContain("sources/" + readdirSync(path.join(directory, "sources")).find((name) => name.startsWith("second-notice-docx-")))
    await writeSourceArtifacts(directory, undefined, { rebuild: true })
    expect(sourceArtifactFiles(directory)).toEqual(staged)
  }, 120_000)

  test("restages in the lock when the registry changed after staging", async () => {
    const removed = path.join(directory, "Corpus 03.docx")
    const extract = spyOn(identifierExtraction, "extractIdentifiers")
    try {
      await ingestDocument(directory, "Third Notice.docx", await docx(["3. Receipt", "The invented receipt."]), {
        updateRegistry: (db) => deleteDocument(db, removed),
      })
      // Once before the lock, and again for the registry the commit holds.
      expect(extract).toHaveBeenCalledTimes(2)
    } finally {
      extract.mockRestore()
    }
    expect(sourceArtifactsMatchRegistry(directory)).toBe(true)
    const index = JSON.parse(readFileSync(path.join(directory, "source-index.json"), "utf8")) as { sources: { document: string }[] }
    expect(index.sources.map((source) => source.document)).not.toContain("Corpus 03.docx")
    expect(index.sources.map((source) => source.document)).toContain("Third Notice.docx")
    const restaged = sourceArtifactFiles(directory)
    await writeSourceArtifacts(directory, undefined, { rebuild: true })
    expect(sourceArtifactFiles(directory)).toEqual(restaged)
    // No staged snapshot is left behind.
    expect(existsSync(path.join(directory, ".dochaus", "document-mutations")) &&
      readdirSync(path.join(directory, ".dochaus", "document-mutations")).length).toBeFalsy()
  }, 120_000)
})
