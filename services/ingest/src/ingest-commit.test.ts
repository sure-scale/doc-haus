import { afterAll, beforeAll, describe, expect, setSystemTime, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import * as fs from "node:fs"
import { existsSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import * as atomicWrite from "./atomic-document-write"
import { deleteDocument, getMeta, insertChunk, openDb, setMeta, upsertDocument } from "./db"
import { copyWordDocument, writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { deleteMatterDocument } from "./document-delete"
import { MODEL } from "./embed"
import * as identifierExtraction from "./identifiers"
import { ingestDocument } from "./ingest"
import { stampIngestLineage } from "./lineage"
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

// Longer than a request connection's wait for a write lock (200 ms), well
// within a commit connection's (500 ms).
const COMMIT_HOLD_MS = 300

// Runs `action` while another process takes legal.db's write lock just before
// the action's registry transaction begins and holds it COMMIT_HOLD_MS.
async function againstHeldWriteLock(directory: string, action: () => Promise<unknown>) {
  let holder: ReturnType<typeof Bun.spawn> | undefined
  const { commitDocumentWriteAsync, commitDocumentDeleteAsync } = atomicWrite
  const write = spyOn(atomicWrite, "commitDocumentWriteAsync").mockImplementation(async (...args) => {
    holder = await holdWriteLockInAnotherProcess(directory, COMMIT_HOLD_MS)
    return commitDocumentWriteAsync(...args)
  })
  const remove = spyOn(atomicWrite, "commitDocumentDeleteAsync").mockImplementation(async (...args) => {
    holder = await holdWriteLockInAnotherProcess(directory, COMMIT_HOLD_MS)
    return commitDocumentDeleteAsync(...args)
  })
  try {
    await action()
  } finally {
    write.mockRestore()
    remove.mockRestore()
    if (holder) await holder.exited
  }
  expect(holder?.exitCode).toBe(0)
  const db = openDb(directory)
  try {
    expect(getMeta(db, "synthetic-holder")).toBe("1")
    return db.query("SELECT name, role FROM documents ORDER BY created_at").all() as { name: string; role: string }[]
  } finally {
    db.close()
  }
}

describe("a delete's or deliverable's legal.db commit waits briefly for another process's write lock", () => {
  const directories: string[] = []
  const matter = () => {
    const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-commit-wait-")))
    directories.push(directory)
    return directory
  }
  afterAll(() => directories.forEach((directory) => rmSync(directory, { recursive: true, force: true })))

  test("deleting a source document", async () => {
    const directory = matter()
    writeCorpus(directory, 2)
    expect(await againstHeldWriteLock(directory, () => deleteMatterDocument(directory, "Corpus 00.docx")))
      .toEqual([{ name: "Corpus 01.docx", role: "source" }])
  }, 60_000)

  test("deleting a deliverable", async () => {
    const directory = matter()
    await writeDeliverable(directory, "Invented Memo.docx", await docx(["The invented memo."]))
    expect(await againstHeldWriteLock(directory, () => deleteMatterDocument(directory, "Invented Memo.docx"))).toEqual([])
  }, 60_000)

  test("writing a deliverable", async () => {
    const directory = matter()
    const bytes = await docx(["The invented memo."])
    expect(await againstHeldWriteLock(directory, () => writeDeliverable(directory, "Invented Memo.docx", bytes)))
      .toEqual([{ name: "Invented Memo.docx", role: "deliverable" }])
  }, 60_000)

  test("writing a deliverable's revision", async () => {
    const directory = matter()
    await writeDeliverable(directory, "Invented Memo.docx", await docx(["The invented memo."]))
    const bytes = await docx(["The revised invented memo."])
    expect(await againstHeldWriteLock(directory, () => writeDeliverableRevision(directory, "Invented Memo.docx", bytes)))
      .toEqual([{ name: "Invented Memo.docx", role: "deliverable" }])
  }, 60_000)

  test("copying a Word document", async () => {
    const directory = matter()
    await writeDeliverable(directory, "Invented Memo.docx", await docx(["The invented memo."]))
    expect(await againstHeldWriteLock(directory, () => copyWordDocument(directory, "Invented Memo.docx", "Invented Memo Copy.docx")))
      .toEqual([{ name: "Invented Memo.docx", role: "deliverable" }, { name: "Invented Memo Copy.docx", role: "deliverable" }])
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

describe("a source delete from a multi-document matter", () => {
  let directory = ""
  beforeAll(() => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-delete-corpus-")))
    writeCorpus(directory, 12)
  })
  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  const indexedSources = () =>
    (JSON.parse(readFileSync(path.join(directory, "source-index.json"), "utf8")) as { sources: { document: string }[] })
      .sources.map((source) => source.document)

  test("holds the legal.db write lock for milliseconds, so the engine writes throughout", async () => {
    const writer = await startEngineWriter(directory)
    const extract = spyOn(identifierExtraction, "extractIdentifiers")
    let writes: { at: number; ms: number; ok: boolean }[] = []
    let started = 0
    let finished = 0
    try {
      started = Date.now()
      await deleteMatterDocument(directory, "Corpus 00.docx")
      finished = Date.now()
      // The remaining matter's identifiers are extracted once, before the lock.
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
    expect(indexedSources()).not.toContain("Corpus 00.docx")
  }, 120_000)

  test("installs the snapshot staged before the lock, byte for byte the one built in the lock", async () => {
    const extract = spyOn(identifierExtraction, "extractIdentifiers")
    try {
      await deleteMatterDocument(directory, "Corpus 01.docx")
      expect(extract).toHaveBeenCalledTimes(1)
    } finally {
      extract.mockRestore()
    }
    expect(indexedSources()).not.toContain("Corpus 01.docx")
    expect(indexedSources()).toContain("Corpus 02.docx")
    const staged = sourceArtifactFiles(directory)
    await writeSourceArtifacts(directory, undefined, { rebuild: true })
    expect(sourceArtifactFiles(directory)).toEqual(staged)
  }, 120_000)

  test("restages in the lock when the registry changed after staging", async () => {
    const removed = path.join(directory, "Corpus 03.docx")
    const commitDocumentDeleteAsync = atomicWrite.commitDocumentDeleteAsync
    // Another writer drops a source between the staging and the transaction.
    const commit = spyOn(atomicWrite, "commitDocumentDeleteAsync").mockImplementation(async (...args) => {
      const db = openDb(directory)
      try {
        deleteDocument(db, removed)
      } finally {
        db.close()
      }
      return commitDocumentDeleteAsync(...args)
    })
    const extract = spyOn(identifierExtraction, "extractIdentifiers")
    try {
      await deleteMatterDocument(directory, "Corpus 02.docx")
      // Once before the lock, and again for the registry the commit holds.
      expect(extract).toHaveBeenCalledTimes(2)
    } finally {
      extract.mockRestore()
      commit.mockRestore()
    }
    expect(sourceArtifactsMatchRegistry(directory)).toBe(true)
    expect(indexedSources()).not.toContain("Corpus 02.docx")
    expect(indexedSources()).not.toContain("Corpus 03.docx")
    expect(indexedSources()).toContain("Corpus 04.docx")
    const restaged = sourceArtifactFiles(directory)
    await writeSourceArtifacts(directory, undefined, { rebuild: true })
    expect(sourceArtifactFiles(directory)).toEqual(restaged)
    // No staged snapshot is left behind.
    expect(existsSync(path.join(directory, ".dochaus", "document-mutations")) &&
      readdirSync(path.join(directory, ".dochaus", "document-mutations")).length).toBeFalsy()
  }, 120_000)
})

type LineageRow = {
  artifact: string
  content_sha256: string
  bytes: number
  run: string
  provenance: { documents: string[]; document_mutation_id?: string }
}

const lineageFile = (directory: string) => path.join(directory, ".dochaus", "lineage.jsonl")

const lineageRows = (directory: string) =>
  readFileSync(lineageFile(directory), "utf8").trim().split("\n").map((line) => JSON.parse(line) as LineageRow)

// Every installed artifact an ingest stamps, with its bytes' hash and size.
function installedArtifacts(directory: string) {
  return [".dochaus/legal.db", ...Object.keys(sourceArtifactFiles(directory))]
    .map((artifact) => {
      const bytes = readFileSync(path.join(directory, artifact))
      return { artifact, content_sha256: createHash("sha256").update(bytes).digest("hex"), bytes: bytes.length }
    })
    .toSorted((left, right) => left.artifact.localeCompare(right.artifact))
}

function isSourceArtifact(directory: string, file: unknown) {
  if (typeof file !== "string") return false
  const relative = path.relative(directory, file).split(path.sep).join("/")
  return relative.startsWith("sources/") ||
    relative.startsWith("source-chunks/") ||
    ["source-index.json", "source-map.md", "source-map.json", "source-ledger.json", "matter-identifiers.json", "matter-identifiers.md"]
      .includes(relative)
}

// Runs a source mutation and records the source artifacts it read back with
// readFileSync and the source-artifact transactions it removed synchronously
// while they still held the replaced snapshot.
async function observeSourceMutation(directory: string, mutate: () => Promise<unknown>) {
  const { readFileSync: read, rmSync: remove } = fs
  const reads: string[] = []
  const snapshotRemovals: string[] = []
  const readSpy = spyOn(fs, "readFileSync").mockImplementation(((file: fs.PathOrFileDescriptor, options?: never) => {
    if (isSourceArtifact(directory, file)) reads.push(path.relative(directory, String(file)))
    return read(file, options)
  }) as typeof fs.readFileSync)
  const removeSpy = spyOn(fs, "rmSync").mockImplementation(((target: fs.PathLike, options?: fs.RmOptions) => {
    if (typeof target === "string" && existsSync(path.join(target, "backup"))) snapshotRemovals.push(path.relative(directory, target))
    return remove(target, options)
  }) as typeof fs.rmSync)
  try {
    await mutate()
  } finally {
    readSpy.mockRestore()
    removeSpy.mockRestore()
  }
  return { reads, snapshotRemovals }
}

// Waits for the replaced snapshot's removal; true once no transaction is left.
async function transactionsSettled(directory: string) {
  const roots = [".dochaus/document-mutations", ".dochaus/source-artifact-transactions"].map((root) => path.join(directory, root))
  for (let waited = 0; waited <= 10_000; waited += 25) {
    if (roots.every((root) => !existsSync(root) || !readdirSync(root).length)) return true
    await Bun.sleep(25)
  }
  return false
}

describe("a source mutation's lineage and replaced snapshot", () => {
  let directory = ""
  beforeAll(async () => {
    directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-source-lineage-")))
    writeCorpus(directory, 4)
    await writeSourceArtifacts(directory)
  })
  afterAll(() => rmSync(directory, { recursive: true, force: true }))

  for (const [label, mutate] of [
    ["an upload", async () => ingestDocument(directory, "Lineage Notice.docx", await docx(["1. Scope", "The invented scope of Example Holdings 7 Ltd."]))],
    ["a delete", () => deleteMatterDocument(directory, "Corpus 00.docx")],
  ] as const) {
    test(`${label} stamps lineage from the bytes it staged and removes the replaced snapshot off the event loop`, async () => {
      const before = lineageRows(directory).length
      // The replaced snapshot is moved aside and removed off the event loop.
      expect(await observeSourceMutation(directory, mutate)).toEqual({ reads: [], snapshotRemovals: [] })
      expect(
        lineageRows(directory).slice(before)
          .map(({ artifact, content_sha256, bytes }) => ({ artifact, content_sha256, bytes }))
          .toSorted((left, right) => left.artifact.localeCompare(right.artifact)),
      ).toEqual(installedArtifacts(directory))
      expect(await transactionsSettled(directory)).toBe(true)
    }, 120_000)
  }

  test("an upload and a delete append the same lineage bytes as stamping the installed files", async () => {
    // A frozen clock and fixed random bytes make two stamps of the same
    // artifacts comparable byte for byte (entry ids and timestamps).
    setSystemTime(new Date(Date.now()))
    const random = spyOn(globalThis.crypto, "getRandomValues")
      .mockImplementation(((array: Uint8Array) => array.fill(7)) as typeof globalThis.crypto.getRandomValues)
    try {
      for (const mutate of [
        async () => ingestDocument(directory, "Second Lineage Notice.docx", await docx(["2. Transport", "Sample Services 8 LLC moves the invented box."])),
        () => deleteMatterDocument(directory, "Corpus 01.docx"),
      ]) {
        const before = readFileSync(lineageFile(directory))
        await mutate()
        const after = readFileSync(lineageFile(directory))
        const [first] = after.subarray(before.length).toString("utf8").trim().split("\n").map((line) => JSON.parse(line) as LineageRow)
        // The stamp as an ingest always made it: every present artifact read and hashed from disk.
        writeFileSync(lineageFile(directory), before)
        stampIngestLineage(directory, first!.provenance.documents, first!.run.slice("ingest:".length), first!.provenance.document_mutation_id)
        expect(readFileSync(lineageFile(directory)).equals(after)).toBe(true)
      }
    } finally {
      random.mockRestore()
      setSystemTime()
    }
  }, 120_000)
})
