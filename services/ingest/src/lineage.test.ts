import { describe, expect, setSystemTime, spyOn, test } from "bun:test"
import * as fs from "node:fs"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { createHash } from "node:crypto"
import { tmpdir } from "node:os"
import path from "node:path"
import { stampIngestLineage, stampSourceRebuildLineage } from "./lineage"

describe("ingest lineage appender", () => {
  test("stamps one PERMANENT row per changed artifact with supersession", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "ingest-lineage-"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      writeFileSync(path.join(dir, ".dochaus", "legal.db"), "")
      writeFileSync(path.join(dir, "source-index.json"), "{}")
      writeFileSync(path.join(dir, "source-map.json"), "{}")
      mkdirSync(path.join(dir, "source-chunks"), { recursive: true })
      writeFileSync(path.join(dir, "source-chunks", "agreement-c001.txt"), "chunk1")
      const first = stampIngestLineage(dir, ["Agreement.docx"], "corr-1")
      expect(first).toBe(4)
      const rows1 = ledgerRows(dir)
      expect(rows1.every((r: any) => r.tier === "PERMANENT" && r.op === "ingest" && r.subagent === "ingest" && r.run === "ingest:corr-1")).toBe(true)
      expect(rows1.some((r: any) => r.artifact === ".dochaus/legal.db")).toBe(true)
      expect(rows1.every((r: any) => r.supersedes === null)).toBe(true)

      // An unchanged re-ingest stamps nothing.
      expect(stampIngestLineage(dir, ["Agreement.docx"], "corr-2")).toBe(0)

      // A re-ingest that changed two artifacts stamps those two, each superseding its prior entry.
      const priorIds = new Map(rows1.map((r: any) => [r.artifact, r.entry_id]))
      writeFileSync(path.join(dir, "source-index.json"), "{\"documents\":1}")
      writeFileSync(path.join(dir, "source-chunks", "agreement-c001.txt"), "chunk1 revised")
      expect(stampIngestLineage(dir, ["Agreement.docx"], "corr-3")).toBe(2)
      const latest = ledgerRows(dir).filter((r: any) => r.run === "ingest:corr-3")
      expect(latest.map((r: any) => r.artifact)).toEqual(["source-index.json", "source-chunks/agreement-c001.txt"])
      for (const r of latest) expect(r.supersedes).toBe(priorIds.get(r.artifact))
      expect(staleArtifacts(dir)).toEqual([])
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("a second upload stamps only what it changed and lists the documents on one row", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "ingest-lineage-"))
    const write = (file: string, text: string) => writeFileSync(path.join(dir, file), text)
    try {
      for (const folder of [".dochaus", "sources", "source-chunks"]) mkdirSync(path.join(dir, folder), { recursive: true })
      write(".dochaus/legal.db", "registry 1")
      write("source-index.json", "{\"documents\":1}")
      write("sources/agreement.txt", "Clause 1.\n\nClause 2.\n")
      write("source-chunks/agreement-c001.txt", "Clause 1.")
      write("source-chunks/agreement-c002.txt", "Clause 2.")
      stampIngestLineage(dir, ["Agreement.docx"], "upload-1", "mutation-1")
      const firstUpload = ledgerRows(dir).length

      // The second upload adds a schedule and leaves the agreement's files untouched.
      write(".dochaus/legal.db", "registry 2")
      write("source-index.json", "{\"documents\":2}")
      write("sources/schedule.txt", "Schedule 1.\n")
      write("source-chunks/schedule-c001.txt", "Schedule 1.")
      stampIngestLineage(dir, ["Agreement.docx", "Schedule.docx"], "upload-2", "mutation-2")
      const added = ledgerRows(dir).slice(firstUpload)
      expect({
        untouched: added.filter((row: any) => row.artifact.includes("agreement")).map((row: any) => row.artifact),
        listingDocuments: added.filter((row: any) => "documents" in row.provenance).length,
      }).toEqual({ untouched: [], listingDocuments: 1 })
      expect(added.map((row: any) => [row.artifact, row.provenance])).toEqual([
        [".dochaus/legal.db", { origin: "upload", documents: ["Agreement.docx", "Schedule.docx"], document_mutation_id: "mutation-2" }],
        ["source-index.json", { origin: "upload", document_mutation_id: "mutation-2" }],
        ["source-chunks/schedule-c001.txt", { origin: "upload", document_mutation_id: "mutation-2" }],
        ["sources/schedule.txt", { origin: "upload", document_mutation_id: "mutation-2" }],
      ])
      expect(staleArtifacts(dir)).toEqual([])
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("a stamp retried after a partial append records only the artifacts still unrecorded", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "ingest-lineage-"))
    const ledger = path.join(dir, ".dochaus", "lineage.jsonl")
    const write = (file: string, text: string) => writeFileSync(path.join(dir, file), text)
    try {
      for (const folder of [".dochaus", "sources", "source-chunks"]) mkdirSync(path.join(dir, folder), { recursive: true })
      write(".dochaus/legal.db", "registry 1")
      write("sources/agreement.txt", "Clause 1.\n")
      write("source-chunks/agreement-c001.txt", "Clause 1.")
      stampIngestLineage(dir, ["Agreement.docx"], "upload-1", "mutation-1")
      const firstUpload = readFileSync(ledger, "utf8")

      write(".dochaus/legal.db", "registry 2")
      write("source-index.json", "{\"documents\":2}")
      write("sources/schedule.txt", "Schedule 1.\n")
      write("source-chunks/schedule-c001.txt", "Schedule 1.")
      expect(stampIngestLineage(dir, ["Agreement.docx", "Schedule.docx"], "upload-2", "mutation-2")).toBe(4)
      // The process stopped mid-append: two whole rows and part of a third reached the ledger.
      const appended = readFileSync(ledger, "utf8").slice(firstUpload.length).split("\n")
      writeFileSync(ledger, `${firstUpload}${appended[0]}\n${appended[1]}\n${appended[2]!.slice(0, 40)}`)

      expect(stampIngestLineage(dir, ["Agreement.docx", "Schedule.docx"], "upload-2", "mutation-2")).toBe(2)
      expect(ledgerRows(dir).slice(-2).map((row: any) => row.artifact)).toEqual(["source-chunks/schedule-c001.txt", "sources/schedule.txt"])
      expect(staleArtifacts(dir)).toEqual([])
      // Once complete, a further retry appends nothing.
      expect(stampIngestLineage(dir, ["Agreement.docx", "Schedule.docx"], "upload-2", "mutation-2")).toBe(0)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("a source rebuild stamps only artifacts whose bytes differ from their latest row, as a rebuild", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "ingest-lineage-"))
    try {
      const rows = () => readFileSync(path.join(dir, ".dochaus", "lineage.jsonl"), "utf8").trim().split("\n").map((row) => JSON.parse(row))
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      mkdirSync(path.join(dir, "source-chunks"), { recursive: true })
      mkdirSync(path.join(dir, "sources"), { recursive: true })
      writeFileSync(path.join(dir, ".dochaus", "legal.db"), "")
      writeFileSync(path.join(dir, "source-map.json"), "{}")
      writeFileSync(path.join(dir, "sources", "agreement.txt"), "text")
      writeFileSync(path.join(dir, "source-chunks", "agreement-c001.txt"), "chunk1")
      stampIngestLineage(dir, ["Agreement.docx"], "corr-1")
      const uploaded = rows()
      const chunkEntry = uploaded.find((row: any) => row.artifact === "source-chunks/agreement-c001.txt").entry_id

      expect(stampSourceRebuildLineage(dir, ["Agreement.docx"], "rebuild-1")).toBe(0)

      // The rebuild rewrote one chunk file. legal.db changed too, but a rebuild
      // never writes it, so it is not the rebuild's to stamp.
      writeFileSync(path.join(dir, "source-chunks", "agreement-c001.txt"), "chunk1\nrewrapped")
      writeFileSync(path.join(dir, ".dochaus", "legal.db"), "migrated")
      expect(stampSourceRebuildLineage(dir, ["Agreement.docx"], "rebuild-2")).toBe(1)
      expect(rows().slice(uploaded.length)).toEqual([
        expect.objectContaining({
          tier: "PERMANENT",
          artifact: "source-chunks/agreement-c001.txt",
          op: "write",
          run: "ingest:rebuild-2",
          stage: null,
          subagent: "ingest",
          content_sha256: createHash("sha256").update("chunk1\nrewrapped").digest("hex"),
          bytes: "chunk1\nrewrapped".length,
          supersedes: chunkEntry,
          provenance: { origin: "source-rebuild", documents: ["Agreement.docx"] },
        }),
      ])
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("appends exactly the bytes of one row per stamped artifact, in the ledger's format", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "ingest-lineage-"))
    const sha256 = (text: string) => createHash("sha256").update(text).digest("hex")
    const ledger = path.join(dir, ".dochaus", "lineage.jsonl")
    const appended = (before: string) => readFileSync(ledger, "utf8").slice(before.length)
    const ids = (text: string) => text.trim().split("\n").map((row) => (JSON.parse(row) as { entry_id: string }).entry_id)
    setSystemTime(new Date("2026-10-19T12:00:00.000Z"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      mkdirSync(path.join(dir, "source-chunks"), { recursive: true })
      mkdirSync(path.join(dir, "sources"), { recursive: true })
      const files = {
        ".dochaus/legal.db": "registry",
        "source-index.json": "{}\n",
        "source-map.json": "{\"documents\":[]}\n",
        "source-chunks/agreement-c001.txt": "Clause 1.",
        "source-chunks/agreement-c002.txt": "Clause 2 — “quoted”.",
        "sources/agreement.txt": "Clause 1.\n\nClause 2 — “quoted”.\n",
      }
      for (const [file, text] of Object.entries(files)) writeFileSync(path.join(dir, file), text)
      // An earlier ledger: superseded and current rows for source-map.json, a
      // corrupt and a blank line, a row this mutation already stamped for the
      // first chunk (its bytes are unchanged, so a retried finalize skips it),
      // and no final newline.
      const earlier = [
        JSON.stringify({ entry_id: "01EARLIER0000000000000000A", artifact: "source-map.json", content_sha256: "0".repeat(64) }),
        "{not json",
        "",
        JSON.stringify({
          entry_id: "01EARLIER0000000000000000B",
          artifact: "source-chunks/agreement-c001.txt",
          content_sha256: sha256(files["source-chunks/agreement-c001.txt"]),
          provenance: { document_mutation_id: "mutation-1" },
        }),
        JSON.stringify({ entry_id: "01EARLIER0000000000000000C", artifact: "source-map.json", content_sha256: "1".repeat(64) }),
      ].join("\n")
      writeFileSync(ledger, earlier)

      expect(stampIngestLineage(dir, ["Agreement.docx"], "corr-1", "mutation-1")).toBe(5)
      const ingest = appended(earlier)
      const [legalDb, index, map, chunk, source] = ids(ingest)
      // Only the first row a stamp appends lists the documents.
      const row = (
        entry_id: string | undefined,
        artifact: keyof typeof files,
        supersedes: string | null,
        provenance: object = { origin: "upload", document_mutation_id: "mutation-1" },
      ) => JSON.stringify({
        entry_id,
        ts: "2026-10-19T12:00:00.000Z",
        tier: "PERMANENT",
        artifact,
        op: "ingest",
        run: "ingest:corr-1",
        stage: null,
        subagent: "ingest",
        content_sha256: sha256(files[artifact]),
        bytes: Buffer.byteLength(files[artifact]),
        supersedes,
        provenance,
      })
      expect(ingest).toBe(`\n${[
        row(legalDb, ".dochaus/legal.db", null, { origin: "upload", documents: ["Agreement.docx"], document_mutation_id: "mutation-1" }),
        row(index, "source-index.json", null),
        row(map, "source-map.json", "01EARLIER0000000000000000C"),
        row(chunk, "source-chunks/agreement-c002.txt", null),
        row(source, "sources/agreement.txt", null),
      ].join("\n")}\n`)

      // A rebuild that rewrote the source text stamps it alone, as a write.
      const afterIngest = readFileSync(ledger, "utf8")
      writeFileSync(path.join(dir, "sources", "agreement.txt"), "Clause 1.\n")
      expect(stampSourceRebuildLineage(dir, ["Agreement.docx"], "rebuild-1")).toBe(1)
      const rebuild = appended(afterIngest)
      expect(rebuild).toBe(`${JSON.stringify({
        entry_id: ids(rebuild)[0],
        ts: "2026-10-19T12:00:00.000Z",
        tier: "PERMANENT",
        artifact: "sources/agreement.txt",
        op: "write",
        run: "ingest:rebuild-1",
        stage: null,
        subagent: "ingest",
        content_sha256: sha256("Clause 1.\n"),
        bytes: Buffer.byteLength("Clause 1.\n"),
        supersedes: source,
        provenance: { origin: "source-rebuild", documents: ["Agreement.docx"] },
      })}\n`)
    } finally {
      setSystemTime()
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("a stamp reads the ledger once, however long it is", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "ingest-lineage-"))
    const ledger = path.join(dir, ".dochaus", "lineage.jsonl")
    const { readFileSync: read } = fs
    let ledgerReads = 0
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      mkdirSync(path.join(dir, "sources"), { recursive: true })
      writeFileSync(path.join(dir, ".dochaus", "legal.db"), "registry")
      writeFileSync(path.join(dir, "sources", "agreement.txt"), "text")
      stampIngestLineage(dir, ["Agreement.docx"], "corr-1", "mutation-1")
      writeFileSync(path.join(dir, "sources", "agreement.txt"), "rebuilt text")
      const spy = spyOn(fs, "readFileSync").mockImplementation(((file: fs.PathOrFileDescriptor, options?: never) => {
        if (file === ledger) ledgerReads++
        return read(file, options)
      }) as typeof fs.readFileSync)
      try {
        expect(stampIngestLineage(dir, ["Agreement.docx"], "corr-2", "mutation-2")).toBe(1)
        expect(stampSourceRebuildLineage(dir, ["Agreement.docx"], "rebuild-1")).toBe(0)
      } finally {
        spy.mockRestore()
      }
      expect(ledgerReads).toBe(2)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})

// The ledger's rows; a line a stop cut short is skipped, as the appender skips it.
function ledgerRows(dir: string): any[] {
  return readFileSync(path.join(dir, ".dochaus", "lineage.jsonl"), "utf8").split("\n").flatMap((line) => {
    try {
      return line.trim() ? [JSON.parse(line)] : []
    } catch {
      return []
    }
  })
}

// Artifacts on disk whose latest lineage row does not record their bytes.
function staleArtifacts(dir: string) {
  const latest = new Map(ledgerRows(dir).map((row) => [row.artifact, row.content_sha256]))
  return [
    ".dochaus/legal.db", "source-index.json", "source-map.json",
    ...[...new Bun.Glob("{source-chunks,sources}/*.txt").scanSync({ cwd: dir })].map((file) => file.replace(/\\/g, "/")),
  ]
    .filter((file) => existsSync(path.join(dir, file)))
    .filter((file) => latest.get(file) !== createHash("sha256").update(readFileSync(path.join(dir, file))).digest("hex"))
}
