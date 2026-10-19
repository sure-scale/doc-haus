import { describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { createHash } from "node:crypto"
import { tmpdir } from "node:os"
import path from "node:path"
import { stampIngestLineage, stampSourceRebuildLineage } from "./lineage"

describe("ingest lineage appender", () => {
  test("stamps one PERMANENT row per present artifact with supersession", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "ingest-lineage-"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      writeFileSync(path.join(dir, ".dochaus", "legal.db"), "")
      writeFileSync(path.join(dir, "source-index.json"), "{}")
      writeFileSync(path.join(dir, "source-map.json"), "{}")
      mkdirSync(path.join(dir, "source-chunks"), { recursive: true })
      writeFileSync(path.join(dir, "source-chunks", "agreement-c001.txt"), "chunk1")
      const first = stampIngestLineage(dir, ["Agreement.docx"], "corr-1")
      expect(first).toBeGreaterThan(0)
      const rows1 = readFileSync(path.join(dir, ".dochaus", "lineage.jsonl"), "utf8").trim().split("\n").map((row) => JSON.parse(row))
      expect(rows1.every((r: any) => r.tier === "PERMANENT" && r.op === "ingest" && r.subagent === "ingest" && r.run === "ingest:corr-1")).toBe(true)
      expect(rows1.some((r: any) => r.artifact === ".dochaus/legal.db")).toBe(true)
      expect(rows1.every((r: any) => r.supersedes === null)).toBe(true)

      // re-ingest: each artifact supersedes its prior entry
      const priorIds = new Map(rows1.map((r: any) => [r.artifact, r.entry_id]))
      const second = stampIngestLineage(dir, ["Agreement.docx"], "corr-2")
      expect(second).toBe(first)
      const rows2 = readFileSync(path.join(dir, ".dochaus", "lineage.jsonl"), "utf8").trim().split("\n").map((row) => JSON.parse(row))
      const latest = rows2.filter((r: any) => r.run === "ingest:corr-2")
      for (const r of latest) expect(r.supersedes).toBe(priorIds.get(r.artifact))
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
})
