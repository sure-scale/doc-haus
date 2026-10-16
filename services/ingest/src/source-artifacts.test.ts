import { describe, expect, test } from "bun:test"
import { existsSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { insertChunk, openDb, upsertDocument } from "./db"
import { writeSourceArtifacts } from "./source-artifacts"

describe("source artifacts", () => {
  test("writes compact source maps and refreshes stale generated files", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      const text = "ARTICLE 1\nAlpha obligations.\n\nARTICLE 2\nBeta obligations."
      const db = openDb(dir)
      const docPath = path.join(dir, "Agreement.docx")
      const documentId = upsertDocument(db, docPath, "Agreement.docx", Date.now(), null)
      insertChunk(db, {
        documentId,
        docPath,
        docName: "Agreement.docx",
        section: "1",
        structuralPath: "Agreement > Article 1",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      const first = await writeSourceArtifacts(dir)
      expect(first.documents).toBe(1)
      expect(first.chunks).toBeGreaterThan(0)
      expect(existsSync(path.join(dir, "source-index.json"))).toBe(true)
      expect(existsSync(path.join(dir, "source-map.md"))).toBe(true)
      const sourceMap = await Bun.file(path.join(dir, "source-map.json")).json() as {
        documents: Array<{ chunks: Array<{ text_path: string }> }>
      }
      expect(existsSync(path.join(dir, sourceMap.documents[0]!.chunks[0]!.text_path))).toBe(true)
      expect((await Bun.file(path.join(dir, "source-map.md")).text()).includes("Agreement.docx")).toBe(true)
      const sourceIndex = await Bun.file(path.join(dir, "source-index.json")).json() as {
        sources: Array<{ document: string }>
      }
      expect(sourceIndex.sources[0]!.document).toBe("Agreement.docx")
      const identifiersJson = await Bun.file(path.join(dir, "matter-identifiers.json")).text()
      const identifiersMarkdown = await Bun.file(path.join(dir, "matter-identifiers.md")).text()

      // #96: ingest stamps PERMANENT lineage rows.
      const ledger = path.join(dir, ".dochaus", "lineage.jsonl")
      expect(existsSync(ledger)).toBe(true)
      const rows = (await Bun.file(ledger).text()).trim().split("\n").map((l) => JSON.parse(l) as { tier: string; artifact: string; op: string; subagent: string })
      expect(rows.length).toBeGreaterThan(0)
      expect(rows.every((r) => r.tier === "PERMANENT" && r.op === "ingest" && r.subagent === "ingest")).toBe(true)
      expect(rows.some((r) => r.artifact === ".dochaus/legal.db")).toBe(true)

      await Bun.write(path.join(dir, "source-chunks", "stale.txt"), "stale")
      await Bun.sleep(5)
      await writeSourceArtifacts(dir)
      expect(existsSync(path.join(dir, "source-chunks", "stale.txt"))).toBe(false)
      expect(await Bun.file(path.join(dir, "matter-identifiers.json")).text()).toBe(identifiersJson)
      expect(await Bun.file(path.join(dir, "matter-identifiers.md")).text()).toBe(identifiersMarkdown)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("refreshes identifier metadata when the derived projection changes", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      const db = openDb(dir)
      const firstPath = path.join(dir, "First.docx")
      const firstText = "Morgan Ellis signed the first notice."
      const firstId = upsertDocument(db, firstPath, "First.docx", Date.now(), null)
      insertChunk(db, {
        documentId: firstId,
        docPath: firstPath,
        docName: "First.docx",
        section: "1",
        structuralPath: "First",
        party: "",
        chunkIndex: 0,
        text: firstText,
        charStart: 0,
        charEnd: firstText.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      await writeSourceArtifacts(dir)
      const before = await Bun.file(path.join(dir, "matter-identifiers.json")).json() as { generated_at: string; documents: number }
      await Bun.sleep(5)

      const updated = openDb(dir)
      const secondPath = path.join(dir, "Second.docx")
      const secondText = "Taylor Morgan signed the second notice."
      const secondId = upsertDocument(updated, secondPath, "Second.docx", Date.now(), null)
      insertChunk(updated, {
        documentId: secondId,
        docPath: secondPath,
        docName: "Second.docx",
        section: "1",
        structuralPath: "Second",
        party: "",
        chunkIndex: 0,
        text: secondText,
        charStart: 0,
        charEnd: secondText.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      updated.close()

      await writeSourceArtifacts(dir)
      const after = await Bun.file(path.join(dir, "matter-identifiers.json")).json() as { generated_at: string; documents: number }
      expect(after.documents).toBe(2)
      expect(after.generated_at).not.toBe(before.generated_at)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("writes matter-identifiers.json and matter-identifiers.md at ingest", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      // Authored synthetic procurement record; no real matter or personal data.
      const text = [
        "Procurement review meeting on March 17, 2026.",
        "Synthetic Case No. 00-TEST-0000001 was filed by Example Procurement Services.",
        "Coordinator: Taylor R. Morgan. Reviewer: Alex Morgan Bennett.",
        "Authority: 15 U.S.C. § 7001.",
      ].join("\n")
      const db = openDb(dir)
      const docPath = path.join(dir, "Synthetic Procurement.docx")
      const documentId = upsertDocument(db, docPath, "Synthetic Procurement.docx", Date.now(), null)
      insertChunk(db, {
        documentId,
        docPath,
        docName: "Synthetic Procurement.docx",
        section: "1",
        structuralPath: "Procurement",
        party: "",
        chunkIndex: 0,
        text,
        charStart: 0,
        charEnd: text.length,
        embedding: new Float32Array([1, 2, 3]),
        flagged: false,
      })
      db.close()

      await writeSourceArtifacts(dir)

      const identifiersJsonPath = path.join(dir, "matter-identifiers.json")
      const identifiersMdPath = path.join(dir, "matter-identifiers.md")
      expect(existsSync(identifiersJsonPath)).toBe(true)
      expect(existsSync(identifiersMdPath)).toBe(true)

      const parsed = await Bun.file(identifiersJsonPath).json()
      expect(parsed.schema_version).toBe(1)
      expect(parsed.rows.length).toBeGreaterThan(0)
      const values = parsed.rows.map((row: { value: string }) => row.value)
      expect(values).toEqual(expect.arrayContaining([
        expect.stringMatching(/00-TEST-0000001/),
        expect.stringMatching(/15 U\.S\.C\. § 7001/),
      ]))

      const md = await Bun.file(identifiersMdPath).text()
      expect(md).toContain("# Matter Identifiers")
      expect(md).toContain("case_reference")
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("stamps a legal.db PERMANENT lineage row even on a zero-document ingest", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "dochaus-source-artifacts-"))
    try {
      openDb(dir).close()
      const result = await writeSourceArtifacts(dir)
      expect(result.documents).toBe(0)
      const ledger = path.join(dir, ".dochaus", "lineage.jsonl")
      expect(existsSync(ledger)).toBe(true)
      const rows = (await Bun.file(ledger).text()).trim().split("\n").map((l) => JSON.parse(l) as { tier: string; artifact: string })
      expect(rows.some((r) => r.tier === "PERMANENT" && r.artifact === ".dochaus/legal.db")).toBe(true)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
