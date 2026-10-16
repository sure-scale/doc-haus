import { describe, expect, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { appendLineageRow } from "../lib/lineage"
import { buildManifest, chainFor, currentEntry, entriesForRun, verifyDrift } from "../script/lineage"

const row = (overrides: Partial<Parameters<typeof appendLineageRow>[1]> = {}) =>
  ({
    tier: "VERSIONED",
    artifact: "spine.md",
    op: "write",
    run: "run_a",
    stage: "spine",
    subagent: "spine-builder",
    content_sha256: "hash",
    bytes: 10,
    supersedes: null,
    ...overrides,
  }) as const

describe("lineage cli (pure)", () => {
  test("chainFor returns artifact rows in append order", () => {
    const rows = [
      { entry_id: "a", ts: "1", ...row() },
      { entry_id: "b", ts: "2", ...row({ op: "edit", supersedes: "a" }) },
      { entry_id: "c", ts: "3", ...row({ artifact: "register.md", run: "run_a" }) },
    ] as any
    expect(chainFor(rows, "spine.md").map((r) => r.entry_id)).toEqual(["a", "b"])
  })

  test("currentEntry returns the last row for an artifact", () => {
    const rows = [
      { entry_id: "a", ts: "1", ...row() },
      { entry_id: "b", ts: "2", ...row({ supersedes: "a" }) },
    ] as any
    expect(currentEntry(rows, "spine.md")?.entry_id).toBe("b")
    expect(currentEntry(rows, "missing.md")).toBeUndefined()
  })

  test("entriesForRun filters by run", () => {
    const rows = [
      { entry_id: "a", ts: "1", ...row({ run: "run_a" }) },
      { entry_id: "b", ts: "2", ...row({ run: "run_b" }) },
    ] as any
    expect(entriesForRun(rows, "run_a").map((r) => r.entry_id)).toEqual(["a"])
  })

  test("verifyDrift flags unledgered and drifted files", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-cli-"))
    try {
      writeFileSync(path.join(dir, "spine.md"), "content") // ledgered below
      writeFileSync(path.join(dir, "rogue.md"), "external edit") // present, no row
      const rows = [
        { entry_id: "a", ts: "1", ...row({ content_sha256: "deadbeef".repeat(8), bytes: 7 }) },
      ] as any
      const drift = verifyDrift(dir, rows)
      // spine.md: ledgered but sha differs (live != deadbeef...) -> drifted
      expect(drift.drifted.map((d) => d.artifact)).toContain("spine.md")
      // rogue.md: present, no ledger row -> unledgered
      expect(drift.unledgered.map((d) => d.artifact)).toContain("rogue.md")
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("buildManifest classifies on-disk artifacts", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-cli-"))
    try {
      writeFileSync(path.join(dir, "spine.md"), "x")
      writeFileSync(path.join(dir, "matter.json"), "{}")
      const manifest = buildManifest(dir)
      const tiers = Object.fromEntries(manifest.map((m) => [m.artifact, m.tier]))
      expect(tiers["spine.md"]).toBe("VERSIONED")
      expect(tiers["matter.json"]).toBe("PERMANENT")
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
