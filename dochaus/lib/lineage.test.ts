import { describe, expect, test } from "bun:test"
import { existsSync, lstatSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  appendLineageRow,
  lastEntryFor,
  lineageFile,
  mintRunId,
  readLineage,
  readRunOwnerSessionID,
  resolveRunId,
  resolveTier,
  sha256File,
  ulid,
} from "./lineage"

describe("lineage lib", () => {
  test("resolveTier classifies every known pattern and the unknown defaults", () => {
    const cases: [string, "PERMANENT" | "VERSIONED" | "FRESH"][] = [
      ["source-chunks/agreement-c001.txt", "PERMANENT"],
      ["sources/agreement.txt", "PERMANENT"],
      ["source-index.json", "PERMANENT"],
      ["source-map.json", "PERMANENT"],
      ["source-map.md", "PERMANENT"],
      ["source-ledger.json", "PERMANENT"],
      ["matter-identifiers.json", "PERMANENT"],
      ["matter-identifiers.md", "PERMANENT"],
      ["matter.json", "PERMANENT"],
      [".dochaus/legal.db", "PERMANENT"],
      [".dochaus/grid.json", "FRESH"],
      [".dochaus/telemetry.db", "FRESH"],
      [".dochaus/lineage.jsonl", "FRESH"],
      [".dochaus-run-plan-state.json", "FRESH"],
      ["spine.md", "VERSIONED"],
      ["register.md", "VERSIONED"],
      ["section-001.md", "VERSIONED"],
      ["arbitrary-run-output.md", "VERSIONED"],
      [".dochaus/unknown-cache.json", "FRESH"],
    ]
    for (const [artifact, tier] of cases) expect(resolveTier(artifact)).toBe(tier)
  })

  test("ulid is unique and 26 chars across a burst", () => {
    const ids = new Set<string>()
    for (let i = 0; i < 1000; i++) ids.add(ulid())
    expect(ids.size).toBe(1000)
    for (const id of ids) expect(id).toHaveLength(26)
  })

  test("appendLineageRow writes JSONL and readLineage parses it back", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      const r1 = appendLineageRow(dir, {
        tier: "VERSIONED", artifact: "spine.md", op: "write", run: "run_x",
        stage: "spine", subagent: "spine-builder", content_sha256: "a", bytes: 1, supersedes: null,
      })
      expect(r1.entry_id).toHaveLength(26)
      expect(r1.supersedes).toBeNull()
      expect(existsSync(lineageFile(dir))).toBe(true)

      const r2 = appendLineageRow(dir, {
        tier: "VERSIONED", artifact: "spine.md", op: "edit", run: "run_x",
        stage: "spine-amend", subagent: "spine-builder", content_sha256: "b", bytes: 2, supersedes: r1.entry_id,
      })
      const rows = readLineage(dir)
      expect(rows).toHaveLength(2)
      expect(rows[1].entry_id).toBe(r2.entry_id)
      expect(lastEntryFor(rows, "spine.md")?.entry_id).toBe(r2.entry_id)
      expect(lastEntryFor(rows, "missing.md")).toBeUndefined()
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("repeated create plus second append stays anchored across many matters", () => {
    // Regression: raw Darwin variadic openat with O_CREAT produced nondeterministic
    // mode bits (e.g. 0040), so the second append's reopen returned EACCES. The fix
    // opens non-creating first, creates with O_CREAT|O_EXCL + fchmod 0600, and
    // verifies dev/ino with the same non-creating append flags. Stress both the
    // create path (first append) and the existing-file path (second append) across
    // independent matters, and assert the created mode is exactly 0600.
    const iterations = 40
    for (let i = 0; i < iterations; i++) {
      const dir = mkdtempSync(path.join(tmpdir(), "lineage-stress-"))
      try {
        const first = appendLineageRow(dir, {
          tier: "VERSIONED", artifact: "spine.md", op: "write", run: `run_${i}`,
          stage: "spine", subagent: "spine-builder", content_sha256: "a", bytes: 1, supersedes: null,
        })
        const second = appendLineageRow(dir, {
          tier: "VERSIONED", artifact: "spine.md", op: "edit", run: `run_${i}`,
          stage: "spine-amend", subagent: "spine-builder", content_sha256: "b", bytes: 2, supersedes: first.entry_id,
        })
        const rows = readLineage(dir)
        expect(rows).toHaveLength(2)
        expect(rows[1]!.entry_id).toBe(second.entry_id)
        expect(lastEntryFor(rows, "spine.md")?.entry_id).toBe(second.entry_id)
        expect(statSync(lineageFile(dir)).mode & 0o777).toBe(0o600)
      } finally {
        rmSync(dir, { recursive: true, force: true })
      }
    }
  })

  test("readLineage returns empty after a cached file is removed", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      writeFileSync(lineageFile(dir), JSON.stringify({ entry_id: "x".repeat(26) }) + "\n")
      expect(readLineage(dir)).toHaveLength(1)

      rmSync(lineageFile(dir))

      expect(readLineage(dir)).toEqual([])
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("readLineage tolerates a corrupt line", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      writeFileSync(lineageFile(dir), "not json\n" + JSON.stringify({ entry_id: "x".repeat(26) }) + "\n")
      const rows = readLineage(dir)
      expect(rows).toHaveLength(1)
      expect(rows[0].entry_id).toBe("x".repeat(26))
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("readLineage rejects a FIFO within 7.5s and preserves it", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-read-fifo-"))
    let child: ReturnType<typeof spawnReadLineageChild> | undefined
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      const fifo = lineageFile(dir)
      expect(Bun.spawnSync(["mkfifo", fifo]).exitCode).toBe(0)

      child = spawnReadLineageChild(dir)

      expect(await boundedChildExit(child.process)).toBe(child.acceptedExitCode)
      expect(lstatSync(fifo).isFIFO()).toBe(true)
    } finally {
      if (child?.process.exitCode === null) child.process.kill(9)
      if (child) await child.process.exited
      rmSync(dir, { recursive: true, force: true })
    }
  }, 10_000)

  test("readLineage rejects a terminal symlink to an external sentinel without reading it", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-read-symlink-"))
    const external = mkdtempSync(path.join(tmpdir(), "lineage-external-sentinel-"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      const sentinel = path.join(external, "lineage.jsonl")
      const contents = JSON.stringify({ entry_id: "x".repeat(26) }) + "\n"
      writeFileSync(sentinel, contents)
      const link = lineageFile(dir)
      symlinkSync(sentinel, link)

      expect(readLineage(dir)).toEqual([])
      expect(readFileSync(sentinel, "utf8")).toBe(contents)
      expect(lstatSync(link).isSymbolicLink()).toBe(true)
    } finally {
      rmSync(dir, { recursive: true, force: true })
      rmSync(external, { recursive: true, force: true })
    }
  })

  test("sha256File returns hash+bytes or null when absent", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      const f = path.join(dir, "f.txt")
      writeFileSync(f, "hello")
      const h = sha256File(f)
      expect(h?.bytes).toBe(5)
      expect(h?.sha256).toHaveLength(64)
      expect(sha256File(path.join(dir, "nope"))).toBeNull()
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("sha256File fails closed promptly on a FIFO without changing its type", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-fifo-"))
    let child: ReturnType<typeof spawnSha256FileChild> | undefined
    try {
      const fifo = path.join(dir, "hash-target")
      expect(Bun.spawnSync(["mkfifo", fifo]).exitCode).toBe(0)
      child = spawnSha256FileChild(dir, fifo)

      expect(await boundedChildExit(child.process)).toBe(
        child.rejectedExitCode,
      )
      expect(lstatSync(fifo).isFIFO()).toBe(true)
    } finally {
      if (child?.process.exitCode === null) child.process.kill(9)
      if (child) await child.process.exited
      rmSync(dir, { recursive: true, force: true })
    }
  }, 20_000)

  test("lineage append rejects a FIFO within 7.5s and preserves it", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-append-fifo-"))
    let child: ReturnType<typeof spawnAppendFifoChild> | undefined
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      const fifo = lineageFile(dir)
      expect(Bun.spawnSync(["mkfifo", fifo]).exitCode).toBe(0)

      child = spawnAppendFifoChild(dir)

      expect(await boundedChildExit(child.process)).toBe(child.rejectedExitCode)
      expect(lstatSync(fifo).isFIFO()).toBe(true)
    } finally {
      if (child?.process.exitCode === null) child.process.kill(9)
      if (child) await child.process.exited
      rmSync(dir, { recursive: true, force: true })
    }
  }, 10_000)

  test("appendLineageRow rejects a terminal symlink and preserves the sentinel", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-append-symlink-"))
    try {
      mkdirSync(path.join(dir, ".dochaus"), { recursive: true })
      const sentinel = path.join(dir, "sentinel.txt")
      writeFileSync(sentinel, "sentinel")
      const link = lineageFile(dir)
      symlinkSync(sentinel, link)

      const before = readFileSync(sentinel, "utf8")
      expect(() =>
        appendLineageRow(dir, {
          tier: "VERSIONED", artifact: "spine.md", op: "write", run: "run_symlink",
          stage: "spine", subagent: "spine-builder", content_sha256: "a", bytes: 1, supersedes: null,
        }),
      ).toThrow()
      expect(readFileSync(sentinel, "utf8")).toBe(before)
      expect(lstatSync(link).isSymbolicLink()).toBe(true)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("resolveRunId mints once and reuses", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      const a = resolveRunId(dir)
      const b = resolveRunId(dir)
      expect(a).toBe(b)
      expect(a.startsWith("run_")).toBe(true)
      const raw = JSON.parse(readFileSync(path.join(dir, ".dochaus-run-state.json"), "utf8"))
      expect(raw.run_id).toBe(a)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("mintRunId persists and reads the exact owner session", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      const run = mintRunId(dir, "ses_exact_owner")
      writeFileSync(
        path.join(dir, ".dochaus-run-trace.jsonl"),
        JSON.stringify({ run_id: run, agent: "draft", sessionID: "ses_competing_owner" }) + "\n",
      )

      expect(readRunOwnerSessionID(dir)).toBe("ses_exact_owner")
      expect(JSON.parse(readFileSync(path.join(dir, ".dochaus-run-state.json"), "utf8"))).toMatchObject({
        run_id: run,
        owner_session_id: "ses_exact_owner",
      })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("legacy owner discovery accepts only a current-run draft trace row without mutating run state", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      writeFileSync(
        path.join(dir, ".dochaus-run-state.json"),
        JSON.stringify({ run_id: "run_current", started_at: "2026-07-12T00:00:00.000Z" }),
      )
      writeFileSync(
        path.join(dir, ".dochaus-run-trace.jsonl"),
        [
          { run_id: "run_prior", agent: "draft", sessionID: "ses_prior_owner" },
          { run_id: "run_current", agent: "research-planner", sessionID: "ses_child" },
          { run_id: "run_current", agent: "draft", sessionID: "ses_current_owner" },
        ].map((row) => JSON.stringify(row)).join("\n") + "\n",
      )

      const before = readFileSync(path.join(dir, ".dochaus-run-state.json"), "utf8")
      expect(readRunOwnerSessionID(dir)).toBe("ses_current_owner")
      expect(readFileSync(path.join(dir, ".dochaus-run-state.json"), "utf8")).toBe(before)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("legacy owner discovery never adopts a trace row from a prior run", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "lineage-lib-"))
    try {
      const newer = { run_id: "run_newer", started_at: "2026-07-12T01:00:00.000Z" }
      writeFileSync(path.join(dir, ".dochaus-run-state.json"), JSON.stringify(newer))
      writeFileSync(
        path.join(dir, ".dochaus-run-trace.jsonl"),
        JSON.stringify({ run_id: "run_prior", agent: "draft", sessionID: "ses_prior_owner" }) + "\n",
      )

      expect(readRunOwnerSessionID(dir)).toBeUndefined()
      expect(JSON.parse(readFileSync(path.join(dir, ".dochaus-run-state.json"), "utf8"))).toEqual(newer)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})

function spawnSha256FileChild(directory: string, file: string) {
  const script = path.join(directory, "sha256-file-child.ts")
  const rejectedExitCode = 89
  writeFileSync(
    script,
    `import { sha256File } from ${JSON.stringify(path.join(import.meta.dir, "lineage.ts"))}\n` +
      `try {\n` +
      `  sha256File(process.argv[2]!)\n` +
      `} catch {\n` +
      `  process.exit(${rejectedExitCode})\n` +
      `}\n` +
      `process.exit(90)\n`,
  )
  return {
    process: Bun.spawn([
      process.execPath,
      script,
      file,
    ], { stderr: "inherit" }),
    rejectedExitCode,
  }
}

function spawnReadLineageChild(directory: string) {
  const script = path.join(directory, "read-lineage-child.ts")
  const acceptedExitCode = 0
  writeFileSync(
    script,
    `import { readLineage } from ${JSON.stringify(path.join(import.meta.dir, "lineage.ts"))}\n` +
      `process.exit(readLineage(process.argv[2]!).length === 0 ? ${acceptedExitCode} : 90)\n`,
  )
  return {
    process: Bun.spawn([
      process.execPath,
      script,
      directory,
    ], { stderr: "inherit" }),
    acceptedExitCode,
  }
}

function spawnAppendFifoChild(directory: string) {
  const script = path.join(directory, "append-fifo-child.ts")
  const rejectedExitCode = 89
  writeFileSync(
    script,
    `import { appendLineageRow } from ${JSON.stringify(path.join(import.meta.dir, "lineage.ts"))}\n` +
      `const row = {\n` +
      `  tier: "VERSIONED",\n` +
      `  artifact: "agreement.md",\n` +
      `  op: "write",\n` +
      `  run: "run_fifo",\n` +
      `  stage: null,\n` +
      `  subagent: "drafter",\n` +
      `  content_sha256: ${JSON.stringify("a".repeat(64))},\n` +
      `  bytes: 1,\n` +
      `  supersedes: null,\n` +
      `} as const\n` +
      `try {\n` +
      `  appendLineageRow(process.argv[2]!, row)\n` +
      `} catch {\n` +
      `  process.exit(${rejectedExitCode})\n` +
      `}\n` +
      `process.exit(90)\n`,
  )
  return {
    process: Bun.spawn([
      process.execPath,
      script,
      directory,
    ], { stderr: "inherit" }),
    rejectedExitCode,
  }
}

function boundedChildExit(
  process: Bun.Subprocess,
) {
  return new Promise<number>((resolve, reject) => {
    const timeout = setTimeout(
      () => reject(new Error("Timed out waiting for bounded child.")),
      7_500,
    )
    process.exited.then(
      (code) => {
        clearTimeout(timeout)
        resolve(code)
      },
      (error) => {
        clearTimeout(timeout)
        reject(error)
      },
    )
  })
}
