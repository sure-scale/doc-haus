// Matter knowledge lineage (#96): tier classification + append-only lineage ledger.
// Pure library — no engine/plugin dependency. Transfer-safe: only paths, run/stage
// ids, hashes, byte counts. No legal content.
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, statSync } from "node:fs"
import path from "node:path"
import { readDurableJson, writeDurableJson } from "./durable-json"
import { withArtifactMutationLock } from "../capabilities/artifacts/mutation-lock"
import {
  appendContainedRegularFile,
  readRegularFile,
  readRegularFileSnapshot,
} from "./regular-file"

export type Tier = "PERMANENT" | "VERSIONED" | "FRESH"
export type LineageOp = "write" | "edit" | "append" | "ingest" | "delete"

export type LineageRow = {
  entry_id: string
  ts: string
  tier: Tier
  artifact: string
  op: LineageOp
  run: string
  stage: string | null
  subagent: string
  content_sha256: string | null
  bytes: number | null
  supersedes: string | null
  provenance?: Record<string, unknown>
}

const lineageReadCache = new Map<string, { identity: string; rows: LineageRow[] }>()
const fileHashReadSnapshots: Array<Map<string, ReturnType<typeof sha256File>>> = []

const PERMANENT_PATTERNS = [
  /^source-chunks\//,
  /^sources\//,
  /^source-index\.json$/,
  /^source-map\.json$/,
  /^source-map\.md$/,
  /^source-ledger\.json$/,
  /^matter-identifiers\.json$/,
  /^matter-identifiers\.md$/,
  /^matter\.json$/,
  /^\.dochaus\/legal\.db$/,
]

const FRESH_PATTERNS = [
  /^\.dochaus\/grid\.json$/,
  /^\.dochaus\/telemetry\.db$/,
  /^\.dochaus\/context-budget\.jsonl$/,
  /^\.dochaus\/lineage\.jsonl$/,
  /^\.dochaus\/run-trace\.jsonl$/,
  /^\.dochaus-run-plan-state\.json$/,
  /^\.dochaus-run-state\.json$/,
  /^\.dochaus-run-initialization\.json$/,
  /^\.dochaus-deliverable\.json$/,
  /^\.dochaus-fanout-completions\.json$/,
  /^\.dochaus-handoff-blocks\.json$/,
  /^\.dochaus-degraded-stages\.json$/,
  /^\.dochaus-completed-stages\.json$/,
  /^\.dochaus-stage-state\.json$/,
  /^\.dochaus-workflow-frontier\.json$/,
  /^decision-directions\//,
  /^material-decision-directions-[a-f0-9]{64}\.md$/,
]

export function resolveTier(artifact: string): Tier {
  const norm = artifact.replace(/^\.\//, "").replace(/\\/g, "/")
  if (PERMANENT_PATTERNS.some((re) => re.test(norm))) return "PERMANENT"
  if (FRESH_PATTERNS.some((re) => re.test(norm))) return "FRESH"
  if (norm.startsWith("section-") && norm.endsWith(".md")) return "VERSIONED"
  if (norm.startsWith(".dochaus/")) return "FRESH"
  return "VERSIONED"
}

export function lineageFile(directory: string): string {
  return path.join(directory, ".dochaus", "lineage.jsonl")
}

export function readLineage(directory: string): LineageRow[] {
  const file = lineageFile(directory)
  try {
    const cached = lineageReadCache.get(file)
    const snapshot = readRegularFileSnapshot(file, cached?.identity)
    if (!snapshot) {
      lineageReadCache.delete(file)
      return []
    }
    if (!snapshot.content) return cached?.rows ?? []
    const rows = snapshot.content.toString("utf8")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => parseLineageRow(line))
      .filter((row): row is LineageRow => row !== null)
    lineageReadCache.set(file, { identity: snapshot.identity, rows })
    return rows
  } catch {
    lineageReadCache.delete(file)
    return []
  }
}

function parseLineageRow(line: string): LineageRow | null {
  try {
    const parsed = JSON.parse(line) as unknown
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return null
    const row = parsed as Record<string, unknown>
    if (typeof row.entry_id !== "string") return null
    return row as unknown as LineageRow
  } catch {
    return null
  }
}

export function lastEntryFor(rows: readonly LineageRow[], artifact: string): LineageRow | undefined {
  for (let i = rows.length - 1; i >= 0; i--) if (rows[i]!.artifact === artifact) return rows[i]
  return undefined
}

export function appendLineageRow(directory: string, partial: Omit<LineageRow, "entry_id" | "ts">): LineageRow {
  return appendLineageRowExact(directory, partial, { entry_id: ulid(), ts: new Date().toISOString() })
}

export function appendLineageRowExact(
  directory: string,
  partial: Omit<LineageRow, "entry_id" | "ts">,
  identity: Pick<LineageRow, "entry_id" | "ts">,
): LineageRow {
  if (
    !/^[0-9A-HJKMNP-TV-Z]{26}$/.test(identity.entry_id) ||
    !Number.isFinite(Date.parse(identity.ts))
  ) throw new Error("Exact lineage identity is malformed.")
  const row: LineageRow = { ...partial, ...identity }
  withArtifactMutationLock(directory, ".dochaus/lineage.jsonl", () => {
    mkdirSync(path.join(directory, ".dochaus"), { recursive: true })
    appendContainedRegularFile(directory, ".dochaus/lineage.jsonl", JSON.stringify(row) + "\n")
    lineageReadCache.delete(lineageFile(directory))
  })
  return row
}

export function sha256File(absPath: string): { sha256: string; bytes: number } | null {
  const snapshot = fileHashReadSnapshots.at(-1)
  if (snapshot?.has(absPath)) return snapshot.get(absPath) ?? null
  const content = readRegularFile(absPath)
  if (!content) {
    snapshot?.set(absPath, null)
    return null
  }
  const value = { sha256: createHash("sha256").update(content).digest("hex"), bytes: content.length }
  snapshot?.set(absPath, value)
  return value
}

export function withFileHashReadSnapshot<T>(callback: () => T) {
  // Pure synchronous derivations need a coherent file view and must not hash
  // the same artifact repeatedly. The scope ends before any asynchronous work.
  if (fileHashReadSnapshots.length) return callback()
  fileHashReadSnapshots.push(new Map())
  try {
    return callback()
  } finally {
    fileHashReadSnapshots.pop()
  }
}

// --- run id ---
type RunState = { run_id: string; started_at: string; owner_session_id?: string }

export function runStateFile(directory: string): string {
  return path.join(directory, ".dochaus-run-state.json")
}

export function createRunId() {
  return `run_${ulid()}`
}

export function persistRunId(directory: string, run_id: string, owner_session_id?: string) {
  writeDurableJson(runStateFile(directory), {
    run_id,
    started_at: new Date().toISOString(),
    ...(owner_session_id ? { owner_session_id } : {}),
  } satisfies RunState)
  return run_id
}

// Unconditionally mint + persist a fresh run_id (overwrites any prior run-state).
// Called by the plugin when a drafting run kickoff is detected, so every VERSIONED
// write in the new run — including the kickoff checkpoint itself — carries the new id.
export function mintRunId(directory: string, owner_session_id?: string): string {
  const run_id = createRunId()
  try {
    persistRunId(directory, run_id, owner_session_id)
  } catch {
    /* best-effort; the returned id still groups this run's writes */
  }
  return run_id
}

// v1: mint on first drafting activity (run-state absent), reuse thereafter. A second
// drafting run on an already-drafted matter shares the run_id in v1 — documented
// limitation; the per-artifact `supersedes` chain remains exact regardless.
export function resolveRunId(directory: string): string {
  const file = runStateFile(directory)
  const current = readRunId(directory)
  if (current) return current
  const run_id = createRunId()
  try {
    writeDurableJson(file, { run_id, started_at: new Date().toISOString() } satisfies RunState)
  } catch {
    /* best-effort; the returned id still groups this process's writes */
  }
  return run_id
}

export function readRunId(directory: string): string | undefined {
  const parsed = readDurableJson(runStateFile(directory))
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return
  return typeof (parsed as RunState).run_id === "string" ? (parsed as RunState).run_id : undefined
}

export function readRunOwnerSessionID(directory: string): string | undefined {
  const parsed = readDurableJson(runStateFile(directory))
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return
  const state = parsed as RunState
  if (state.owner_session_id?.trim()) return state.owner_session_id
  const trace = path.join(directory, ".dochaus-run-trace.jsonl")
  if (!state.run_id || !existsSync(trace)) return
  for (const line of readFileSync(trace, "utf8").split(/\r?\n/)) {
    if (!line.trim()) continue
    try {
      const row = JSON.parse(line) as Record<string, unknown>
      if (row.run_id !== state.run_id || row.agent !== "draft" || typeof row.sessionID !== "string" || !row.sessionID.trim())
        continue
      return row.sessionID
    } catch {
      continue
    }
  }
}

// --- ulid: 10-char Crockford-base32 ms timestamp + 16-char random (26 chars) ---
const CROCKFORD = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"

export function ulid(): string {
  const ts = Date.now()
  let time = ""
  let t = ts
  for (let i = 0; i < 10; i++) {
    time = CROCKFORD[t % 32] + time
    t = Math.floor(t / 32)
  }
  const randomBytes = new Uint8Array(16)
  ;(globalThis.crypto as Crypto).getRandomValues(randomBytes)
  let random = ""
  for (let i = 0; i < 16; i++) random += CROCKFORD[randomBytes[i]! % 32]!
  return time + random
}
