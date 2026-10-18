import { appendFileSync, existsSync, mkdirSync, readFileSync } from "node:fs"
import type { DocID } from "./plan"
import { leanDir, leanFile, type StageID } from "./state"

// Record-only event log of a lean drafting run: one JSON line per stage event in
// .dochaus/lean/events.jsonl. Nothing reads it to decide anything. Appending never
// throws, so a full disk or a removed directory cannot block or fail a stage.

export type LeanEventName =
  | "started"
  | "skipped"
  | "model_call"
  | "checker"
  | "retry"
  | "done"
  | "blocked"
  | "failed"
  | "delivered"
  | "restored"
  | "request_received"
  | "request_processed"
  | "paused"
  | "stopped"
  /** The draft step's recorded plan coverage: its status as `outcome`, its issues. */
  | "coverage"
  /** The draft step's recorded requested files: whether the lawyer's words state
   * them as `outcome` ("stated" or "not_stated"), the mismatches the outline
   * still leaves as `issues`, and the file names given to the plan as `reason`;
   * or "unchecked", with no record and the reason, when the files job's reply
   * never passed its check. */
  | "requested_files"

export type LeanEvent = {
  stage: StageID | "run"
  /** The planned document a per-document step works on. */
  document?: DocID
  event: LeanEventName
  /** Which call within the stage, for stages that make more than one. */
  step?: string
  attempt?: number
  session_id?: string
  issues?: string[]
  reason?: string
  outputs?: string[]
  /** Lawyer request events: the queued request's id, its kind and how it was processed. */
  request_id?: string
  kind?: string
  outcome?: string
  /** Checker events: milliseconds the attempt awaited the model's reply; for a
   * prefetched review, only the wait left when its step began. Timing only. */
  wait_ms?: number
  /** Milliseconds of deterministic work: on checker events the reply's check, on
   * started and skipped events keying the step and deciding whether to skip it. */
  check_ms?: number
}

// The latest event time per matter in this process, so status can report when
// a running coordinator last made progress without reading the log back.
const lastEventAt = new Map<string, string>()

export function lastLeanEventAt(directory: string) {
  return lastEventAt.get(directory)
}

/** Milliseconds since `from` (a performance.now() reading), to 0.1 ms. */
export function leanElapsedMs(from: number, to = performance.now()) {
  return Math.round(Math.max(0, to - from) * 10) / 10
}

export function appendLeanEvent(directory: string, event: LeanEvent) {
  const at = new Date().toISOString()
  lastEventAt.set(directory, at)
  try {
    mkdirSync(leanDir(directory), { recursive: true, mode: 0o700 })
    appendFileSync(leanFile(directory, "events.jsonl"), `${JSON.stringify({ at, ...event })}\n`, { mode: 0o600 })
  } catch {
    // Diagnostic record only; losing a line must never affect the run.
  }
}

/** Reads the log back for diagnostics and tests. Unparseable lines are skipped. */
export function readLeanEvents(directory: string): (LeanEvent & { at: string })[] {
  const file = leanFile(directory, "events.jsonl")
  if (!existsSync(file)) return []
  return readFileSync(file, "utf8").split("\n").filter(Boolean).flatMap(line => {
    try {
      return [JSON.parse(line)]
    } catch {
      return []
    }
  })
}
