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
}

export function appendLeanEvent(directory: string, event: LeanEvent) {
  try {
    mkdirSync(leanDir(directory), { recursive: true, mode: 0o700 })
    appendFileSync(leanFile(directory, "events.jsonl"), `${JSON.stringify({ at: new Date().toISOString(), ...event })}\n`, { mode: 0o600 })
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
