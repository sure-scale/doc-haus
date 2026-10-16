import { openQuestions } from "../drafting/model"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import { activeStep } from "./activity"

// When the narrator writes. It summarises each checkpoint once (a stage that
// finished, failed or paused, the drafting plan, a question for the lawyer, the
// run settling) and, during a long stage, gives a progress update every few
// minutes while there is new activity to report. Pure, so the hook can call it
// on every tick and the tests can drive the clock.

/** Quiet time after a new checkpoint so a burst of stage changes is one update. */
export const NARRATION_DEBOUNCE_MS = 8_000
/** A progress update during a long stage, at most this often. */
export const NARRATION_PROGRESS_MS = 210_000
/** The least time between two updates while the run is working. */
export const NARRATION_MIN_GAP_MS = 30_000
/** Consecutive narrator failures after which the narration stops calling it. */
export const NARRATION_FAILURE_LIMIT = 3

export type NarrationPlan = {
  /** Every checkpoint this update covers; recorded with the prompt so a reload
   * never narrates the same checkpoint twice. */
  keys: string[]
  /** The newest checkpoint, or the active step for a progress update. */
  focus: string
  reason: "checkpoint" | "progress"
}

export function checkpointKeys(status: LeanDraftingStatus) {
  const plan = status.plan?.documents.length ? [`plan:${status.plan.documents.map((row) => row.id).join(",")}`] : []
  const stages = Object.entries(status.stages)
    .filter(([, row]) => row.status === "done" || row.status === "failed" || row.status === "blocked")
    .map(([step, row]) => `${step}:${row.status}:${row.finished_at ?? row.started_at ?? ""}`)
  const questions = openQuestions(status)
    .map((question) => `question:${question.id}`)
  const settled =
    status.status !== "running" || !status.in_flight
      ? [`run:${status.status === "running" ? "paused" : status.status}:${status.updated_at}`]
      : []
  return [...plan, ...stages, ...questions, ...settled]
}

export function planNarration(input: {
  status: LeanDraftingStatus
  narrated: ReadonlySet<string>
  /** When each checkpoint key was first seen by this view. */
  seenAt: ReadonlyMap<string, number>
  now: number
  inFlight: boolean
  failures: number
  /** When the last update was asked for; 0 before the first. */
  lastCallAt: number
  /** Activity lines seen since the last update. */
  activitySince: number
}): NarrationPlan | undefined {
  if (input.inFlight || input.failures >= NARRATION_FAILURE_LIMIT) return undefined
  const working = input.status.status === "running" && input.status.in_flight
  const pending = checkpointKeys(input.status).filter((key) => !input.narrated.has(key))
  if (pending.length) {
    const newest = Math.max(...pending.map((key) => input.seenAt.get(key) ?? input.now))
    if (input.now - newest < NARRATION_DEBOUNCE_MS) return undefined
    if (working && input.now - input.lastCallAt < NARRATION_MIN_GAP_MS) return undefined
    return { keys: pending, focus: pending.at(-1)!, reason: "checkpoint" }
  }
  const step = activeStep(input.status)
  if (!working || !step || input.activitySince === 0) return undefined
  const started = Date.parse(input.status.stages[step]?.started_at ?? "") || input.now
  const since = Math.max(input.lastCallAt, started)
  if (input.now - since < NARRATION_PROGRESS_MS) return undefined
  const window = Math.floor((input.now - started) / NARRATION_PROGRESS_MS)
  return { keys: [`${step}:progress:${started}:${window}`], focus: step, reason: "progress" }
}
