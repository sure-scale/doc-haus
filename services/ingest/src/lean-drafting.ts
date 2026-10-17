import { checklistRefinableParents } from "../../../dochaus/workflows/lean/checklist-flow"
import { draftingRunStopped, draftingUpdateAdmission } from "../../../dochaus/lib/drafting-update-admission"
import { mkdirSync, renameSync } from "node:fs"
import path from "node:path"
import { Hono } from "hono"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { researchStatus } from "../../../dochaus/workflows/lean/research"
import { readLeanDocuments, type LeanDocuments } from "../../../dochaus/workflows/lean/deliverables"
import { sha256File } from "../../../dochaus/lib/lineage"
import { lastLeanEventAt } from "../../../dochaus/workflows/lean/events"
import { readLeanPlan } from "../../../dochaus/workflows/lean/plan"
import { enqueueLeanRequest, pendingLeanRequests, leanUserUpdates, queuedAnswerIDs, isChecklistRefinement } from "../../../dochaus/workflows/lean/requests"
import { runLean, stageInputSha, stopLeanRun } from "../../../dochaus/workflows/lean/runner"
import { leanSteps } from "../../../dochaus/workflows/lean/plan"
import { withSourceSnapshot } from "../../../dochaus/workflows/lean/quotes"
import { engineSessionPort, type LeanDeps } from "../../../dochaus/workflows/lean/stages"
import {
  LeanQuestionError,
  currentOpenItems,
  initLeanState,
  leanDir,
  leanStateExists,
  matterRelative,
  readLeanJson,
  readLeanState,
  readLeanText,
  readQuestions,
  validateAnswerText,
  writeLeanState,
  type LeanModel,
  type LeanState,
} from "../../../dochaus/workflows/lean/state"
import { registerLeanDeliverable } from "./lean-deliverables"

// Web wiring for the lean drafting profile. Start creates the owner session and
// the lean state, then runs the coordinator in the background. Continue and
// answers only enqueue a request (202): the coordinator is the single writer of
// the run and drains the queue itself, so a request never waits on, or races,
// the stage in flight. One coordinator per matter directory runs at a time in
// this process; a request that arrives as it finishes starts it again. A new
// request on a matter whose run completed or failed archives that run and
// starts a fresh one; a blocked or paused run is continued, not replaced.
// Pause is the one request that does not queue: it stops the coordinator at
// once (see runLean), and Continue later resumes from the unfinished step.
// Stop settles a paused or blocked run that nothing is working on as failed: it
// takes no further Continue, answer or update, and the next drafting request
// archives it like any failed run. Nothing in the matter is deleted.

export type DraftingProfile = "lean"

export class LeanDraftingConflictError extends Error {
  constructor(message: string, readonly code: "lean_run_in_progress" | "lean_run_awaiting_continue" | "lean_run_stopped" | "drafting_profile_conflict") {
    super(message)
  }
}

const running = new Set<string>()
/** The coordinator working on each matter, and how to pause it. */
type Coordinator = { pause: AbortController; stopped: Promise<void>; queuedAtPause?: Set<string> }
const coordinators = new Map<string, Coordinator>()

/** The requested drafting profile: lean is the only profile; undefined when
 * any other profile is named. */
export function draftingStartProfile(value: unknown): DraftingProfile | undefined {
  if (value === undefined || value === null || value === "lean") return "lean"
  return undefined
}

export async function startLeanDraftingRun(
  directory: string,
  input: { request: string; agent: string; model: LeanModel },
  deps: LeanDeps = defaultDeps(),
) {
  if (running.has(directory)) throw new LeanDraftingConflictError("A lean drafting run is already in progress for this matter.", "lean_run_in_progress")
  const previous = readLeanState(directory)
  // "running" with no coordinator here is a run an ingest restart paused.
  if (previous && previous.status !== "completed" && previous.status !== "failed")
    throw new LeanDraftingConflictError(
      "This matter's lean drafting run has not finished; continue it or answer its questions instead of starting another.",
      "lean_run_awaiting_continue",
    )
  if (!previous && leanStateExists(directory))
    throw new LeanDraftingConflictError("This matter's lean drafting state cannot be read.", "drafting_profile_conflict")
  running.add(directory)
  // No coordinator runs for this matter (checked above, claimed here), so
  // nothing else writes the finished run while it is moved aside.
  if (previous) archiveLeanRun(directory, previous)
  const ownerSessionID = await deps.sessions.createSession({
    directory, agent: input.agent, title: "Lean drafting", model: input.model, metadata: { dochaus_role: "lean-drafting-owner" },
  }).catch((error: unknown) => {
    running.delete(directory)
    throw error
  })
  initLeanState(directory, { request: input.request, ownerSessionID, model: input.model })
  launch(directory, deps)
  return { profile: "lean" as const, mode: "started" as const, ownerSessionID, status: leanDraftingStatus(directory) }
}

/** Continue: queues a request for the coordinator and starts it if it is not
 * running. On a blocked gate the coordinator starts one repair round; otherwise
 * it resumes, skipping stages whose inputs are unchanged. */
export function resumeLeanDraftingRun(directory: string, deps: LeanDeps = defaultDeps()) {
  if (!leanStateExists(directory)) return undefined
  refuseStopped(readLeanState(directory))
  const request = enqueueLeanRequest(directory, { kind: "continue" })
  if (!running.has(directory)) launch(directory, deps)
  return { profile: "lean" as const, mode: "queued" as const, request_id: request.id, status: leanDraftingStatus(directory) }
}

/** Queues the lawyer's answer to an open question. The coordinator records it
 * at its next stage boundary and re-runs the stages that consume answers. */
export function answerLeanQuestion(directory: string, questionID: string, answer: unknown, deps: LeanDeps = defaultDeps()) {
  if (!leanStateExists(directory)) return undefined
  refuseStopped(readLeanState(directory))
  const text = validateAnswerText(typeof answer === "string" ? answer : "")
  if (!readQuestions(directory).some(question => question.id === questionID))
    throw new LeanQuestionError("No such question in this matter's lean drafting run.", "not_found")
  const request = enqueueLeanRequest(directory, { kind: "answer", question_id: questionID, answer: text })
  if (!running.has(directory)) launch(directory, deps)
  return { profile: "lean" as const, mode: "queued" as const, request_id: request.id, status: leanDraftingStatus(directory) }
}

/** One synchronous read of the run: every step key it takes shares one read of
 * the matter's sources and processed requests. */
export function leanDraftingStatus(directory: string) {
  return withSourceSnapshot(() => readLeanDraftingStatus(directory))
}

function readLeanDraftingStatus(directory: string) {
  const state = readLeanState(directory)
  if (!state) return undefined
  const refinable = checklistRefinableParents(directory)
  const gate = readLeanJson(directory, "gate.json") ?? null
  const documents = readLeanDocuments(directory) ?? null
  const event = lastLeanEventAt(directory)
  return {
    ...state,
    in_flight: running.has(directory),
    // The later of the last state write and this process's last run event.
    last_progress_at: event && event > state.updated_at ? event : state.updated_at,
    pending_requests: pendingLeanRequests(directory).length,
    queued_answer_ids: [...queuedAnswerIDs(directory)],
    updates: draftingUpdates(directory, state),
    open_items: currentOpenItems(directory),
    questions: readQuestions(directory),
    gate,
    documents,
    changed_after_check: changedAfterCheck(directory, gate, documents),
    plan: plannedDocuments(directory),
    research: researchStatus(directory),
    checklist_refinement: {
      available: refinable.length > 0 && !running.has(directory) && !pendingLeanRequests(directory).length && state.status === "blocked" && !state.paused_at && state.stages["checklist-check"]?.status === "blocked",
      parents: refinable,
    },
    // The lawyer's own request, so the chat can show it again after a reload.
    request: readLeanText(directory, "request.md")?.trim() ?? null,
  }
}

/** Admission is not application. Only a completed, current final check can
 * substantiate that saved instructions have made it through this run. */
function draftingUpdates(directory: string, state: LeanState) {
  const updates = leanUserUpdates(directory)
  if (!updates.length) return updates
  const included = updates.some(update => update.status === "included")
  const stale = included && leanSteps(directory).some(step => {
    const saved = state.stages[step]
    return saved?.status === "done" && saved.input_sha !== stageInputSha(directory, step)
  })
  const final = state.stages["package-review"]
  const applied = included && !stale && state.status === "completed" && final?.status === "done" &&
    final.input_sha === stageInputSha(directory, "package-review")
  return updates.map(update => ({ ...update, progress: update.status === "queued" ? "received" as const
    : applied ? "applied" as const
    : running.has(directory) && state.status === "running" ? "applying" as const
    : stale ? "requires-rework" as const : undefined }))
}

/** Delivered drafts whose matter file no longer holds the bytes the final check
 * saw: a change accepted after delivery, or a newer output kept from replacing
 * such a change. Disclosure only: the check is not re-run and its verdict is
 * passed through unchanged, so the lawyer is told to review the draft. */
function changedAfterCheck(directory: string, gate: unknown, documents: LeanDocuments | null) {
  if (documents?.status !== "completed") return null
  const changed = documents.documents.flatMap(row => {
    if (row.kind !== "body" || !row.deliverable || row.deliverable !== path.basename(row.deliverable)) return []
    const current = sha256File(path.join(directory, row.deliverable))?.sha256
    return row.conflict || (current && current !== row.sha256) ? [row.deliverable] : []
  })
  if (!changed.length) return null
  const since = gate && typeof gate === "object" && "checked_at" in gate && typeof gate.checked_at === "string" ? gate.checked_at : null
  return { since, documents: changed }
}

/** The run's planned documents without their parts, or null before the outline. */
function plannedDocuments(directory: string) {
  const plan = readLeanPlan(directory)
  return plan ? { documents: plan.documents.map(row => ({ id: row.id, title: row.title, purpose: row.purpose })) } : null
}

/** Pause: stops the coordinator working on this matter. It interrupts the
 * native sessions it is waiting on, drops any reply that arrives later, and
 * settles the run as blocked with the step in progress back to pending;
 * finished steps, answers and repair allowances stay as they were. Requests
 * already queued wait for the next Continue. Resolves once the coordinator has
 * stopped (mode "paused") or, after `settleMs`, while it is still stopping
 * (mode "pausing"); undefined when no coordinator is working on this matter. */
export async function pauseLeanDraftingRun(directory: string, settleMs = 5_000) {
  const coordinator = coordinators.get(directory)
  if (!coordinator) return undefined
  if (!coordinator.pause.signal.aborted) {
    coordinator.queuedAtPause = new Set(pendingLeanRequests(directory))
    coordinator.pause.abort()
  }
  let timer: ReturnType<typeof setTimeout> | undefined
  const stopped = await Promise.race([
    coordinator.stopped.then(() => true),
    new Promise<false>(resolve => { timer = setTimeout(() => resolve(false), settleMs) }),
  ])
  clearTimeout(timer)
  return { profile: "lean" as const, mode: stopped ? "paused" as const : "pausing" as const, status: leanDraftingStatus(directory)! }
}

const STOPPED = { error: "This draft was stopped. Send a new drafting request to start a new draft.", code: "lean_run_stopped" } as const

/** A stopped run is finished: it takes no further Continue, answer or update. */
function refuseStopped(state: LeanState | undefined) {
  if (state && draftingRunStopped(state)) throw new LeanDraftingConflictError(STOPPED.error, STOPPED.code)
}

/** Resolves once the background coordinator for this directory has stopped. */
export async function leanDraftingIdle(directory: string, pollMs = 10) {
  while (running.has(directory)) await Bun.sleep(pollMs)
}

export function leanDraftingRoutes(options: { directory: (id: string) => string; deps?: () => LeanDeps }) {
  const deps = options.deps ?? defaultDeps
  const app = new Hono()
  app.post("/:id/lean/checklist/refine", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const input = await c.req.json().catch(() => null)
    if (!input || typeof input !== "object" || Array.isArray(input) ||
      !isChecklistRefinement({ kind: "refine_checklist", parents: input.parents }))
      return c.json({ error: "Send ownerSessionID and unique parent_id/scope selections.", code: "invalid_checklist_refinement" }, 400)
    const state = readLeanState(directory)
    if (!state || input.ownerSessionID !== state.owner_session_id)
      return c.json({ error: "The active drafting run changed. Reload before refining its checklist.", code: "lean_run_changed" }, 409)
    if (running.has(directory) || pendingLeanRequests(directory).length || state.status !== "blocked" || state.paused_at || state.stages["checklist-check"]?.status !== "blocked")
      return c.json({ error: "Refinement requires an idle run blocked at checklist checking.", code: "checklist_not_settled" }, 409)
    const request = { kind: "refine_checklist", parents: input.parents }
    const comparable = new Set(checklistRefinableParents(directory).map(parent => parent.parent_id))
    if (!isChecklistRefinement(request) || request.parents.some(parent => !comparable.has(parent.parent_id)))
      return c.json({ error: "Select existing adverse checklist items; completed findings cannot be silently replaced.", code: "invalid_checklist_parent" }, 400)
    const queued = enqueueLeanRequest(directory, request)
    launch(directory, deps())
    return c.json({ profile: "lean", mode: "queued", request_id: queued.id, status: leanDraftingStatus(directory) }, 202)
  })
  app.post("/:id/lean/resume", (c) => {
    const directory = options.directory(c.req.param("id"))
    const resumed = guard(() => resumeLeanDraftingRun(directory, deps()))
    if (resumed instanceof LeanDraftingConflictError) return c.json({ error: resumed.message, code: resumed.code }, 409)
    if (!resumed) return c.json({ error: "This matter has no lean drafting run.", code: "lean_run_not_found" }, 404)
    return c.json(resumed, 202)
  })
  app.post("/:id/lean/pause", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const input: { ownerSessionID?: unknown } | null = await c.req.json().catch(() => null)
    const state = readLeanState(directory)
    if (!state || input?.ownerSessionID !== state.owner_session_id)
      return c.json({ error: "The active drafting run changed. Reload before pausing it.", code: "lean_run_changed" }, 409)
    const paused = await pauseLeanDraftingRun(directory)
    if (!paused) return c.json({ error: "This drafting run is not working, so there is nothing to pause.", code: "lean_run_not_working" }, 409)
    return c.json(paused, paused.mode === "paused" ? 200 : 202)
  })
  // Stop: only a run that is paused or blocked and that no coordinator is
  // working on. The check and the write run in one synchronous step, so no
  // coordinator can start in between; the run's files all stay in place.
  app.post("/:id/lean/stop", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const input: { ownerSessionID?: unknown } | null = await c.req.json().catch(() => null)
    const state = readLeanState(directory)
    if (!state || input?.ownerSessionID !== state.owner_session_id)
      return c.json({ error: "The active drafting run changed. Reload before stopping it.", code: "lean_run_changed" }, 409)
    if (running.has(directory) || state.status !== "blocked")
      return c.json({ error: "Only a paused or held drafting run can be stopped. Pause it first, or reload to see where it stands.", code: "lean_run_not_stoppable" }, 409)
    stopLeanRun(directory, state)
    return c.json({ profile: "lean", mode: "stopped", status: leanDraftingStatus(directory)! })
  })
  app.get("/:id/lean/state", (c) => {
    const status = leanDraftingStatus(options.directory(c.req.param("id")))
    if (!status) return c.json({ error: "This matter has no lean drafting run.", code: "lean_run_not_found" }, 404)
    return c.json(status)
  })
  app.post("/:id/lean/updates", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const body: unknown = await c.req.json().catch(() => null)
    if (!body || typeof body !== "object" || Array.isArray(body))
      return c.json({ error: "Send a drafting update as an object with text and ownerSessionID.", code: "invalid_update" }, 400)
    const input = body as { text?: unknown; ownerSessionID?: unknown }
    const state = readLeanState(directory)
    if (!state) return c.json({ error: "This matter has no drafting run.", code: "lean_run_not_found" }, 404)
    if (input.ownerSessionID !== state.owner_session_id)
      return c.json({ error: "The active drafting run changed. Reload before sending this update.", code: "lean_run_changed" }, 409)
    if (draftingRunStopped(state)) return c.json(STOPPED, 409)
    const admission = draftingUpdateAdmission({ ...state, in_flight: running.has(directory), pending_requests: pendingLeanRequests(directory).length })
    if (admission === "closed") return c.json({ error: "This update was not sent. Late clarification requires an idle draft held at checklist checking. Wait for the current work to settle; completed drafts need an amendment. You can still answer specific questions.", code: "lean_update_too_late" }, 409)
    const text = typeof input.text === "string" ? input.text.replace(/\r\n/g, "\n").trim() : ""
    if (!text || text.length > 20_000)
      return c.json({ error: "An update must contain 1–20,000 characters.", code: "invalid_update" }, 400)
    const request = enqueueLeanRequest(directory, { kind: admission, text })
    if (!running.has(directory)) launch(directory, deps())
    return c.json({ profile: "lean", mode: "queued", request_id: request.id, status: leanDraftingStatus(directory) }, 202)
  })
  app.post("/:id/lean/questions/:qid", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const input: { answer?: unknown } = await c.req.json().catch(() => ({}))
    const queued = guard(() => answerLeanQuestion(directory, c.req.param("qid"), input.answer, deps()))
    if (queued instanceof LeanDraftingConflictError) return c.json({ error: queued.message, code: queued.code }, 409)
    if (queued instanceof LeanQuestionError)
      return c.json({ error: queued.message, code: queued.code }, queued.code === "not_found" ? 404 : 400)
    if (!queued) return c.json({ error: "This matter has no lean drafting run.", code: "lean_run_not_found" }, 404)
    return c.json(queued, 202)
  })
  return app
}

function launch(directory: string, deps: LeanDeps) {
  running.add(directory)
  const queuedAtLaunch = new Set(pendingLeanRequests(directory))
  const coordinator: Coordinator = { pause: new AbortController(), stopped: Promise.resolve() }
  coordinators.set(directory, coordinator)
  coordinator.stopped = runLean(directory, { ...deps, pause: coordinator.pause.signal })
    .then(() => undefined, (error: unknown) => {
      console.error("lean drafting run stopped", { matter_ref: draftingTraceReference(directory), ...diagnosticErrorMetadata(error) })
      // Stage errors are recorded by the coordinator; this covers a failure
      // outside any stage so the state never claims a run that is not there.
      const state = readLeanState(directory)
      if (state?.status === "running")
        writeLeanState(directory, { ...state, status: "failed", stop_reason: `stopped by a runtime error: ${matterRelative(directory, error instanceof Error ? error.message : String(error))}` })
    })
    .finally(() => {
      if (coordinators.get(directory) === coordinator) coordinators.delete(directory)
      // The coordinator re-checks the queue before it returns; a request that
      // landed after that check starts it again. Requests it already saw and
      // left (a run that stopped on an error) do not, so a failure cannot loop;
      // nor do requests a pause left waiting for the next Continue.
      if (pendingLeanRequests(directory).some(name => !queuedAtLaunch.has(name) && !coordinator.queuedAtPause?.has(name))) return launch(directory, deps)
      running.delete(directory)
    })
}

/** Moves a finished run to .dochaus/lean-archive/<created>-<owner session>/.
 * Its registered deliverables are matter documents and stay in the matter root. */
function archiveLeanRun(directory: string, state: LeanState) {
  const root = path.join(directory, ".dochaus", "lean-archive")
  mkdirSync(root, { recursive: true, mode: 0o700 })
  renameSync(leanDir(directory), path.join(root, `${state.created_at}-${state.owner_session_id}`.replace(/[^A-Za-z0-9_-]/g, "")))
}

function defaultDeps(): LeanDeps {
  return { sessions: engineSessionPort(), registerDeliverable: registerLeanDeliverable }
}

/** Returns a known refusal as a value so routes map it to a status code. */
function guard<T>(work: () => T): T | LeanDraftingConflictError | LeanQuestionError {
  try {
    return work()
  } catch (error) {
    if (error instanceof LeanDraftingConflictError || error instanceof LeanQuestionError) return error
    throw error
  }
}
