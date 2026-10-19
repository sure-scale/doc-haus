import { checklistRefinableParents } from "../../../dochaus/workflows/lean/checklist-flow"
import { draftingRunStopped, draftingUpdateAdmission } from "../../../dochaus/lib/drafting-update-admission"
import { mkdirSync, renameSync } from "node:fs"
import path from "node:path"
import { Hono } from "hono"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { researchStatus } from "../../../dochaus/workflows/lean/research"
import { readLeanDocuments, requestedMemo, type LeanDocuments } from "../../../dochaus/workflows/lean/deliverables"
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
  answerWith,
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
  type AnswerForm,
  type LeanState,
  type Question,
} from "../../../dochaus/workflows/lean/state"
import { registeredCombinedParts } from "./combined-output"
import { deliverableDocumentName } from "./deliverable"
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
// Stop settles an unfinished run that nothing is working on (paused, blocked, or
// left running by a restart) as failed: it takes no further Continue, answer or
// update, and the next drafting request archives it like any failed run.
// Nothing in the matter is deleted.

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

/** Queues the lawyer's answer to a question. The body carries exactly one of
 * `answer` text, the `option_id` of an offered option, or `decide_for_me: true`
 * to adopt the recommended option. It is admitted only if the question, as its
 * queued answers will leave it, takes that form (answerWith): an unknown
 * question, empty words or an option it does not offer is refused, and so is an
 * adoption of a question answered, recorded or still queued, or one with no
 * recommendation to adopt. The coordinator records it at its next stage
 * boundary, checking again since a stage may have replaced the question
 * meanwhile, and re-runs the stages that consume answers. Check and enqueue run
 * synchronously, so no other request in this process lands between. */
export function answerLeanQuestion(directory: string, questionID: string, input: unknown, deps: LeanDeps = defaultDeps()) {
  if (!leanStateExists(directory)) return undefined
  refuseStopped(readLeanState(directory))
  const body = input && typeof input === "object" && !Array.isArray(input) ? input as Record<string, unknown> : {}
  const given = ["answer", "option_id", "decide_for_me"].filter(key => key in body)
  if (given.length !== 1 || ("option_id" in body && typeof body.option_id !== "string") || ("decide_for_me" in body && body.decide_for_me !== true))
    throw new LeanQuestionError("Send exactly one of answer text, option_id or decide_for_me: true.", "invalid_answer")
  const form: AnswerForm = "decide_for_me" in body ? { adopt: true } : "option_id" in body ? { option_id: body.option_id as string } : { text: typeof body.answer === "string" ? body.answer : "" }
  const question = readQuestions(directory).find(question => question.id === questionID)
  answerWith(afterQueuedAnswers(directory, question), form)
  const queued = enqueueLeanRequest(directory, "adopt" in form ? { kind: "answer", question_id: questionID, adopt: true }
    : "option_id" in form ? { kind: "answer", question_id: questionID, option_id: form.option_id }
    : { kind: "answer", question_id: questionID, answer: validateAnswerText(form.text) })
  if (!running.has(directory)) launch(directory, deps)
  return { profile: "lean" as const, mode: "queued" as const, request_id: queued.id, status: leanDraftingStatus(directory) }
}

/** `question` as the drain will leave it: answered once any answer is queued for it. */
function afterQueuedAnswers(directory: string, question: Question | undefined): Question | undefined {
  return question && queuedAnswerIDs(directory).has(question.id) ? { ...question, status: "answered" } : question
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
    // Derived, never stored: the lawyer stopped this run (the reason is display only).
    stopped: draftingRunStopped(state),
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
    // The name the plan gives the memo the lawyer asked for, which the run's
    // notes are written and delivered as; null when no memo was requested.
    memo: requestedMemo(directory)?.name ?? null,
    combined_output: combinedOutput(directory),
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

/** Hash of a delivered draft read inside the matter: a plain document name,
 * opened without following links, regular files only. Anything else reads as
 * missing, the same as an absent file. */
function matterFileSha(directory: string, name: string) {
  try { return sha256File(path.join(directory, name))?.sha256 } catch { return undefined }
}

/** Delivered drafts whose matter file no longer holds the bytes the final check
 * saw: a change accepted after delivery, a newer output kept from replacing
 * such a change, or a draft no longer in the matter. `missing` is the subset
 * with no readable file under the name (deleted, or replaced by anything but a
 * regular file), so the lawyer is not told to review a draft that is gone.
 * Disclosure only: the check is not re-run and its verdict is passed through
 * unchanged. */
function changedAfterCheck(directory: string, gate: unknown, documents: LeanDocuments | null) {
  if (documents?.status !== "completed") return null
  const missing: string[] = []
  const changed = documents.documents.flatMap(row => {
    if (row.kind !== "body" || !row.deliverable || !deliverableDocumentName(row.deliverable)) return []
    const current = matterFileSha(directory, row.deliverable)
    if (current === undefined) missing.push(row.deliverable)
    return row.conflict || current !== row.sha256 ? [row.deliverable] : []
  })
  if (!changed.length) return null
  const since = gate && typeof gate === "object" && "checked_at" in gate && typeof gate.checked_at === "string" ? gate.checked_at : null
  return { since, documents: changed, missing }
}

/** Derived, never stored: the one file a combined document downloads as, once
 * the summary registers it and both its parts under the current plan's names
 * (registeredCombinedParts), with
 * the document's id, the combined file's registered matter name ("<name>.docx"
 * or a " (n)" suffix) and the exact registered names of its body and notes, so a
 * client matches those rows rather than guessing them from the name. Null
 * otherwise, including when no combined row is registered. */
function combinedOutput(directory: string): { document: string; name: string; file: string; body: string; notes: string } | null {
  const parts = registeredCombinedParts(directory)
  return parts ? { document: parts.document, name: parts.name, file: parts.file, body: parts.body, notes: parts.notes } : null
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
  // Stop: only an unfinished run that no coordinator is working on, the same
  // check the status reports as in_flight. The check and the write run in one
  // synchronous step, so no coordinator can start in between; the run's files
  // all stay in place.
  app.post("/:id/lean/stop", async (c) => {
    const directory = options.directory(c.req.param("id"))
    const input: { ownerSessionID?: unknown } | null = await c.req.json().catch(() => null)
    const state = readLeanState(directory)
    if (!state || input?.ownerSessionID !== state.owner_session_id)
      return c.json({ error: "The active drafting run changed. Reload before stopping it.", code: "lean_run_changed" }, 409)
    if (running.has(directory))
      return c.json({ error: "This drafting run is working. Pause it before stopping it.", code: "lean_run_not_stoppable" }, 409)
    if (state.status !== "blocked" && state.status !== "running")
      return c.json({ error: "This drafting run has already finished, so there is nothing to stop.", code: "lean_run_not_stoppable" }, 409)
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
    const input: unknown = await c.req.json().catch(() => null)
    const queued = guard(() => answerLeanQuestion(directory, c.req.param("qid"), input, deps()))
    if (queued instanceof LeanDraftingConflictError) return c.json({ error: queued.message, code: queued.code }, 409)
    if (queued instanceof LeanQuestionError)
      return c.json({ error: queued.message, code: queued.code }, queued.code === "not_found" ? 404 : queued.code === "invalid_answer" ? 400 : 409)
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
