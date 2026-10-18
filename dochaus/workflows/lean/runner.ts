import { runFocusedChangeReview } from "./focused-change-review"
import { readFirmDraftingPreferences } from "../../lib/drafting-preferences"
import { DRAFTING_STOPPED_REASON } from "../../lib/drafting-update-admission"
import { renderMarkdownToDocx } from "../../lib/render"
import { existsSync, readFileSync, rmSync } from "node:fs"
import { runChangeReconciliation } from "./change-reconciliation"
import { focusedChecklistRepairs, runFocusedRepairAdjudication } from "./checklist-repair-adjudication"
import { classifyChecklistFinding } from "./checklist-classify"
import { checklistFindingRecovery } from "./checklist-contract"
import { buildChecklistRefinementContext } from "./checklist-refinement-context"
import { runChecklistRefinementFlow, readEffectiveChecklistItems, runChecklistExtractionFlow, runChecklistCheckFlow, type ChecklistCheckResult } from "./checklist-flow"
import { selectedDeliveryPlan } from "./delivery-selection"
import { changePlanPrompt, changeReviewPrompt, readChangeReview, runChangePreparation } from "./change-pipeline"
import { packageReviewPrompt, readPackageReview } from "./package-review"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, checklistRepairItems, readRepairRound, repairedChecklistItems, startRepairRound } from "./repairs"
import { acceptEvidenceReply, evidencePrompt } from "./evidence"
import { researchPermissions, acceptResearchReply, researchPrompt, startResearchGapRound } from "./research"
import { runApply } from "./apply-stage"
import { DELIVERED_NOTES, forgetUnplannedDocuments, recordLeanDocuments } from "./deliverables"
import { runDraftStage } from "./draft-sections"
import { appendLeanEvent, leanElapsedMs } from "./events"
import { checklistFindingLabel, deliveredNotes, finalGate, notesCoverageReport } from "./gate"
import { readLeanPlan, reviewStages, leanSteps, stepDocument, stepStage, type StepID } from "./plan"
import { retryPrompt, stagePrompt } from "./prompts"
import { excerpt, inDocument, openItem, withSourceSnapshot } from "./quotes"
import { checklistRefinementSelections, drainLeanRequests, pendingLeanRequests, withheldChecklistRepair } from "./requests"
import {
  LEAN_CHILD_PERMISSIONS,
  LEAN_MAX_RETRIES,
  LEAN_PROMPT_VERSION,
  LEAN_STAGE_AGENTS,
  acceptStageReply,
  stageInputs,
  stageOutputValid,
  stepOutputs,
  stepSessionTitle,
  type LeanDeps,
  type ModelStageID,
} from "./stages"
import {
  isReviewStage,
  hashStageInputs,
  keepAskedQuestions,
  leanFile,
  matterRelative,
  readLeanState,
  readLeanJson,
  readLeanText,
  readOpenItems,
  readQuestions,
  replaceStageOpenItems,
  replaceStageQuestions,
  snapshotStage,
  syncLeanSteps,
  updateLeanStage,
  writeOpenItems,
  writeQuestions,
  writeLeanBytes,
  writeLeanState,
  writeLeanText,
  type LeanStageState,
  type LeanState,
  type OpenItem,
  type Question,
} from "./state"

// Lean drafting coordinator. It walks the run's steps in order (each planned
// document's draft, review, changes and apply, then finalize and the gate),
// skips a step whose inputs are unchanged since it last completed, and
// otherwise runs it: one native child session per model step, a deterministic
// check of its output, and up to three retries that hand the checker's issues
// back to the same session. The first document's draft saves the plan, which
// fixes the remaining steps. The coordinator knows where each draft docx is,
// where the applier is and which changes to apply; it never reads or judges
// their legal content. Open items are recorded, tagged with their document, as
// they arise. Unanswered questions pause the walk before any further step.
//
// The coordinator is the only writer of the run's files. Lawyer requests
// (Continue, answers) wait in the request queue: answers are drained at every
// stage boundary, and a Continue once the pass has settled. Before it returns,
// the coordinator checks the queue again and runs another pass for anything
// that arrived, so no request is left behind.
//
// A pause (deps.pause) stops the coordinator at once: it interrupts the native
// sessions it is waiting on, accepts no reply that arrives later, returns the
// step in progress to pending and settles the run as blocked. Finished steps,
// answers, approvals and repair allowances are left exactly as they were, so a
// later Continue resumes from the first unfinished step.

type StageRun =
  | { status: "restart" }
  | { status: "done" | "blocked"; outputs: string[]; openItems: OpenItem[]; questions?: Question[]; reason?: string }
  | { status: "failed"; reason: string }

type PreparedReview = { inputSha: string; sessionID: string; reply: Promise<string> }

type PassResult = { state: LeanState; restart?: boolean; paused?: boolean }

type Progress = (patch: Partial<LeanStageState>) => void

// Substantive repairs share a run-wide allowance; checker retries remain separate.
export { LEAN_MAX_SUBSTANTIVE_REPAIRS } from "./repairs"

export async function runLean(directory: string, deps: LeanDeps): Promise<LeanState> {
  const initial = readLeanState(directory)
  if (!initial) throw new Error("This matter has no lean drafting run.")
  // Completed legacy runs keep their original accepted package until the
  // lawyer supplies an update or answer; installing new stages is not a repair.
  if (initial.status === "completed" && !initial.stages.research) {
    const changes = drainLeanRequests(directory, initial)
    if (!changes.answered && !changes.updated) return initial
  }
  // The state as last settled; a Continue is applied against it.
  const pass = { state: initial, settled: initial as LeanState | undefined, reinspectChecklist: false }
  for (;;) {
    if (deps.pause?.aborted) return settlePause(directory, pass.state)
    // A paused run never settled at checklist checking, so its Continue is
    // not a request to reinspect.
    pass.reinspectChecklist ||= pass.settled?.status === "blocked" && !pass.settled.paused_at && pass.settled.stages["checklist-check"]?.status === "blocked" &&
      pendingLeanRequests(directory).some(name => (readLeanJson(directory, `requests/${name}`) as { kind?: string } | undefined)?.kind === "continue")
    drainLeanRequests(directory, pass.settled)
    const result: PassResult = await runPass(directory, writeLeanState(directory, { ...pass.state, status: "running", stop_reason: null, paused_at: undefined }), deps, pass.reinspectChecklist)
      .catch((error: unknown) => deps.pause?.aborted
        ? { state: settlePause(directory, pass.state), paused: true }
        : { state: writeLeanState(directory, { ...(readLeanState(directory) ?? pass.state), status: "failed", stop_reason: `Drafting stopped by a runtime error: ${message(directory, error)}` }), restart: false })
    pass.state = readLeanState(directory) ?? result.state
    if (result.paused) return pass.state
    pass.settled = result.restart ? undefined : pass.state
    if (!result.restart) pass.reinspectChecklist = false
    if (!result.restart && !pendingLeanRequests(directory).length) return pass.state
  }
}

/** One walk over the steps. An answer drained at a step boundary that changes
 * the inputs of a step this pass already settled restarts the walk, so the
 * earliest step that consumes the answer runs next. The step list is read
 * again at every boundary because the first draft writes the plan. */
async function runPass(directory: string, initial: LeanState, sourceDeps: LeanDeps, reinspectChecklist = false): Promise<PassResult> {
  const preferences = readFirmDraftingPreferences()
  const preferenceChange = { observed: false }
  const preferencesChanged = () => preferenceChange.observed ||= readFirmDraftingPreferences() !== preferences
  const pause = sourceDeps.pause
  const paused = () => pause?.aborted === true
  // A pause interrupts every native session with a prompt in flight.
  const live = new Set<string>()
  const interruptLive = () => {
    for (const sessionID of live) void sourceDeps.sessions.interrupt?.({ directory, sessionID }).catch(() => undefined)
  }
  pause?.addEventListener("abort", interruptLive, { once: true })
  // Every model exchange, including focused helpers, shares this pass's guard.
  // A later return to the original value must not clear an observed change.
  // Once paused, no session or prompt starts and no reply is returned, even
  // one that arrives afterwards.
  const deps: LeanDeps = { ...sourceDeps, sessions: { ...sourceDeps.sessions, createSession(input) {
    return paused() ? Promise.reject(new LeanPauseRequested()) : sourceDeps.sessions.createSession(input)
  }, async prompt(input) {
    if (paused()) throw new LeanPauseRequested()
    if (preferencesChanged()) throw new ChecklistRefreshRestart()
    live.add(input.sessionID)
    try {
      return await untilPaused(sourceDeps.sessions.prompt(pause ? { ...input, signal: pause } : input), pause)
    } finally {
      live.delete(input.sessionID)
      if (paused()) throw new LeanPauseRequested()
      if (preferencesChanged()) throw new ChecklistRefreshRestart()
    }
  } } }
  let state = followPlan(directory, initial)
  const settled: StepID[] = []
  const prepared = new Map<StepID, PreparedReview>()
  try {
  for (let index = 0; index < leanSteps(directory).length; index += 1) {
    if (paused()) return { state: settlePause(directory, state), paused: true }
    // Settings saves are outside the matter queue. At safe stage boundaries,
    // restart the content-keyed walk before accepting work from mixed settings.
    if (preferencesChanged()) return { state, restart: true }
    const step = leanSteps(directory)[index]!
    const stage = stepStage(step)
    const document = stepDocument(step) ?? undefined
    const changes = drainLeanRequests(directory)
    // The gate asks its question after the walk; it holds back no step.
    const unanswered = readQuestions(directory).filter(question => question.status === "open" && question.stage !== "gate")
    if (unanswered.length)
      return { state: writeLeanState(directory, { ...state, status: "blocked", stop_reason: `Waiting for your answers: ${unanswered.map(question => question.id).join(", ")}` }) }
    if ((changes.answered || changes.updated) &&
      withSourceSnapshot(() => settled.some(done => stageInputSha(directory, done) !== state.stages[done]?.input_sha)))
      return { state, restart: true }
    // A resumed draft may be skipped, but source/target changes must still
    // refresh extraction before any body is consumed or checking begins.
    if (stage === "draft" && readLeanPlan(directory)) {
      await ensureChecklistExtraction(directory, state, deps)
      if (paused()) return { state: settlePause(directory, state), paused: true }
      if (preferencesChanged()) return { state, restart: true }
      const queued = drainLeanRequests(directory)
      if (queued.answered || queued.updated) return { state, restart: true }
    }
    // The first draft can write the plan its own key reads; its key is then
    // taken again once the plan is saved.
    const keyed = performance.now()
    const recorded = state.stages[step]
    // Keying and the skip decision are synchronous: they share one source read.
    const decision = withSourceSnapshot(() => {
      // Package review may take time; do not classify or register a package
      // against checklist results for Word bytes that have since changed.
      if (stage === "gate" && state.stages["checklist-check"]?.input_sha !== stageInputSha(directory, "checklist-check")) return undefined
      const key = stageInputSha(directory, step)
      // Done from these very inputs; the saved output decides whether it is skipped.
      const current = recorded?.status === "done" && recorded.input_sha === key
      return { key, current, unchanged: current && stageOutputValid(directory, step) && (step !== "package-review" || readPackageReview(directory)?.status === "pass") && (stage !== "change-review" || !document || readChangeReview(directory, document)?.status === "pass") }
    })
    if (!decision) return { state, restart: true }
    settled.push(step)
    let inputSha = decision.key
    // Time to key the step and decide whether it can be skipped.
    const check_ms = leanElapsedMs(keyed)
    if (decision.unchanged) {
      // A durable accepted review may have preceded a crash before its
      // delivery selection was applied. Reconcile it before any registration.
      if (step === "package-review") {
        const repair = await reconcilePackageReview(directory, state, deps)
        if (repair) return repair
      }
      appendLeanEvent(directory, { stage, document, event: "skipped", check_ms })
      // A review accepted before a crash may still owe its gap round.
      if (step === "research-review" && startResearchGapRound(directory)) return { state, restart: true }
      continue
    }
    const prefetched = prepared.get(step)
    const currentPrefetch = prefetched?.inputSha === inputSha
    if (!currentPrefetch) appendLeanEvent(directory, { stage, document, event: "started", check_ms })
    state = updateLeanStage(directory, state, step, {
      status: "running", input_sha: null, outputs: [], session_ids: [], session_id: currentPrefetch ? prefetched.sessionID : null, attempts: 0, calls: 0, feedback: [], started_at: currentPrefetch ? state.stages[step]?.started_at ?? now() : now(), finished_at: null,
    })
    const progress: Progress = patch => {
      state = updateLeanStage(directory, state, step, patch)
    }
    const planned = async () => {
      state = followPlan(directory, state)
      await ensureChecklistExtraction(directory, state, deps)
      const queued = drainLeanRequests(directory)
      if (queued.answered || queued.updated || preferencesChanged()) throw new ChecklistRefreshRestart()
      inputSha = stageInputSha(directory, step)
    }
    // Fan out the remaining reviews before accepting any reply. Only this
    // writer saves replies in stage order; every review feeds the one changes
    // stage. Queued updates invalidate prefetched replies by their input keys.
    if (isReviewStage(stage) && document) {
      // Keying the remaining reviews is synchronous: they share one source read.
      // A source change before a prompt is built still fails its key at its own step.
      const fanout = withSourceSnapshot(() => reviewStages(directory, document).slice(reviewStages(directory, document).indexOf(stage) + 1).flatMap(review => {
        const independent: StepID = `${document}.${review}`
        if (prepared.has(independent)) return []
        const prefetchKeyed = performance.now()
        const saved = state.stages[independent]
        const key = stageInputSha(directory, independent)
        return saved?.status === "done" && saved.input_sha === key && stageOutputValid(directory, independent) ? [] : [{ review, independent, key, prefetchCheck: leanElapsedMs(prefetchKeyed) }]
      }))
      for (const { review, independent, key, prefetchCheck } of fanout) {
        const agent = LEAN_STAGE_AGENTS[review]
        const creation = await deps.sessions.createSession({ directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, independent), model: state.model, permissions: LEAN_CHILD_PERMISSIONS })
          .then(sessionID => ({ sessionID }), (error: unknown) => ({ error }))
        if ("error" in creation) {
          if (paused()) return { state: settlePause(directory, state), paused: true }
          if (preferencesChanged()) return { state, restart: true }
          const reason = `${independent}: stopped by a runtime error: ${message(directory, creation.error)}`
          state = updateLeanStage(directory, state, step, { status: "pending", input_sha: null, outputs: [], finished_at: null })
          state = updateLeanStage(directory, state, independent, { status: "failed", finished_at: now() })
          appendLeanEvent(directory, { stage: review, document, event: "failed", reason })
          return { state: writeLeanState(directory, { ...state, status: "failed", stop_reason: reason }) }
        }
        const sessionID = creation.sessionID
        state = updateLeanStage(directory, state, independent, { status: "running", session_id: sessionID, input_sha: null, outputs: [], attempts: 0, calls: 1, feedback: [], started_at: now(), finished_at: null })
        appendLeanEvent(directory, { stage: review, document, event: "started", session_id: sessionID, check_ms: prefetchCheck })
        const reply = deps.sessions.prompt({ directory, sessionID, agent, text: stagePrompt(directory, independent) })
        // Consume native errors when accepting this stage, without an unhandled
        // rejection while another independent review is running.
        void reply.catch(() => undefined)
        prepared.set(independent, { inputSha: key, sessionID, reply })
      }
    }
    if (prefetched && prefetched.inputSha !== inputSha) await prefetched.reply.catch(() => undefined)
    prepared.delete(step)
    const result = await runStep(directory, state, step, deps, progress, planned, prefetched?.inputSha === inputSha ? prefetched : undefined, reinspectChecklist, decision.current)
      .catch((error: unknown): StageRun => error instanceof ChecklistRefreshRestart ? { status: "restart" } : ({ status: "failed", reason: `${step}: stopped by a runtime error: ${message(directory, error)}` }))
    // Whatever the step returned once paused (helpers turn a refused prompt into
    // their own failure) is not accepted; the step is run again on Continue.
    if (paused()) return { state: settlePause(directory, state), paused: true }
    if (result.status === "restart" || preferencesChanged()) {
      state = updateLeanStage(directory, state, step, { status: "pending", input_sha: null, outputs: [], finished_at: null })
      return { state, restart: true }
    }
    appendLeanEvent(directory, result.status === "failed"
      ? { stage, document, event: "failed", reason: result.reason }
      : { stage, document, event: result.status, outputs: result.outputs, reason: result.reason })
    if (result.status === "failed") {
      state = updateLeanStage(directory, state, step, { status: "failed", finished_at: now() })
      return { state: writeLeanState(directory, { ...state, status: "failed", stop_reason: result.reason }) }
    }
    replaceStageOpenItems(directory, stage, document ?? null, document ? result.openItems.map(item => inDocument(item, document)) : result.openItems)
    if (result.questions)
      replaceStageQuestions(directory, stage, document ?? null, result.questions.map(question => document ? { ...question, document } : question))
    snapshotStage(directory, state.revision, step, result.outputs)
    // These steps write some of their own inputs (the gate its question); their key is taken again.
    state = updateLeanStage(directory, state, step, { status: result.status, input_sha: step === "checklist-check" || step === "gate" ? stageInputSha(directory, step) : inputSha, outputs: result.outputs, finished_at: now() })
    // Researchable gaps from the first review of this research go back to it once.
    if (step === "research-review" && result.status === "done" && startResearchGapRound(directory)) return { state, restart: true }
    if (stage === "change-review" && document && readChangeReview(directory, document)?.status === "changes-required") {
      const review = readChangeReview(directory, document)!
      if ((readRepairRound(directory)?.substantive_attempts ?? 0) >= LEAN_MAX_SUBSTANTIVE_REPAIRS)
        return { state: writeLeanState(directory, { ...state, status: "blocked", stop_reason: `${step}: material corrections remain after ${LEAN_MAX_SUBSTANTIVE_REPAIRS} substantive repair rounds: ${review.findings.map(row => row.detail).join(" ")}` }) }
      startRepairRound(directory, review.findings.map(row => openItem("change-review", "omission", row.detail,
        { repair_target: row.target ?? "body", document: row.document ?? document, ...(row.pid ? { pid: row.pid } : {}) })), true)
      return { state, restart: true }
    }
    if (stage === "package-review") {
      const repair = await reconcilePackageReview(directory, state, deps)
      if (repair) return repair
    }
    if (result.status === "blocked") {
      await recordLeanDocuments(directory, deps, false)
      if (preferencesChanged()) return { state, restart: true }
      return { state: writeLeanState(directory, { ...state, status: "blocked", stop_reason: result.reason ?? null }) }
    }
  }
  if (preferencesChanged()) return { state, restart: true }
  const delivered = await recordLeanDocuments(directory, deps, true).then(() => undefined, (error: unknown) => message(directory, error))
  if (preferencesChanged()) return { state, restart: true }
  if (delivered) return { state: writeLeanState(directory, { ...state, status: "failed", stop_reason: `deliverables: ${delivered}` }) }
  return { state: writeLeanState(directory, { ...state, status: "completed", stop_reason: null }) }
  } catch (error) {
    // Helpers may translate a rejected prompt into their own failure result.
    // A pause, then preference invalidation, takes precedence over that work.
    if (paused()) return { state: settlePause(directory, state), paused: true }
    if (preferencesChanged()) return { state, restart: true }
    throw error
  } finally {
    pause?.removeEventListener("abort", interruptLive)
    // Read-only native reviews must settle even when the other review fails or
    // a later update restarts the pass. Their session history remains durable;
    // stale replies are never saved as accepted outputs.
    await Promise.all([...prepared.values()].map(row => row.reply.catch(() => undefined)))
    const latest = readLeanState(directory)
    if (latest && prepared.size) {
      // A settled but unaccepted reply has no output to resume from. Retain its
      // native session for inspection and show the stage as pending, not live.
      writeLeanState(directory, { ...latest, stages: { ...latest.stages,
        ...Object.fromEntries([...prepared].flatMap(([step, row]) => latest.stages[step]?.status === "running" && latest.stages[step]?.session_id === row.sessionID
          ? [[step, { ...latest.stages[step], status: "pending", input_sha: null, outputs: [], finished_at: null }]] : [])),
      } })
    }
  }
}

/** The existing semantic review selects the requested artifact. The sole
 * writer applies that selection atomically, then refreshes notes/review with
 * the corrected delivery context. Content stages consume operative inputs. */
async function reconcilePackageReview(directory: string, state: LeanState, deps: LeanDeps) {
  const review = readPackageReview(directory)
  if (!review) return
  const plan = selectedDeliveryPlan(directory, review.delivery)
  if (review.status === "pass" && !plan) return
  const prior = readRepairRound(directory)
  if ((prior?.substantive_attempts ?? 0) >= LEAN_MAX_SUBSTANTIVE_REPAIRS) {
    await recordLeanDocuments(directory, deps, false)
    return { state: writeLeanState(directory, { ...state, status: "blocked", stop_reason: `package-review: material corrections remain after ${LEAN_MAX_SUBSTANTIVE_REPAIRS} substantive repair rounds: ${review.findings.map(row => row.detail).join(" ")}${plan ? " Requested delivery selection remains unapplied." : ""}` }) }
  }
  const findings = review.findings.map(row => openItem("package-review", "omission", row.detail,
    // A separately planned advice/note document has its own editing stages.
    // Its explicit owner takes precedence over the reviewer's prose label.
    { repair_target: row.target === "memo" && row.document ? "body" : row.target ?? (row.document ? "body" : "memo"), ...(row.document ? { document: row.document } : {}), ...(row.pid ? { pid: row.pid } : {}) }))
  // Delivery and notes-only corrections charge the allowance without replacing
  // accepted operative/research repair context or changing its generation key.
  if (prior && findings.every(item => item.ref.repair_target === "memo"))
    writeLeanText(directory, "repairs.json", `${JSON.stringify({ ...prior,
      // A checklist notes repair stays required beside the new review findings.
      items: [...prior.items.filter(item => item.stage !== "package-review" || item.ref.repair_target !== "memo"), ...findings],
      substantive_attempts: (prior.substantive_attempts ?? 0) + 1,
    }, null, 2)}\n`)
  else startRepairRound(directory, findings, true)
  if (plan) writeLeanText(directory, "plan.json", `${JSON.stringify(plan, null, 2)}\n`)
  return { state, restart: true }
}

class ChecklistRefreshRestart extends Error {}

export const LEAN_PAUSED_REASON = "Paused at your request."

/** The lawyer paused the run: no further model work starts or is accepted. */
class LeanPauseRequested extends Error {
  constructor() {
    super(LEAN_PAUSED_REASON)
  }
}

/** The reply, unless the run is paused first. A reply that arrives after the
 * pause is dropped: nothing waits for it any more. */
function untilPaused<T>(work: Promise<T>, pause: AbortSignal | undefined): Promise<T> {
  if (!pause) return work
  return new Promise<T>((resolve, reject) => {
    const stop = () => reject(new LeanPauseRequested())
    pause.addEventListener("abort", stop, { once: true })
    if (pause.aborted) stop()
    void work.then(resolve, reject).finally(() => pause.removeEventListener("abort", stop))
  })
}

/** Settles a paused run as blocked. A step still running returns to pending,
 * keeping its native session for inspection; nothing else is changed. */
function settlePause(directory: string, fallback: LeanState): LeanState {
  const latest = readLeanState(directory) ?? fallback
  appendLeanEvent(directory, { stage: "run", event: "paused", reason: LEAN_PAUSED_REASON })
  return writeLeanState(directory, { ...latest, status: "blocked", stop_reason: LEAN_PAUSED_REASON, paused_at: now(), stages: unfinished(latest.stages) })
}

/** A step left running, which no coordinator now works on, back to pending; its
 * session and attempts stay for inspection. Other steps are returned as they are. */
function unfinished(stages: LeanState["stages"]) {
  return Object.fromEntries(Object.entries(stages).map(([step, row]) => [step, row?.status === "running"
    ? { ...row, status: "pending", input_sha: null, outputs: [], finished_at: null } : row])) as LeanState["stages"]
}

/** The lawyer stopped a run no coordinator is working on (the caller checks):
 * paused, held, or left running by an ingest restart. It settles as failed, so
 * the next drafting request archives it and starts a new run. As on a pause, a
 * step left running returns to pending; every other step, output, answer,
 * repair allowance and queued request stays exactly as it was: nothing is
 * marked done and no budget is reset. */
export function stopLeanRun(directory: string, state: LeanState): LeanState {
  appendLeanEvent(directory, { stage: "run", event: "stopped", reason: DRAFTING_STOPPED_REASON })
  return writeLeanState(directory, { ...state, status: "failed", stop_reason: DRAFTING_STOPPED_REASON, paused_at: undefined, stages: unfinished(state.stages) })
}

async function ensureChecklistExtraction(directory: string, state: LeanState, deps: LeanDeps) {
  const result = await runChecklistExtractionFlow(directory, { sessions: deps.sessions, model: state.model, parentID: state.owner_session_id })
  if (result.status !== "succeeded") throw new Error(`Source checklist extraction is incomplete: ${result.sources.filter(row => row.status !== "succeeded").map(row => `${row.key}: ${row.issues.join(" ")}`).join("; ")}. See checklist/extraction.json.`)
}

async function runChecklistStage(directory: string, state: LeanState, deps: LeanDeps, reinspectAdverse = false): Promise<StageRun> {
  const refinementOutputs: string[] = []
  for (const selection of checklistRefinementSelections(directory)) {
    const refined = await runChecklistRefinementFlow(directory, { sessions: deps.sessions, model: state.model, parentID: state.owner_session_id }, buildChecklistRefinementContext(directory, selection))
    refinementOutputs.push(...refined.outputs)
  }
  const key = stageInputSha(directory, "checklist-check")
  const checkDeps = { sessions: deps.sessions, model: state.model, parentID: state.owner_session_id }
  const checked = await runChecklistCheckFlow(directory, checkDeps, { reinspectAdverse, pause: deps.pause })
  // A finding the check could not establish gets its one reinspection here,
  // limited to those items and only with no integrity failure; it is then
  // disclosed, never repaired. A recorded reinspection is never repeated.
  const uninspected = checked.status === "succeeded" && checked.checks.every(row => row.reinspection?.status !== "failed") && !deps.pause?.aborted
    ? missingInspections(directory, checked.checks) : new Set<string>()
  const result = uninspected.size ? await runChecklistCheckFlow(directory, checkDeps, { reinspectItems: uninspected, pause: deps.pause }) : checked
  // An answer is queued while read-only jobs run; it cannot become accepted
  // evidence for the snapshot they inspected. Refresh before proceeding.
  const queued = drainLeanRequests(directory)
  if (queued.answered || queued.updated) return { status: "restart" }
  // The fresh key and the effective items share one source read.
  const current = withSourceSnapshot(() => ({ key: stageInputSha(directory, "checklist-check"), effective: readEffectiveChecklistItems(directory) }))
  if (current.key !== key) return { status: "restart" }
  const itemMap = new Map(current.effective.items.map(record => [record.id, record.item] as const))
  // Integrity: a current item without one completed finding. Only these block;
  // the gate discloses every completed adverse finding with the review draft.
  const integrity = result.checks.filter(row => row.status !== "succeeded" || row.reinspection?.status === "failed" || !row.finding || !itemMap.has(row.item_id))
  const adverse = result.checks.flatMap(row => {
    const item = itemMap.get(row.item_id)
    if (integrity.includes(row) || !item || !row.finding) return []
    const classified = classifyChecklistFinding(item, row.finding, current.effective.items)
    return classified ? [{ row, item, finding: row.finding, recovery: checklistFindingRecovery(item, row.finding), class: classified.class }] : []
  })
  // A repair the current source adjudication declined is disclosed with its
  // reason; neither the runner nor Continue resubmits it unchanged.
  const declined = new Map(adverse.flatMap(({ row, item, recovery }) => {
    if (recovery !== "repair") return []
    const adjudication = withheldChecklistRepair(directory, state, row.item_id, item)
    return adjudication ? [[row.item_id, excerpt(adjudication.claims.map(claim => claim.reason).join(" "))] as const] : []
  }))
  const details = integrity.map(row => `${row.item_id}: ${row.finding?.status ?? row.status}: ${row.finding?.reason ?? row.issues.join(" ")}${row.reinspection
    ? " One reinspection has been used for this evidence; a remaining defect or inspection gap requires changed evidence or a focused substantive fix, not another unchanged Continue." : ""}`)
  if (!result.refinement_complete) details.push(...current.effective.issues)
  if (!result.extraction_complete) details.unshift("Source extraction is incomplete or stale; not every current source item was checked.")
  // An omission gets one checklist repair round, started here without a
  // Continue: only with no integrity failure, before the run's checklist round
  // is used (a later review round does not restore it), and within the shared
  // substantive allowance. The walk restarts and this stage re-runs from the
  // repaired draft.
  const omissions = new Set(adverse.filter(row => row.class === "omission" && row.recovery === "repair" && !declined.has(row.row.item_id)).map(row => row.row.item_id))
  const round = readRepairRound(directory)
  if (!details.length && omissions.size && !repairedChecklistItems(round).size && (round?.substantive_attempts ?? 0) < LEAN_MAX_SUBSTANTIVE_REPAIRS) {
    const repairs = checklistRepairItems(directory, item_id => !omissions.has(item_id))
    if (repairs.length) {
      startRepairRound(directory, repairs, true)
      return { status: "restart" }
    }
  }
  const disclosed = [
    ...adverse.map(({ row, item, finding, class: kind }) => `${row.item_id}: ${checklistFindingLabel(directory, state, { item_id: row.item_id, item, class: kind, reinspected: !!row.reinspection })}: ${finding.reason}${declined.has(row.item_id) ? ` Adjudicator's reason: ${declined.get(row.item_id)}` : ""}`),
    ...result.parent_coverage.filter(parent => ["unsatisfied", "indeterminate"].includes(parent.status))
      .map(parent => `${parent.parent_id}: refined checks ${parent.status}: ${parent.issues.join(" ") || "A refined child remains adverse."}`),
  ]
  const report = ["# Checklist handoff findings", "", "Each result addresses its unchanged source item and declared target. Open findings remain unresolved; they are not satisfied findings. Comments, deleted text and layout may be uninspected; see each finding's scope and limitation.", "", ...result.checks.flatMap(row => [`## ${row.item_id}`, `Target: ${itemMap.get(row.item_id)?.target.document ?? "unresolved"}; treatment: ${itemMap.get(row.item_id)?.target.treatment ?? "unresolved"}`, itemMap.get(row.item_id)?.check ?? "", ...(row.reinspection ? [`Original finding: ${row.reinspection.original_finding.status}: ${row.reinspection.original_finding.reason}`, `Reinspection: ${row.reinspection.status}; ${row.reinspection.issues.join(" ")}`] : []), `Status: ${row.finding?.status ?? row.status}`, row.finding?.reason ?? row.issues.join(" "), ...(row.finding ? [`Inspection: ${row.finding.inspection_scope}`, `Limitation: ${row.finding.limitation}`] : []), ""]), ...result.parent_coverage.flatMap(parent => [`Parent: ${parent.parent_id}; coverage: ${parent.scope}; status: ${parent.status}`, `Children: ${parent.children.join(", ")}`, ...parent.issues, ""]), ...details.map(detail => `Delivery blocked: ${detail}`), ...disclosed.map(detail => `Disclosed, not satisfied: ${detail}`)].join("\n")
  writeLeanText(directory, "checklist/report.md", `${report}\n`)
  const openItems = details.map(detail => openItem("checklist-check", "omission", detail, {}))
  // Informational `open` remains in the authoritative report, rather than being
  // coerced to satisfied or to a delivery-blocking generic omission.
  return { status: details.length ? "blocked" : "done", outputs: [...new Set(["checklist/extraction.json", ...refinementOutputs, ...result.outputs, "checklist/report.md"])], openItems,
    ...(details.length ? { reason: `Checklist handoff blocked: ${details.join("; ")}. See checklist/report.md.` } : {}) }
}

/** Completed findings classed as a missing inspection that have not had their one reinspection. */
function missingInspections(directory: string, checks: ChecklistCheckResult[]) {
  const items = readEffectiveChecklistItems(directory).items
  return new Set(checks.flatMap(row => {
    const item = items.find(record => record.id === row.item_id)?.item
    return item && row.finding && !row.reinspection && classifyChecklistFinding(item, row.finding, items)?.class === "missing_inspection" ? [row.item_id] : []
  }))
}

/** One synchronous read of the matter's sources serves every input of the key. */
export function stageInputSha(directory: string, step: StepID) {
  return withSourceSnapshot(() => hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)))
}

function runStep(directory: string, state: LeanState, step: StepID, deps: LeanDeps, progress: Progress, planned: () => void | Promise<void>, prepared?: PreparedReview, reinspectChecklist = false, current = false): Promise<StageRun> {
  const stage = stepStage(step)
  const document = stepDocument(step)
  if (stage === "apply" && document) return runApply(directory, state, deps, document)
  if (stage === "gate") return runGate(directory)
  if (stage === "checklist-check") return runChecklistStage(directory, state, deps, reinspectChecklist)
  if (stage === "change-plan" && document && focusedChecklistRepairs(directory, document)) return runFocusedRepairAdjudication(directory, state, deps, document, progress)
  if (stage === "change-review" && document && focusedChecklistRepairs(directory, document)) return runFocusedChangeReview(directory, state, deps, document, progress)
  if (stage === "changes" && document && focusedChecklistRepairs(directory, document)) return runChangeReconciliation(directory, state, deps, document, progress)
  if (stage === "change-prepare" && document) return runChangePreparation(directory, state, deps, document, progress)
  if (stage === "draft" && document) return runDraftStage(directory, state, deps, progress, document, planned, current)
  return runModelStage(directory, state, step, deps, progress, prepared)
}

/** One child session; the first prompt carries the step's inputs, each retry the
 * checker's issues. On the last attempt the changes step withholds amendments
 * whose quotations still fail (they become open items); any other fourth failed
 * check stops the run with those issues. */
async function runModelStage(directory: string, state: LeanState, step: StepID, deps: LeanDeps, progress: Progress, prepared?: PreparedReview): Promise<StageRun> {
  const stage = stepStage(step) as ModelStageID
  const document = stepDocument(step) ?? undefined
  const agent = LEAN_STAGE_AGENTS[stage]
  const sessionID = prepared?.sessionID ?? await deps.sessions.createSession({
    directory, parentID: state.owner_session_id, agent, title: stepSessionTitle(directory, step), model: state.model, permissions: stage === "research" ? researchPermissions : LEAN_CHILD_PERMISSIONS,
  })
  progress({ session_id: sessionID })
  const attempt = async (index: number, text: string): Promise<StageRun> => {
    appendLeanEvent(directory, { stage, document, event: index ? "retry" : "model_call", attempt: index, session_id: sessionID })
    const asked = performance.now()
    const reply = await (index === 0 && prepared ? prepared.reply : deps.sessions.prompt({ directory, sessionID, agent, text }))
    const answered = performance.now()
    const accepted = stage === "discrepancy" || stage === "spine"
      ? acceptEvidenceReply(directory, stage, reply)
      : stage === "intake" || stage === "research" || stage === "research-review"
      ? acceptResearchReply(directory, stage, reply)
      : await acceptStageReply(directory, step, reply, { final: index >= LEAN_MAX_RETRIES })
    const timing = { wait_ms: leanElapsedMs(asked, answered), check_ms: leanElapsedMs(answered) }
    progress({ attempts: index + 1, calls: index + 1, feedback: accepted.issues })
    appendLeanEvent(directory, { stage, document, event: "checker", attempt: index, issues: accepted.issues, ...timing })
    if (!accepted.issues.length)
      return { status: (stage === "intake" || stage === "change-plan") && accepted.questions?.length ? "blocked" : "done", outputs: accepted.outputs, openItems: accepted.openItems, questions: accepted.questions,
        ...((stage === "intake" || stage === "change-plan") && accepted.questions?.length ? { reason: stage === "intake" ? "Waiting for your answers before legal research and drafting." : "Waiting for your answers before drafting amendments." } : {}) }
    if (index >= LEAN_MAX_RETRIES)
      return { status: "failed", reason: `${step}: checker failed after ${LEAN_MAX_RETRIES} retries: ${accepted.issues.join(" ")}` }
    return attempt(index + 1, retryPrompt(stage, accepted.issues))
  }
  return attempt(0, stage === "change-plan" && document ? changePlanPrompt(directory, document) : stage === "change-review" && document ? changeReviewPrompt(directory, document) : stage === "package-review" ? `${packageReviewPrompt(directory)}\n\nFocused checklist findings are saved at .dochaus/lean/checklist/report.md and checklist/checks.json; use targeted read-record lookup when relevant. Open findings remain unresolved, not satisfied. In that report, a "Disclosed, not satisfied" line is an unresolved finding: it is disclosed with the review draft, is not a blocker for this review to clear, and is never treated as satisfied. Do not re-run every checklist item in this package review.` : stage === "discrepancy" || stage === "spine" ? evidencePrompt(directory, stage) : stage === "intake" || stage === "research" || stage === "research-review" ? researchPrompt(directory, stage) : stagePrompt(directory, step))
}

/** The state's step rows follow the plan, whether the first draft wrote it or
 * a lawyer edited it. Summary rows of a title no longer planned are dropped.
 * When the planned documents differ from the rows, the open items and open
 * questions of a document no longer planned go with its steps; a question
 * another document's item still asks stays with it. */
function followPlan(directory: string, state: LeanState) {
  const steps = leanSteps(directory)
  forgetUnplannedDocuments(directory)
  const documents = (rows: string[]) => [...new Set(rows.flatMap(step => stepDocument(step as StepID) ?? []))].join(" ")
  if (documents(Object.keys(state.stages)) === documents(steps)) return syncLeanSteps(directory, state, steps)
  const planned = new Set<string>(steps.flatMap(step => stepDocument(step) ?? []))
  const kept = (document: string | undefined) => !document || planned.has(document)
  writeOpenItems(directory, readOpenItems(directory).filter(item => kept(item.ref.document)))
  writeQuestions(directory, keepAskedQuestions(readQuestions(directory), readOpenItems(directory), question => kept(question.document)))
  return syncLeanSteps(directory, state, steps)
}

async function runGate(directory: string): Promise<StageRun> {
  // The gate's checks are synchronous: they share one source read. Its
  // questions replace the ones it asked before, or withdraw them.
  const { gate: { questions = [], ...gate }, notes } = withSourceSnapshot(() => {
    const gate = finalGate(directory)
    return { gate, notes: deliveredNotes(directory, gate) }
  })
  writeLeanText(directory, "gate.json", `${JSON.stringify(gate, null, 2)}\n`)
  writeLeanText(directory, "drafting-notes-coverage.md", notesCoverageReport(gate))
  // The drafting notes are delivered as the reader gets them, with any
  // unresolved checklist findings and disclosed requirements left open in plain
  // words: whenever any of them exists, as the gate's validity requires
  // (stageOutputValid). Notes left without a heading are delivered under a
  // plain title, so they are never empty. With no drafting notes and no
  // finding there is nothing to deliver, so earlier delivered notes are
  // removed. Rendering is not byte-stable (a list gets a fresh relationship
  // id), so unchanged notes keep their earlier file; the .docx is written
  // before the .md, so a stop between the two re-renders.
  const previous = leanFile(directory, DELIVERED_NOTES.docx)
  if (notes !== undefined) {
    if (!existsSync(previous) || readLeanText(directory, DELIVERED_NOTES.markdown) !== notes) {
      writeLeanBytes(directory, DELIVERED_NOTES.docx, (await renderMarkdownToDocx(notes, existsSync(previous) ? readFileSync(previous) : undefined)).bytes)
      writeLeanText(directory, DELIVERED_NOTES.markdown, notes)
    }
  } else for (const name of [DELIVERED_NOTES.markdown, DELIVERED_NOTES.docx]) rmSync(leanFile(directory, name), { force: true })
  if (gate.status === "pass") return { status: "done", outputs: stepOutputs("gate"), openItems: [], questions }
  return {
    status: "blocked",
    outputs: stepOutputs("gate"),
    openItems: [],
    questions,
    reason: `gate: blocked by open items ${gate.blocking.map(item => item.id).join(", ")}`,
  }
}

function now() {
  return new Date().toISOString()
}

function message(directory: string, error: unknown) {
  return matterRelative(directory, error instanceof Error ? error.message : String(error))
}
