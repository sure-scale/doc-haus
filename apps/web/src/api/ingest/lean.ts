import { INGEST_URL } from "../../config"
import { POSTURES, type Posture } from "./preferences"

// The lean drafting profile as ingest serves it. Its drafting status is the run's
// state.json plus the coordinator's projection (in_flight, open questions and
// items, gate result, delivered documents). The types mirror
// dochaus/workflows/lean/{state,gate,deliverables}.ts, which the browser cannot
// import. Continue and answers queue a request for the run's own coordinator;
// Pause stops it at once, and Stop ends a paused or held run for good.

/** Stage keys are steps in walk order: drafting and four reviews, then
 * change-plan, change-prepare, changes, change-review and apply per document;
 * finalize (drafting notes), package-review and gate finish the run. */
export type LeanStep = string

export type LeanStageStatus = {
  status: "pending" | "running" | "done" | "failed" | "blocked"
  /** Tries of the stage's hardest step; above 1 means it was retried. */
  attempts: number
  /** Model calls the stage made in all. */
  calls?: number
  /** The stage's latest native session; parallel preparation can have only
   * worker sessions, with no single stage session. */
  session_id?: string | null
  /** Native sessions for parallel workers belonging to this stage. */
  session_ids?: string[]
  /** Output files of the stage, relative to the run's directory. */
  outputs?: string[]
  /** The deterministic checker's notes on the stage's last try. */
  feedback?: string[]
  started_at: string | null
  finished_at: string | null
}

/** All native sessions represented by a stage, including older singular rows. */
export function leanStageSessions(row: LeanStageStatus | undefined) {
  return [...new Set([...(row?.session_ids ?? []), ...(row?.session_id ? [row.session_id] : [])])]
}

export type LeanOpenItem = {
  id: string
  stage: string
  kind: string
  detail: string
  status: "open" | "resolved" | "disclosed"
  /** A `checklist_unresolved` item names its checklist item and, when known, why
   * it stayed unresolved (`checklist_class`: departure, source_conflict,
   * not_inspectable, missing_inspection or omission). */
  ref?: { document?: string; checklist_item?: string; checklist_class?: string }
}

/** What a question asks the lawyer for: a fact, what their instructions mean, a
 * professional judgment, the package's scope, a conflict between sources, or a
 * file's name. */
export type LeanQuestionKind = "fact" | "instruction" | "judgment" | "scope" | "conflict" | "file_name"

/** One answer the lawyer may pick, in the order the run gives them. */
export type LeanQuestionOption = { id: string; text: string }

/** The run's advice on an open question: the option it recommends, why, and the
 * passages that support it (`source_id` is "request" for the lawyer's own
 * instructions, else a matter file's name or a research URL); or why it
 * recommends none. `options` are the advice's own, absent when the question
 * carries options. It is never an answer. */
export type LeanQuestionAdvice =
  | { options?: LeanQuestionOption[]; recommendation: { option: string; rationale: string; basis: { source_id: string; quote: string }[] } }
  | { options?: LeanQuestionOption[]; unavailable: { reason: string; reason_code?: string } }

export type LeanQuestion = {
  id: string
  stage: string
  document?: string
  question: string
  context: string
  status: "open" | "answered"
  answer?: string
  kind?: LeanQuestionKind
  /** A scope question that also asks for missing facts or approvals. */
  mixed?: true
  /** The answers the question itself offers, beside the lawyer's own words. */
  options?: LeanQuestionOption[]
  advice?: LeanQuestionAdvice
}

export type LeanDocument = {
  /** "combined": a combined document's one file, saved in the matter at delivery. */
  kind: "body" | "redline" | "drafting_notes" | "drafting_notes_markdown" | "combined"
  /** Planned document id ("d1"...); null for the drafting notes, which cover the run. */
  document: string | null
  title: string | null
  path: string
  sha256: string
  deliverable: string | null
  /** The deliverable holds the lawyer's accepted Word revision; this output was not registered over it. */
  conflict?: "accepted_revision"
}

export type LeanPlannedDocument = { id: string; title: string; purpose: string }

export type LeanDraftingStatus = {
  schema: "dochaus.lean.state/2"
  profile: "lean"
  revision: number
  status: "running" | "blocked" | "failed" | "completed"
  stop_reason: string | null
  /** Set while the run is blocked only because the lawyer paused it. */
  paused_at?: string
  /** Set while the run is failed only because the model provider or setup needs
   * the lawyer in Settings (dochaus/lib/provider-stop.ts); Continue resumes. */
  provider_stop?: "auth" | "setup" | "model"
  owner_session_id: string
  model: { providerID: string; modelID: string } | null
  created_at: string
  updated_at: string
  research?: { status: "completed" | "limited" | "not-needed"; findings: unknown[]; limitations: string[] } | null
  stages: Record<LeanStep, LeanStageStatus>
  /** True while ingest runs this matter's coordinator. */
  in_flight: boolean
  /** The lawyer stopped this run; derived by ingest at read time. `stop_reason` is display only. */
  stopped?: boolean
  pending_requests: number
  /** Answers durably received, waiting for the coordinator to consume them. */
  queued_answer_ids?: string[]
  updates?: { kind?: "clarification"; id: string; text: string; status: "queued" | "included"; progress?: "received" | "applying" | "applied" | "requires-rework" }[]
  open_items: LeanOpenItem[]
  questions: LeanQuestion[]
  /** The final check's result; `items` holds every open item as the check settled it
   * (open = still blocking, disclosed in the drafting notes, or resolved), and
   * each checklist finding left open, disclosed as kind `checklist_open`.
   * `delivery` is `execution_ready`, `review_draft`, or `review_draft_with_findings`:
   * a review draft delivered with checklist requirements it does not satisfy,
   * each disclosed as kind `checklist_unresolved`. */
  gate: {
    status: "pass" | "blocked"
    delivery: string | null
    blocking: LeanOpenItem[]
    disclosed?: LeanOpenItem[]
    items?: LeanOpenItem[]
  } | null
  documents: { status: "blocked" | "completed"; documents: LeanDocument[] } | null
  /** Delivered drafts whose matter file differs from what the final check saw
   * (`since` is when it ran); `missing` is the subset no longer in the matter.
   * The gate's verdict above is not re-evaluated. */
  changed_after_check?: { since: string | null; documents: string[]; missing: string[] } | null
  /** The documents the run plans to draft, in order, whom the lawyer acts
   * for (`stated` false, with no party, when the request did not say) and the
   * Posture the run drafts to; null before the outline, `acting_for` and
   * `posture` absent from an older ingest. */
  plan: { documents: LeanPlannedDocument[]; acting_for?: { party: string; stated: boolean }; posture?: Posture } | null
  /** The name the plan gives the memo the lawyer asked for, which the run's
   * drafting notes are written and delivered as; null when none was requested,
   * absent from an older ingest. */
  memo?: string | null
  /** The one file (`file`, "<name>.docx" or its next free name) a combined
   * document is delivered as: the requested memo, then the body. `document` is
   * the planned document's id, and `body` and `notes` the exact names its two
   * parts are registered under (the body as "<name> - clean" or its next free
   * name). Ingest reports it only once both parts are delivered; the file is
   * saved in the matter at delivery and rebuilt when a change is accepted into
   * a part. Null otherwise, absent from an older ingest. */
  combined_output?: { document: string; name: string; file: string; body: string; notes: string } | null
  /** The lawyer's drafting request as the run recorded it. */
  request?: string | null
  /** Explicit adverse original requirements that can be checked against the full draft. */
  checklist_refinement?: {
    available: boolean
    parents: {
      parent_id: string
      check: string
      target: { document: string; role: string; treatment: string }
      finding: { status: "unsatisfied" | "indeterminate"; reason: string }
      scopes: "target_body"[]
    }[]
  }
}

export type LeanDraftingStart = { profile: "lean"; mode: "started"; ownerSessionID: string; status: LeanDraftingStatus }
export type LeanDraftingQueued = { profile: "lean"; mode: "queued"; request_id: string; status: LeanDraftingStatus }
/** "pausing" while the run is still stopping; its status then follows. */
export type LeanDraftingPaused = { profile: "lean"; mode: "paused" | "pausing"; status: LeanDraftingStatus }
export type LeanDraftingStopped = { profile: "lean"; mode: "stopped"; status: LeanDraftingStatus }

const RUN_STATUSES = ["running", "blocked", "failed", "completed"]
const STAGE_STATUSES = ["pending", "running", "done", "failed", "blocked"]

export function isLeanDraftingStatus(value: unknown): value is LeanDraftingStatus {
  if (!record(value) || value.schema !== "dochaus.lean.state/2" || value.profile !== "lean") return false
  const stages = value.stages
  const documents = value.documents
  return (
    RUN_STATUSES.includes(String(value.status)) &&
    (value.stop_reason === null || typeof value.stop_reason === "string") &&
    typeof value.owner_session_id === "string" &&
    typeof value.updated_at === "string" &&
    typeof value.in_flight === "boolean" &&
    (value.stopped === undefined || typeof value.stopped === "boolean") &&
    Number.isInteger(value.pending_requests) &&
    Array.isArray(value.open_items) &&
    Array.isArray(value.questions) &&
    (value.gate === null || record(value.gate)) &&
    (documents === null || (record(documents) && Array.isArray(documents.documents))) &&
    (value.plan === null ||
      (record(value.plan) &&
        Array.isArray(value.plan.documents) &&
        (value.plan.acting_for === undefined ||
          (record(value.plan.acting_for) && typeof value.plan.acting_for.party === "string" && typeof value.plan.acting_for.stated === "boolean")) &&
        (value.plan.posture === undefined || POSTURES.map((row) => row.value).includes(value.plan.posture as Posture)))) &&
    (value.memo == null || typeof value.memo === "string") &&
    (value.combined_output == null ||
      (record(value.combined_output) &&
        ["document", "name", "file", "body", "notes"].every((key) => typeof (value.combined_output as Record<string, unknown>)[key] === "string"))) &&
    record(stages) &&
    "finalize" in stages &&
    "gate" in stages &&
    Object.values(stages).every(
      (row) => record(row) && STAGE_STATUSES.includes(String(row.status)) && Number.isInteger(row.attempts),
    )
  )
}

/** The lean start response, checked against the request the browser sent. */
export function leanDraftingStart(value: unknown, model: { providerID: string; modelID: string }): LeanDraftingStart {
  if (
    !record(value) ||
    value.profile !== "lean" ||
    value.mode !== "started" ||
    typeof value.ownerSessionID !== "string" ||
    !value.ownerSessionID ||
    !isLeanDraftingStatus(value.status) ||
    value.status.owner_session_id !== value.ownerSessionID ||
    value.status.model?.providerID !== model.providerID ||
    value.status.model?.modelID !== model.modelID
  )
    throw new Error("Drafting run start response is invalid.")
  return value as LeanDraftingStart
}

/** The lawyer-facing reason ingest refused a new drafting request because the
 * matter's lean run is still open; undefined for any other refusal. A run that
 * completed or failed never refuses: ingest archives it and starts afresh. */
export function leanStartRefusal(body: unknown) {
  if (!record(body)) return undefined
  if (body.code === "lean_run_in_progress")
    return "Drafting is still running for this matter. Wait until it finishes to start a new one."
  if (body.code === "lean_run_awaiting_continue")
    return "This matter's drafting run has not finished. Open its conversation to continue it, or wait until it finishes to start a new one."
  return undefined
}

/** Continue: the coordinator resumes from the first stage whose inputs changed,
 * or starts one repair round on a blocked gate. */
export async function resumeLeanDraftingRun(matterId: string): Promise<LeanDraftingQueued> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/lean/resume`, { method: "POST" })
  return leanQueued(res, "Failed to continue the drafting run")
}

/** Pause: the run stops its current work at once. Finished steps are kept and
 * Continue picks up from the step that was in progress. */
export async function pauseLeanDraftingRun(matterId: string, ownerSessionID: string): Promise<LeanDraftingPaused> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/lean/pause`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ownerSessionID }),
  })
  if (!res.ok) {
    const body = (await res.json().catch(() => undefined)) as { error?: unknown } | undefined
    throw new Error(typeof body?.error === "string" ? body.error : `Could not pause the drafting run (${res.status})`)
  }
  const body: unknown = await res.json()
  if (!record(body) || body.profile !== "lean" || (body.mode !== "paused" && body.mode !== "pausing") || !isLeanDraftingStatus(body.status))
    throw new Error("Drafting run response is invalid.")
  return body as LeanDraftingPaused
}

/** Stop: an unfinished run nothing is working on ends as failed and takes no
 * Continue. Its files stay in the matter, and the next drafting request starts a new run. */
export async function stopLeanDraftingRun(matterId: string, ownerSessionID: string): Promise<LeanDraftingStopped> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/lean/stop`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ownerSessionID }),
  })
  if (!res.ok) {
    const body = (await res.json().catch(() => undefined)) as { error?: unknown } | undefined
    throw new Error(typeof body?.error === "string" ? body.error : `Could not stop the drafting run (${res.status})`)
  }
  const body: unknown = await res.json()
  if (!record(body) || body.profile !== "lean" || body.mode !== "stopped" || !isLeanDraftingStatus(body.status))
    throw new Error("Drafting run response is invalid.")
  return body as LeanDraftingStopped
}

/** Answers a question in the lawyer's own words. */
export function answerLeanQuestion(matterId: string, questionID: string, answer: string) {
  return postLeanQuestion(matterId, questionID, { answer }, "Failed to send the answer")
}

/** Answers a question with one of the options it offers. */
export function chooseLeanOption(matterId: string, questionID: string, optionID: string) {
  return postLeanQuestion(matterId, questionID, { option_id: optionID }, "Failed to send the answer")
}

/** Decide for me: answers the question with the option its advice recommends. */
export function decideLeanQuestion(matterId: string, questionID: string) {
  return postLeanQuestion(matterId, questionID, { decide_for_me: true }, "Failed to send the answer")
}

// Ingest's 409 and 400 codes for a question that cannot take this answer, in the lawyer's words.
const QUESTION_REFUSALS: Record<string, string> = {
  not_delegable: "Decide for me isn't available for this question. Choose an option or write your own answer.",
  already_answered: "This question is already answered.",
  invalid_answer: "This answer can't be sent. Choose one of the options shown, or write an answer of up to 20,000 characters.",
}

async function postLeanQuestion(
  matterId: string,
  questionID: string,
  body: { answer: string } | { option_id: string } | { decide_for_me: true },
  fallback: string,
): Promise<LeanDraftingQueued> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/lean/questions/${encodeURIComponent(questionID)}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  })
  if (res.status === 409 || res.status === 400) {
    const refusal: unknown = await res.clone().json().catch(() => undefined)
    const message = record(refusal) && typeof refusal.code === "string" ? QUESTION_REFUSALS[refusal.code] : undefined
    if (message) throw new Error(message)
  }
  return leanQueued(res, fallback)
}

/** Ask the run to compare one explicitly selected requirement with the full draft. */
export async function refineLeanChecklist(
  matterId: string,
  ownerSessionID: string,
  parentID: string,
): Promise<LeanDraftingQueued> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/lean/checklist/refine`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ownerSessionID, parents: [{ parent_id: parentID, scope: "target_body" }] }),
  })
  return leanQueued(res, "Could not request the full draft check")
}

/** The Word files a run registered in the matter folder, once it completed. */
export function leanDeliveredDocuments(status: LeanDraftingStatus) {
  if (status.documents?.status !== "completed") return []
  return status.documents.documents.flatMap((document) => (document.deliverable ? [document.deliverable] : []))
}

async function leanQueued(res: Response, fallback: string) {
  if (!res.ok) {
    const body = (await res.json().catch(() => undefined)) as { error?: unknown } | undefined
    throw new Error(typeof body?.error === "string" ? body.error : `${fallback} (${res.status})`)
  }
  const body: unknown = await res.json()
  if (
    !record(body) ||
    body.profile !== "lean" ||
    body.mode !== "queued" ||
    typeof body.request_id !== "string" ||
    !isLeanDraftingStatus(body.status)
  )
    throw new Error("Drafting run response is invalid.")
  return body as LeanDraftingQueued
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

export async function updateLeanDraftingRun(
  matterId: string,
  ownerSessionID: string,
  text: string,
): Promise<LeanDraftingQueued> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/lean/updates`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ownerSessionID, text }),
  })
  return leanQueued(res, "Could not send the drafting update. Your text has been kept.")
}
