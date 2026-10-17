import { INGEST_URL } from "../../config"

// The lean drafting profile as ingest serves it. Its drafting status is the run's
// state.json plus the coordinator's projection (in_flight, open questions and
// items, gate result, delivered documents). The types mirror
// dochaus/workflows/lean/{state,gate,deliverables}.ts, which the browser cannot
// import. Continue and answers queue a request for the run's own coordinator;
// Pause stops it at once.

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
  ref?: { document?: string }
}

export type LeanQuestion = {
  id: string
  stage: string
  document?: string
  question: string
  context: string
  status: "open" | "answered"
  answer?: string
}

export type LeanDocument = {
  kind: "body" | "redline" | "drafting_notes" | "drafting_notes_markdown"
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
  owner_session_id: string
  model: { providerID: string; modelID: string } | null
  created_at: string
  updated_at: string
  research?: { status: "completed" | "limited" | "not-needed"; findings: unknown[]; limitations: string[] } | null
  stages: Record<LeanStep, LeanStageStatus>
  /** True while ingest runs this matter's coordinator. */
  in_flight: boolean
  pending_requests: number
  /** Answers durably received, waiting for the coordinator to consume them. */
  queued_answer_ids?: string[]
  updates?: { kind?: "clarification"; id: string; text: string; status: "queued" | "included"; progress?: "received" | "applying" | "applied" | "requires-rework" }[]
  open_items: LeanOpenItem[]
  questions: LeanQuestion[]
  /** The final check's result; `items` holds every open item as the check settled it
   * (open = still blocking, disclosed in the drafting notes, or resolved), and
   * each checklist finding left open, disclosed as kind `checklist_open`. */
  gate: { status: "pass" | "blocked"; delivery: string | null; blocking: LeanOpenItem[]; items?: LeanOpenItem[] } | null
  documents: { status: "blocked" | "completed"; documents: LeanDocument[] } | null
  /** Delivered drafts whose matter file differs from what the final check saw
   * (`since` is when it ran). The gate's verdict above is not re-evaluated. */
  changed_after_check?: { since: string | null; documents: string[] } | null
  /** The documents the run plans to draft, in order; null before the outline. */
  plan: { documents: LeanPlannedDocument[] } | null
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
    Number.isInteger(value.pending_requests) &&
    Array.isArray(value.open_items) &&
    Array.isArray(value.questions) &&
    (value.gate === null || record(value.gate)) &&
    (documents === null || (record(documents) && Array.isArray(documents.documents))) &&
    (value.plan === null || (record(value.plan) && Array.isArray(value.plan.documents))) &&
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

export async function answerLeanQuestion(
  matterId: string,
  questionID: string,
  answer: string,
): Promise<LeanDraftingQueued> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/lean/questions/${encodeURIComponent(questionID)}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ answer }),
  })
  return leanQueued(res, "Failed to send the answer")
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
