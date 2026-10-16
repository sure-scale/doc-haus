import path from "node:path"
import type { LeanDocuments } from "../../workflows/lean/deliverables"
import type { LeanState } from "../../workflows/lean/state"

// Synthetic integration-test status adapter for the lean drafting profile. The diagnostic
// client starts one run and reads ingest's drafting status until the run is
// settled. This pure observer reports what the status says; it never continues
// the run, answers a question or judges legal content.

export type LeanDraftingStatus = LeanState & {
  in_flight: boolean
  documents: LeanDocuments | null
  plan: { documents: { id: string; title: string; purpose: string }[] } | null
}
export type LeanDraftingStart = { profile: "lean"; mode: "started"; ownerSessionID: string; status: LeanDraftingStatus }

export function isLeanDraftingStatus(value: unknown): value is LeanDraftingStatus {
  if (!record(value) || value.schema !== "dochaus.lean.state/2" || value.profile !== "lean") return false
  return ["running", "blocked", "failed", "completed"].includes(String(value.status)) &&
    typeof value.owner_session_id === "string" && typeof value.in_flight === "boolean" && record(value.stages)
}

export function isLeanDraftingStart(value: unknown): value is LeanDraftingStart {
  return record(value) && value.profile === "lean" && value.mode === "started" &&
    typeof value.ownerSessionID === "string" && Boolean(value.ownerSessionID) && isLeanDraftingStatus(value.status) &&
    value.status.owner_session_id === value.ownerSessionID
}

/** complete: the run finished and its deliverables are registered. failed and
 * blocked are the product's own stops. inactive: the status still says running
 * but no coordinator runs it (an ingest restart), so nothing will change it. */
export function classifyLeanObservation(status: LeanDraftingStatus) {
  if (status.status === "completed") return "complete"
  if (status.status === "failed") return "failed"
  if (status.status === "blocked") return "blocked"
  if (!status.in_flight) return "inactive"
  return "observing"
}

/** One line of stage progress, for the diagnostic log. */
export function leanProgressLine(status: LeanDraftingStatus) {
  const stages = Object.entries(status.stages).flatMap(([stage, row]) => row
    ? [`${stage}=${row.status}${row.attempts > 1 ? `x${row.attempts}` : ""}${row.calls > Math.max(row.attempts, 1) ? `/${row.calls}calls` : ""}`]
    : [])
  return `${status.status} ${stages.join(" ")}`
}

/** The DOCX deliverables a completed lean run registered in the matter folder. */
export function leanDeliverableDocx(status: LeanDraftingStatus) {
  const documents = status.documents?.documents
  if (status.documents?.status !== "completed" || !Array.isArray(documents))
    throw new Error("completed lean drafting status has no delivered document list")
  return matterRootDocx(documents.map(document => document.deliverable), "completed lean drafting status records no DOCX deliverables")
}

/** Matter-root `.docx` names from a status list, deduplicated; empty is an error. */
export function matterRootDocx(files: unknown[], empty: string) {
  const docx = [...new Set(files.filter((file): file is string =>
    typeof file === "string" && path.basename(file) === file && file.toLowerCase().endsWith(".docx")))]
  if (!docx.length) throw new Error(empty)
  return docx
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
