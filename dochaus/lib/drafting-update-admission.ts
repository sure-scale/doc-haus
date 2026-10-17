/** The reason a run the lawyer stopped records. A stopped run is finished: it
 * settles as failed, takes no further Continue, answer or update, and the next
 * drafting request archives it and starts a new run. */
export const DRAFTING_STOPPED_REASON = "Stopped at your request."

export function draftingRunStopped(run: { status: string; stop_reason?: string | null }) {
  return run.status === "failed" && run.stop_reason === DRAFTING_STOPPED_REASON
}

/** Shared browser/API admission policy. A held checklist may be clarified for
 * inspection, but its existing documents cannot receive a late general edit. */
export function draftingUpdateAdmission(run: {
  status: string
  /** A paused run stopped mid-pass; its checklist row is not a settled hold. */
  paused_at?: string | null
  stages: Record<string, { status: string } | undefined>
  in_flight: boolean
  pending_requests: number
}): "update" | "clarification" | "closed" {
  if (run.status === "completed") return "closed"
  const late = Object.entries(run.stages).some(([step, value]) =>
    (step.endsWith(".apply") || step === "finalize" || step === "gate") && value && value.status !== "pending")
  if (!late) return "update"
  if (run.status === "blocked" && !run.paused_at && run.stages["checklist-check"]?.status === "blocked" &&
    !run.in_flight && run.pending_requests === 0) return "clarification"
  return "closed"
}
