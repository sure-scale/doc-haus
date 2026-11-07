/** The reason a run the lawyer stopped records. A stopped run is finished: it
 * settles as failed, takes no further Continue, answer or update, and the next
 * drafting request archives it and starts a new run. */
export const DRAFTING_STOPPED_REASON = "Stopped at your request."

export function draftingRunStopped(run: { status: string; stop_reason?: string | null }) {
  return run.status === "failed" && run.stop_reason === DRAFTING_STOPPED_REASON
}

type StageRows = Record<string, { status: string } | undefined>

/** The rows the walk has reached: every row up to and including the first step
 * not done (rows are in walk order). A row after it is left from an earlier
 * pass and says nothing about this one. */
function reachedStages(stages: StageRows) {
  const rows = Object.entries(stages)
  const open = rows.findIndex(([, value]) => value?.status !== "done")
  return open < 0 ? rows : rows.slice(0, open + 1)
}

/** Whether the walk has a written draft: a document's draft step it reached is done. */
export function draftingRunDrafted(run: { stages: StageRows }) {
  return reachedStages(run.stages).some(([step, value]) => step.endsWith(".draft") && value?.status === "done")
}

/** Shared browser/API admission policy. A held checklist may be clarified for
 * inspection, but its existing documents cannot receive a late general edit.
 * Late means the walk has reached an apply, finalize or gate step. */
export function draftingUpdateAdmission(run: {
  status: string
  /** A paused run stopped mid-pass; its checklist row is not a settled hold. */
  paused_at?: string | null
  stages: StageRows
  in_flight: boolean
  pending_requests: number
}): "update" | "clarification" | "closed" {
  if (run.status === "completed") return "closed"
  const late = reachedStages(run.stages).some(([step, value]) =>
    (step.endsWith(".apply") || step === "finalize" || step === "gate") && value && value.status !== "pending")
  if (!late) return "update"
  if (run.status === "blocked" && !run.paused_at && run.stages["checklist-check"]?.status === "blocked" &&
    !run.in_flight && run.pending_requests === 0) return "clarification"
  return "closed"
}
