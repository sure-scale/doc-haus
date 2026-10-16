/** Shared browser/API admission policy. A held checklist may be clarified for
 * inspection, but its existing documents cannot receive a late general edit. */
export function draftingUpdateAdmission(run: {
  status: string
  stages: Record<string, { status: string } | undefined>
  in_flight: boolean
  pending_requests: number
}): "update" | "clarification" | "closed" {
  if (run.status === "completed") return "closed"
  const late = Object.entries(run.stages).some(([step, value]) =>
    (step.endsWith(".apply") || step === "finalize" || step === "gate") && value && value.status !== "pending")
  if (!late) return "update"
  if (run.status === "blocked" && run.stages["checklist-check"]?.status === "blocked" &&
    !run.in_flight && run.pending_requests === 0) return "clarification"
  return "closed"
}
