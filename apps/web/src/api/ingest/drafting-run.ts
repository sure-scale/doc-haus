import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"
import { leanDraftingStart, leanStartRefusal, type LeanDraftingStart } from "./lean"

export type ContinuationDispatchInput = {
  kind: "malformed-tool"
  progressToken: string
  sessionID: string
}

export type ContinuationDispatchResult = {
  accepted: boolean
  attempted: boolean
  existing: boolean
  executing: boolean
  persisted: boolean
  uncertain: boolean
  ownerState?: "idle" | "busy" | "retry" | "unavailable"
}

export async function startDraftingRun(
  matterId: string,
  input: {
    request: string
    agent: string
    model: { providerID: string; modelID: string }
  },
): Promise<LeanDraftingStart> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/drafting-runs/start`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  })
  if (!res.ok)
    throw new Error(leanStartRefusal(await res.clone().json().catch(() => undefined)) ?? (await ingestError(res, "Failed to start drafting run")))
  return leanDraftingStart(await res.json(), input.model)
}

export async function dispatchMatterContinuation(
  matterId: string,
  input: ContinuationDispatchInput,
): Promise<ContinuationDispatchResult> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/continuations/dispatch`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Failed to dispatch workflow continuation"))
  const result = (await res.json()) as ContinuationDispatchResult
  if (
    !result ||
    ["accepted", "attempted", "existing", "executing", "persisted", "uncertain"].some(
      (field) => typeof result[field as keyof ContinuationDispatchResult] !== "boolean",
    ) ||
    (result.ownerState !== undefined && !["idle", "busy", "retry", "unavailable"].includes(result.ownerState))
  ) throw new Error("Workflow continuation dispatch response is invalid.")
  return result
}
