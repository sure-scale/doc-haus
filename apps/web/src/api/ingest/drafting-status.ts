import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"
import { isLeanDraftingStatus, type LeanDraftingStatus } from "./lean"

const pendingDraftingStatuses = new Map<string, Promise<LeanDraftingStatus | undefined>>()

// The matter's drafting run, or undefined when the matter has none.
export function getDraftingStatus(matterId: string): Promise<LeanDraftingStatus | undefined> {
  const pending = pendingDraftingStatuses.get(matterId)
  if (pending) return pending
  // Multiple panels poll this projection. Share only the in-flight read,
  // including when a caller stops waiting; never cache a settled status.
  const request = fetch(`${INGEST_URL}/matters/${matterId}/drafting-status`)
    .then(async (res) => {
      if (!res.ok) throw new Error(await ingestError(res, "Failed to load drafting workflow status"))
      const status: unknown = await res.json()
      if (status === null) return undefined
      if (isLeanDraftingStatus(status)) return status
      throw new Error("Drafting workflow status response is invalid.")
    })
    .finally(() => { pendingDraftingStatuses.delete(matterId) })
  pendingDraftingStatuses.set(matterId, request)
  return request
}
