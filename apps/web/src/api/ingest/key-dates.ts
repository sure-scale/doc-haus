import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

// A key-date row recorded during a review. `date` is a resolved ISO calendar
// date; `formula` is the document's verbatim relative wording, kept separate
// because it is not a resolved date. Fields the review could not establish stay
// null rather than becoming a guess.
export type KeyDateRegisterRow = {
  id: number
  document: string
  kind: "renewal" | "termination_notice" | "expiry" | "payment" | "deliverable" | "other"
  date: string | null
  formula: string | null
  noticeDays: number | null
  actBy: string | null
  party: string | null
  quote: string
  reason: string
  author: string
  recordedAt: number
}

export async function getKeyDates(matterId: string, signal?: AbortSignal): Promise<KeyDateRegisterRow[]> {
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matterId)}/key-dates`, { signal })
  if (!response.ok) throw new Error(await ingestError(response, "Could not read the key-date register."))
  const data = (await response.json().catch(() => undefined)) as { rows?: unknown } | undefined
  if (!Array.isArray(data?.rows)) throw new Error("The key-date register response was not readable.")
  return data.rows as KeyDateRegisterRow[]
}
