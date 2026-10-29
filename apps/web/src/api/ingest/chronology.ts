import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"
import type {
  Chronology,
  ChronologyAction,
  ChronologyDecision,
  ChronologyEvent,
  ChronologyPrivilege,
} from "../../../../../dochaus/lib/chronology"
export type {
  Chronology,
  ChronologyAction,
  ChronologyAccount,
  ChronologyCoverage,
  ChronologyDate,
  ChronologyDecision,
  ChronologyEvent,
  ChronologyEvidence,
  ChronologyPrivilege,
  ChronologyProposal,
} from "../../../../../dochaus/lib/chronology"
export { chronologyCounters, chronologyNeedsAttention } from "../../../../../dochaus/lib/chronology"

export type ChronologyPreflight = {
  sources: {
    sourceId: string
    documentName: string
    version: string
    estimatedSections: number
    kind: "source" | "supplied-chronology"
    extraction: "pending" | "readable" | "degraded" | "unreadable" | "unavailable"
    detail?: string
  }[]
  model: { providerID: string; modelID: string } | null
  framing: string
  framingKnown: boolean
  estimatedSections: number
  workExplanation: string
}
export class ChronologyConflict extends Error {
  constructor(
    public current?: Chronology,
    detail?: string,
  ) {
    super(detail ?? "The chronology changed elsewhere. Review the latest version before trying again.")
  }
}
async function request<T>(id: string, suffix = "", body?: unknown): Promise<T> {
  const response = await fetch(
    `${INGEST_URL}/matters/${encodeURIComponent(id)}/chronology${suffix}`,
    body === undefined
      ? undefined
      : {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        },
  )
  if (response.status === 409) {
    const data = (await response.json().catch(() => ({}))) as { chronology?: Chronology; error?: string }
    throw new ChronologyConflict(data.chronology, typeof data.error === "string" ? data.error : undefined)
  }
  if (!response.ok) throw new Error(await ingestError(response, "Could not update the chronology."))
  return response.json() as Promise<T>
}
export const getChronology = (id: string) => request<Chronology>(id)
export const assessChronology = (id: string) => request<Chronology>(id, "/assess", {})
export const chronologyPreflight = (id: string) => request<ChronologyPreflight>(id, "/preflight")
export const chronologyAction = (id: string, expectedRevision: number, action: ChronologyAction) =>
  request<Chronology>(id, "/actions", { expectedRevision, action })
export const previewChronologyAction = (id: string, expectedRevision: number, action: ChronologyAction) =>
  request<{ preview: ChronologyEvent[] }>(id, "/actions", { expectedRevision, action: { ...action, preview: true } })
export const decideChronologyProposal = (
  id: string,
  expectedRevision: number,
  proposalId: string,
  decision: ChronologyDecision,
) => request<Chronology>(id, `/proposals/${encodeURIComponent(proposalId)}/decision`, { expectedRevision, decision })
export const batchChronologyAdditions = (id: string, expectedRevision: number, proposalIds: string[]) =>
  request<Chronology>(id, "/proposals/batch", { expectedRevision, proposalIds })
export const runChronology = (
  id: string,
  expectedRevision: number,
  privilege: ChronologyPrivilege,
  mode: "initial" | "refresh" | "resume",
) => request<Chronology>(id, "/run", { expectedRevision, privilege, mode, requestId: crypto.randomUUID() })
export const stopChronology = (id: string, expectedRevision: number) =>
  request<Chronology>(id, "/run/stop", { expectedRevision })
