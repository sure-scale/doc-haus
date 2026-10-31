import type { ChronologyCoverage, ChronologyEvent } from "./chronology"

/** Export is a dated view of saved entries. Proposal text is never an event. */
export type ChronologyExportRequest = {
  expectedRevision: number
  scope: "all-visible" | "filtered"
  eventIds?: string[]
  includeHidden: boolean
}

export type ChronologyExportSnapshot = {
  schemaVersion: 1
  matterId: string
  title: string
  snapshotAt: number
  revision: number
  scope: ChronologyExportRequest["scope"]
  includeHidden: boolean
  events: ChronologyEvent[]
  coverage: ChronologyCoverage[]
  totalEvents: number
  excludedHidden: number
  pendingProposals: number
  warnings: string[]
}

/** This identifies derived work, never authority to read or amend a matter. */
export const CHRONOLOGY_EXPORT_PROPERTY = "doc.haus.origin"
export const CHRONOLOGY_EXPORT_ORIGIN = "chronology-snapshot-v1"
export const CHRONOLOGY_EXPORT_LABEL = "Chronology export snapshot"

export function validChronologyExportRequest(value: unknown): value is ChronologyExportRequest {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const input = value as Record<string, unknown>
  if (
    !Number.isSafeInteger(input.expectedRevision) ||
    (input.expectedRevision as number) < 0 ||
    !["all-visible", "filtered"].includes(input.scope as string) ||
    typeof input.includeHidden !== "boolean"
  ) return false
  if (input.scope === "all-visible") return input.eventIds === undefined
  return Array.isArray(input.eventIds) && input.eventIds.length <= 50_000 && input.eventIds.every(
    (id) => typeof id === "string" && id.length > 0 && id.length <= 1000,
  ) && new Set(input.eventIds).size === input.eventIds.length
}
