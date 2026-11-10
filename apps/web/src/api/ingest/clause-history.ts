import type { ClauseHistoryRequest, ClauseHistoryResult, ClauseHistorySourceRequest, ClauseHistorySourceResult } from "../../../../../dochaus/lib/clause-history-contract"
import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

async function historyRequest<T>(matterId: string, name: string, route: string, input: object, signal?: AbortSignal): Promise<T> {
  const response = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(matterId)}/documents/clause-history${route}?name=${encodeURIComponent(name)}`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(input), signal,
  })
  if (!response.ok) throw new Error(await ingestError(response, "Could not read saved clause history"))
  return response.json()
}

export function fetchClauseHistory(matterId: string, name: string, request: ClauseHistoryRequest, signal?: AbortSignal): Promise<ClauseHistoryResult> {
  return historyRequest(matterId, name, "", request, signal)
}

export function fetchClauseHistorySource(matterId: string, name: string, request: ClauseHistorySourceRequest, signal?: AbortSignal): Promise<ClauseHistorySourceResult> {
  return historyRequest(matterId, name, "/source", request, signal)
}
