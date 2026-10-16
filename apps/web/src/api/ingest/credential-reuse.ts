import { INGEST_URL } from "../../config"

export type ReusableConnection = { id: string; integrationID: string; label: string; authType: "key" | "oauth" }
export type ReusableSource = { id: string; label: string; status: "ready" | "unavailable"; unavailableCount: number; connections: ReusableConnection[]; message?: string }

export async function listReusableSignIns(integrationID?: string): Promise<ReusableSource[]> {
  const query = integrationID === undefined ? "" : `?integrationID=${encodeURIComponent(integrationID)}`
  const response = await fetch(`${INGEST_URL}/auth/reuse${query}`, { credentials: "omit", signal: AbortSignal.timeout(10_000) })
  if (!response.ok) throw new Error("Could not discover existing sign-ins.")
  const result = await response.json() as { sources: ReusableSource[] }
  if (!Array.isArray(result.sources)) throw new Error("Could not discover existing sign-ins.")
  return result.sources
}

export async function reuseSignIn(sourceId: string, connectionId: string, expectedIntegrationID?: string): Promise<ReusableConnection> {
  const response = await fetch(`${INGEST_URL}/auth/reuse`, {
    method: "POST", credentials: "omit", signal: AbortSignal.timeout(expectedIntegrationID === undefined ? 15_000 : 25_000),
    headers: { "Content-Type": "application/json" }, body: JSON.stringify({ sourceId, connectionId, ...(expectedIntegrationID === undefined ? {} : { expectedIntegrationID }) }),
  })
  if (!response.ok) throw new Error("The selected sign-in could not be copied. Retry or connect a provider below.")
  const result = await response.json() as { connection: ReusableConnection }
  if (!result.connection || typeof result.connection.integrationID !== "string") throw new Error("The sign-in response could not be checked. Refresh the provider catalog before retrying.")
  if (expectedIntegrationID !== undefined && result.connection.integrationID !== expectedIntegrationID) throw new Error("The copied sign-in did not match the selected provider. Refresh the provider catalog before continuing.")
  return result.connection
}
