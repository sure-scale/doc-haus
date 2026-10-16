import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

export type TelemetryLean = {
  current_activity: { cost_usd: number; turns: number; since_ts: number } | null
  today: { cost_usd: number; turns: number }
  matter_total: { cost_usd: number; turns: number; sessions: number; since_ts: number }
  current_session: { tokens: number; status: "ok" | "warn" | "fail"; agent: string; model: string } | null
  rate_card_as_of: string | null
  unknown_models: string[]
}

export type TelemetryDetail = {
  by_agent: Array<{ agent: string; cost_usd: number; turns: number; tool_calls: number }>
  by_model: Array<{ model: string; cost_usd: number; turns: number }>
  by_tool: Array<{ tool: string; calls: number; total_duration_ms: number; avg_duration_ms: number }>
  recent_turns: Array<{ ts: number; agent: string; model: string; tokens: number; cost_usd: number | null; status: string }>
  recent_sessions: Array<{ id: string; agent: string; status: string; last_activity_at: number; cost_usd: number }>
  sparkline: Array<{ ts: number; cost_usd: number }>
}

export type TelemetryStorage = {
  size_bytes: number
  matter_count: number
  oldest_ts: number | null
  unknown_models: string[]
  rate_card_as_of: string | null
  rate_source: string
  rate_reference_url: string
}

export function getTelemetry(matterId: string, session?: string): Promise<TelemetryLean> {
  const qs = session ? `?session=${encodeURIComponent(session)}` : ""
  return readJson(`${INGEST_URL}/matters/${matterId}/telemetry${qs}`, "Could not load usage")
}

export function getTelemetryDetail(matterId: string): Promise<TelemetryDetail> {
  return readJson(`${INGEST_URL}/matters/${matterId}/telemetry/detail`, "Could not load usage detail")
}

export function getTelemetryStorage(): Promise<TelemetryStorage> {
  return readJson(`${INGEST_URL}/telemetry/storage`, "Could not load storage usage")
}

export async function clearTelemetry(matterId: string): Promise<{ ok: boolean }> {
  const res = await fetch(`${INGEST_URL}/matters/${matterId}/telemetry`, { method: "DELETE" })
  if (!res.ok) throw new Error(`Failed to clear telemetry (${res.status})`)
  return res.json()
}

async function readJson(url: string, fallback: string) {
  const res = await fetch(url)
  if (!res.ok) throw new Error(await ingestError(res, fallback))
  return res.json()
}
