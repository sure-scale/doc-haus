import { INGEST_URL } from "../../config"
import { ingestError } from "./errors"

// Costs are provider-reported per model step, read from native sessions.
export type UsageGroup = "background" | "work"

export type TelemetryLean = {
  current_activity: { cost_usd: number; turns: number; since_ts: number } | null
  today: { cost_usd: number; turns: number }
  matter_total: { cost_usd: number; turns: number; sessions: number; since_ts: number }
  current_session: { tokens: number; status: "ok" | "warn" | "fail"; agent: string; label: string; model: string } | null
  unpriced_turns: number
  partial: boolean
}

export type TelemetryDetail = {
  by_group: Array<{ group: UsageGroup; label: string; cost_usd: number; turns: number }>
  by_agent: Array<{ agent: string; label: string; group: UsageGroup; cost_usd: number; turns: number; tool_calls: number }>
  // Conversation upkeep (titles, compaction) has no model; its row carries a
  // plain label so the rows add up to the matter total.
  by_model: Array<{ model: string | null; label?: string; cost_usd: number; turns: number; unpriced_turns: number }>
  by_tool: Array<{ tool: string; label: string; calls: number; total_duration_ms: number; avg_duration_ms: number }>
  recent_turns: Array<{ ts: number; agent: string; label: string; group: UsageGroup; model: string | null; tokens: number; cost_usd: number | null }>
  recent_sessions: Array<{ id: string; title: string; agent: string; label: string; group: UsageGroup; last_activity_at: number; cost_usd: number }>
  sparkline: Array<{ ts: number; cost_usd: number }>
}

export type TelemetryStorage = {
  size_bytes: number
  matter_count: number
  oldest_ts: number | null
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
