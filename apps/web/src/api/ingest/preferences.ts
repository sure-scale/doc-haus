import { INGEST_URL } from "../../config"

// Firm-wide drafting preferences (Settings → Drafting). Stored on the engine
// host and rendered into standing instructions for drafting and explanation
// turns, including when drafting resumes or completed work is revisited.
export type DraftingPreferences = {
  attorney: string
  firm: string
  posture: "client-favorable" | "balanced" | "conservative"
  formality: "formal" | "plain"
  detail: "concise" | "detailed"
  strategyApproval: "off" | "interactive"
  materialDecisionCheckpoints: "off" | "interactive"
  dateFormat: "month-day-year" | "day-month-year" | "iso"
  numberStyle: "words-and-numerals" | "numerals"
  houseStyle: string
  webResearch: "official" | "open"
  searchApiKey: string
  caseLawApiKey: string
  courtlistenerMcpEnabled: boolean
  approvedSources: string[]
  telemetry: TelemetryPreferences
  // Page layout (Settings → Drafting). Flat keys: the ingest save merges top-level
  // keys only. Read when a document renders, never sent to the model. Legends
  // are the firm's own text, empty (none) by default.
  pageNumbers: "plain" | "page" | "off"
  pageNumberPosition: "center" | "right"
  draftLine: boolean
  agreementLegend: string
  adviceLegend: string
  notesMarking: boolean
}

export type RateOverride = Partial<{
  input_per_mtok: number
  output_per_mtok: number
  cache_read_per_mtok?: number
}>

export type TelemetryPreferences = {
  retention_days: number
  retention_days_source?: "user" | "default"
  rate_overrides: Record<string, RateOverride>
}

export async function getPreferences(): Promise<DraftingPreferences> {
  const res = await fetch(`${INGEST_URL}/preferences`, { signal: AbortSignal.timeout(15_000) })
  if (!res.ok) throw new Error(`Failed to load preferences (${res.status})`)
  return res.json()
}

export async function savePreferences(prefs: Partial<DraftingPreferences>): Promise<DraftingPreferences> {
  const res = await fetch(`${INGEST_URL}/preferences`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(prefs),
    signal: AbortSignal.timeout(15_000),
  })
  if (!res.ok) throw new Error(`Failed to save preferences (${res.status})`)
  return res.json()
}

export async function testResearchConnection(service: "exa" | "courtlistener" | "courtlistener-mcp"): Promise<{ status: "verified" | "failed" | "not-configured"; message: string }> {
  const res = await fetch(`${INGEST_URL}/research/${service}/test`, { method: "POST", signal: AbortSignal.timeout(15_000) })
  if (!res.ok) throw new Error(`Research connection test unavailable (${res.status}).`)
  return res.json()
}
