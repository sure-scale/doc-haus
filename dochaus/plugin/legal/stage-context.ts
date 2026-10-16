import { loadJurisdiction, readMatterJurisdictionCodes, type JurisdictionPack } from "../../lib/jurisdiction"

export type JurisdictionContextMode = "declared" | "none"
export type DraftingPreferencesMode = "full" | "explanation" | "none"

// The draft orchestrator and router route work, and the narrator restates a
// drafting run's progress, rather than make jurisdiction-sensitive judgments, so
// they receive no jurisdiction packs. The narrator receives only preferences
// for human-facing explanations; substantive stages receive preferences in full.
const CONTEXT_FREE_AGENTS = new Set(["draft", "router", "narrator"])

export function jurisdictionContextMode(agent: string | undefined): JurisdictionContextMode {
  if (agent && CONTEXT_FREE_AGENTS.has(agent)) return "none"
  return "declared"
}

export function draftingPreferencesMode(agent: string | undefined): DraftingPreferencesMode {
  if (agent === "narrator") return "explanation"
  if (agent && CONTEXT_FREE_AGENTS.has(agent)) return "none"
  return "full"
}

export async function loadStageJurisdictions(directory: string, mode: JurisdictionContextMode): Promise<JurisdictionPack[]> {
  if (mode === "none") return []
  return (await Promise.all(readMatterJurisdictionCodes(directory).map(loadJurisdiction))).filter(
    (pack): pack is JurisdictionPack => Boolean(pack),
  )
}
