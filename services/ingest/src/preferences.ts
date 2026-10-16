import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { WORKSPACE_ROOT } from "./matter"
import { DEFAULT_DRAFTING, normalizeApprovedSources, normalizePreferences, renderInstructions, type DraftingPreferences } from "../../../dochaus/lib/drafting-preferences"

export { DEFAULT_DRAFTING, type DraftingPreferences, type RateOverride, type TelemetryPreferences } from "../../../dochaus/lib/drafting-preferences"

// Structured preferences are the shared source for settings and model context.
// Keep a rendered markdown export alongside them for inspection. The engine
// renders the JSON/defaults on every turn, so reading Settings need not write.
const PREFS_DIR = () => path.join(WORKSPACE_ROOT, ".preferences")
const JSON_FILE = () => path.join(PREFS_DIR(), "preferences.json")
const INSTRUCTIONS_FILE = () => path.join(PREFS_DIR(), "drafting.md")

export function readDraftingPreferences(): DraftingPreferences {
  if (!existsSync(JSON_FILE())) return DEFAULT_DRAFTING
  return normalizePreferences(readSavedPreferences())
}

export function writeDraftingPreferences(input: unknown): DraftingPreferences {
  validateResearchPreferences(input)
  const prefs = normalizePreferences({ ...readDraftingPreferences(), ...input as Record<string, unknown> })
  const instructions = renderInstructions(prefs)
  const json = JSON.stringify(prefs, null, 2) + "\n"
  mkdirSync(PREFS_DIR(), { recursive: true })
  writeFileSync(INSTRUCTIONS_FILE(), instructions)
  // Publish authority only after its inspection export succeeds. Same-directory
  // rename lets readers see either complete JSON version on every model turn.
  const temporary = `${JSON_FILE()}.${crypto.randomUUID()}.tmp`
  try {
    writeFileSync(temporary, json, { flag: "wx" })
    renameSync(temporary, JSON_FILE())
  } finally {
    rmSync(temporary, { force: true })
  }
  return prefs
}

export class PreferencesValidationError extends Error {}
export class PreferencesUnavailableError extends Error {}

function readSavedPreferences() {
  try {
    const input: unknown = JSON.parse(readFileSync(JSON_FILE(), "utf8"))
    if (!input || typeof input !== "object" || Array.isArray(input)) throw new Error("Invalid preferences")
    return input
  } catch {
    // JSON parser exceptions can quote credential-bearing source bytes.
    throw new PreferencesUnavailableError("Saved preferences could not be read. Repair the preferences file before saving changes.")
  }
}

export function validateResearchPreferences(input: unknown): asserts input is Record<string, unknown> {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new PreferencesValidationError("Preferences must be an object.")
  const patch = input as Record<string, unknown>
  if (Object.hasOwn(patch, "webResearch") && patch.webResearch !== "official" && patch.webResearch !== "open")
    throw new PreferencesValidationError("Choose approved sources or the entire web.")
  for (const key of ["searchApiKey", "caseLawApiKey"]) {
    if (Object.hasOwn(patch, key) && (typeof patch[key] !== "string" || patch[key].length > 4096 || /[\r\n\x00]/.test(patch[key])))
      throw new PreferencesValidationError("Research credentials must be a single line of text.")
  }
  if (Object.hasOwn(patch, "courtlistenerMcpEnabled") && typeof patch.courtlistenerMcpEnabled !== "boolean")
    throw new PreferencesValidationError("Choose whether to enable the CourtListener connector.")
  if (Object.hasOwn(patch, "approvedSources")) {
    if (!Array.isArray(patch.approvedSources) || patch.approvedSources.length > 100 || patch.approvedSources.some(value =>
      typeof value !== "string" || !normalizeApprovedSources([value]).length))
      throw new PreferencesValidationError("Approved sources must be bare hostnames, one per line, with no URL, path, port or wildcard.")
  }
}

export function readDraftingTelemetryRetentionPreference(): number | undefined {
  if (!existsSync(JSON_FILE())) return undefined
  const input = readSavedPreferences() as { telemetry?: { retention_days?: unknown; retention_days_source?: unknown } }
  const retention = Number(input.telemetry?.retention_days)
  if (!Number.isFinite(retention) || retention <= 0) return undefined
  if (input.telemetry?.retention_days_source === "user") return retention
  if (!input.telemetry?.retention_days_source && retention !== DEFAULT_DRAFTING.telemetry.retention_days) return retention
  return undefined
}
