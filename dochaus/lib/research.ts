import { readFirmPreferences } from "./preferences"
import { normalizeWebResearch } from "./drafting-preferences"
import { OFFICIAL_SOURCE_HOSTS, OFFICIAL_SOURCE_SUFFIXES } from "./research-sources"

// Web research defaults to the entire web. Firms can restrict it to official
// primary sources in Settings; the catalog also covers jurisdiction-pack sources.
// Source verification and untrusted-content handling apply in either scope.

export function isOfficialSource(url: string) {
  const host = new URL(url).hostname
  if (OFFICIAL_SOURCE_HOSTS.some((h) => host === h || host.endsWith("." + h)) || OFFICIAL_SOURCE_SUFFIXES.some((s) => host.endsWith(s))) return true
  // Firm-added approved hosts (Settings > Research) extend the built-in list,
  // matched the same way. Read from preferences.json per call so a settings save
  // applies on the next fetch, no restart.
  return approvedSources().some((h) => host === h || host.endsWith("." + h))
}

function approvedSources(): string[] {
  const list = readFirmPreferences().approvedSources
  return Array.isArray(list) ? list.filter((h): h is string => typeof h === "string" && h.trim().length > 0) : []
}

// The firm-wide web-research scope set in the web settings. Lives in the same
// preferences.json the ingest service writes for drafting preferences; read
// from disk on every fetch so a settings save applies immediately, no restart.
export function researchScope(): "official" | "open" {
  return normalizeWebResearch(readFirmPreferences())
}

export function allowRedirect(from: string, to: string): boolean {
  if (researchScope() === "open") return true
  const fromHost = new URL(from).hostname
  const toHost = new URL(to).hostname
  if (toHost === fromHost || toHost.endsWith("." + fromHost)) return true
  if (fromHost.endsWith("." + toHost)) return true
  return isOfficialSource(to)
}

// The firm's optional Exa key (Settings > Drafting). Same per-call disk read
// as researchScope, so a saved key applies on the next call. Without a key,
// search uses Exa's rate-limited keyless MCP; /contents retrieval still needs one.
export function readSearchKey(): string | undefined {
  const key = readFirmPreferences().searchApiKey
  if (typeof key === "string" && key.trim()) return key.trim()
  return process.env.EXA_API_KEY?.trim() || undefined
}

// The firm's CourtListener API token (Settings > Research). Same per-call disk
// read as readSearchKey, so pasting a token in Settings raises the case-law rate
// limit on the very next turn. Undefined until the firm configures one — the
// case-law tool runs unauthenticated until then, never failing for lack of it.
export function readCaseLawKey(): string | undefined {
  const key = readFirmPreferences().caseLawApiKey
  return typeof key === "string" && key.trim() ? key.trim() : undefined
}
