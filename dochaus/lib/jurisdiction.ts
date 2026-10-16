import { existsSync, readFileSync } from "node:fs"
import path from "node:path"

// A jurisdiction pack is a config-only bundle under dochaus/jurisdiction/<code>/:
//   profile.json — machine fields (code, name, citationStyle, preferredModel)
//   prompt.md    — the system-prompt fragment injected for matters in this
//                  jurisdiction (citation conventions, statutory framing, which
//                  authority is binding versus persuasive).
// Packs are swappable: a contributor adds a jurisdiction by dropping a new
// directory here, with no code change (issue #18). The loader resolves packs
// relative to this file so it works regardless of the engine's process cwd.

export type JurisdictionProfile = {
  code: string
  name: string
  citationStyle: string
  // The model a matter in this jurisdiction should run on once sovereign routing
  // (issues #11/#15) lands. Declarative today: it is recorded but not yet
  // enforced — there is only one (global) inference region, so nothing routes on
  // it. Kept here so the routing layer has the per-jurisdiction preference ready.
  preferredModel?: string | null
}

export type JurisdictionPack = JurisdictionProfile & { prompt: string }

const ROOT = path.join(import.meta.dir, "..", "jurisdiction")

// A pack code becomes a directory name joined to a filesystem path. It arrives
// from a matter's matter.json, so reject anything that could climb out of the
// jurisdiction directory before it reaches the path.
function packDir(code: string) {
  if (!/^[A-Za-z0-9][A-Za-z0-9_-]*$/.test(code)) return undefined
  return path.join(ROOT, code)
}

export async function loadJurisdiction(code: string): Promise<JurisdictionPack | undefined> {
  const dir = packDir(code)
  const profile = readJurisdictionProfile(code)
  if (!dir || !profile) return undefined
  const promptPath = path.join(dir, "prompt.md")
  if (!existsSync(promptPath)) {
    // A pack without a prompt fragment still loads (profile fields keep working),
    // but the matter loses the pack's reasoning/citation steering — make that
    // visible instead of silently injecting nothing.
    console.warn(`[jurisdiction] pack "${profile.code}" has no prompt.md (${promptPath}); injecting profile only`)
    return { ...profile, prompt: "" }
  }
  return { ...profile, prompt: await Bun.file(promptPath).text() }
}

export function readMatterJurisdictionCodes(directory: string) {
  return readMatterJurisdictions(directory)
    .filter((code, index, codes) => codes.indexOf(code) === index)
    .flatMap((code) => readJurisdictionProfile(code) ? [code] : [])
}

// The jurisdiction codes a matter carries, read from its matter.json. A matter
// can span several (a cross-border deal), so this is a list. The ingest service
// owns that file; the engine only reads it to steer reasoning.
export function readMatterJurisdictions(directory: string) {
  const file = path.join(directory, "matter.json")
  if (!existsSync(file)) return []
  return (JSON.parse(readFileSync(file, "utf8")) as { jurisdictions?: string[] }).jurisdictions ?? []
}

function readJurisdictionProfile(code: string): JurisdictionProfile | undefined {
  const dir = packDir(code)
  if (!dir || !existsSync(path.join(dir, "profile.json"))) return undefined
  const profile = JSON.parse(readFileSync(path.join(dir, "profile.json"), "utf8")) as JurisdictionProfile
  if (profile.code !== code) {
    console.warn(`[jurisdiction] pack directory "${code}" declares profile code "${profile.code}"; ignoring it`)
    return undefined
  }
  return profile
}
