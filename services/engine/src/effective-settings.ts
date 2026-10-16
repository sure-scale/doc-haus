import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { dataPaths } from "./data-paths"

export type EngineSettings = {
  model?: string
  small_model?: string
  agent?: Record<string, Record<string, unknown>>
  permission?: Record<string, unknown>
  provider?: Record<string, Record<string, unknown>>
  [key: string]: unknown
}

export const settingsRepository = path.resolve(import.meta.dir, "../../..")
export function settingsPaths(repository = settingsRepository, environment = process.env) {
  return {
    shipped: path.join(environment.DOCHAUS_DIR ?? environment.OPENCODE_CONFIG_DIR ?? path.join(repository, "dochaus"), "opencode.json"),
    overrides: path.join(dataPaths(repository, environment).runtime, "config", "opencode", "opencode.json"),
  }
}

export function isSettingsObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

// JSON maps merge by key; arrays are deliberate whole-field replacements.
// null is used only by PATCH and removes the corresponding override.
export function mergeSettings(base: Record<string, unknown>, patch: Record<string, unknown>): EngineSettings {
  const result = { ...base }
  for (const [key, value] of Object.entries(patch)) {
    if (["__proto__", "prototype", "constructor"].includes(key)) throw new Error("Invalid settings field.")
    if (value === null) delete result[key]
    else result[key] = isSettingsObject(value) ? mergeSettings(isSettingsObject(result[key]) ? result[key] : key === "skill" && typeof result[key] === "string" ? { "*": result[key] } : {}, value) : value
  }
  return result as EngineSettings
}

export function readSettingsFile(file: string): EngineSettings {
  if (!existsSync(file)) return {}
  let value: unknown
  try { value = JSON.parse(readFileSync(file, "utf8")) } catch { throw new Error("Settings file is unreadable or contains invalid JSON. Repair it before continuing.") }
  if (!isSettingsObject(value)) throw new Error("Settings file must contain an object.")
  return mergeSettings({}, value)
}

// Synchronous reads deliberately avoid location caches: model policy and plugin
// hooks see the same shipped-plus-runtime values as the settings API.
export function readEffectiveSettings(paths = settingsPaths()): EngineSettings {
  const effective = mergeSettings(readSettingsFile(paths.shipped), readSettingsFile(paths.overrides))
  for (const key of ["model", "small_model"]) if (effective[key] !== undefined && (typeof effective[key] !== "string" || !/^[^\s/]+\/[^\s]+$/.test(effective[key] as string))) throw new Error("Saved model settings are invalid. Select a provider and model in Settings.")
  return effective
}
