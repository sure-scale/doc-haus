import { chmodSync, closeSync, fsyncSync, mkdirSync, openSync, renameSync, unlinkSync, writeFileSync } from "node:fs"
import path from "node:path"
import { randomUUID } from "node:crypto"
import { isSettingsObject, mergeSettings, readEffectiveSettings, readSettingsFile, settingsPaths, type EngineSettings } from "../../engine/src/effective-settings"

export function readEngineConfig(): EngineSettings { return readEffectiveSettings() }
export const engineConfigFile = () => settingsPaths().overrides

// Ingest is the sole writer. These synchronous read/merge/rename operations run
// to completion on its event loop, so concurrent HTTP requests cannot lose edits.
export function writeEngineOverrides(config: EngineSettings, file = engineConfigFile()) {
  mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 })
  chmodSync(path.dirname(file), 0o700)
  const temporary = `${file}.${randomUUID()}.tmp`
  let descriptor: number | undefined
  try {
    descriptor = openSync(temporary, "wx", 0o600)
    writeFileSync(descriptor, JSON.stringify(config, null, 2) + "\n")
    fsyncSync(descriptor)
    closeSync(descriptor)
    descriptor = undefined
    renameSync(temporary, file)
  } finally {
    if (descriptor !== undefined) closeSync(descriptor)
    try { unlinkSync(temporary) } catch {}
  }
}

function object(value: unknown, message = "Settings must be an object."): Record<string, unknown> {
  if (!isSettingsObject(value)) throw new Error(message)
  return value
}
function keys(value: Record<string, unknown>, allowed: string[]) {
  if (Object.keys(value).some(key => !allowed.includes(key))) throw new Error("Unsupported engine setting.")
}
function identifier(value: string) {
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(value)) throw new Error("Invalid settings identifier.")
}
function text(value: unknown) {
  if (typeof value !== "string" || !value.trim() || value.length > 2048 || /[\r\n\0]/.test(value)) throw new Error("Invalid setting value.")
}
function model(value: unknown) {
  if (value === null) return
  text(value)
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]*\/[^\s]+$/.test(value as string)) throw new Error("Select a provider and model.")
}
function endpoint(value: unknown) {
  text(value)
  let url: URL
  try { url = new URL(value as string) } catch { throw new Error("Provide an HTTP or HTTPS endpoint.") }
  if (!["http:", "https:"].includes(url.protocol) || url.username || url.password || url.search || url.hash) throw new Error("Provide an HTTP or HTTPS endpoint without credentials, query or fragment.")
}
const cloudOptions: Record<string, string[]> = {
  "google-vertex": ["project", "location"],
  "google-vertex-anthropic": ["project", "location"],
  "amazon-bedrock": ["region", "profile", "endpoint"],
  azure: ["resourceName", "baseURL", "apiVersion"],
  "azure-openai": ["resourceName", "baseURL", "apiVersion"],
}

export function validateSettingsPatch(patch: unknown): Record<string, unknown> {
  const value = object(patch)
  keys(value, ["model", "small_model", "disabled_providers", "provider", "agent", "permission"])
  for (const key of ["model", "small_model"]) if (key in value) model(value[key])
  if ("disabled_providers" in value && value.disabled_providers !== null) {
    if (!Array.isArray(value.disabled_providers)) throw new Error("Disabled providers must be a list of provider IDs.")
    value.disabled_providers.forEach(id => { if (typeof id !== "string") throw new Error("Invalid provider ID."); identifier(id) })
  }
  if ("agent" in value) for (const [id, raw] of Object.entries(object(value.agent))) {
    identifier(id)
    if (raw === null) continue
    const entry = object(raw); keys(entry, ["model", "disable"])
    if ("model" in entry) model(entry.model)
    if ("disable" in entry && entry.disable !== null && typeof entry.disable !== "boolean") throw new Error("Agent disable must be boolean.")
  }
  if ("permission" in value) {
    const permission = object(value.permission); keys(permission, ["skill"])
    if ("skill" in permission) for (const [id, action] of Object.entries(object(permission.skill))) {
      identifier(id)
      if (action !== null && action !== "deny" && action !== "allow") throw new Error("Invalid skill toggle.")
    }
  }
  if ("provider" in value) for (const [id, raw] of Object.entries(object(value.provider))) {
    identifier(id)
    if (raw === null) continue
    const provider = object(raw); keys(provider, ["name", "npm", "options", "models"])
    if ("name" in provider && provider.name !== null) text(provider.name)
    if ("npm" in provider && provider.npm !== null && provider.npm !== "@ai-sdk/openai-compatible") throw new Error("Only the bundled OpenAI-compatible provider is supported.")
    if ("options" in provider && provider.options !== null) {
      const options = object(provider.options); keys(options, [...(cloudOptions[id] ?? []), "baseURL"])
      for (const [key, option] of Object.entries(options)) {
        if (option === null) continue
        if (["baseURL", "endpoint"].includes(key)) endpoint(option)
        else text(option)
      }
    }
    if ("models" in provider && provider.models !== null) for (const [modelID, rawModel] of Object.entries(object(provider.models))) {
      if (!modelID.trim() || /[\s\0]/.test(modelID) || modelID.length > 256) throw new Error("Invalid model ID.")
      if (rawModel === null) continue
      const entry = object(rawModel); keys(entry, ["name"])
      if ("name" in entry && entry.name !== null) text(entry.name)
    }
  }
  // Also rejects prototype keys at every depth before touching disk.
  mergeSettings({}, value)
  return value
}

export function patchEngineConfig(patch: unknown) {
  const value = validateSettingsPatch(patch)
  writeEngineOverrides(mergeSettings(readSettingsFile(engineConfigFile()), value))
  return readEngineSettings()
}
export function disabledSkills(): Set<string> {
  const skill = readEngineConfig().permission?.skill
  return new Set(Object.entries(isSettingsObject(skill) ? skill : {}).filter(([name, action]) => name !== "*" && action === "deny").map(([name]) => name))
}
export function setSkillDisabled(name: string, disabled: boolean) { patchEngineConfig({ permission: { skill: { [name]: disabled ? "deny" : null } } }) }
export function disabledAgents(): Set<string> {
  return new Set(Object.entries(readEngineConfig().agent ?? {}).filter(([, agent]) => agent.disable === true).map(([name]) => name))
}
export function setAgentDisabled(name: string, disabled: boolean) {
  if (disabled) { patchEngineConfig({ agent: { [name]: { disable: true } } }); return }
  const overrides = readSettingsFile(engineConfigFile())
  const entry = overrides.agent?.[name]
  if (entry) { delete entry.disable; if (!Object.keys(entry).length) delete overrides.agent![name] }
  writeEngineOverrides(overrides)
}

// Browser reads expose settings only, never policy, module paths or secrets.
export function readEngineSettings(): EngineSettings {
  const effective = readEngineConfig()
  const result: EngineSettings = {}
  for (const key of ["model", "small_model", "disabled_providers"]) if (effective[key] !== undefined) result[key] = effective[key]
  if (effective.agent) result.agent = Object.fromEntries(Object.entries(effective.agent).flatMap(([id, entry]) => {
    const fields = Object.fromEntries(Object.entries(entry).filter(([key]) => ["model", "disable"].includes(key)))
    return Object.keys(fields).length ? [[id, fields]] : []
  }))
  if (effective.permission?.skill) result.permission = { skill: effective.permission.skill }
  if (effective.provider) result.provider = Object.fromEntries(Object.entries(effective.provider).map(([id, entry]) => {
    const fields: Record<string, unknown> = {}
    if (typeof entry.name === "string") fields.name = entry.name
    if (entry.npm === "@ai-sdk/openai-compatible") fields.npm = entry.npm
    if (isSettingsObject(entry.options)) fields.options = Object.fromEntries(Object.entries(entry.options).filter(([key]) => [...(cloudOptions[id] ?? []), "baseURL"].includes(key)))
    if (isSettingsObject(entry.models)) fields.models = Object.fromEntries(Object.entries(entry.models).map(([name, model]) => [name, isSettingsObject(model) && typeof model.name === "string" ? { name: model.name } : {}]))
    return [id, fields]
  }))
  return result
}
