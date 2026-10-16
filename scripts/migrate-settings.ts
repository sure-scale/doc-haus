import { chmodSync, copyFileSync, existsSync, lstatSync, mkdirSync, writeFileSync } from "node:fs"
import path from "node:path"
import { mergeSettings, readSettingsFile, type EngineSettings } from "../services/engine/src/effective-settings"
import { validateSettingsPatch, writeEngineOverrides } from "../services/ingest/src/engine-config"

// Explicit one-off only. Dry-run prints field names, never source values.
export function migrateSettings(input: { source: string; runtime: string; apply?: boolean }) {
  if (!path.isAbsolute(input.source) || !path.isAbsolute(input.runtime)) throw new Error("Source and runtime must be explicit absolute paths.")
  if (!lstatSync(input.source).isFile() || lstatSync(input.source).isSymbolicLink()) throw new Error("Source must be a regular file.")
  const directory = path.join(input.runtime, "config", "opencode"), target = path.join(directory, "opencode.json")
  if (path.resolve(input.source) === target) throw new Error("Source must differ from runtime overrides.")
  const marker = path.join(directory, "legacy-settings-migrated.json")
  if (existsSync(marker)) return { status: "already-migrated", fields: [] as string[], skipped: [] as string[] }
  const source = readSettingsFile(input.source), migrated: EngineSettings = {}
  const fields: string[] = [], skipped: string[] = []
  function candidate(patch: Record<string, unknown>, label: string) {
    try { validateSettingsPatch(patch); Object.assign(migrated, mergeSettings(migrated, patch)); fields.push(label) }
    catch { skipped.push(label) }
  }
  for (const key of ["model", "small_model", "disabled_providers"]) if (source[key] !== undefined) candidate({ [key]: source[key] }, key)
  for (const [id, agent] of Object.entries(source.agent ?? {})) for (const key of ["model", "disable"]) if (agent[key] !== undefined) candidate({ agent: { [id]: { [key]: agent[key] } } }, `agent.${id}.${key}`)
  const skills = source.permission?.skill
  if (typeof skills === "object" && skills !== null && !Array.isArray(skills)) for (const [id, action] of Object.entries(skills)) if (id !== "*") candidate({ permission: { skill: { [id]: action } } }, `permission.skill.${id}`)
  for (const [id, provider] of Object.entries(source.provider ?? {})) {
    for (const key of ["name", "npm"]) if (provider[key] !== undefined) candidate({ provider: { [id]: { [key]: provider[key] } } }, `provider.${id}.${key}`)
    if (typeof provider.options === "object" && provider.options !== null) for (const [key, value] of Object.entries(provider.options)) candidate({ provider: { [id]: { options: { [key]: value } } } }, `provider.${id}.options.${key}`)
    if (typeof provider.models === "object" && provider.models !== null) for (const [idModel, model] of Object.entries(provider.models)) if (model && typeof model === "object" && "name" in model) candidate({ provider: { [id]: { models: { [idModel]: { name: model.name } } } } }, `provider.${id}.models.${idModel}.name`)
  }
  if (!input.apply) return { status: "dry-run", fields, skipped }
  const overrides = mergeSettings(migrated, readSettingsFile(target)) // existing local choices win
  validateSettingsPatch(overrides)
  mkdirSync(directory, { recursive: true, mode: 0o700 }); chmodSync(directory, 0o700)
  const backup = path.join(directory, "legacy-settings.backup.json")
  if (!existsSync(backup)) { copyFileSync(input.source, backup, 1); chmodSync(backup, 0o600) }
  writeEngineOverrides(overrides, target)
  if (JSON.stringify(readSettingsFile(target)) !== JSON.stringify(overrides)) throw new Error("Runtime settings readback failed; migration marker was not written.")
  writeFileSync(marker, JSON.stringify({ version: 1 }) + "\n", { flag: "wx", mode: 0o600 })
  return { status: "migrated", fields, skipped }
}
if (import.meta.main) {
  try {
    const args = process.argv.slice(2)
    if (args.some((arg, index) => !["--source", "--runtime", "--apply"].includes(arg) && !["--source", "--runtime"].includes(args[index - 1]))) throw new Error("Unsupported argument")
    const source = args[args.indexOf("--source") + 1], runtime = args[args.indexOf("--runtime") + 1]
    if (!args.includes("--source") || !args.includes("--runtime") || !source || !runtime) throw new Error("Explicit --source and --runtime required")
    console.log(JSON.stringify(migrateSettings({ source, runtime, apply: args.includes("--apply") })))
  } catch { console.error("Settings migration failed. Check explicit paths and supported fields; no source values are logged."); process.exitCode = 1 }
}
