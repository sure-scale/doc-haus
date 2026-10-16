import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { isSettingsObject, readEffectiveSettings, settingsPaths } from "./effective-settings"

// Read installation policy synchronously at admission. A cached Location model
// catalog must never substitute a model after the user removes that selection.
export function generationAdmissionError(paths = settingsPaths()): string | undefined {
  try {
    const settings = readEffectiveSettings(paths)
    const primary = typeof settings.model === "string" ? settings.model.split("/")[0] : undefined
    const fast = typeof settings.small_model === "string" ? settings.small_model.split("/")[0] : undefined
    if (!primary || !fast || primary !== fast)
      return "Select primary and fast models from the same provider in Settings before generating."
    if (Array.isArray(settings.disabled_providers) && settings.disabled_providers.includes(primary) ||
        Array.isArray(settings.enabled_providers) && !settings.enabled_providers.includes(primary))
      return "The selected provider is disabled. Enable it in Settings before generating."
    const setup = path.join(path.dirname(paths.overrides), "setup.json")
    if (!existsSync(setup)) return undefined
    const saved: unknown = JSON.parse(readFileSync(setup, "utf8"))
    if (!isSettingsObject(saved) || saved.version !== 1 || !Array.isArray(saved.completed) || !Array.isArray(saved.skipped))
      return "Finish installation setup, including choosing or skipping the optional steps, before generating."
    const completed = saved.completed
    const skipped = saved.skipped
    if (!completed.includes("models") || !["models", "details", "research"].every(step => completed.includes(step) || skipped.includes(step)))
      return "Finish installation setup, including choosing or skipping the optional steps, before generating."
  } catch {
    return "Saved installation settings could not be read. Repair Settings before generating."
  }
}
