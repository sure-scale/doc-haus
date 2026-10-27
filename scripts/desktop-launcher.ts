// Refresh the installed desktop entry after an update or explicit Word setup.
// This changes only this user's launch integration; it never starts the stack.
import os from "node:os"
import path from "node:path"
import { readFileSync, realpathSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { dataPaths } from "../services/engine/src/data-paths"

export const desktopProfileKeys = ["DOCHAUS_DATA_ROOT", "WORKSPACE_ROOT", "DOCHAUS_RUNTIME_DIR", "DOC_HAUS_ENGINE_PORT", "DOC_HAUS_INGEST_PORT", "DOC_HAUS_WEB_PORT"] as const

export function selectedDesktopEnvironment(environment: Record<string, string | undefined>) {
  const selected: Record<string, string> = {}
  for (const name of desktopProfileKeys) {
    const key = Object.keys(environment).find(key => key.toUpperCase() === name)
    if (key && environment[key]) selected[name] = environment[key]!
  }
  return selected
}

type Options = {
  environment?: Record<string, string | undefined>
  platform?: NodeJS.Platform
  home?: string
  startup?: boolean
  command?: (command: string[], environment: Record<string, string | undefined>) => Promise<{ code: number; stdout: string }>
}

function canonical(file: string) {
  try { return realpathSync.native(file) } catch { return path.resolve(file) }
}

async function execute(command: string[], environment: Record<string, string | undefined>, options: Options) {
  if (options.command) return options.command(command, environment)
  const child = Bun.spawn(command, { env: environment, stdin: "ignore", stdout: "pipe", stderr: "ignore", windowsHide: true })
  const [code, stdout] = await Promise.all([child.exited, new Response(child.stdout).text()])
  return { code, stdout }
}

// Legacy app bundles contain only our fixed shell exports. Verify the signed
// bundle before reading them; clear inherited selectors so a second stack
// cannot supply values the original installation did not select.
const legacySelection = `set -eu
unset DOCHAUS_DATA_ROOT WORKSPACE_ROOT DOCHAUS_RUNTIME_DIR DOC_HAUS_ENGINE_PORT DOC_HAUS_INGEST_PORT DOC_HAUS_WEB_PORT
. "$1"
cd "$DOCHAUS_INSTALL_DIR"
. scripts/data-paths.sh
configure_data_paths "$PWD"
printf '%s\\0%s\\0' "$(pwd -P)" "$DOCHAUS_RUNTIME_DIR"`

async function alreadySelected(platform: string, home: string, source: Record<string, string | undefined>, environment: Record<string, string | undefined>, expected: { repository: string; runtime: string }, options: Options) {
  try {
    let selected: { repository: string; runtime: string }
    if (platform === "win32") {
      const local = source.LOCALAPPDATA || path.join(source.USERPROFILE || home, "AppData", "Local")
      const record = JSON.parse(readFileSync(path.join(local, "doc.haus/launch-profile.json"), "utf8"))
      if (typeof record?.repository !== "string" || !record.environment || typeof record.environment !== "object" || Array.isArray(record.environment)) return false
      if (Object.values(record.environment).some(value => typeof value !== "string")) return false
      selected = { repository: record.repository, runtime: dataPaths(record.repository, selectedDesktopEnvironment(record.environment)).runtime }
    } else {
      const app = path.join(home, "Applications/doc.haus.app")
      if ((await execute(["/usr/bin/codesign", "--verify", "--deep", "--strict", app], environment, options)).code !== 0) return false
      const resources = path.join(app, "Contents/Resources")
      try {
        selected = JSON.parse(readFileSync(path.join(resources, "installation.json"), "utf8"))
      } catch {
        const result = await execute(["/bin/bash", "-c", legacySelection, "desktop-selection", path.join(resources, "installation.sh")], environment, options)
        const fields = result.stdout.split("\0")
        if (result.code !== 0 || fields.length !== 3 || fields[2] !== "") return false
        selected = { repository: fields[0], runtime: fields[1] }
      }
    }
    return typeof selected?.repository === "string" && typeof selected.runtime === "string" &&
      path.isAbsolute(selected.repository) && path.isAbsolute(selected.runtime) &&
      canonical(selected.repository) === expected.repository && canonical(selected.runtime) === expected.runtime
  } catch { return false }
}

export async function refreshDesktopLauncher(options: Options = {}) {
  const platform = options.platform ?? process.platform
  if (platform !== "darwin" && platform !== "win32") return
  const repository = canonical(fileURLToPath(new URL("../", import.meta.url)))
  const source = options.environment ?? process.env
  const home = options.home ?? os.homedir()
  const environment: Record<string, string | undefined> = { ...selectedDesktopEnvironment(source), HOME: home, BUN_BIN: process.execPath }
  for (const key of ["PATH", "USERPROFILE", "LOCALAPPDATA", "SystemRoot", "TEMP", "TMP", "TMPDIR"])
    environment[key] = Object.entries(source).find(([name]) => name.toUpperCase() === key.toUpperCase())?.[1]
  const selected = { repository, runtime: canonical(dataPaths(repository, selectedDesktopEnvironment(source)).runtime) }
  // The installers already set this flag when explicitly selecting a copy.
  // Ordinary startup may refresh only the desktop entry it already owns.
  if (options.startup && source.DOCHAUS_SETUP_WORD_ADDIN !== "1" && !await alreadySelected(platform, home, source, environment, selected, options)) return
  const command = platform === "darwin"
    ? ["/bin/bash", path.join(repository, "scripts/macos-app.sh"), repository, path.join(home, "Applications/doc.haus.app")]
    : [process.execPath, path.join(repository, "scripts/windows/register-protocol.ts")]
  if (platform === "darwin") {
    environment.DOCHAUS_DESKTOP_HELPER = "1"
    environment.DOCHAUS_DESKTOP_METADATA = JSON.stringify(selected)
  }
  const result = await execute(command, environment, options)
  if (result.code !== 0) throw new Error("Desktop Word links could not be refreshed. Use Settings > Add-ins to try again.")
  if (platform === "darwin") {
    // Reassert the association even for an unchanged bundle, so a previous
    // failed registration (or a removed association) is repaired on retry.
    const registration = ["/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister", "-f", path.join(home, "Applications/doc.haus.app")]
    const code = options.command ? (await options.command(registration, environment)).code :
      await Bun.spawn(registration, { env: environment, stdin: "ignore", stdout: "ignore", stderr: "ignore" }).exited
    if (code !== 0) throw new Error("Desktop Word links could not be registered. Use Settings > Add-ins to try again.")
  }
}

if (import.meta.main) {
  try { await refreshDesktopLauncher({ startup: process.argv.slice(2).includes("--startup") }) }
  catch { console.error("doc.haus: desktop Word links could not be refreshed. Use Settings > Add-ins to try again."); process.exitCode = 1 }
}
