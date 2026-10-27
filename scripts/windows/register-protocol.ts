// Shared by the Windows installer, normal startup and explicit Word setup.
// reg.exe receives structured arguments; no shell evaluates URI placeholders.
import { existsSync, mkdirSync, readFileSync, realpathSync, renameSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { selectedDesktopEnvironment } from "../desktop-launcher"

type Options = {
  repository: string
  environment?: Record<string, string | undefined>
  bun?: string
  command?: (args: string[]) => Promise<number>
}

export async function registerWindowsProtocol(options: Options) {
  const environment = options.environment ?? process.env
  const repository = realpathSync.native(options.repository)
  const bun = options.bun ?? process.execPath
  const local = environment.LOCALAPPDATA || path.join(environment.USERPROFILE || os.homedir(), "AppData", "Local")
  const directory = path.join(local, "doc.haus")
  const profile = path.join(directory, "launch-profile.json")
  const content = JSON.stringify({ repository, environment: selectedDesktopEnvironment(environment) }, null, 2) + "\n"
  mkdirSync(directory, { recursive: true })
  let previous = ""
  try { previous = readFileSync(profile, "utf8") } catch {}
  if (previous !== content) {
    const temporary = `${profile}.${crypto.randomUUID()}.tmp`
    try { writeFileSync(temporary, content, { mode: 0o600, flag: "wx" }); renameSync(temporary, profile) }
    finally { rmSync(temporary, { force: true }) }
  }
  const key = "HKCU\\Software\\Classes\\doc.haus"
  const handler = path.join(repository, "scripts/windows/protocol.ts")
  const values = [[key, undefined, "URL:doc.haus"], [key, "URL Protocol", ""],
    [`${key}\\shell\\open\\command`, undefined, `"${bun}" "${handler}" "%1"`]] as (string | undefined)[][]
  const icon = path.join(repository, "scripts/icons/doc-haus.ico")
  if (existsSync(icon)) values.push([`${key}\\DefaultIcon`, undefined, `"${icon}",0`])
  for (const [entry, name, value] of values) {
    const args = ["reg.exe", "add", entry!, ...(name === undefined ? ["/ve"] : ["/v", name]), "/t", "REG_SZ", "/d", value!, "/f"]
    const status = options.command ? await options.command(args) : await Bun.spawn(args, { stdin: "ignore", stdout: "ignore", stderr: "ignore", windowsHide: true }).exited
    if (status !== 0) throw new Error("The doc.haus desktop link registration could not be refreshed.")
  }
}

if (import.meta.main) {
  try {
    if (process.platform !== "win32") throw new Error("Windows only")
    await registerWindowsProtocol({ repository: path.resolve(import.meta.dir, "../..") })
  } catch { console.error("doc.haus: desktop links could not be registered. Use Settings > Add-ins to try again."); process.exitCode = 1 }
}
