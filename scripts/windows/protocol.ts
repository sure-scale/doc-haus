// Per-user doc.haus://launch entry point. The registry invokes Bun directly:
// no part of the supplied URI is evaluated by cmd.exe or PowerShell.
import { existsSync, readFileSync, realpathSync } from "node:fs"
import path from "node:path"
import { dataPaths } from "../../services/engine/src/data-paths"
import { listeningPids, parseEnvFile, windowsLauncherRunning } from "./start"
import { openWordDocument, parseWordOpenUri, WordOpenError } from "../word-open"
import { desktopProfileKeys, selectedDesktopEnvironment } from "../desktop-launcher"

class LaunchError extends Error {}

export function isLaunchURI(args: string[]) {
  return args.length === 1 && (args[0] === "doc.haus://launch" || args[0] === "doc.haus://launch/")
}

export function resolveProfile(repository: string, saved: unknown, inherited: Record<string, string | undefined>, local = "") {
  const profile = saved as { repository?: unknown; environment?: unknown } | undefined
  let savedRepository: string | undefined
  try { if (typeof profile?.repository === "string") savedRepository = realpathSync.native(profile.repository) } catch {}
  if (savedRepository !== repository || !profile?.environment || typeof profile.environment !== "object" || Array.isArray(profile.environment))
    throw new LaunchError("The doc.haus installation changed. Run the installer again.")
  for (const [key, value] of Object.entries(profile.environment))
    if (!desktopProfileKeys.includes(key as typeof desktopProfileKeys[number]) || typeof value !== "string")
      throw new LaunchError("The doc.haus launch settings are not valid. Run the installer again.")
  const environment = { ...inherited }
  // Browser processes can retain another installation's old environment.
  // Only this install's recorded selection (then its .env.local) chooses paths.
  for (const key of Object.keys(environment))
    if (desktopProfileKeys.includes(key.toUpperCase() as typeof desktopProfileKeys[number])) delete environment[key]
  Object.assign(environment, profile.environment as Record<string, string>)
  // Use the same local override precedence as start.ts, without retaining any
  // provider credentials in the launch profile or exposing them in messages.
  const configuration = { ...selectedDesktopEnvironment(environment), ...selectedDesktopEnvironment(parseEnvFile(local)) }
  const ports = ["DOC_HAUS_ENGINE_PORT", "DOC_HAUS_INGEST_PORT", "DOC_HAUS_WEB_PORT"].map((key, index) => {
    const port = configuration[key] || ["4097", "4501", "5174"][index]
    if (!/^\d+$/.test(port) || Number(port) < 1 || Number(port) > 65535)
      throw new LaunchError("A doc.haus port setting is not valid. Check this installation's settings.")
    return port
  })
  return { repository, environment, runtime: dataPaths(repository, configuration).runtime, ports, webURL: `http://localhost:${ports[2]}` }
}

export function listenerBelongsTo(pid: string, owner: string, parents: Map<string, string>) {
  const visited = new Set<string>()
  while (/^\d+$/.test(pid) && Number(pid) > 1 && visited.size < 64 && !visited.has(pid)) {
    if (pid === owner) return true
    visited.add(pid)
    pid = parents.get(pid) ?? ""
  }
  return false
}

type Profile = ReturnType<typeof resolveProfile>
type System = {
  read: (file: string) => string
  exists: (file: string) => boolean
  running: (pid: string) => boolean
  ownerRunning: (runtime: string, pid: number) => Promise<boolean>
  netstat: () => Promise<string>
  parents: () => Promise<Map<string, string>>
  ready: (url: string) => Promise<boolean>
  open: (url: string) => Promise<void>
  start: (profile: Profile) => Promise<void>
}

export async function launchFromProtocol(args: string[], loadProfile: () => Profile, system: System) {
  // Validate before reading settings, querying processes or invoking anything.
  if (!isLaunchURI(args)) throw new LaunchError("This doc.haus link is not supported.")
  const profile = loadProfile()
  const lock = path.join(profile.runtime, "engine.lock")
  const owner = system.read(path.join(lock, "pid"))
  const engine = system.read(path.join(profile.runtime, "engine-process.lock"))
  if (system.running(owner)) {
    if (!await system.ownerRunning(profile.runtime, Number(owner)))
      throw new LaunchError("Another process owns this doc.haus runtime. Check its console before opening this installation.")
    const listeners = listeningPids(await system.netstat(), profile.ports[2])
    // A verified launcher may still be starting. Its normal startup owns the
    // eventual browser open; clicking the recovery link must leave it alone.
    if (!listeners.length) return
    const parents = await system.parents()
    if (!listeners.every(pid => listenerBelongsTo(pid, owner, parents)))
      throw new LaunchError("This installation's web port belongs to another process. Check its original port settings.")
    if (system.read(path.join(lock, "pid")) !== owner || !await system.ownerRunning(profile.runtime, Number(owner)))
      throw new LaunchError("The doc.haus runtime owner changed. Try opening it again.")
    if (await system.ready(profile.webURL)) await system.open(profile.webURL)
    return
  }
  if (system.running(engine))
    throw new LaunchError("A doc.haus engine is still running without its launcher. Check its console before opening this installation.")
  if (system.exists(lock) && (!/^\d+$/.test(owner) || Number(owner) <= 1))
    throw new LaunchError("This doc.haus runtime owner could not be identified. Check its console before opening this installation.")
  const netstat = await system.netstat()
  if (profile.ports.some(port => listeningPids(netstat, port).length))
    throw new LaunchError("A port needed by this doc.haus installation is already in use. Check its console before starting it.")
  // start.cmd retains responsibility for stale locks, tool selection and boot.
  await system.start(profile)
}

export async function dispatchProtocol(args: string[], loadProfile: () => Profile, system: System,
  openWord: typeof openWordDocument = openWordDocument) {
  if (isLaunchURI(args)) return launchFromProtocol(args, loadProfile, system)
  if (args.length !== 1) throw new LaunchError("This doc.haus link is not supported.")
  // Parse the shared strict schema before reading any installation settings
  // or performing process/network/native actions. No caller text is a command.
  parseWordOpenUri(args[0])
  const profile = loadProfile()
  await openWord({ uri: args[0], runtime: profile.runtime })
}

function read(file: string) {
  try { return readFileSync(file, "utf8").trim() } catch { return "" }
}

function running(pid: string) {
  if (!/^\d+$/.test(pid) || Number(pid) <= 1) return false
  try { process.kill(Number(pid), 0); return true } catch (error) { return (error as NodeJS.ErrnoException).code === "EPERM" }
}

async function output(command: string[], environment?: NodeJS.ProcessEnv) {
  const child = Bun.spawn(command, { env: environment, stdin: "ignore", stdout: "pipe", stderr: "ignore", windowsHide: true, timeout: 5000 })
  const [text, status] = await Promise.all([new Response(child.stdout).text(), child.exited])
  if (status !== 0) throw new LaunchError("doc.haus could not check or open this installation. Start it from the Start menu.")
  return text
}

const system: System = {
  read,
  exists: existsSync,
  running,
  ownerRunning: windowsLauncherRunning,
  netstat: () => output(["netstat.exe", "-ano"]),
  parents: async () => {
    // Process IDs and parent IDs only: no command lines, windows or contents.
    const text = await output(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
      "Get-CimInstance Win32_Process -ErrorAction Stop | Select-Object ProcessId,ParentProcessId | ConvertTo-Json -Compress"])
    const value = JSON.parse(text)
    const records = Array.isArray(value) ? value : [value]
    return new Map(records.filter(record => Number.isSafeInteger(record?.ProcessId) && Number.isSafeInteger(record?.ParentProcessId))
      .map(record => [String(record.ProcessId), String(record.ParentProcessId)]))
  },
  ready: async url => {
    try { return (await fetch(url, { signal: AbortSignal.timeout(1000), redirect: "manual" })).ok } catch { return false }
  },
  open: async url => { await output(["cmd.exe", "/d", "/c", "start", "", url]) },
  start: async profile => {
    // The command text is fixed. WorkingDirectory carries the installation
    // path as a value, so spaces, %, & and quotes are never shell program text.
    await output(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
      "Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\\cmd.exe') -ArgumentList '/d','/k','start.cmd' -WorkingDirectory $env:DOCHAUS_PROTOCOL_DIRECTORY -ErrorAction Stop"],
    { ...profile.environment, BUN_BIN: process.execPath, DOCHAUS_PROTOCOL_DIRECTORY: profile.repository })
  },
}

if (import.meta.main) {
  try {
    if (process.platform !== "win32") throw new LaunchError("Use the doc.haus app to open this link on this computer.")
    await dispatchProtocol(process.argv.slice(2), () => {
      const repository = realpathSync.native(path.resolve(import.meta.dir, "../.."))
      const profileFile = path.join(process.env.LOCALAPPDATA || path.join(process.env.USERPROFILE || "", "AppData", "Local"), "doc.haus", "launch-profile.json")
      return resolveProfile(repository, JSON.parse(read(profileFile).replace(/^\uFEFF/, "")), process.env, read(path.join(repository, "dochaus", ".env.local")))
    }, system)
  } catch (error) {
    // Only our fixed messages are safe to show; native exceptions may contain
    // supplied URI text or configuration details.
    console.error(error instanceof LaunchError || error instanceof WordOpenError ? error.message : "doc.haus could not open this link. Start it from the Start menu, or run the installer again.")
    process.exitCode = 1
  }
}
