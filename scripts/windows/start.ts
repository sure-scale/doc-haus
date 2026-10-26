// doc.haus launcher for Windows, run by start.cmd. It follows start.sh step by
// step. Windows has no POSIX signals, lsof or readable working directories of
// other processes, so: Ctrl+C reaches every process on the shared console and
// taskkill ends what remains, netstat reports the ports, the engine is released
// through its standard input, ingest's update request is polled, and the lock
// records this launcher's creation time so it can be identified.
import { existsSync, mkdirSync, readFileSync, realpathSync, rmdirSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import type { Subprocess } from "bun"
import { embeddingModelCacheDir } from "../../dochaus/lib/search/model-cache"
import { dataPaths } from "../../services/engine/src/data-paths"
import { installDependencies } from "../install-dependencies"
import { ensureWindowsHostTools, windowsToolPath } from "./requirements"

// Windows paths keep the case they were typed in (Command Prompt keeps a typed
// lowercase drive letter), while realpath reports the stored spelling that
// ingest records. Every process gets the stored spelling, so prefix checks agree.
const repository = realpathSync.native(path.resolve(import.meta.dir, "../.."))
// Windows environment names are case-insensitive, but plain objects are not.
const pathKey = Object.keys(process.env).find(key => key.toUpperCase() === "PATH") ?? "PATH"

// Read dochaus/.env.local as a dotenv file: KEY=value lines, optional export,
// quotes and # comments. Double quotes expand \n; nothing else is expanded.
export function parseEnvFile(text: string) {
  const values: Record<string, string> = {}
  for (const line of text.split(/\r?\n/)) {
    const match = /^\s*(?:export\s+)?([A-Za-z_]\w*)\s*=\s*(.*?)\s*$/.exec(line)
    if (!match) continue
    const quoted = /^(["'])(.*)\1(?:\s+#.*)?$/.exec(match[2])
    values[match[1]] = quoted
      ? quoted[1] === '"' ? quoted[2].replace(/\\n/g, "\n") : quoted[2]
      : match[2].replace(/\s+#.*$/, "")
  }
  return values
}

// netstat localises its state column, but a listening TCP socket always has
// the wildcard foreign address 0.0.0.0:0 or [::]:0.
export function listeningPids(netstat: string, port: string) {
  const pids = new Set<string>()
  for (const line of netstat.split(/\r?\n/)) {
    const [protocol, local, foreign, ...rest] = line.trim().split(/\s+/)
    if (protocol === "TCP" && local?.endsWith(`:${port}`) && (foreign === "0.0.0.0:0" || foreign === "[::]:0") && rest.length)
      pids.add(rest[rest.length - 1])
  }
  return [...pids]
}

function processRunning(pid: number) {
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM"
  }
}

function running(pid: string) {
  return /^\d+$/.test(pid) && Number(pid) > 1 && processRunning(Number(pid))
}

function read(file: string) {
  try {
    return readFileSync(file, "utf8").trim()
  } catch {
    return ""
  }
}

async function output(command: string[]) {
  try {
    const child = Bun.spawn(command, { stdin: "ignore", stdout: "pipe", stderr: "ignore" })
    const [text, status] = await Promise.all([new Response(child.stdout).text(), child.exited])
    return status === 0 ? text : ""
  } catch {
    return ""
  }
}

async function taskkill(pid: number) {
  await output(["taskkill", "/PID", String(pid), "/T", "/F"])
}

// CIM reports a creation time for every process; a live PID with another
// creation time has been reused by a different process.
async function processCreated(pid: number) {
  const created = (await output(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
    `(Get-CimInstance Win32_Process -Filter 'ProcessId = ${pid}' -ErrorAction Stop).CreationDate.ToFileTimeUtc()`])).trim()
  return /^\d+$/.test(created) ? created : undefined
}

// Whether the launcher recorded in the lock is running, gone (its PID exited
// or now belongs to another process) or cannot be identified.
async function ownerState(lock: string, owner: string): Promise<"running" | "gone" | "unknown"> {
  if (/^\d+$/.test(owner) && Number(owner) > 1 && !processRunning(Number(owner))) return "gone"
  let record: { pid?: unknown; started?: unknown } | undefined
  try {
    record = JSON.parse(readFileSync(path.join(lock, "launcher.json"), "utf8"))
  } catch {}
  if (String(record?.pid) !== owner || typeof record?.started !== "string" || !record.started) return "unknown"
  const created = await processCreated(Number(owner))
  return created === undefined ? "unknown" : created === record.started ? "running" : "gone"
}

// Used by ingest before it requests a one-click update.
export async function windowsLauncherRunning(runtime: string, pid: number) {
  return await ownerState(path.join(runtime, "engine.lock"), String(pid)) === "running"
}

function removeLock(lock: string) {
  rmSync(path.join(lock, "launcher.json"), { force: true })
  rmSync(path.join(lock, "pid"), { force: true })
  rmdirSync(lock)
}

// Claim this launcher runtime, or let an interactive caller restart its owner,
// as scripts/start-runtime.sh does. Return 2 when the existing stack should
// simply be left running.
async function claimRuntime(runtime: string, interactive: boolean) {
  const lock = path.join(runtime, "engine.lock")
  const engineRecord = path.join(runtime, "engine-process.lock")
  const started = await processCreated(process.pid) ?? ""
  while (true) {
    try {
      mkdirSync(lock)
      break
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error
    }
    const owner = read(path.join(lock, "pid"))
    const state = await ownerState(lock, owner)
    if (state === "gone") {
      const engine = read(engineRecord)
      if (running(engine)) {
        console.error(`doc.haus: the engine is still running (pid ${engine}), but its launcher has exited. Stop that engine before restarting.`)
        return 1
      }
      console.error("doc.haus: clearing a stale launcher lock")
      removeLock(lock)
      continue
    }
    if (state === "unknown") {
      console.error(`doc.haus: another process owns this runtime (${lock}); it could not be identified as this stack's launcher.`)
      return 1
    }
    console.error(`doc.haus: already running (launcher pid ${owner}).`)
    if (!interactive) {
      console.error("doc.haus: leaving it running. Run start.cmd in a console to choose whether to restart it.")
      return 2
    }
    if (!/^(y|yes)$/i.test(prompt("Stop that stack and restart it? [y/N]")?.trim() ?? "")) {
      console.error("doc.haus: leaving the existing stack running.")
      return 2
    }
    // Re-check ownership after the answer; never stop a replacement owner.
    if (read(path.join(lock, "pid")) !== owner || await ownerState(lock, owner) !== "running") {
      console.error("doc.haus: runtime ownership changed; run start.cmd again.")
      return 1
    }
    console.error("doc.haus: stopping the existing stack...")
    const engine = read(engineRecord)
    // Ctrl+C cannot be sent to another console, so end that launcher's tree.
    await taskkill(Number(owner))
    for (let attempt = 0; attempt < 60 && (running(owner) || running(engine)); attempt++) await Bun.sleep(250)
    if (running(owner) || running(engine)) {
      console.error("doc.haus: the existing stack has not stopped yet. Retry once shutdown finishes.")
      return 1
    }
  }
  const engine = read(engineRecord)
  if (running(engine)) {
    rmdirSync(lock)
    console.error(`doc.haus: the engine is still running (pid ${engine}), but its launcher has exited. Stop that engine before restarting.`)
    return 1
  }
  // The engine uses exclusive file creation, so a dead PID's record must also
  // be removed before launch. Preserve malformed or changed ownership records.
  if (existsSync(engineRecord)) {
    if (!/^\d+$/.test(engine) || Number(engine) <= 1 || read(engineRecord) !== engine) {
      rmdirSync(lock)
      console.error(`doc.haus: engine ownership could not be verified (${engineRecord}).`)
      return 1
    }
    console.error(`doc.haus: clearing a stale engine lock (pid ${engine})`)
    try {
      rmSync(engineRecord)
    } catch {
      rmdirSync(lock)
      return 1
    }
  }
  writeFileSync(path.join(lock, "launcher.json"), JSON.stringify({ pid: process.pid, started, exe: process.execPath }))
  writeFileSync(path.join(lock, "pid"), String(process.pid))
  return 0
}

// Hand children exactly one PATH; `undefined` removes a variable.
function environment(extra: Record<string, string | undefined> = {}) {
  const env: Record<string, string> = {}
  for (const [key, value] of Object.entries({ ...process.env, ...extra }))
    if (value !== undefined && key.toUpperCase() !== "PATH") env[key] = value
  env.PATH = process.env[pathKey] ?? ""
  return env
}

async function run(command: string[], cwd = ".") {
  return await Bun.spawn(command, { cwd, env: environment(), stdin: "inherit", stdout: "inherit", stderr: "inherit" }).exited
}

async function main(args: string[]) {
  process.chdir(repository)
  const bun = process.execPath
  process.env.BUN_BIN = bun
  // Children that run `bun` from PATH get the Bun start.cmd selected.
  process.env[pathKey] = `${path.dirname(bun)};${process.env[pathKey] ?? ""}`
  // --smoke is a boot check for CI: no prompts or browser, stop once the whole
  // stack is ready, and fail when it is not or another stack owns the runtime.
  const smoke = args.includes("--smoke")
  if (smoke) process.env.NO_OPEN = "1"

  // Provider credentials for all three processes. As when start.sh sources the
  // file, its values replace inherited ones. Gitignored.
  const local = path.join("dochaus", ".env.local")
  if (existsSync(local)) Object.assign(process.env, parseEnvFile(readFileSync(local, "utf8")))
  if (args.includes("--demo"))
    console.error("doc.haus: --demo was removed: choose 'Add a sample matter' during setup, or POST /sample-matter to the ingest service")

  if (process.env.DOCHAUS_DATA_ROOT) process.env.DOCHAUS_DATA_ROOT = path.resolve(repository, process.env.DOCHAUS_DATA_ROOT)
  const paths = dataPaths(repository)
  for (const directory of [path.join(paths.workspace, ".skills"), path.join(paths.workspace, ".playbooks"), ...["data", "config", "state", "cache"].map(name => path.join(paths.runtime, name))])
    mkdirSync(directory, { recursive: true })
  const workspace = realpathSync.native(paths.workspace), runtime = realpathSync.native(paths.runtime)
  process.env.WORKSPACE_ROOT = workspace
  process.env.DOCHAUS_RUNTIME_DIR = runtime

  const enginePort = process.env.DOC_HAUS_ENGINE_PORT || "4097"
  const ingestPort = process.env.DOC_HAUS_INGEST_PORT || "4501"
  const webPort = process.env.DOC_HAUS_WEB_PORT || "5174"
  const engineHost = process.env.OPENCODE_HOSTNAME || "127.0.0.1"
  const ingestHost = process.env.INGEST_HOST || "127.0.0.1"
  const webHost = process.env.WEB_HOST || engineHost
  const viteEngineUrl = process.env.VITE_OPENCODE_URL || `http://localhost:${enginePort}`
  const viteIngestUrl = process.env.VITE_INGEST_URL || `http://localhost:${ingestPort}`

  const lock = path.join(runtime, "engine.lock")
  const engineRecord = path.join(runtime, "engine-process.lock")
  const services: { name: string; child: Subprocess }[] = []
  let stopping = false
  // Ctrl+C, Ctrl+Break and closing the window reach every process on this
  // console, so those stops give the services a grace period first.
  const stopServices = async (graceful: boolean) => {
    const stopped = services.splice(0).filter(({ child }) => child.exitCode === null)
    if (graceful) await Promise.race([Promise.all(stopped.map(({ child }) => child.exited)), Bun.sleep(10_000)])
    for (const { child } of stopped) if (child.exitCode === null) await taskkill(child.pid)
    await Promise.all(stopped.map(({ child }) => child.exited))
    // A forced stop skips the engine's own exit handler.
    for (const { name, child } of stopped)
      if (name === "engine" && read(engineRecord) === String(child.pid)) rmSync(engineRecord, { force: true })
  }
  const releaseLock = () => {
    if (read(path.join(lock, "pid")) === String(process.pid)) removeLock(lock)
  }
  const stop = async (status: number, graceful = false): Promise<never> => {
    if (!stopping) {
      stopping = true
      await stopServices(graceful)
      releaseLock()
      process.exit(status)
    }
    return await new Promise<never>(() => {})
  }
  process.on("SIGINT", () => void stop(130, true))
  process.on("SIGBREAK", () => void stop(130, true))
  process.on("SIGHUP", () => void stop(129, true))
  process.on("exit", () => {
    for (const { child } of services) if (child.exitCode === null) Bun.spawnSync(["taskkill", "/PID", String(child.pid), "/T", "/F"])
    releaseLock()
  })
  process.env.DOCHAUS_LAUNCHER_PID = String(process.pid)

  // A second stack must never open the same durable session database.
  const claimed = await claimRuntime(runtime, !smoke && process.stdin.isTTY === true)
  if (claimed !== 0) process.exit(claimed !== 2 ? claimed : smoke ? 1 : 0)

  process.env.DOCHAUS_MODEL_CACHE_DIR = embeddingModelCacheDir()
  const webUrl = `http://localhost:${webPort}`
  const updateRequest = path.join(runtime, "update-request.json")
  // An update restarts the stack in this loop, keeping the lock throughout.
  while (true) {
    // Ingest writes a request only after confirming this launcher; one left
    // from an earlier run must never start an update.
    rmSync(updateRequest, { force: true })
    // Refuse occupied ports instead of terminating another process.
    const netstat = await output(["netstat", "-ano"])
    for (const port of [enginePort, ingestPort, webPort]) {
      const pids = listeningPids(netstat, port)
      if (pids.length) {
        console.error(`doc.haus: port ${port} is already in use (pid ${pids.join(" ")})`)
        return await stop(1)
      }
    }

    process.env[pathKey] = windowsToolPath()
    if (!process.env.SKIP_INSTALL) {
      await ensureWindowsHostTools(runtime, smoke ? { interactive: false } : {})
      process.env[pathKey] = windowsToolPath()
      for (const directory of [".", "services/ingest"]) {
        const status = await installDependencies(directory)
        if (status !== 0) return await stop(status)
      }
    }

    const paneStatus = await run([bun, "run", "--cwd", "apps/word-addin", "build"], ".")
    if (paneStatus !== 0) return await stop(paneStatus)
    if (process.env.DOCHAUS_SETUP_WORD_ADDIN === "1") {
      delete process.env.DOCHAUS_SETUP_WORD_ADDIN
      if (await run([bun, "run", "scripts/setup-word-addin.ts", "--offer"], ".") !== 0)
        console.error("doc.haus: Word setup can be retried with bun scripts/setup-word-addin.ts")
    }

    // Every upload and search needs the local embedding model; see start.sh.
    console.log(`doc.haus: checking the embedding model (${process.env.DOCHAUS_MODEL_CACHE_DIR})`)
    for (let attempt = 1; await run([bun, "run", "src/warm-embedding-model.ts"], "services/ingest") !== 0; attempt++) {
      if (attempt >= 3) {
        console.error(`doc.haus: the embedding model could not be loaded after ${attempt} attempts.
  It downloads once (about 50 MB) from https://huggingface.co into
  ${process.env.DOCHAUS_MODEL_CACHE_DIR}. Check that this machine can reach huggingface.co
  (HTTPS_PROXY is honoured) and that the directory is writable, then run
  start.cmd again. Set DOCHAUS_MODEL_CACHE_DIR to use a different directory.`)
        return await stop(1)
      }
      console.error(`doc.haus: retrying the embedding model in ${attempt * 5}s`)
      await Bun.sleep(attempt * 5000)
    }

    console.log(`doc.haus: starting engine, ingest, web (workspace: ${workspace})`)
    console.log(`doc.haus: ports engine=${enginePort} ingest=${ingestPort} web=${webPort}`)
    console.log(`doc.haus: web UI will be at ${webUrl}`)
    process.env.INGEST_URL = `http://127.0.0.1:${ingestPort}`
    process.env.DOC_HAUS_INGEST_PORT = ingestPort
    process.env.DOC_HAUS_ENGINE_PORT = enginePort

    const engine = Bun.spawn([bun, "run", "services/engine/src/server.ts"], {
      env: environment({
        // Preserve only the native credential source for explicit, selected reuse.
        DOCHAUS_NATIVE_OPENCODE_DATA_DIR: process.env.DOCHAUS_NATIVE_OPENCODE_DATA_DIR ||
          path.join(process.env.XDG_DATA_HOME || path.join(os.homedir(), ".local", "share"), "opencode"),
        XDG_DATA_HOME: path.join(runtime, "data"),
        XDG_CONFIG_HOME: path.join(runtime, "config"),
        XDG_STATE_HOME: path.join(runtime, "state"),
        XDG_CACHE_HOME: path.join(runtime, "cache"),
        // Configuration comes only from the isolated runtime and the shipped
        // doc.haus config directory, never from the developer's shell.
        OPENCODE_CONFIG: undefined,
        OPENCODE_CONFIG_CONTENT: undefined,
        OPENCODE_PERMISSION: undefined,
        DOCHAUS_MANAGED_RESTART: "1",
      }),
      stdin: "pipe", stdout: "inherit", stderr: "inherit",
    })
    services.push({ name: "engine", child: engine })
    // Native inventories populate after HTTP listens; wait for the legal plugin.
    const engineReady = await run([bun, "run", "services/engine/script/wait-ready.ts"])
    if (engineReady !== 0) return await stop(engineReady)

    services.push({ name: "ingest", child: Bun.spawn([bun, "run", "src/server.ts"], {
      cwd: "services/ingest", env: environment({ INGEST_HOST: ingestHost, INGEST_PORT: ingestPort }),
      stdin: "ignore", stdout: "inherit", stderr: "inherit",
    }) })
    services.push({ name: "web", child: Bun.spawn([bun, "run", "dev", "--", "--host", webHost, "--port", webPort], {
      cwd: "apps/web", env: environment({ VITE_OPENCODE_URL: viteEngineUrl, VITE_INGEST_URL: viteIngestUrl }),
      stdin: "ignore", stdout: "inherit", stderr: "inherit",
    }) })

    // Retained native work starts only after the whole managed stack can serve it.
    const stackReady = await run([bun, "run", "services/engine/script/wait-stack-ready.ts"])
    if (stackReady !== 0) return await stop(stackReady)
    if (read(engineRecord) !== String(engine.pid)) {
      console.error("doc.haus: managed restart release lost its exact engine process owner")
      return await stop(1)
    }
    engine.stdin.write("release\n")
    engine.stdin.end()
    if (smoke) {
      console.log("doc.haus: the stack is ready; stopping (--smoke)")
      return await stop(0)
    }

    // Say how to start doc.haus again, through the Start menu entry the
    // installer adds when there is one.
    const startMenu = path.join(process.env.APPDATA ?? "", "Microsoft", "Windows", "Start Menu", "Programs", "doc.haus.lnk")
    console.log(`doc.haus: ready at ${webUrl}. Keep this window open while you use doc.haus; press Ctrl+C to stop it.`)
    console.log(`doc.haus: to start it again later, ${existsSync(startMenu) ? "open doc.haus from the Start menu, or " : ""}run ${path.join(repository, "start.cmd")}.`)

    // wait-stack-ready has already loaded the web UI. Opt out with NO_OPEN=1.
    if (!process.env.NO_OPEN) {
      console.log(`doc.haus: opening ${webUrl}`)
      void output(["cmd", "/c", "start", "", webUrl])
    }

    // Run until interrupted (Ctrl+C) or a managed process exits on its own; the
    // stack cannot serve without all three. Windows has no SIGUSR1, so poll for
    // ingest's update request too.
    while (!existsSync(updateRequest)) {
      for (const { name, child } of services) {
        if (child.exitCode !== null) {
          console.error(`doc.haus: ${name} exited (status ${child.exitCode}); stopping the stack`)
          return await stop(1)
        }
      }
      await Bun.sleep(1000)
    }
    console.log("doc.haus: installing the requested update...")
    // Let ingest return its accepted response first, as its signal delay does.
    await Bun.sleep(1000)
    // Keep the launcher lock until its updater finishes. No second stack may
    // start against the session database during installation.
    await stopServices(false)
    await run([bun, "run", "scripts/updater.ts", "--apply"])
    process.env.NO_OPEN = "1"
  }
}

if (import.meta.main) await main(process.argv.slice(2))
