import { afterEach, expect, test } from "bun:test"
import { chmod, copyFile, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

const repository = path.resolve(import.meta.dir, "../..")
const roots: string[] = []
const processes: ReturnType<typeof Bun.spawn>[] = []
const servers: ReturnType<typeof Bun.serve>[] = []
afterEach(async () => {
  for (const server of servers.splice(0)) server.stop(true)
  for (const child of processes.splice(0)) { if (child.exitCode === null) child.kill("SIGTERM"); await child.exited }
  for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true })
})

async function fixture() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-install-launcher-"))
  roots.push(root)
  const home = path.join(root, "home"), bin = path.join(root, "bin")
  // Spaces and shell syntax in a selected installation must remain literal.
  const checkout = path.join(root, "doc haus $(touch unexpected-checkout)")
  for (const directory of [home, bin, path.join(checkout, "scripts")]) await mkdir(directory, { recursive: true })
  for (const file of ["data-paths.sh", "start-runtime.sh", "macos-app.sh", "macos-app.applescript", "macos-app-launch.sh"])
    await copyFile(path.join(repository, "scripts", file), path.join(checkout, "scripts", file))
  // These OS tools are stubbed: no app is launched, registered or trusted.
  // The actual compiler gets a separate compile-only smoke check on macOS.
  for (const [file, source] of [
    ["git", "#!/bin/sh\ncase \"$1\" in --version) echo 'git version 2.50.0';; *) echo 'https://github.com/sure-scale/doc-haus.git';; esac\n"],
    ["uname", "#!/bin/sh\necho Darwin\n"],
    ["osacompile", `#!/bin/bash
set -eu
[ "$1" = -s ] && [ "$2" = -o ]
[ ! -f "$HOME/fail-compile" ] || exit 9
mkdir -p "$3/Contents/MacOS" "$3/Contents/Resources"
printf '<plist version="1.0"><dict><key>CFBundleExecutable</key><string>applet</string><key>OSAAppletStayOpen</key><true/></dict></plist>' > "$3/Contents/Info.plist"
printf '#!/bin/sh\\nexit 0\\n' > "$3/Contents/MacOS/applet"
chmod +x "$3/Contents/MacOS/applet"
`],
    ["plutil", "#!/bin/sh\n[ \"$1\" != -extract ]\n"],
    ["codesign", "#!/bin/sh\nexit 0\n"],
    ["open", "#!/bin/bash\nprintf '%s\\0' \"$@\" >> \"$HOME/open-calls\"\n"],
  ] as const) {
    await writeFile(path.join(bin, file), source)
    await chmod(path.join(bin, file), 0o755)
  }
  await writeFile(path.join(checkout, "start.sh"), `#!/bin/bash
set -eu
cd "$(dirname "$0")"
. scripts/data-paths.sh
configure_data_paths "$PWD"
printf '%s\\0' "$PWD" "\${DOCHAUS_DATA_ROOT:-}" "$WORKSPACE_ROOT" "$DOCHAUS_RUNTIME_DIR" "\${DOC_HAUS_ENGINE_PORT:-4097}" "\${DOC_HAUS_INGEST_PORT:-4501}" "\${DOC_HAUS_WEB_PORT:-5174}" "\${OPENCODE_SERVER_PASSWORD:-}" > profile
`)
  await chmod(path.join(checkout, "start.sh"), 0o755)
  const environment = { HOME: home, PATH: `${bin}:/usr/bin:/bin:/usr/sbin:/sbin` }
  const app = path.join(home, "Applications/doc.haus.app"), resources = path.join(app, "Contents/Resources")
  const command = path.join(resources, "doc.haus.command")
  async function run(script: string, env = environment, args: string[] = []) {
    const child = Bun.spawn(["/bin/bash", script, ...args], { cwd: root, env, stdin: "ignore", stdout: "pipe", stderr: "pipe" })
    const [status, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()])
    return { status, stdout, stderr }
  }
  async function checked(script: string, env = environment, args: string[] = []) {
    const result = await run(script, env, args)
    if (result.status !== 0) throw new Error(`Launcher exited ${result.status}: ${result.stderr}`)
  }
  async function profile() { return (await readFile(path.join(checkout, "profile"), "utf8")).split("\0").slice(0, -1) }
  async function install(overrides: Record<string, string> = {}) {
    await checked(path.join(repository, "scripts/install.sh"), { ...environment, DOCHAUS_DIR: checkout, ...overrides })
    return profile()
  }
  async function reopenCommand() { await checked(command); return profile() }
  const dispatch = (action = "launch") => run(path.join(resources, "open-doc-haus.sh"), environment, [action])
  const refresh = () => run(path.join(repository, "scripts/macos-app.sh"), environment, [checkout, app])
  const openCalls = async () => (await readFile(path.join(home, "open-calls"), "utf8").catch(() => "")).split("\0").filter(Boolean)
  async function launchOwner(web = true) {
    const runtime = path.join(checkout, ".dochaus-runtime"), script = path.join(root, "start.sh"), childScript = path.join(root, "web.ts")
    const reservation = web ? undefined : Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response("unused") })
    const unusedPort = reservation?.port ?? 0
    reservation?.stop(true)
    await writeFile(childScript, `const server = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response("synthetic web") }); await Bun.write(process.argv[2]!, String(server.port))`)
    await writeFile(script, `#!/bin/bash
mkdir -p "$1/engine.lock"
echo "$$" > "$1/engine.lock/pid"
child=
trap 'if [ -n "$child" ]; then kill -TERM "$child" 2>/dev/null; wait "$child" 2>/dev/null; fi; exit 0' TERM
if [ "$4" = yes ]; then "$2" "$3" "$1/web-port" & child=$!; fi
while true; do sleep 0.1; done
`)
    const child = Bun.spawn(["/bin/bash", script, runtime, process.execPath, childScript, web ? "yes" : "no"], { cwd: checkout, stdout: "ignore", stderr: "ignore" })
    processes.push(child)
    for (let attempt = 0; attempt < 100; attempt++) {
      const owner = (await readFile(path.join(runtime, "engine.lock/pid"), "utf8").catch(() => "")).trim()
      const port = Number(await readFile(path.join(runtime, "web-port"), "utf8").catch(() => ""))
      if (owner === String(child.pid) && (!web || port)) return { child, port: web ? port : unusedPort }
      await Bun.sleep(20)
    }
    throw new Error("Fixture launcher did not claim its runtime")
  }
  return { root, home, checkout, app, resources, command, environment, install, reopenCommand, dispatch, refresh, openCalls, launchOwner }
}

test("the Mac app reopens the selected data paths and ports without retaining credentials", async () => {
  const f = await fixture()
  const selected = {
    DOCHAUS_DATA_ROOT: path.join(f.root, "data $(touch unexpected-data) 'quoted'"),
    WORKSPACE_ROOT: "selected matters",
    DOCHAUS_RUNTIME_DIR: path.join(f.root, "runtime\nwith newline"),
    DOC_HAUS_ENGINE_PORT: "49197", DOC_HAUS_INGEST_PORT: "49201", DOC_HAUS_WEB_PORT: "49274",
  }
  const initial = await f.install({ ...selected, OPENCODE_SERVER_PASSWORD: "synthetic-install-only" })
  const reopened = await f.reopenCommand()
  expect(reopened.slice(0, -1)).toEqual(initial.slice(0, -1))
  expect(reopened).toEqual([f.checkout, selected.DOCHAUS_DATA_ROOT, path.join(f.checkout, selected.WORKSPACE_ROOT),
    selected.DOCHAUS_RUNTIME_DIR, "49197", "49201", "49274", ""])
  const source = await readFile(path.join(f.resources, "installation.sh"), "utf8")
  expect(source).not.toContain("OPENCODE_SERVER_PASSWORD")
  expect(source).not.toContain("synthetic-install-only")
  expect(await Bun.file(path.join(f.root, "unexpected-data")).exists()).toBe(false)
  expect(await Bun.file(path.join(f.root, "unexpected-checkout")).exists()).toBe(false)
})

test("the Mac app keeps repository defaults when no installation overrides are selected", async () => {
  const f = await fixture()
  const initial = await f.install()
  expect(await f.reopenCommand()).toEqual(initial)
  expect(initial).toEqual([f.checkout, "", path.join(f.checkout, "workspace"), path.join(f.checkout, ".dochaus-runtime"), "4097", "4501", "5174", ""])
})

test("clicking an app with its own running stack opens only the browser", async () => {
  const f = await fixture()
  const owner = await f.launchOwner()
  await f.install({ DOC_HAUS_WEB_PORT: String(owner.port) })
  expect((await f.dispatch()).status).toBe(0)
  expect((await f.dispatch("waiting")).status).toBe(0)
  expect(await f.openCalls()).toEqual([`http://localhost:${owner.port}`, `http://localhost:${owner.port}`])
  expect(owner.child.exitCode).toBeNull()
})

test("cold start requests Terminal once during the brief debounce and can resume later", async () => {
  const f = await fixture()
  await f.install()
  expect((await f.dispatch()).stdout.trim()).toBe("started")
  expect((await f.dispatch("waiting")).status).toBe(0)
  expect(await f.openCalls()).toEqual(["-a", "Terminal", f.command])
  // A later click, after the applet's debounce expires, can start a stopped stack.
  expect((await f.dispatch()).stdout.trim()).toBe("started")
  expect(await f.openCalls()).toEqual(["-a", "Terminal", f.command, "-a", "Terminal", f.command])
})

test("a verified launcher still warming up is left alone", async () => {
  const f = await fixture()
  const owner = await f.launchOwner(false)
  await f.install({ DOC_HAUS_WEB_PORT: String(owner.port) })
  expect((await f.dispatch()).status).toBe(0)
  expect(await f.openCalls()).toEqual([])
  expect(owner.child.exitCode).toBeNull()
})

test("a different profile's HTTP listener is not opened even when this launcher is alive", async () => {
  const f = await fixture()
  const owner = await f.launchOwner()
  const unrelated = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response("unrelated profile") })
  servers.push(unrelated)
  await f.install({ DOC_HAUS_WEB_PORT: String(unrelated.port) })
  const result = await f.dispatch()
  expect(result.status).toBe(1)
  expect(result.stderr).toContain("web port belongs to another process")
  expect(await f.openCalls()).toEqual([])
  expect(owner.child.exitCode).toBeNull()
  expect((await fetch(`http://127.0.0.1:${unrelated.port}`)).status).toBe(200)
})

test("an unrelated live runtime owner is preserved and never opened or restarted", async () => {
  const f = await fixture()
  await f.install()
  const unrelated = Bun.spawn(["sleep", "30"], { cwd: f.checkout })
  processes.push(unrelated)
  const lock = path.join(f.checkout, ".dochaus-runtime/engine.lock")
  await mkdir(lock, { recursive: true })
  await writeFile(path.join(lock, "pid"), String(unrelated.pid))
  const result = await f.dispatch()
  expect(result.status).toBe(1)
  expect(result.stderr).toContain("Another process owns")
  expect(await f.openCalls()).toEqual([])
  expect(unrelated.exitCode).toBeNull()
  expect((await readFile(path.join(lock, "pid"), "utf8")).trim()).toBe(String(unrelated.pid))
})

test("bundle refresh uses candidate resources for an older checkout and retains recovery", async () => {
  const f = await fixture()
  await f.install()
  for (const file of ["macos-app.sh", "macos-app.applescript", "macos-app-launch.sh"]) await rm(path.join(f.checkout, "scripts", file))
  const result = await f.refresh()
  expect(result.status).toBe(0)
  expect(result.stdout).toContain("previous app kept")
  expect((await readdir(path.dirname(f.app))).some(file => file.startsWith("doc.haus.app.previous-"))).toBe(true)
  expect(await f.reopenCommand()).toEqual([f.checkout, "", path.join(f.checkout, "workspace"), path.join(f.checkout, ".dochaus-runtime"), "4097", "4501", "5174", ""])
  expect(await f.openCalls()).toEqual([])
})

test("a compilation failure leaves the previous app intact", async () => {
  const f = await fixture()
  await f.install()
  const previous = await readFile(path.join(f.resources, "installation.sh"), "utf8")
  await writeFile(path.join(f.home, "fail-compile"), "")
  expect((await f.refresh()).status).not.toBe(0)
  expect(await readFile(path.join(f.resources, "installation.sh"), "utf8")).toBe(previous)
  expect((await readdir(path.dirname(f.app))).filter(file => file.startsWith(".doc.haus-build."))).toEqual([])
  expect(await f.openCalls()).toEqual([])
})
