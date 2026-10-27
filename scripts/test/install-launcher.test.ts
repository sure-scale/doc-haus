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
  const refresh = (overrides: Record<string, string> = {}) => run(path.join(repository, "scripts/macos-app.sh"), { ...environment, ...overrides }, [checkout, app])
  const wordOpen = (uri: string) => run(path.join(resources, "open-word.sh"), environment, [uri])
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
  return { root, home, checkout, app, resources, command, environment, install, reopenCommand, dispatch, refresh, wordOpen, openCalls, launchOwner }
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

test("a running stack returns only its verified URL for the applet's browser selection", async () => {
  const f = await fixture()
  const owner = await f.launchOwner()
  await f.install({ DOC_HAUS_WEB_PORT: String(owner.port) })
  const ready = `ready http://localhost:${owner.port}`
  expect((await f.dispatch()).stdout.trim()).toBe(ready)
  expect((await f.dispatch("waiting")).stdout.trim()).toBe(ready)
  expect(await f.openCalls()).toEqual([])
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
  expect((await readdir(path.join(path.dirname(f.app), ".doc.haus-recovery"))).some(file => file.startsWith("doc.haus.app.previous-"))).toBe(true)
  expect((await readdir(path.dirname(f.app))).some(file => file.startsWith("doc.haus.app.previous-"))).toBe(false)
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

test("an unchanged desktop refresh skips compilation and retains the current bundle", async () => {
  const f = await fixture()
  await f.install()
  expect((await f.refresh()).status).toBe(0)
  const recovery = path.join(path.dirname(f.app), ".doc.haus-recovery")
  const previous = await readdir(recovery)
  await writeFile(path.join(f.home, "fail-compile"), "")
  const result = await f.refresh()
  expect(result.status).toBe(0)
  expect(result.stdout).toContain("already current")
  expect(await readdir(recovery)).toEqual(previous)
  expect(await f.openCalls()).toEqual([])
})

test("the bundled Word dispatcher preserves the fixed Bun path, selected runtime and literal URI", async () => {
  const f = await fixture()
  await f.install()
  const bun = path.join(f.root, "bun $(touch unexpected-helper)")
  await writeFile(bun, `#!/bin/bash
if [ "$1" = build ]; then
  [ ! -f "$HOME/fail-helper" ] || exit 9
  while [ "$1" != --outfile ]; do shift; done
  printf '// synthetic self-contained Word helper\\n' > "$2"
else
  printf '%s\\0' "$@" > "$HOME/word-helper-calls"
fi
`)
  await chmod(bun, 0o755)
  const data = path.join(f.root, "selected data 'quotes' $(touch unexpected-data)")
  expect((await f.refresh({ BUN_BIN: bun, DOCHAUS_DATA_ROOT: data, DOCHAUS_DESKTOP_HELPER: "1" })).status).toBe(0)
  const uri = "doc.haus://word/open?source=" + encodeURIComponent("https://localhost:4511/word/open/" + "A".repeat(43))
  expect((await f.wordOpen(uri)).status).toBe(0)
  expect((await readFile(path.join(f.home, "word-helper-calls"), "utf8")).split("\0").slice(0, -1))
    .toEqual([path.join(f.resources, "word-open.js"), "--runtime", path.join(data, "runtime"), uri])
  expect(await Bun.file(path.join(f.root, "unexpected-helper")).exists()).toBe(false)
  expect(await f.openCalls()).toEqual([])
  const previous = await readFile(path.join(f.resources, "installation.sh"), "utf8")
  const helper = await readFile(path.join(f.resources, "word-open.js"), "utf8")
  const recovery = path.join(path.dirname(f.app), ".doc.haus-recovery")
  const backups = await readdir(recovery)
  await writeFile(path.join(f.home, "fail-helper"), "")
  // A pre-dependency installer refresh must never downgrade an existing full
  // desktop helper to a launch-only bundle when the new build is unavailable.
  const optional = await f.refresh({ BUN_BIN: bun })
  expect(optional.status).toBe(0)
  expect(optional.stdout).toContain("keeping the existing desktop Word helper")
  expect((await f.refresh({ BUN_BIN: bun, DOCHAUS_DESKTOP_HELPER: "1" })).status).not.toBe(0)
  expect(await readFile(path.join(f.resources, "installation.sh"), "utf8")).toBe(previous)
  expect(await readFile(path.join(f.resources, "word-open.js"), "utf8")).toBe(helper)
  expect(await readdir(recovery)).toEqual(backups)
})

test("a failed app move restores the previous bundle from hidden recovery", async () => {
  const f = await fixture()
  await f.install()
  const previous = await readFile(path.join(f.resources, "installation.sh"), "utf8")
  const move = path.join(f.root, "bin/mv")
  await writeFile(move, `#!/bin/bash
if [[ "$1" == *"/.doc.haus-build."* ]] && [ "$2" = "$HOME/Applications/doc.haus.app" ]; then exit 9; fi
exec /bin/mv "$@"
`)
  await chmod(move, 0o755)
  expect((await f.refresh()).status).not.toBe(0)
  expect(await readFile(path.join(f.resources, "installation.sh"), "utf8")).toBe(previous)
  expect(await readdir(path.join(path.dirname(f.app), ".doc.haus-recovery"))).toEqual([])
  expect(await f.openCalls()).toEqual([])
})

test.skipIf(process.platform !== "darwin")("the applet isolates desktop browser instances and focuses only a matching tab", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-browser-selection-"))
  roots.push(root)
  const compiled = path.join(root, "launcher.scpt"), driver = path.join(root, "selection.applescript")
  // Loading a compiled script does not run its launch/browser handlers. Every
  // value below is synthetic; this test sends no browser or permission events.
  const compile = Bun.spawn(["/usr/bin/osacompile", "-o", compiled, path.join(repository, "scripts/macos-app.applescript")], { stdout: "pipe", stderr: "pipe" })
  const [compileStatus, compileError] = await Promise.all([compile.exited, new Response(compile.stderr).text()])
  expect(compileError).toBe("")
  expect(compileStatus).toBe(0)
  await writeFile(driver, `use framework "Foundation"
use scripting additions

on collectionOf(values)
    script collection
        property collectionValues : {}
        on |get|()
            return collectionValues
        end |get|
    end script
    set collection's collectionValues to values
    return collection
end collectionOf

on browserInstance(identifier, launchedThroughLS, policy, terminated, appURL)
    script instance
        property syntheticIdentifier : 0
        property launchDateValue : missing value
        property policyValue : 0
        property terminatedValue : false
        property bundleURLValue : missing value
        on launchDate()
            return launchDateValue
        end launchDate
        on activationPolicy()
            return policyValue
        end activationPolicy
        on isTerminated()
            return terminatedValue
        end isTerminated
        on bundleURL()
            return bundleURLValue
        end bundleURL
    end script
    set instance's syntheticIdentifier to identifier
    if launchedThroughLS then set instance's launchDateValue to current date
    set instance's policyValue to policy
    set instance's terminatedValue to terminated
    set instance's bundleURLValue to appURL
    return instance
end browserInstance

on browserTab(address, tabIndex)
    script browserTab
        property addressValue : ""
        property indexValue : 0
        on |URL|()
            return addressValue
        end |URL|
        on |index|()
            return indexValue
        end |index|
    end script
    -- ScriptingBridge returns Foundation strings for native properties.
    set browserTab's addressValue to current application's NSString's stringWithString:address
    set browserTab's indexValue to tabIndex
    return browserTab
end browserTab

on browserWindow(tabValues, selectedIndex, windowMode)
    script browserWindow
        property tabCollection : missing value
        property selected : 0
        property minimizedValue : true
        property indexValue : 2
        property modeValue : "normal"
        property familyValue : "chromium"
        on |mode|()
            return modeValue
        end |mode|
        on |tabs|()
            if modeValue is "incognito" then error "Private tabs must never be queried"
            return tabCollection
        end |tabs|
        on activeTabIndex()
            return selected
        end activeTabIndex
        on currentTab()
            return item selected of (tabCollection's |get|())
        end currentTab
        on setActiveTabIndex:chosenIndex
            set selected to chosenIndex
        end setActiveTabIndex:
        on setCurrentTab:chosenTab
            set selected to chosenTab's |index|()
        end setCurrentTab:
        on setMinimized:minimized
            if familyValue is not "chromium" then error "Safari uses Cocoa miniaturized, not Chromium minimized"
            set minimizedValue to minimized
        end setMinimized:
        on setMiniaturized:minimized
            if familyValue is not "safari" then error "Chromium uses minimized, not Cocoa miniaturized"
            set minimizedValue to minimized
        end setMiniaturized:
        on setIndex:windowIndex
            set indexValue to windowIndex
        end setIndex:
    end script
    set browserWindow's tabCollection to my collectionOf(tabValues)
    set browserWindow's selected to selectedIndex
    set browserWindow's modeValue to windowMode
    return browserWindow
end browserWindow

on browserWithWindows(windowValues)
    script browser
        property windowCollection : missing value
        on |windows|()
            return windowCollection
        end |windows|
    end script
    set browser's windowCollection to my collectionOf(windowValues)
    return browser
end browserWithWindows

on run argv
    set launcherFile to POSIX file (item 1 of argv)
    set launcher to load script launcherFile
    repeat with launchURL in {"doc.haus://launch", "doc.haus://launch/"}
        if not (launcher's isLaunchURL(contents of launchURL)) then error "Launch route was rejected"
    end repeat
    repeat with launchURL in {"doc.haus://launch?instance=https%3A%2F%2Flocalhost", "doc.haus://launch#fragment", "doc.haus://user@launch", "doc.haus://launch:80", "doc.haus://launch/extra", "doc.haus://word/open", "https://launch", "DOC.HAUS://launch", "doc.haus://launch%2F"}
        if (launcher's isLaunchURL(contents of launchURL)) then error "Unsupported launch route was accepted"
    end repeat
    set targetOrigin to "http://localhost:5274"
    repeat with candidateURL in {targetOrigin, targetOrigin & "/", targetOrigin & "/matter/synthetic", targetOrigin & "?setup", targetOrigin & "#matter"}
        if not launcher's matchesOrigin(contents of candidateURL, targetOrigin) then error "Expected configured origin to match"
    end repeat
    repeat with candidateURL in {"http://localhost:52740/", "http://localhost:5274.example/", "http://localhost:5274@unrelated.example/", "https://localhost:5274/", "http://127.0.0.1:5274/", "http://localhost:5275/", "unrelated"}
        if launcher's matchesOrigin(contents of candidateURL, targetOrigin) then error "Refused origin matched"
    end repeat
    set candidates to {{targetOrigin & "/matter/first", "https://unrelated.example/"}, {"https://other.example/", targetOrigin & "/matter/current"}}
    if (launcher's matchingTab(candidates, {2, 2}, targetOrigin)) is not {2, 2} then error "Current matter tab was not preferred"
    if (launcher's matchingTab(candidates, {2, 1}, targetOrigin)) is not {1, 1} then error "Existing matter tab was not found"
    if (launcher's matchingTab({{"http://localhost:52740/"}}, {1}, targetOrigin)) is not {0, 0} then error "Foreign tab was selected"
    set tabAddresses to {}
    repeat 27 times
        set end of tabAddresses to "https://unrelated.example/"
    end repeat
    set end of tabAddresses to targetOrigin & "/matter/current"
    set snapshots to {}
    copy tabAddresses to end of snapshots
    set item 28 of tabAddresses to "https://changed.example/"
    if (count item 1 of snapshots) is not 28 then error "Window snapshot must retain every tab"
    if (launcher's matchingTab(snapshots, {28}, targetOrigin)) is not {1, 28} then error "Concrete snapshot lost its active matter tab"
    set addressReference to a reference to item 28 of item 1 of snapshots
    if not (launcher's matchesOrigin(addressReference, targetOrigin)) then error "Referenced native text must match"
    if (launcher's browserFamily("com.google.Chrome")) is not "chromium" then error "Chrome family missing"
    if (launcher's browserFamily("com.apple.Safari")) is not "safari" then error "Safari family missing"
    if (launcher's browserFamily("org.mozilla.firefox")) is not "" then error "Unsupported browser should use normal opening"
    set expectedAppURL to current application's NSURL's fileURLWithPath:"/Applications/Synthetic Browser.app"
    set desktopInstance to my browserInstance(41, true, 0, false, expectedAppURL)
    set directInstance to my browserInstance(42, false, 0, false, expectedAppURL)
    repeat with instances in {{desktopInstance, directInstance}, {directInstance, desktopInstance}, {desktopInstance}}
        set chosenInstance to launcher's selectBrowserInstance(contents of instances, expectedAppURL)
        if chosenInstance's syntheticIdentifier is not 41 then error "Direct-spawn browser displaced the desktop instance"
    end repeat
    set anotherDesktop to my browserInstance(43, true, 0, false, expectedAppURL)
    set backgroundInstance to my browserInstance(44, true, 1, false, expectedAppURL)
    set terminatedInstance to my browserInstance(45, true, 0, true, expectedAppURL)
    set otherBundle to my browserInstance(46, true, 0, false, current application's NSURL's fileURLWithPath:"/Applications/Other Browser.app")
    repeat with instances in {{}, {directInstance}, {desktopInstance, anotherDesktop}, {backgroundInstance}, {terminatedInstance}, {otherBundle}}
        if (launcher's selectBrowserInstance(contents of instances, expectedAppURL)) is not missing value then error "Ambiguous or ineligible browser must open normally"
    end repeat
    set tabValues to {}
    repeat with tabIndex from 1 to 27
        set end of tabValues to my browserTab("https://unrelated.example/", tabIndex)
    end repeat
    set matterTab to my browserTab(targetOrigin & "/matter/synthetic?view=documents", 28)
    set end of tabValues to matterTab
    set normalWindow to my browserWindow(tabValues, 28, "normal")
    set privateWindow to my browserWindow({}, 0, "incognito")
    set browser to my browserWithWindows({privateWindow, normalWindow})
    if not (launcher's focusBrowserTab(browser, "chromium", targetOrigin)) then error "PID browser transport lost the matching tab"
    if normalWindow's selected is not 28 or normalWindow's minimizedValue or normalWindow's indexValue is not 1 then error "Matching window was not focused"
    if (matterTab's addressValue as text) is not (targetOrigin & "/matter/synthetic?view=documents") then error "Matter path was changed"
    set safariWindow to my browserWindow({my browserTab("https://unrelated.example/", 1), my browserTab(targetOrigin & "/matter/safari", 2)}, 1, "normal")
    set safariWindow's familyValue to "safari"
    if not (launcher's focusBrowserTab(my browserWithWindows({safariWindow}), "safari", targetOrigin)) then error "Safari native tab selection failed"
    if safariWindow's selected is not 2 or safariWindow's minimizedValue or safariWindow's indexValue is not 1 then error "Safari did not select and reveal its matching tab"
    set foreignWindow to my browserWindow({my browserTab("http://localhost:52740/", 1)}, 1, "normal")
    if (launcher's focusBrowserTab(my browserWithWindows({foreignWindow}), "chromium", targetOrigin)) then error "Foreign origin was focused"
    if foreignWindow's minimizedValue is not true or foreignWindow's indexValue is not 2 then error "An unrelated window was modified"
    set startedAt to current date
    if (launcher's launchGraceRemaining(missing value, startedAt)) is not 0 then error "Ready interface should exit promptly"
    if (launcher's launchGraceRemaining(startedAt, startedAt)) is not 10 then error "Cold launch needs its debounce grace"
    if (launcher's launchGraceRemaining(startedAt, startedAt + 3)) is not 7 then error "Duplicate opens must not extend the grace"
    if (launcher's launchGraceRemaining(startedAt, startedAt + 10)) is not 0 then error "Cold launch grace should end"
    if (launcher's launchGraceRemaining(startedAt, startedAt + 20)) is not 0 then error "Expired launcher should exit"
    return "passed"
end run
`)
  const child = Bun.spawn(["/usr/bin/osascript", driver, compiled], { stdout: "pipe", stderr: "pipe" })
  const [status, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()])
  expect(stderr).toBe("")
  expect(status).toBe(0)
  expect(stdout.trim()).toBe("passed")
})
