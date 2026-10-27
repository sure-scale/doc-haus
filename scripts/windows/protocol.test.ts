import { expect, test } from "bun:test"
import path from "node:path"
import { dispatchProtocol, isLaunchURI, launchFromProtocol, listenerBelongsTo, resolveProfile } from "./protocol"

const repository = path.resolve(import.meta.dir, "../..")
const profile = () => resolveProfile(repository, {
  repository,
  environment: { DOCHAUS_RUNTIME_DIR: "selected runtime", DOC_HAUS_WEB_PORT: "5274" },
}, {})

function fixture() {
  const selected = profile()
  const effects: string[] = []
  const files = new Map([[path.join(selected.runtime, "engine.lock/pid"), "100"], [path.join(selected.runtime, "engine-process.lock"), "101"]])
  const live = new Set(["100", "101"])
  const listeners = (pid = "300", port = "5274") => `  TCP    127.0.0.1:${port}    0.0.0.0:0    LISTENING    ${pid}\r\n`
  let snapshot = listeners()
  let verified = true, ready = true, lockExists = true
  const parents = new Map([["300", "200"], ["200", "100"]])
  const system: Parameters<typeof launchFromProtocol>[2] = {
    read: file => files.get(file) ?? "",
    exists: () => lockExists,
    running: pid => live.has(pid),
    ownerRunning: async () => verified,
    netstat: async () => snapshot,
    parents: async () => parents,
    ready: async () => ready,
    open: async url => { effects.push(`open ${url}`) },
    start: async value => { effects.push(`start ${value.repository}`) },
  }
  return { selected, system, effects, files, live, parents, listeners,
    setSnapshot: (value: string) => { snapshot = value },
    setVerified: (value: boolean) => { verified = value },
    setReady: (value: boolean) => { ready = value },
    cold: () => { live.clear(); files.clear(); lockExists = false },
  }
}

test("only the no-argument launch route is accepted before any settings or system action", async () => {
  expect(isLaunchURI(["doc.haus://launch"])).toBe(true)
  expect(isLaunchURI(["doc.haus://launch/"])).toBe(true)
  const f = fixture()
  let loaded = false
  for (const args of [[], ["doc.haus://launch?instance=other"], ["doc.haus://launch#hash"], ["doc.haus://user@launch"],
    ["doc.haus://launch:80"], ["doc.haus://launch/extra"], ["doc.haus://launch%2F"], ["DOC.HAUS://launch"],
    ["doc.haus://word/open"], ["doc.haus://launch", "--eval", "arbitrary program"], ["doc.haus://launch\" & start unrelated & \""]]) {
    expect(isLaunchURI(args)).toBe(false)
    await expect(launchFromProtocol(args, () => { loaded = true; return f.selected }, f.system)).rejects.toThrow("not supported")
  }
  expect(loaded).toBe(false)
  expect(f.effects).toEqual([])
})

test("selected profile paths stay literal and local settings have the launcher's precedence", () => {
  const environment = { DOCHAUS_DATA_ROOT: "data & %literal%", DOC_HAUS_WEB_PORT: "5274" }
  const selected = resolveProfile(repository, { repository, environment }, { DOC_HAUS_WEB_PORT: "9999" }, "DOC_HAUS_WEB_PORT=5374\r\nPROVIDER_SECRET=not-retained\r\n")
  expect(selected.runtime).toBe(path.join(repository, "data & %literal%", "runtime"))
  expect(selected.webURL).toBe("http://localhost:5374")
  expect(selected.environment.DOC_HAUS_WEB_PORT).toBe("5274")
  expect(selected.environment).not.toHaveProperty("PROVIDER_SECRET")
  const defaults = resolveProfile(repository, { repository, environment: {} }, { DOCHAUS_DATA_ROOT: "another profile", DOC_HAUS_WEB_PORT: "9999" })
  expect(defaults.runtime).toBe(path.join(repository, ".dochaus-runtime"))
  expect(defaults.webURL).toBe("http://localhost:5174")
  expect(() => resolveProfile(repository, { repository: "another installation", environment }, {})).toThrow("installation changed")
  expect(() => resolveProfile(repository, { repository, environment: { PROVIDER_SECRET: "not-allowed" } }, {})).toThrow("settings are not valid")
  for (const port of ["0", "65536", "5374 & arbitrary", "http://foreign.example"])
    expect(() => resolveProfile(repository, { repository, environment: { DOC_HAUS_WEB_PORT: port } }, {})).toThrow("port setting")
})

test("Word dispatch validates the shared schema before settings or native actions and passes only the selected runtime", async () => {
  const f = fixture()
  let loaded = 0
  const load = () => { loaded++; return f.selected }
  const native: { uri: string; runtime: string }[] = []
  const openWord: Parameters<typeof dispatchProtocol>[3] = async options => {
    native.push({ uri: options.uri, runtime: options.runtime })
    return { path: "synthetic.docx", reused: false }
  }
  for (const args of [["doc.haus://word/open?source=https%3A%2F%2Funrelated.example%2Fother"],
    ["doc.haus://word/open?source=malformed#fragment"], ["doc.haus://word/open", "--eval", "program"]])
    await expect(dispatchProtocol(args, load, f.system, openWord)).rejects.toThrow()
  expect(loaded).toBe(0)
  const uri = "doc.haus://word/open?source=" + encodeURIComponent("https://localhost:4511/word/open/" + "A".repeat(43))
  await dispatchProtocol([uri], load, f.system, openWord)
  expect(native).toEqual([{ uri, runtime: f.selected.runtime }])
  expect(f.effects).toEqual([])
})

test("a verified running stack opens only its own ready web port without restarting it", async () => {
  const f = fixture()
  await launchFromProtocol(["doc.haus://launch/"], () => f.selected, f.system)
  expect(f.effects).toEqual(["open http://localhost:5274"])
  f.effects.length = 0
  f.setReady(false)
  await launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)
  expect(f.effects).toEqual([])
  f.setSnapshot("")
  await launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)
  expect(f.effects).toEqual([])
})

test("foreign listeners, uncertain ownership and ownership changes leave every process alone", async () => {
  const f = fixture()
  f.setSnapshot(f.listeners() + f.listeners("400"))
  await expect(launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)).rejects.toThrow("another process")
  f.setSnapshot(f.listeners())
  f.setVerified(false)
  await expect(launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)).rejects.toThrow("Another process owns")
  let checks = 0
  f.system.ownerRunning = async () => ++checks === 1
  await expect(launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)).rejects.toThrow("owner changed")
  expect(f.effects).toEqual([])
  expect(listenerBelongsTo("300", "100", new Map([["300", "400"], ["400", "300"]]))).toBe(false)
})

test("a stopped installation delegates once to ordinary startup while orphan engines and occupied ports block it", async () => {
  const f = fixture()
  f.cold()
  f.setSnapshot("")
  await launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)
  expect(f.effects).toEqual([`start ${repository}`])
  f.effects.length = 0
  f.setSnapshot(f.listeners("700", "4097"))
  await expect(launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)).rejects.toThrow("already in use")
  f.setSnapshot("")
  f.files.set(path.join(f.selected.runtime, "engine-process.lock"), "101")
  f.live.add("101")
  await expect(launchFromProtocol(["doc.haus://launch"], () => f.selected, f.system)).rejects.toThrow("without its launcher")
  expect(f.effects).toEqual([])
})

test.skipIf(process.platform !== "win32")("the Windows installer remains valid PowerShell without running it", async () => {
  const child = Bun.spawn(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
    "$tokens = $null; $errors = $null; [System.Management.Automation.Language.Parser]::ParseFile($env:DOCHAUS_TEST_INSTALLER, [ref]$tokens, [ref]$errors) | Out-Null; if ($errors.Count) { exit 1 }"],
  { env: { ...process.env, DOCHAUS_TEST_INSTALLER: path.join(repository, "scripts/install.ps1") }, stdout: "pipe", stderr: "pipe", windowsHide: true })
  const [status, errors] = await Promise.all([child.exited, new Response(child.stderr).text()])
  expect(errors).toBe("")
  expect(status).toBe(0)
})
