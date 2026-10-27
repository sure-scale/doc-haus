import { afterEach, expect, test } from "bun:test"
import { mkdir, mkdtemp, readFile, realpath, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { refreshDesktopLauncher } from "../desktop-launcher"
import { registerWindowsProtocol } from "../windows/register-protocol"
import { setupWordAddin } from "../setup-word-addin"

const roots: string[] = []
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }) })

test("desktop refresh uses fixed native helpers and forwards only selected paths and required process environment", async () => {
  const calls: { args: string[]; environment: Record<string, string | undefined> }[] = []
  const options = { home: path.join(os.tmpdir(), "synthetic 'home'"), environment: { DOCHAUS_DATA_ROOT: "selected data", DOC_HAUS_WEB_PORT: "5274", PROVIDER_SECRET: "never-forward", BUN_BIN: "unrelated-bun", PATH: "synthetic-path" },
    command: async (args: string[], environment: Record<string, string | undefined>) => { calls.push({ args, environment }); return { code: 0, stdout: "" } } }
  await refreshDesktopLauncher({ ...options, platform: "darwin" })
  expect(calls[0].args[0]).toBe("/bin/bash")
  expect(calls[0].args.at(-1)).toBe(path.join(options.home, "Applications/doc.haus.app"))
  expect(calls[0].environment).toMatchObject({ DOCHAUS_DATA_ROOT: "selected data", DOC_HAUS_WEB_PORT: "5274", BUN_BIN: process.execPath, DOCHAUS_DESKTOP_HELPER: "1" })
  expect(calls[0].environment).not.toHaveProperty("PROVIDER_SECRET")
  expect(calls[1].args.slice(1)).toEqual(["-f", path.join(options.home, "Applications/doc.haus.app")])
  await refreshDesktopLauncher({ ...options, platform: "win32" })
  expect(calls[2].args[0]).toBe(process.execPath)
  expect(calls[2].args[1]).toEndWith(path.join("scripts", "windows", "register-protocol.ts"))
  await refreshDesktopLauncher({ ...options, platform: "linux" })
  expect(calls).toHaveLength(3)
  let unchangedCalls = 0
  await refreshDesktopLauncher({ ...options, platform: "darwin", command: async () => { unchangedCalls++; return { code: 0, stdout: "doc.haus: Mac app is already current.\n" } } })
  expect(unchangedCalls).toBe(2)
  await expect(refreshDesktopLauncher({ ...options, platform: "darwin", command: async () => ({ code: 1, stdout: "" }) })).rejects.toThrow("Settings > Add-ins")
})

test("Windows registration uses per-user registry argv and persists only the literal selected profile", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-register-"))
  roots.push(root)
  const repository = path.join(root, "doc haus & %literal%"), local = path.join(root, "local app data")
  await mkdir(path.join(repository, "scripts/icons"), { recursive: true })
  await writeFile(path.join(repository, "scripts/icons/doc-haus.ico"), "synthetic icon")
  const calls: string[][] = []
  const bun = "C:\\Program Files\\Bun & tools\\bun.exe"
  const options = { repository, bun, environment: { LOCALAPPDATA: local, doc_haus_web_port: "5374", DOCHAUS_DATA_ROOT: "literal data & %value%", PROVIDER_SECRET: "never-persist" },
    command: async (args: string[]) => { calls.push(args); return 0 } }
  await registerWindowsProtocol(options)
  const canonicalRepository = await realpath(repository)
  const profile = JSON.parse(await readFile(path.join(local, "doc.haus/launch-profile.json"), "utf8"))
  expect(profile.environment).toEqual({ DOCHAUS_DATA_ROOT: "literal data & %value%", DOC_HAUS_WEB_PORT: "5374" })
  expect(calls.every(args => args[0] === "reg.exe" && args[2].startsWith("HKCU\\Software\\Classes\\doc.haus"))).toBe(true)
  expect(calls.find(args => args[2].endsWith("\\command"))?.at(-2))
    .toBe(`"${bun}" "${path.join(canonicalRepository, "scripts/windows/protocol.ts")}" "%1"`)
  expect(calls.find(args => args[2].endsWith("\\DefaultIcon"))?.at(-2)).toBe(`"${path.join(canonicalRepository, "scripts/icons/doc-haus.ico")}",0`)
  await expect(registerWindowsProtocol({ ...options, command: async () => 1 })).rejects.toThrow("could not be refreshed")
})

test("automatic desktop refresh preserves the selected repository and runtime on both platforms", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-desktop-selection-"))
  roots.push(root)
  const repository = await realpath(path.resolve(import.meta.dir, "../.."))
  const runtime = path.join(root, "selected runtime"), otherRuntime = path.join(root, "other runtime")
  const otherRepository = path.join(root, "other checkout")
  await mkdir(otherRepository)
  for (const platform of ["darwin", "win32"] as const) {
    const home = path.join(root, platform), local = path.join(home, "local app data")
    const metadata = platform === "darwin" ? path.join(home, "Applications/doc.haus.app/Contents/Resources/installation.json") : path.join(local, "doc.haus/launch-profile.json")
    await mkdir(path.dirname(metadata), { recursive: true })
    const selection = (selectedRepository = repository) => JSON.stringify(platform === "darwin" ? { repository: selectedRepository, runtime } : { repository: selectedRepository, environment: { DOCHAUS_RUNTIME_DIR: runtime } })
    const calls: string[][] = []
    const options = { platform, home, startup: true, environment: { LOCALAPPDATA: local, DOCHAUS_RUNTIME_DIR: runtime },
      command: async (args: string[]) => { calls.push(args); return { code: 0, stdout: "" } } }
    const mutations = () => calls.filter(args => args[0] !== "/usr/bin/codesign" && args[0] !== "/bin/bash" || args[1]?.endsWith("macos-app.sh"))
    // A different data profile and a different checkout must both leave the
    // actual persisted desktop selection untouched.
    await writeFile(metadata, selection())
    await refreshDesktopLauncher({ ...options, environment: { ...options.environment, DOCHAUS_RUNTIME_DIR: otherRuntime } })
    expect(mutations()).toEqual([])
    expect(await readFile(metadata, "utf8")).toBe(selection())
    await writeFile(metadata, selection(otherRepository))
    await refreshDesktopLauncher(options)
    expect(mutations()).toEqual([])
    expect(await readFile(metadata, "utf8")).toBe(selection(otherRepository))
    // The same resolved runtime can refresh without changing the selection.
    await writeFile(metadata, selection())
    await refreshDesktopLauncher(options)
    expect(mutations()).toHaveLength(platform === "darwin" ? 2 : 1)
    calls.length = 0
    await rm(metadata)
    await refreshDesktopLauncher(options)
    expect(mutations()).toEqual([])
    // Settings and the existing explicit installer flag may select a copy.
    await refreshDesktopLauncher({ ...options, startup: false })
    expect(mutations()).toHaveLength(platform === "darwin" ? 2 : 1)
    calls.length = 0
    await refreshDesktopLauncher({ ...options, environment: { ...options.environment, DOCHAUS_SETUP_WORD_ADDIN: "1" } })
    expect(mutations()).toHaveLength(platform === "darwin" ? 2 : 1)
  }
})

test.skipIf(process.platform === "win32")("legacy signed Mac selection clears the new stack's inherited runtime before comparison", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-legacy-selection-"))
  roots.push(root)
  const repository = await realpath(path.resolve(import.meta.dir, "../.."))
  const resources = path.join(root, "Applications/doc.haus.app/Contents/Resources")
  await mkdir(resources, { recursive: true })
  // This fixture is the same literal export shape written by the old builder.
  const quote = (value: string) => "'" + value.replaceAll("'", "'\\''") + "'"
  await writeFile(path.join(resources, "installation.sh"), `export DOCHAUS_INSTALL_DIR=${quote(repository)}\n`)
  const calls: string[][] = []
  let trusted = true
  const command = async (args: string[], environment: Record<string, string | undefined>) => {
    calls.push(args)
    if (args[0] === "/usr/bin/codesign") return { code: trusted ? 0 : 1, stdout: "" }
    if (args[0] === "/bin/bash" && args[1] === "-c") {
      // Only the synthetic signed-profile reader runs. No build, registration,
      // app launch, browser event or product stack is executed by this test.
      const child = Bun.spawn(args, { env: environment, stdout: "pipe", stderr: "pipe" })
      return { code: await child.exited, stdout: await new Response(child.stdout).text() }
    }
    return { code: 0, stdout: "" }
  }
  const options = { platform: "darwin" as const, home: root, startup: true, command }
  await refreshDesktopLauncher({ ...options, environment: { DOCHAUS_RUNTIME_DIR: path.join(root, "other runtime") } })
  expect(calls.map(args => args[1])).toEqual(["--verify", "-c"])
  calls.length = 0
  await refreshDesktopLauncher({ ...options, environment: {} })
  expect(calls).toHaveLength(4)
  expect(calls[2][1]).toEndWith("macos-app.sh")
  calls.length = 0
  trusted = false
  await refreshDesktopLauncher({ ...options, environment: {} })
  expect(calls).toHaveLength(1)
})

test("explicit Word setup refreshes an existing desktop installation through the same bounded command seam", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-word-desktop-refresh-"))
  roots.push(root)
  const calls: string[][] = []
  await setupWordAddin([], { platform: "darwin", home: root,
    environment: { DOCHAUS_DATA_ROOT: root, DOCHAUS_WORD_ORIGIN: "https://word.synthetic.example", DOCHAUS_WORD_EXTERNAL_HTTPS: "1" },
    command: async args => { calls.push(args); return { code: 0, stdout: "" } } })
  expect(calls).toHaveLength(2)
  expect(calls[0][0]).toBe("/bin/bash")
  expect(calls[0][1]).toEndWith(path.join("scripts", "macos-app.sh"))
  expect(calls[0].at(-1)).toBe(path.join(root, "Applications/doc.haus.app"))
})
