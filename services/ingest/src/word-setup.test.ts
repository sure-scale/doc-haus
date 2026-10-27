import { afterEach, expect, test } from "bun:test"
import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { prepareLocalCertificate } from "../../../scripts/setup-word-addin"
import { wordConfigurationFile, wordHttpsConfiguration } from "./word-configuration"
import { startWordHttps, stopWordHttps, wordHttpsAvailable } from "./word-https"
import { localWordSetupRequest, wordSetupRoutes } from "./word-setup"

const roots: string[] = []
afterEach(() => {
  stopWordHttps()
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})

function fixture() {
  const root = mkdtempSync(path.join(os.tmpdir(), "dochaus-word-setup-"))
  roots.push(root)
  const assets = path.join(root, "assets")
  for (const file of ["pane.html", "pane.js", "pane.css", "desk.html", "desk.js", "desk.css", "assets/icon.png"]) {
    mkdirSync(path.dirname(path.join(assets, file)), { recursive: true })
    writeFileSync(path.join(assets, file), file === "pane.html" ? "<script src='office.js'></script>" : "fixture")
  }
  mkdirSync(path.join(root, "Applications/Microsoft Word.app"), { recursive: true })
  const reservation = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
  const port = reservation.port
  reservation.stop(true)
  const environment = { DOCHAUS_DATA_ROOT: root, DOCHAUS_WORD_ORIGIN: `https://localhost:${port}` }
  return { root, assets, environment }
}

function request(method = "GET", origin = "http://localhost:5174", body = "{}") {
  return new Request("http://127.0.0.1:4501/word/setup", { method, headers: { Origin: origin, "Content-Type": "application/json" }, ...(method === "POST" ? { body } : {}) })
}

test("setup requires a local explicit JSON request and refuses foreign origins and options", async () => {
  const item = fixture()
  let commands = 0
  const app = wordSetupRoutes({ ...item, home: item.root, platform: "darwin", fetch: () => new Response(), command: async () => { commands++; return { code: 1, stdout: "" } } })
  expect((await app.fetch(request("POST", "https://unrelated.example"))).status).toBe(403)
  expect((await app.fetch(new Request("http://unrelated.example/word/setup", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" }))).status).toBe(403)
  expect((await app.fetch(request("POST", "null"))).status).toBe(403)
  expect((await app.fetch(request("POST", undefined, '{"origin":"https://unrelated.example"}'))).status).toBe(400)
  expect((await app.fetch(new Request("http://127.0.0.1/word/setup", { method: "POST", body: "{}" }))).status).toBe(415)
  expect((await app.fetch(request("POST", undefined, " ".repeat(129)))).status).toBe(400)
  expect(commands).toBe(0)
  expect(localWordSetupRequest(new Request("http://localhost/word/setup", { headers: { "Sec-Fetch-Site": "cross-site" } }))).toBe(false)
  expect(localWordSetupRequest(new Request("http://127.0.0.1/word/setup", { headers: { Origin: "http://localhost:5174", "Sec-Fetch-Site": "cross-site" } }))).toBe(true)
})

test("explicit setup confirms trust, repairs registration, activates immediately and is idempotent", async () => {
  const item = fixture()
  let trusted = false
  const commands: string[][] = []
  const app = wordSetupRoutes({ ...item, home: item.root, platform: "darwin", fetch: () => new Response(), command: async args => {
    commands.push(args)
    if (args[1] === "verify-cert") return { code: trusted ? 0 : 1, stdout: "" }
    if (args[1] === "add-trusted-cert") trusted = true
    return { code: 0, stdout: "" }
  } })
  expect(await (await app.fetch(request())).json()).toMatchObject({ status: "not_configured", available: false, canSetup: true })
  expect(commands).toHaveLength(0)
  const responses = await Promise.all([app.fetch(request("POST")), app.fetch(request("POST"))])
  expect(responses.map(response => response.status)).toEqual([200, 200])
  const ready = await responses[0].json()
  expect(ready).toMatchObject({ status: "ready", available: true, trusted: true, registered: true, running: true, activatedWithoutRestart: true, restartWord: true })
  expect(commands.filter(args => args[0] === "osascript")).toHaveLength(1)
  expect(commands.filter(args => args[1] === "add-trusted-cert")).toHaveLength(1)
  expect(commands.findIndex(args => args[0] === "osascript")).toBeLessThan(commands.findIndex(args => args[1] === "add-trusted-cert"))
  const file = wordConfigurationFile(item.environment)
  const saved = readFileSync(file, "utf8")
  const certificate = readFileSync(JSON.parse(saved).certificate, "utf8")
  expect(await (await app.fetch(request("POST"))).json()).toMatchObject({ status: "ready", available: true, restartWord: false })
  expect(readFileSync(file, "utf8")).toBe(saved)
  expect(readFileSync(JSON.parse(saved).certificate, "utf8")).toBe(certificate)
  expect(commands.filter(args => args[0] === "osascript")).toHaveLength(1)
  expect(JSON.stringify(ready)).not.toContain(item.root)
  trusted = false
  expect(await (await app.fetch(request())).json()).toMatchObject({ available: false, trusted: false, status: "needs_repair" })
  expect(commands.filter(args => args[0] === "osascript")).toHaveLength(1)
})

test("a cancelled native confirmation does not trust, register, or mark setup ready", async () => {
  const item = fixture()
  const commands: string[][] = []
  const app = wordSetupRoutes({ ...item, home: item.root, platform: "darwin", fetch: () => new Response(), command: async args => {
    commands.push(args)
    return { code: 1, stdout: "private native diagnostics" }
  } })
  const response = await app.fetch(request("POST"))
  expect(response.status).toBe(503)
  const result = await response.json()
  expect(result).toMatchObject({ available: false, registered: false, trusted: false })
  expect(commands.some(args => args[1] === "add-trusted-cert")).toBe(false)
  expect(JSON.stringify(result)).not.toContain(item.root)
  expect(JSON.stringify(result)).not.toContain("private native")
})

test("a saved local configuration is unavailable until its actual listener and assets are active", async () => {
  const item = fixture()
  const tls = await prepareLocalCertificate(path.join(item.root, "certificates"))
  const file = wordConfigurationFile(item.environment)
  mkdirSync(path.dirname(file), { recursive: true })
  writeFileSync(file, JSON.stringify({ origin: item.environment.DOCHAUS_WORD_ORIGIN, external: false, ...tls }))
  const configuration = wordHttpsConfiguration(item.environment)!
  expect(wordHttpsAvailable(configuration)).toBe(false)
  const server = startWordHttps(() => new Response(), item.environment, item.assets)
  expect(server).toBeDefined()
  const verified = async (route: string) => {
    const child = Bun.spawn(["curl", "--silent", "--show-error", "--noproxy", "*", "--max-time", "5", "--cacert", tls.certificate,
      "--write-out", "\\n%{http_code}", `${item.environment.DOCHAUS_WORD_ORIGIN}${route}`], { stdout: "pipe", stderr: "pipe" })
    const [code, body, error] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()])
    expect(error).toBe("")
    expect(code).toBe(0)
    return body
  }
  expect(await verified("/word-addin/pane.html")).toContain("office.js")
  expect(await verified("/word/setup")).toEndWith("404")
  expect(wordHttpsAvailable(configuration)).toBe(true)
  expect(startWordHttps(() => new Response(), item.environment, item.assets)).toBe(server)
  rmSync(path.join(item.assets, "pane.js"))
  expect(wordHttpsAvailable(configuration)).toBe(false)
  stopWordHttps()
  expect(wordHttpsAvailable(configuration)).toBe(false)
})

test("UI setup preserves administrator HTTPS configuration and refuses public listener overrides", async () => {
  const item = fixture()
  const file = wordConfigurationFile(item.environment)
  mkdirSync(path.dirname(file), { recursive: true })
  const saved = JSON.stringify({ origin: "https://127.0.0.1:1", external: true })
  writeFileSync(file, saved)
  let commands = 0
  const common = { ...item, home: item.root, platform: "darwin" as const, fetch: () => new Response(), command: async () => { commands++; return { code: 1, stdout: "" } } }
  const app = wordSetupRoutes({ ...common, environment: { DOCHAUS_DATA_ROOT: item.root } })
  expect(await (await app.fetch(request())).json()).toMatchObject({ status: "external", canSetup: false, available: false })
  expect((await app.fetch(request("POST"))).status).toBe(409)
  expect(readFileSync(file, "utf8")).toBe(saved)
  rmSync(file)
  const publicListener = wordSetupRoutes({ ...common, environment: { ...item.environment, DOCHAUS_WORD_HOST: "0.0.0.0" } })
  expect(await (await publicListener.fetch(request())).json()).toMatchObject({ canSetup: false, available: false })
  expect((await publicListener.fetch(request("POST"))).status).toBe(409)
  expect(commands).toBe(0)
})

test("a network-facing ingest bind refuses native setup despite a spoofed localhost Host", async () => {
  const item = fixture()
  const commands: string[][] = []
  const app = wordSetupRoutes({ ...item, environment: { ...item.environment, INGEST_HOST: "0.0.0.0" }, home: item.root, platform: "darwin",
    fetch: () => new Response(), command: async args => { commands.push(args); return { code: 0, stdout: "" } } })
  const server = Bun.serve({ hostname: "0.0.0.0", port: 0, fetch: app.fetch })
  try {
    const target = `http://127.0.0.1:${server.port}/word/setup`
    const headers = { Host: `localhost:${server.port}`, "Content-Type": "application/json" }
    expect((await fetch(target, { method: "POST", headers, body: "{}" })).status).toBe(403)
    expect(await (await fetch(target, { headers })).json()).toMatchObject({ canSetup: false })
    expect(commands).toHaveLength(0)
    expect(existsSync(wordConfigurationFile(item.environment))).toBe(false)
    expect(existsSync(path.join(item.root, "Library/Containers/com.microsoft.Word/Data/Documents/wef"))).toBe(false)
  } finally { server.stop(true) }
})
