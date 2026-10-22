import { expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { engineConfigFile, patchEngineConfig, readEngineConfig } from "./engine-config"
import { providerSetupRoutes } from "./provider-setup"

const selection = { sessionID: "synthetic-session", primary: "primary", fast: "fast" }
const patch = { model: "synthetic/primary", small_model: "synthetic/fast", provider: { synthetic: { options: {} } } }
type RequestRecord = { action: string; input?: unknown; authorization: string | null; signal: AbortSignal | null | undefined }
function fixture(options: { patch?: unknown; fail?: string; throwOn?: string; rollbackFails?: boolean; writeFails?: boolean } = {}) {
  const calls: RequestRecord[] = []
  const written: unknown[] = []
  const order: string[] = []
  const transport = async (url: string, init: RequestInit) => {
    const action = new URL(url).pathname.split("/").at(-1)!
    order.push(action)
    calls.push({ action, input: init.body ? JSON.parse(String(init.body)) : undefined, authorization: new Headers(init.headers).get("authorization"), signal: init.signal })
    if (options.throwOn === action) throw new Error("synthetic-secret-from-transport")
    if (options.fail === action || (action === "rollback" && options.rollbackFails)) return Response.json({ error: "synthetic-secret-from-engine" }, { status: 400 })
    return Response.json(action === "prepare" ? { patch: options.patch ?? patch } : { ok: true })
  }
  const app = providerSetupRoutes({ engine: "http://engine.invalid", password: () => "synthetic-password", transport, writeSettings(value) {
    order.push("write")
    if (options.writeFails) throw new Error("synthetic-secret-from-writer")
    written.push(value)
  } })
  const save = (input: unknown = selection, headers: Record<string, string> = {}) => app.request("/engine/provider-setup", { method: "POST", headers: { "Content-Type": "application/json", ...headers }, body: JSON.stringify(input) })
  return { app, save, calls, written, order, transport }
}

test("combined Save validates tested models before credential activation and one settings write", async () => {
  const f = fixture()
  const response = await f.save()
  expect(response.status).toBe(200)
  expect(await response.json()).toEqual({ ok: true })
  expect(response.headers.get("cache-control")).toBe("no-store")
  expect(f.order).toEqual(["prepare", "save", "write", "session"])
  expect(f.written).toEqual([patch])
  expect(f.calls.slice(0, 2).map(call => call.input)).toEqual([selection, selection])
  expect(f.calls.every(call => call.authorization === `Basic ${btoa("opencode:synthetic-password")}` && call.signal instanceof AbortSignal)).toBe(true)
})

test("failed prepare never activates credentials or writes settings and hides engine detail", async () => {
  const f = fixture({ fail: "prepare" })
  const response = await f.save()
  expect(response.status).toBe(400)
  expect(await response.text()).not.toContain("synthetic-secret")
  expect(f.order).toEqual(["prepare"])
  expect(f.written).toEqual([])
})

test("invalid settings or mismatched selected models never activate credentials", async () => {
  for (const invalid of [
    { ...patch, provider: { synthetic: { options: { apiKey: "synthetic-secret" } } } },
    { ...patch, model: "synthetic/untested" },
    { ...patch, small_model: "other/fast" },
    { ...patch, permission: { skill: { unrelated: "deny" } } },
  ]) {
    const f = fixture({ patch: invalid })
    const response = await f.save()
    expect(response.status).toBe(400)
    expect(await response.text()).not.toContain("synthetic-secret")
    expect(f.order).toEqual(["prepare"])
    expect(f.written).toEqual([])
  }
})

test("failed credential Save restores any partial activation without a settings write", async () => {
  for (const failure of [{ fail: "save" }, { throwOn: "save" }]) {
    const f = fixture(failure)
    const response = await f.save()
    expect(response.status).toBe(503)
    expect(await response.json()).toMatchObject({ code: "connection_save_failed", partialState: false })
    expect(f.order).toEqual(["prepare", "save", "rollback"])
    expect(f.calls.at(-1)?.input).toEqual({ sessionID: selection.sessionID })
    expect(f.written).toEqual([])
  }
})

test("settings write failure rolls credentials back and suppresses private errors", async () => {
  const f = fixture({ writeFails: true })
  const response = await f.save()
  expect(response.status).toBe(503)
  const body = await response.json()
  expect(body).toMatchObject({ code: "settings_save_failed", partialState: false })
  expect(JSON.stringify(body)).not.toContain("synthetic-secret")
  expect(f.order).toEqual(["prepare", "save", "write", "rollback"])
})

test("rollback failure explicitly reports uncertain partial state", async () => {
  for (const failure of [{ writeFails: true }, { throwOn: "save" }]) {
    const f = fixture({ ...failure, rollbackFails: true })
    const response = await f.save()
    expect(response.status).toBe(503)
    expect(await response.json()).toMatchObject({ code: "setup_partial_state", partialState: true })
    expect(f.order.at(-1)).toBe("rollback")
  }
})

test("completed Save stays successful if candidate cleanup is unavailable", async () => {
  const f = fixture({ fail: "session" })
  const response = await f.save()
  expect(response.status).toBe(200)
  expect(await response.json()).toEqual({ ok: true, sessionClosePending: true })
  expect(f.written).toEqual([patch])
  expect(f.order).not.toContain("rollback")
})

test("origin, malformed fields, browser secrets and oversized inputs fail before any engine call", async () => {
  const f = fixture()
  expect((await f.save(selection, { Origin: "https://foreign.invalid" })).status).toBe(403)
  expect((await f.save(selection, { "Content-Type": "text/plain" })).status).toBe(415)
  for (const input of [{ ...selection, apiKey: "synthetic-secret" }, { ...selection, primary: "" }, { ...selection, fast: "has spaces" }, { ...selection, sessionID: "../other" }, { ...selection, primary: "x".repeat(3000) }, []]) {
    expect((await f.save(input)).status).toBe(400)
  }
  expect(f.calls).toEqual([])
  expect(f.written).toEqual([])
})

test("concurrent Save requests cannot interleave credential and settings transitions", async () => {
  let release!: () => void
  const waiting = new Promise<void>(resolve => { release = resolve })
  let entered!: () => void
  const prepared = new Promise<void>(resolve => { entered = resolve })
  const f = fixture()
  const app = providerSetupRoutes({ transport: async (url, init) => {
    if (url.endsWith("/prepare")) { entered(); await waiting }
    return f.transport(url, init)
  }, writeSettings() {} })
  const save = () => app.request("/engine/provider-setup", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(selection) })
  const first = save()
  await prepared
  expect((await save()).status).toBe(409)
  release()
  expect((await first).status).toBe(200)
  expect(f.order).toEqual(["prepare", "save", "session"])
})

test("normal atomic writer merges settings changed during credential Save", async () => {
  const previous = { DOCHAUS_DIR: process.env.DOCHAUS_DIR, DOCHAUS_RUNTIME_DIR: process.env.DOCHAUS_RUNTIME_DIR }
  const directory = mkdtempSync(path.join(tmpdir(), "dochaus-provider-setup-"))
  try {
    process.env.DOCHAUS_DIR = directory
    process.env.DOCHAUS_RUNTIME_DIR = path.join(directory, "runtime")
    const shipped = path.join(directory, "opencode.json")
    writeFileSync(shipped, JSON.stringify({ agent: { draft: { prompt: "synthetic policy" } } }))
    patchEngineConfig({ provider: { unrelated: { name: "Keep provider" } } })
    const f = fixture()
    const app = providerSetupRoutes({ transport: async (url, init) => {
      const response = await f.transport(url, init)
      if (url.endsWith("/save")) patchEngineConfig({ agent: { unrelated: { disable: true } } })
      return response
    } })
    const response = await app.request("/engine/provider-setup", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(selection) })
    expect(response.status).toBe(200)
    expect(readEngineConfig()).toMatchObject({ model: "synthetic/primary", small_model: "synthetic/fast", agent: { draft: { prompt: "synthetic policy" }, unrelated: { disable: true } }, provider: { unrelated: { name: "Keep provider" } } })
    expect(JSON.parse(readFileSync(engineConfigFile(), "utf8")).agent.unrelated.disable).toBe(true)
    expect(JSON.parse(readFileSync(shipped, "utf8"))).toEqual({ agent: { draft: { prompt: "synthetic policy" } } })
  } finally {
    for (const key of ["DOCHAUS_DIR", "DOCHAUS_RUNTIME_DIR"] as const) if (previous[key] === undefined) delete process.env[key]; else process.env[key] = previous[key]
    rmSync(directory, { recursive: true, force: true })
  }
})
