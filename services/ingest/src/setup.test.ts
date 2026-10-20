import { expect, test } from "bun:test"
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { setupRoutes, validModelPair } from "./setup"
const models = [{ providerID: "cloud", id: "primary", enabled: true }, { providerID: "cloud", id: "fast", enabled: true }, { providerID: "other", id: "model", enabled: true }]
const pair = { model: "cloud/primary", small_model: "cloud/fast" }
function fixture(config: { model?: string; small_model?: string } = {}) {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-setup-"))
  return { directory, app: setupRoutes({ directory, config: () => config, models: async () => models }), dispose: () => rmSync(directory, { recursive: true, force: true }) }
}
function patch(app: ReturnType<typeof setupRoutes>, body: unknown) { return app.request("/setup", { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }) }
test("model readiness requires an enabled saved same-provider pair", () => {
  expect(validModelPair(pair, models)).toBe(true)
  expect(validModelPair({ ...pair, small_model: "other/model" }, models)).toBe(false)
  expect(validModelPair({ ...pair, small_model: "cloud/missing" }, models)).toBe(false)
  expect(validModelPair(pair, models.map(model => ({ ...model, enabled: false })))).toBe(false)
  expect(validModelPair({ model: pair.model }, models)).toBe(false)
})
test("model readiness preserves valid native variants and rejects unknown or malformed variants", () => {
  const catalog = models.map(model => ({ ...model, variants: [{ id: "reasoning" }] }))
  const configured = { model: "cloud/primary#reasoning", small_model: "cloud/fast#reasoning" }
  expect(validModelPair(configured, catalog)).toBe(true)
  expect(configured).toEqual({ model: "cloud/primary#reasoning", small_model: "cloud/fast#reasoning" })
  for (const small_model of ["cloud/fast#unknown", "cloud/fast#", "cloud/fast#reasoning#extra", "other/model#reasoning"])
    expect(validModelPair({ ...configured, small_model }, catalog)).toBe(false)
  expect(validModelPair(configured, models)).toBe(false)
})
test("setup checks the native HTTP catalog with Basic auth only when configured", async () => {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-setup-http-"))
  const previousURL = process.env.OPENCODE_URL
  const previousPassword = process.env.OPENCODE_SERVER_PASSWORD
  const password = "synthetic-setup-fixture"
  let requireAuth = true
  let catalogRequests = 0
  let anonymousRequests = 0
  const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch(request) {
    if (new URL(request.url).pathname !== "/api/model" || request.method !== "GET") return new Response(null, { status: 404 })
    catalogRequests++
    const authorization = request.headers.get("Authorization")
    if (requireAuth && authorization !== `Basic ${btoa(`opencode:${password}`)}`) return new Response(null, { status: 401 })
    if (!requireAuth && authorization !== null) return new Response(null, { status: 400 })
    if (!authorization) anonymousRequests++
    return Response.json({ data: models })
  } })
  const app = setupRoutes({ directory, config: () => pair })
  const ingest = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: app.fetch })
  try {
    process.env.OPENCODE_URL = engine.url.origin
    delete process.env.OPENCODE_SERVER_PASSWORD
    expect((await fetch(new URL("/setup", ingest.url))).status).toBe(503)
    process.env.OPENCODE_SERVER_PASSWORD = password
    const authenticated = await fetch(new URL("/setup", ingest.url))
    expect(authenticated.status).toBe(200)
    expect(await authenticated.json()).toMatchObject({ ready: true, next: "done" })
    requireAuth = false
    delete process.env.OPENCODE_SERVER_PASSWORD
    expect((await fetch(new URL("/setup", ingest.url))).status).toBe(200)
    expect(catalogRequests).toBe(3)
    expect(anonymousRequests).toBe(1)
  } finally {
    if (previousURL === undefined) delete process.env.OPENCODE_URL
    else process.env.OPENCODE_URL = previousURL
    if (previousPassword === undefined) delete process.env.OPENCODE_SERVER_PASSWORD
    else process.env.OPENCODE_SERVER_PASSWORD = previousPassword
    ingest.stop(true)
    engine.stop(true)
    rmSync(directory, { recursive: true, force: true })
  }
})
test("fresh setup cannot skip required models or advance without a saved pair", async () => {
  const setup = fixture()
  try {
    expect(await (await setup.app.request("/setup")).json()).toMatchObject({ next: "models", ready: false })
    expect((await patch(setup.app, { step: "models", status: "skipped" })).status).toBe(400)
    expect((await patch(setup.app, { step: "models", status: "completed" })).status).toBe(400)
    expect((await patch(setup.app, { step: "research", status: "skipped" })).status).toBe(400)
  } finally { setup.dispose() }
})
test("existing configured installations migrate optional steps to skipped", async () => {
  const setup = fixture(pair)
  try { expect(await (await setup.app.request("/setup")).json()).toMatchObject({ completed: ["models"], skipped: ["research", "details"], next: "done", ready: true }) }
  finally { setup.dispose() }
})
test("new setup resumes optional steps and rechecks required models after config changes", async () => {
  const config: { model?: string; small_model?: string } = {}
  const setup = fixture(config)
  try {
    await setup.app.request("/setup")
    Object.assign(config, pair)
    expect(await (await patch(setup.app, { step: "models", status: "completed" })).json()).toMatchObject({ next: "research", skipped: [] })
    expect(await (await patch(setup.app, { step: "research", status: "skipped" })).json()).toMatchObject({ next: "details", skipped: ["research"] })
    expect(await (await patch(setup.app, { step: "details", status: "completed" })).json()).toMatchObject({ next: "done", skipped: ["research"] })
    const saved = JSON.parse(readFileSync(path.join(setup.directory, "setup.json"), "utf8"))
    expect(Object.keys(saved).sort()).toEqual(["completed", "skipped", "version"])
    delete config.small_model
    expect(await (await setup.app.request("/setup")).json()).toMatchObject({ next: "models", ready: false })
  } finally { setup.dispose() }
})
test("setup rejects credential fields", async () => {
  const setup = fixture(pair)
  try { expect((await patch(setup.app, { step: "research", status: "completed", apiKey: "must-not-persist" })).status).toBe(400) }
  finally { setup.dispose() }
})
async function withEngine(respond: () => Response, run: (app: ReturnType<typeof setupRoutes>, file: string) => Promise<void>) {
  const directory = mkdtempSync(path.join(os.tmpdir(), "dochaus-setup-consent-"))
  const previousURL = process.env.OPENCODE_URL
  const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: respond })
  try {
    process.env.OPENCODE_URL = engine.url.origin
    await run(setupRoutes({ directory, config: () => pair }), path.join(directory, "setup.json"))
  } finally {
    if (previousURL === undefined) delete process.env.OPENCODE_URL
    else process.env.OPENCODE_URL = previousURL
    engine.stop(true)
    rmSync(directory, { recursive: true, force: true })
  }
}
test("a step saved while research consent reloads waits for it and reads the catalog once", async () => {
  let requests = 0
  // Saving preferences reloads research consent; the engine refuses inventory reads until it settles.
  await withEngine(() => ++requests === 3 ? Response.json({ data: models }) : Response.json({ ready: false, legalState: "active", consentState: "pending" }, { status: 503 }), async (app, file) => {
    writeFileSync(file, JSON.stringify({ version: 1, completed: ["models"], skipped: [] }))
    const response = await patch(app, { step: "research", status: "completed" })
    expect(response.status).toBe(200)
    expect(await response.json()).toEqual({ version: 1, completed: ["models", "research"], skipped: [], next: "details", ready: true })
    expect(requests).toBe(3)
    expect(JSON.parse(readFileSync(file, "utf8"))).toEqual({ version: 1, completed: ["models", "research"], skipped: [] })
  })
})
test("failed research consent protection is named and never saves first-run progress as not ready", async () => {
  let requests = 0
  await withEngine(() => { requests++; return Response.json({ ready: false, legalState: "active", consentState: "failed" }, { status: 503 }) }, async (app, file) => {
    expect((await app.request("/setup")).status).toBe(503)
    const response = await patch(app, { step: "research", status: "completed" })
    expect(response.status).toBe(400)
    expect((await response.json()).error).toContain("Research protection")
    expect(requests).toBe(2)
    expect(existsSync(file)).toBe(false)
  })
})
