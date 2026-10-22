import { expect, test } from "bun:test"
import { cp, mkdir, mkdtemp, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import { existsSync } from "node:fs"
import os from "node:os"
import path from "node:path"

async function fixture(research: "active" | "missing" | "failed") {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-managed-readiness-"))
  const repository = path.join(temporary, "repo")
  const workspace = path.join(temporary, "workspace")
  const matter = path.join(workspace, "matter")
  const runtime = path.join(temporary, "runtime")
  const override = path.join(runtime, "config", "opencode", "opencode.json")
  const calls: { messages: unknown[]; tools?: unknown[] }[] = []
  const provider = Bun.serve({ port: 0, async fetch(request) {
    const body = await request.json()
    calls.push(body)
    const answer = { id: "chatcmpl_provider_test", object: "chat.completion.chunk", created: 1, model: "neutral", choices: [{ index: 0, delta: { role: "assistant", content: "OK" }, finish_reason: "stop" }] }
    return body.stream === false
      ? Response.json({ id: "chatcmpl_provider_test", object: "chat.completion", created: 1, model: "neutral", choices: [{ index: 0, message: { role: "assistant", content: "OK" }, finish_reason: "stop" }] })
      : new Response(`data: ${JSON.stringify(answer)}\n\ndata: [DONE]\n\n`, { headers: { "content-type": "text/event-stream" } })
  } })
  await mkdir(path.join(repository, "services", "engine"), { recursive: true })
  await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
  await mkdir(path.dirname(override), { recursive: true })
  await mkdir(matter, { recursive: true })
  await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
  await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
  await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
  for (const entry of await readdir(path.join(original, "dochaus"))) {
    if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
    await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
  }
  await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
  if (research !== "missing") {
    const source = await readFile(path.join(original, "dochaus", "plugin", "research-mcp.ts"), "utf8")
    await writeFile(path.join(repository, "dochaus", "plugin", "research-mcp.ts"), research === "failed"
      ? source.replace("setup: async (context) => {", 'setup: async (context) => { context.mcp.transform = async () => { throw new Error("synthetic native registration failure") }') : source)
  }
  const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
  config.mcp = {}
  delete config.model
  delete config.small_model
  config.provider = { "neutral-fixture": { npm: "@ai-sdk/openai-compatible", name: "Neutral test provider", options: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only" }, models: { neutral: { name: "Neutral", limit: { context: 131072, output: 8192 } } } } }
  await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const base = `http://127.0.0.1:${port}`
  const authorization = `Basic ${btoa("opencode:synthetic-password")}`
  const log = path.join(temporary, "engine.log")
  const environment = { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_RUNTIME_DIR: runtime,
    DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: base, OPENCODE_HOSTNAME: "127.0.0.1",
    OPENCODE_SERVER_PASSWORD: "synthetic-password", OPENCODE_DISABLE_MODELS_FETCH: "1" }
  delete environment.DOCHAUS_PROVIDER_CAPTURE_DIR
  for (const name of ["GOOGLE_VERTEX_PROJECT", "GOOGLE_CLOUD_PROJECT", "GCP_PROJECT", "GCLOUD_PROJECT", "GOOGLE_VERTEX_API_KEY",
    "AWS_PROFILE", "AWS_ACCESS_KEY_ID", "AWS_WEB_IDENTITY_TOKEN_FILE", "AWS_CONTAINER_CREDENTIALS_RELATIVE_URI", "AWS_CONTAINER_CREDENTIALS_FULL_URI", "AWS_BEARER_TOKEN_BEDROCK"])
    delete environment[name as keyof typeof environment]
  const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository, env: environment, stdout: Bun.file(log), stderr: Bun.file(log) })
  const request = (route: string, body?: object, directory: string | null = matter, authorized = true) => fetch(base + route, {
    method: body ? "POST" : "GET", headers: { "Content-Type": "application/json", ...(directory ? { "x-opencode-directory": directory } : {}), ...(authorized ? { authorization } : {}), Origin: "http://localhost:5173" }, body: body ? JSON.stringify(body) : undefined,
  })
  const close = async () => {
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    provider.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
  const deadline = Date.now() + 30_000
  while (!(await request("/api/info").catch(() => undefined))?.ok) {
    if (Date.now() > deadline || engine.exitCode !== null) { await close(); throw new Error("Isolated engine did not listen") }
    await Bun.sleep(50)
  }
  const readiness = async (directory: string | null = matter) => {
    while (true) {
      const response = await request("/api/dochaus/ready", undefined, directory)
      const body = await response.json()
      if (body.legalState !== "loading" && body.consentState !== "pending") return { response, body }
      if (Date.now() > deadline) throw new Error("Managed consent readiness did not settle")
      await Bun.sleep(50)
    }
  }
  return { request, readiness, close, workspace, matter, runtime, override, calls, log, base, providerBaseURL: `http://127.0.0.1:${provider.port}/v1` }
}

test("empty first-run is ready after actual legal and research activation; generation requires setup while fixed provider test works", async () => {
  const f = await fixture("active")
  try {
    expect((await f.request("/api/dochaus/ready", undefined, f.matter, false)).status).toBe(401)
    expect(existsSync(path.join(f.workspace, ".preferences"))).toBe(false)
    const { response, body } = await f.readiness(null)
    expect(response.status).toBe(200)
    expect(response.headers.get("access-control-allow-origin")).toBe("http://localhost:5173")
    const preflight = await fetch(f.base + "/api/dochaus/provider-test", { method: "OPTIONS", headers: {
      Origin: "http://localhost:5173", "Access-Control-Request-Method": "POST", "Access-Control-Request-Headers": "content-type",
    } })
    expect(preflight.status).toBe(204)
    expect(preflight.headers.get("access-control-allow-credentials")).toBe("true")
    expect(body).toMatchObject({ ready: true, legalState: "active", consentState: "active", legalPluginCount: 1 })
    expect((await f.request("/api/dochaus/provider-catalog", undefined, null, false)).status).toBe(401)
    const catalog = await (await f.request("/api/dochaus/provider-catalog", undefined, null)).json()
    expect(catalog.data.find((provider: { id: string }) => provider.id === "google-vertex")).toMatchObject({ activation: "auto", host: { type: "google-adc" } })
    expect(catalog.data.find((provider: { id: string }) => provider.id === "amazon-bedrock")).toMatchObject({ activation: "auto", host: { type: "aws-chain" } })
    expect(JSON.stringify(catalog)).not.toContain("synthetic-test-only")
    const available = await (await f.request("/api/provider", undefined, null)).json()
    expect(available.data.some((provider: { id: string }) => ["google-vertex", "amazon-bedrock"].includes(provider.id))).toBe(false)
    expect(f.calls).toHaveLength(0)
    expect((await f.request("/api/session", { id: "ses_readiness", location: { directory: f.matter } })).status).toBe(200)
    for (const route of ["/api/session/ses_readiness/prompt", "/api/session/ses_readiness/generate", "/api/experimental/generate", "/api/dochaus/generate"]) {
      const blocked = await f.request(route, { prompt: "Never sent", text: "Never sent", resume: false })
      // dochaus's existing strict payload check runs before admission.
      expect(blocked.status).toBe(route === "/api/dochaus/generate" ? 400 : 503)
    }
    expect((await f.request("/api/dochaus/generate", { prompt: "Never sent" })).status).toBe(503)
    expect((await (await f.request("/api/session/ses_readiness/inbox")).json()).data).toEqual([])
    expect(f.calls).toHaveLength(0)
    const tested = await f.request("/api/dochaus/provider-test", { providerID: "neutral-fixture", modelID: "neutral" })
    expect(tested.status).toBe(200)
    expect(await tested.json()).toEqual({ ok: true })
    expect(f.calls).toHaveLength(1)
    expect(JSON.stringify(f.calls[0].messages)).toContain("Reply with OK.")
    expect(f.calls[0].tools ?? []).toEqual([])
    expect((await f.request("/api/dochaus/provider-test", { providerID: "neutral-fixture", modelID: "neutral", prompt: "Unapproved text" })).status).toBe(400)
    expect((await f.request("/api/dochaus/provider-test", { providerID: "neutral-fixture", modelID: "unknown-model" })).status).toBe(400)
    expect(f.calls).toHaveLength(1)
    await writeFile(f.override, JSON.stringify({ model: "neutral-fixture/neutral", small_model: "neutral-fixture/neutral" }))
    const setup = path.join(path.dirname(f.override), "setup.json")
    await writeFile(setup, JSON.stringify({ version: 1, completed: ["models"], skipped: [] }))
    expect((await f.request("/api/session/ses_readiness/prompt", { text: "Blocked until optional choices", resume: false })).status).toBe(503)
    await writeFile(setup, JSON.stringify({ version: 1, completed: ["models"], skipped: ["details", "research"] }))
    const prompt = { id: "msg_readiness", text: "Original untouched prompt", resume: false, delivery: "queue" }
    const admitted = await f.request("/api/session/ses_readiness/prompt", prompt, path.join(f.workspace, "unavailable-header-directory"))
    expect(admitted.status).toBe(200)
    expect((await admitted.json()).data.payload.text).toBe(prompt.text)
    expect((await (await f.request("/api/session/ses_readiness/inbox")).json()).data).toHaveLength(1)
    await writeFile(f.override, JSON.stringify({ model: "neutral-fixture/neutral", small_model: "neutral-fixture/neutral", disabled_providers: ["neutral-fixture"] }))
    expect((await f.request("/api/dochaus/provider-test", { providerID: "neutral-fixture", modelID: "neutral" })).status).toBe(403)
    expect(f.calls).toHaveLength(1)
    expect(await readFile(f.log, "utf8")).not.toContain("TRACE CAPTURE DISABLED")
  } finally { await f.close() }
}, 60_000)

for (const research of ["missing", "failed"] as const) test(`${research} research activation cannot claim ready or admit generation`, async () => {
  const f = await fixture(research)
  try {
    const { response, body } = await f.readiness()
    expect(response.status).toBe(503)
    expect(body).toMatchObject({ ready: false, legalState: "active", consentState: research })
    expect((await f.request("/api/session", { id: "ses_rejected", location: { directory: f.matter } })).status).toBe(200)
    for (const route of ["/api/session/ses_rejected/prompt", "/api/session/ses_rejected/generate", "/api/experimental/generate", "/api/dochaus/generate", "/api/dochaus/provider-test"]) {
      expect((await f.request(route, { prompt: "Never sent" })).status).toBe(503)
    }
    expect((await (await f.request("/api/session/ses_rejected/inbox")).json()).data).toEqual([])
    expect(f.calls).toHaveLength(0)
  } finally { await f.close() }
}, 60_000)

test("candidate setup uses authenticated managed routes and discards an exact draft without changing saved settings", async () => {
  const f = await fixture("active")
  let sessionID = ""
  try {
    expect((await f.readiness()).response.status).toBe(200)
    const route = "/api/dochaus/provider-candidate/session"
    const input = { providerID: "candidate-fixture", connection: { type: "key", key: "synthetic-candidate-only" }, local: { name: "Candidate fixture", baseURL: f.providerBaseURL, modelIDs: ["candidate-model"] } }
    expect((await f.request(route, input, f.matter, false)).status).toBe(401)
    expect((await f.request(route, input, path.dirname(f.workspace))).status).toBe(403)
    const preflight = await fetch(f.base + route, { method: "OPTIONS", headers: { Origin: "http://localhost:5173", "Access-Control-Request-Method": "POST", "Access-Control-Request-Headers": "content-type" } })
    expect(preflight.status).toBe(204)
    expect(preflight.headers.get("access-control-allow-credentials")).toBe("true")
    const before = existsSync(f.override) ? await readFile(f.override, "utf8") : undefined
    const started = await f.request(route, input)
    expect(started.status).toBe(200)
    sessionID = (await started.json()).sessionID
    const models = await (await f.request("/api/dochaus/provider-candidate/models", { sessionID })).json()
    expect(models.data.map((model: { id: string }) => model.id)).toEqual(["candidate-model"])
    const tested = await f.request("/api/dochaus/provider-candidate/test", { sessionID, modelID: "candidate-model" })
    expect(tested.status).toBe(200)
    expect(await tested.json()).toEqual({ ok: true })
    const prepared = await (await f.request("/api/dochaus/provider-candidate/prepare", { sessionID, primary: "candidate-model", fast: "candidate-model" })).json()
    expect(prepared.patch.model).toBe("candidate-fixture/candidate-model")
    expect(JSON.stringify(prepared)).not.toContain("synthetic-candidate-only")
    expect(existsSync(f.override) ? await readFile(f.override, "utf8") : undefined).toBe(before)
    expect(f.calls).toHaveLength(1)
    expect(JSON.stringify(f.calls[0].messages)).toContain("Reply with OK.")
    expect(await readFile(f.log, "utf8")).not.toContain("synthetic-candidate-only")
  } finally {
    if (sessionID) await fetch(f.base + "/api/dochaus/provider-candidate/session?sessionID=" + sessionID, { method: "DELETE", headers: { authorization: `Basic ${btoa("opencode:synthetic-password")}`, "x-opencode-directory": f.matter } })
    await f.close()
  }
}, 60_000)
