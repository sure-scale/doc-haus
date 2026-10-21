import { expect, test } from "bun:test"
import { cp, mkdtemp, mkdir, readFile, readdir, realpath, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

test("cold native matter generation uses the selected model/variant and refuses invalid requests and fallback", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await realpath(await mkdtemp(path.join(os.tmpdir(), "dochaus-matter-generation-")))
  const repository = path.join(temporary, "repo")
  const workspace = path.join(temporary, "workspace")
  const matter = path.join(workspace, "matter")
  const outside = path.join(temporary, "outside")
  const escaped = path.join(workspace, "escaped")
  await mkdir(path.join(repository, "services", "engine", "src"), { recursive: true })
  await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
  await mkdir(matter, { recursive: true })
  await mkdir(outside)
  await symlink(outside, escaped)
  await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
  await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
  await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
  for (const entry of await readdir(path.join(original, "dochaus"))) {
    if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
    await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
  }
  await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
  await symlink(path.join(original, "dochaus", "plugin", "research-mcp.ts"), path.join(repository, "dochaus", "plugin", "research-mcp.ts"))
  const requests: { model: string; stream?: boolean; messages: { role: string; content: unknown }[];
    temperature?: number; max_tokens?: number; max_completion_tokens?: number; headers: { variant: string | null; session: string | null } }[] = []
  const fixture = { fail: false, healthStatus: 0 }
  const provider = Bun.serve({ port: 0, async fetch(request) {
    if (new URL(request.url).pathname !== "/v1/chat/completions") return Response.json({ error: { message: "Unexpected fixture route" } }, { status: 404 })
    const body = await request.json()
    requests.push({ ...body, headers: { variant: request.headers.get("x-fixture-variant"), session: request.headers.get("x-opencode-session") } })
    if (fixture.healthStatus) return Response.json({ error: { message: fixture.healthStatus === 401
      ? "Synthetic API key has expired: synthetic-private-value" : "Account credits expired; authentication allowance exceeded" } }, { status: fixture.healthStatus })
    if (fixture.fail) return Response.json({ error: { message: "Synthetic model refusal", type: "invalid_request_error", code: "model_unavailable" } }, { status: 400 })
    // Generate.text returns one string, but native LLM.generate collects a
    // provider stream. Support both protocol modes without changing the engine.
    if (body.stream) return new Response(`data: ${JSON.stringify({ id: "chatcmpl_generation_fixture", object: "chat.completion.chunk", created: 1, model: body.model,
      choices: [{ index: 0, delta: { role: "assistant", content: "The independent working-note check completed." }, finish_reason: "stop" }] })}\n\ndata: ${JSON.stringify({ id: "chatcmpl_generation_fixture", object: "chat.completion.chunk", created: 1, model: body.model,
      choices: [], usage: { prompt_tokens: 12, completion_tokens: 8, total_tokens: 20 } })}\n\ndata: [DONE]\n\n`, { headers: { "content-type": "text/event-stream" } })
    return Response.json({ id: "chatcmpl_generation_fixture", object: "chat.completion", created: 1, model: body.model,
      choices: [{ index: 0, message: { role: "assistant", content: "The independent working-note check completed." }, finish_reason: "stop" }],
      usage: { prompt_tokens: 12, completion_tokens: 8, total_tokens: 20 } })
  } })
  const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
  config.mcp = {}; config.provider = {}
  config.model = config.small_model = "generation-fixture/configured-note#audit"
  config.providers = { "generation-fixture": { package: "@opencode/ai/providers/openai-compatible", name: "Independent generation fixture",
    settings: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only", provider: "generation-fixture" }, models: {
      "configured-note": { name: "Configured note", limit: { context: 131072, output: 8192 },
        variants: [{ id: "audit", headers: { "x-fixture-variant": "audit" }, body: { temperature: 0.25 } }] },
      "alternate-note": { name: "Never selected", limit: { context: 131072, output: 8192 } },
    } } }
  config.agent = Object.fromEntries(Object.entries(config.agent).map(([id, value]) => [id, { ...(value as object), model: config.model }]))
  await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port; probe.stop(true)
  const authorization = `Basic ${btoa("opencode:synthetic-engine-password")}`
  const base = `http://127.0.0.1:${port}`
  const log = path.join(temporary, "engine.log")
  const processes: ReturnType<typeof Bun.spawn>[] = []
  function start(runtime: string) {
    const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository,
      env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_RUNTIME_DIR: path.join(temporary, runtime), DOC_HAUS_ENGINE_PORT: String(port),
        OPENCODE_URL: base, OPENCODE_HOSTNAME: "127.0.0.1", OPENCODE_SERVER_PASSWORD: "synthetic-engine-password", OPENCODE_DISABLE_MODELS_FETCH: "1" },
      stdout: Bun.file(log), stderr: Bun.file(log) })
    processes.push(engine); return engine
  }
  async function stop(engine: ReturnType<typeof Bun.spawn>) {
    if (engine.exitCode !== null) return
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { if (engine.exitCode === null) engine.kill("SIGKILL"); return engine.exited })])
  }
  async function request(body: unknown, directory: string | null = matter, authorized = true) {
    return fetch(base + "/api/dochaus/generate", { method: "POST", headers: { "content-type": "application/json",
      ...(directory ? { "x-opencode-directory": directory } : {}), ...(authorized ? { authorization } : {}) }, body: JSON.stringify(body) })
  }
  async function ready(engine: ReturnType<typeof Bun.spawn>) {
    const deadline = Date.now() + 35_000
    while (true) {
      // Global listener health must not prewarm this matter's plugin graph.
      const response = await fetch(base + "/api/info", { headers: { authorization } }).catch(() => undefined)
      if (response?.ok) return
      if (Date.now() > deadline || engine.exitCode !== null) throw new Error(`Engine listener did not start: ${(await readFile(log, "utf8")).slice(-3000)}`)
      await Bun.sleep(100)
    }
  }
  try {
    const engine = start("runtime")
    await ready(engine)
    const response = await request({ prompt: "Summarize this independent internal working note in one sentence." })
    const generated = await response.json()
    if (response.status !== 200) {
      const catalog = await (await fetch(base + "/api/model", { headers: { authorization, "x-opencode-directory": matter } })).json()
      const configuration = await (await fetch(base + "/api/config", { headers: { authorization, "x-opencode-directory": matter } })).json()
      throw new Error(`Native generation failed: ${JSON.stringify({ response: generated, requests: requests.length,
        catalog: catalog.data?.filter((entry: { providerID: string }) => entry.providerID === "generation-fixture")
          .map((entry: { id: string; package?: string; variants?: { id: string }[] }) => ({ id: entry.id, package: entry.package, variants: entry.variants?.map(v => v.id) })),
        selected: configuration.filter((entry: { type: string }) => entry.type === "document").map((entry: { info: { model?: unknown } }) => entry.info.model),
      })}`)
    }
    expect(generated.data.text).toBe("The independent working-note check completed.")
    expect(generated.data.model).toEqual({ providerID: "generation-fixture", id: "configured-note", variant: "audit" })
    expect(generated.data.elapsed_ms).toBeGreaterThanOrEqual(0)
    expect(generated.data.usage).toBeNull()
    expect(requests).toHaveLength(1)
    expect(requests[0].model).toBe("configured-note")
    expect(requests[0].stream).toBe(true)
    expect(requests[0].headers.variant).toBe("audit")
    expect(requests[0].headers.session).toBeTruthy()
    expect(requests[0].temperature).toBe(0.25)
    expect(JSON.stringify(requests[0].messages)).toContain("independent internal working note")
    const plugins = await (await fetch(base + "/api/plugin", { headers: { authorization, "x-opencode-directory": matter } })).json()
    expect(plugins.data.some((entry: { id: string; state: { status: string } }) => entry.id === "dochaus.legal" && entry.state.status === "active")).toBe(true)
    const denied = [
      { name: "unauthorized", body: { prompt: "Denied auth" }, directory: matter, authorized: false, status: 401 },
      { name: "missing matter", body: { prompt: "Missing directory" }, directory: null, status: 400 },
      { name: "workspace root", body: { prompt: "Root directory" }, directory: workspace, status: 403 },
      { name: "outside workspace", body: { prompt: "Outside directory" }, directory: outside, status: 403 },
      { name: "symlink escape", body: { prompt: "Escaped directory" }, directory: escaped, status: 403 },
      { name: "blank prompt", body: { prompt: " " }, status: 400 },
      { name: "oversized prompt", body: { prompt: "x".repeat(1_048_577) }, status: 400 },
      { name: "invalid prompt type", body: { prompt: 7 }, status: 400 },
      { name: "caller model override", body: { prompt: "Do not substitute the configured model", model: { providerID: "generation-fixture", id: "alternate-note" } }, status: 400 },
    ]
    const refusals = [] as { name: string; status: number; expected: number }[]
    for (const deniedRequest of denied) {
      const refused = await request(deniedRequest.body, deniedRequest.directory === undefined ? matter : deniedRequest.directory, deniedRequest.authorized ?? true)
      refusals.push({ name: deniedRequest.name, status: refused.status, expected: deniedRequest.status })
    }
    const malformed = await fetch(base + "/api/dochaus/generate", { method: "POST", headers: { authorization,
      "content-type": "application/json", "x-opencode-directory": matter }, body: "{" })
    refusals.push({ name: "malformed JSON", status: malformed.status, expected: 400 })
    expect(refusals.filter(refusal => refusal.status !== refusal.expected), JSON.stringify(refusals)).toEqual([])
    expect(requests).toHaveLength(1)
    const testConnection = (body: object = { providerID: "generation-fixture", modelID: "configured-note" }, directory = matter, authorized = true) =>
      fetch(base + "/api/dochaus/provider-test", { method: "POST", headers: { "content-type": "application/json", "x-opencode-directory": directory,
        ...(authorized ? { authorization } : {}) }, body: JSON.stringify(body) })
    expect((await testConnection(undefined, matter, false)).status).toBe(401)
    expect((await testConnection(undefined, outside)).status).toBe(403)
    expect((await testConnection({ providerID: "generation-fixture", modelID: "configured-note", prompt: "No custom prompt" })).status).toBe(400)
    expect(requests).toHaveLength(1)
    const healthy = await Promise.all([testConnection(), testConnection()])
    expect(await healthy[0].json()).toEqual({ ok: true })
    expect(await healthy[1].json()).toEqual({ ok: true })
    expect(requests).toHaveLength(2)
    expect(requests[1].max_completion_tokens ?? requests[1].max_tokens).toBeUndefined()
    expect(requests[1].headers.session).toMatch(/^ses_/)
    expect(JSON.stringify(requests[1].messages)).toContain("Reply with OK.")
    expect(JSON.stringify(requests[1].messages)).not.toContain("working note")
    fixture.healthStatus = 401
    expect(await (await testConnection()).json()).toEqual({ ok: true })
    expect(requests).toHaveLength(2)
    const expired = await testConnection({ providerID: "generation-fixture", modelID: "configured-note", force: true })
    expect(expired.status).toBe(503)
    const expiry = await expired.json()
    expect(expiry).toMatchObject({ ok: false, kind: "auth", expired: true })
    expect(JSON.stringify(expiry)).not.toContain("synthetic-private-value")
    fixture.healthStatus = 429
    expect(await (await testConnection({ providerID: "generation-fixture", modelID: "configured-note", force: true })).json()).toMatchObject({ kind: "unavailable" })
    const beforeUnknown = requests.length
    expect(await (await testConnection({ providerID: "generation-fixture", modelID: "unavailable-note", force: true })).json()).toMatchObject({ kind: "model" })
    expect(requests).toHaveLength(beforeUnknown)
    fixture.healthStatus = 0
    fixture.fail = true
    expect((await request({ prompt: "Exercise a configured provider refusal." })).status).toBe(503)
    expect(requests.every(request => request.model === "configured-note")).toBe(true)
    const count = requests.length
    await stop(engine)
    config.model = "generation-fixture/not-configured"
    await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
    const unavailable = start("runtime-unavailable")
    await ready(unavailable)
    expect((await request({ prompt: "An unavailable configured model must refuse." })).status).toBe(503)
    expect(requests).toHaveLength(count)
  } finally {
    for (const engine of processes) await stop(engine)
    provider.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)
