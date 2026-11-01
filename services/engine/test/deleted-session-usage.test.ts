import { Database } from "bun:sqlite"
import { expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { cp, mkdtemp, mkdir, readFile, readdir, realpath, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

test("deleting a matter session keeps its tree's usage in the matter ledger", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await realpath(await mkdtemp(path.join(os.tmpdir(), "dochaus-deleted-usage-")))
  const repository = path.join(temporary, "repo")
  const matter = path.join(temporary, "workspace", "matter")
  await mkdir(path.join(repository, "services", "engine", "src"), { recursive: true })
  await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
  await mkdir(matter, { recursive: true })
  await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
  await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
  await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
  await cp(path.join(original, "release.json"), path.join(repository, "release.json"))
  for (const entry of await readdir(path.join(original, "dochaus"))) {
    if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
    await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
  }
  await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
  await symlink(path.join(original, "dochaus", "plugin", "research-mcp.ts"), path.join(repository, "dochaus", "plugin", "research-mcp.ts"))
  // Scripted provider: one priced text reply per request, no tools.
  const provider = Bun.serve({ port: 0, async fetch(request) {
    const body = await request.json()
    const chunk = (value: object) => `data: ${JSON.stringify({ id: "chatcmpl_usage", object: "chat.completion.chunk", created: 1, model: body.model, ...value })}\n\n`
    return new Response(chunk({ choices: [{ index: 0, delta: { role: "assistant", content: "The usage check completed." }, finish_reason: "stop" }] }) +
      chunk({ choices: [], usage: { prompt_tokens: 40, completion_tokens: 6, total_tokens: 46 } }) + "data: [DONE]\n\n", { headers: { "content-type": "text/event-stream" } })
  } })
  const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
  config.mcp = {}; config.provider = {}
  config.model = config.small_model = "usage-fixture/priced-note"
  config.providers = { "usage-fixture": { package: "@opencode/ai/providers/openai-compatible", name: "Usage fixture",
    settings: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only", provider: "usage-fixture" },
    models: { "priced-note": { name: "Priced note", limit: { context: 131072, output: 8192 }, cost: { input: 2, output: 10 } } } } }
  config.agent = Object.fromEntries(Object.entries(config.agent).map(([id, value]) => [id, { ...(value as object), model: config.model }]))
  await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
  const probe = Bun.serve({ port: 0, fetch: () => new Response() }); const port = probe.port; probe.stop(true)
  const logPath = path.join(temporary, "engine.log")
  const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository,
    env: { ...process.env, WORKSPACE_ROOT: path.dirname(matter), DOCHAUS_RUNTIME_DIR: path.join(temporary, "runtime"), DOC_HAUS_ENGINE_PORT: String(port),
      OPENCODE_URL: `http://127.0.0.1:${port}`, OPENCODE_HOSTNAME: "127.0.0.1", OPENCODE_SERVER_PASSWORD: "synthetic-engine-password", OPENCODE_DISABLE_MODELS_FETCH: "1" },
    stdout: Bun.file(logPath), stderr: Bun.file(logPath) })
  const base = `http://127.0.0.1:${port}`
  const authorization = `Basic ${btoa("opencode:synthetic-engine-password")}`
  const request = (route: string, init: { method?: string; body?: object; directory?: string | null } = {}) =>
    fetch(base + route, { method: init.method ?? (init.body ? "POST" : "GET"), headers: { authorization, "content-type": "application/json",
      ...(init.directory === null ? {} : { "x-opencode-directory": init.directory ?? matter }) }, body: init.body ? JSON.stringify(init.body) : undefined })
  const model = { providerID: "usage-fixture", id: "priced-note" }
  async function settle(id: string) {
    const deadline = Date.now() + 30_000
    while (true) {
      const messages = (await (await request(`/api/session/${id}/message`)).json()).data
      const active = (await (await request("/api/session/active")).json()).data
      if (messages.some((message: { type: string; finish?: string }) => message.type === "assistant" && message.finish === "stop") && !active[id]) return
      if (Date.now() > deadline) throw new Error(`Turn did not settle: ${JSON.stringify(messages)}\n${await readFile(logPath, "utf8")}`)
      await Bun.sleep(100)
    }
  }
  try {
    const readyDeadline = Date.now() + 35_000
    while (true) {
      const response = await request("/api/plugin").catch(() => undefined)
      if (response?.ok && (await response.json()).data.some((entry: { id: string; state: { status: string } }) => entry.id === "dochaus.legal" && entry.state.status === "active")) break
      if (Date.now() > readyDeadline || engine.exitCode !== null) throw new Error(`Legal plugin did not activate: ${await readFile(logPath, "utf8")}`)
      await Bun.sleep(100)
    }
    expect((await request("/api/session", { body: { id: "ses_usage_parent", agent: "qa", model, location: { directory: matter } } })).status).toBe(200)
    expect((await request("/api/session/ses_usage_parent/prompt", { body: { id: "msg_usage_parent", text: "Report the usage check." } })).status).toBe(200)
    await settle("ses_usage_parent")
    expect((await request("/api/dochaus/session", { body: { parentID: "ses_usage_parent", id: "ses_usage_child", agent: "qa", model, metadata: { kind: "suggestions" } } })).status).toBe(200)
    expect((await request("/api/session/ses_usage_child/prompt", { body: { id: "msg_usage_child", text: "Report the child usage check." } })).status).toBe(200)
    await settle("ses_usage_child")
    const cost = async (id: string) => (await (await request(`/api/session/${id}`)).json()).data.cost as number
    const parentCost = await cost("ses_usage_parent"), childCost = await cost("ses_usage_child")
    expect(parentCost).toBeGreaterThan(0)
    expect(childCost).toBeGreaterThan(0)

    // An unauthenticated delete is refused and keeps nothing.
    expect((await fetch(base + "/api/session/ses_usage_parent", { method: "DELETE", headers: { "x-opencode-directory": matter } })).status).toBe(401)
    const kept = () => {
      if (!existsSync(path.join(matter, ".dochaus", "telemetry.db"))) return 0
      const ledger = new Database(path.join(matter, ".dochaus", "telemetry.db"), { readonly: true })
      try { return (ledger.query("SELECT COUNT(*) AS rows FROM generations").get() as { rows: number }).rows } finally { ledger.close() }
    }
    expect(kept()).toBe(0)

    expect((await request("/api/session/ses_usage_parent", { method: "DELETE" })).status).toBe(204)
    expect((await request("/api/session/ses_usage_parent")).status).toBe(404)
    // The core delete cascades to children; their usage is kept too.
    expect((await request("/api/session/ses_usage_child")).status).toBe(404)
    const db = new Database(path.join(matter, ".dochaus", "telemetry.db"), { readonly: true })
    try {
      const rows = db.query("SELECT session_id, agent, model, input_tokens, output_tokens, cost_usd FROM generations ORDER BY ts").all() as
        { session_id: string; agent: string; model: string | null; input_tokens: number; output_tokens: number; cost_usd: number }[]
      const total = (id: string) => rows.filter(row => row.session_id === id).reduce((sum, row) => sum + row.cost_usd, 0)
      expect(total("ses_usage_parent")).toBeCloseTo(parentCost, 12)
      expect(total("ses_usage_child")).toBeCloseTo(childCost, 12)
      expect(rows).toContainEqual({ session_id: "ses_usage_parent", agent: "qa", model: "usage-fixture/priced-note", input_tokens: 40, output_tokens: 6, cost_usd: expect.closeTo((40 * 2 + 6 * 10) / 1_000_000, 12) })
      expect(rows).toContainEqual(expect.objectContaining({ session_id: "ses_usage_child", agent: "word-suggestions", model: "usage-fixture/priced-note" }))
    } finally { db.close() }
  } finally {
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    provider.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)
