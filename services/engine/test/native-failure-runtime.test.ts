import { expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { cp, mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"

// Stage only a neutral plugin into the existing disposable engine fixture. No
// production fault flag, real provider account, or native-state fabrication.
for (const mode of ["http", "restart", "sink-failure"] as const) {
  test(`native outer failure evidence reaches ${mode} without a provider request`, async () => {
    const original = path.resolve(import.meta.dir, "../../..")
    const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-native-failure-"))
    const repository = path.join(temporary, "repo")
    const matter = path.join(temporary, "workspace", "matter")
    const runtime = path.join(temporary, "runtime")
    const evidenceFile = path.join(runtime, "diagnostics", "native-failures.jsonl")
    const faultFile = path.join(matter, "fault-ready")
    const message = "PRIVATE_NATIVE_FAILURE_SENTINEL"
    let providerRequests = 0
    const engines: ReturnType<typeof Bun.spawn>[] = []
    await mkdir(path.join(repository, "services", "engine"), { recursive: true })
    await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
    await mkdir(matter, { recursive: true })
    await mkdir(runtime)
    await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
    await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
    await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
    for (const entry of await readdir(path.join(original, "dochaus"))) {
      if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
      await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
    }
    await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
    await symlink(path.join(original, "dochaus", "plugin", "research-mcp.ts"), path.join(repository, "dochaus", "plugin", "research-mcp.ts"))
    await writeFile(path.join(repository, "dochaus", "plugin", "failure-fixture.ts"), `
import { Plugin } from "@opencode/plugin"
import { existsSync } from "node:fs"
export default Plugin.define({ id: "neutral.failure", setup: async context => {
  await context.session.hook("context", async () => {
    if (!existsSync(${JSON.stringify(faultFile)})) await new Promise<void>(() => {})
    throw new Error(${JSON.stringify(message)})
  })
} })
`)
    const provider = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => {
      providerRequests++
      return new Response("Unexpected neutral provider request", { status: 500 })
    } })
    const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
    config.mcp = {}
    config.plugin = ["./plugin/failure-fixture.ts"]
    config.model = config.small_model = "local-fixture/local-fixture"
    config.provider = { "local-fixture": { npm: "@ai-sdk/openai-compatible", name: "Neutral failure fixture",
      options: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only" },
      models: { "local-fixture": { name: "Local fixture", tool_call: true, limit: { context: 131072, output: 8192 } } } } }
    config.agent = Object.fromEntries(Object.entries(config.agent).map(([id, value]) => [id, { ...(value as object), model: config.model }]))
    config.agent["failure-fixture"] = { mode: "primary", model: config.model, prompt: "Return a neutral response.", tools: { "*": false } }
    await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
    const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
    const port = probe.port
    probe.stop(true)
    const base = `http://127.0.0.1:${port}`
    const headers = { "content-type": "application/json", "x-opencode-directory": matter }
    const sessionID = "ses_native_failure"
    const env = { ...process.env, WORKSPACE_ROOT: path.dirname(matter), DOCHAUS_RUNTIME_DIR: runtime,
      DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: base, OPENCODE_HOSTNAME: "127.0.0.1",
      OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" }
    async function request(route: string, body?: object) {
      const response = await fetch(base + route, { headers, method: body ? "POST" : "GET",
        body: body ? JSON.stringify(body) : undefined, signal: AbortSignal.timeout(5000) })
      expect(response.status).toBe(200)
      return response.json()
    }
    async function until(check: () => Promise<boolean>) {
      const deadline = Date.now() + 30_000
      while (!await check()) {
        if (Date.now() >= deadline) throw new Error("Neutral native failure boundary was not reached")
        await Bun.sleep(50)
      }
    }
    async function start(managed: boolean) {
      const log = path.join(temporary, `engine-${engines.length}.log`)
      const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository,
        env: { ...env, DOCHAUS_MANAGED_RESTART: managed ? "1" : "" }, stdout: Bun.file(log), stderr: Bun.file(log) })
      engines.push(engine)
      await until(async () => {
        if (engine.exitCode !== null) throw new Error("Neutral engine exited before readiness")
        return (await fetch(base + "/api/info").catch(() => undefined))?.ok === true
      })
      expect(Number(await readFile(path.join(runtime, "engine-process.lock"), "utf8"))).toBe(engine.pid)
      await until(async () => (await request("/api/plugin")).data.some((entry: { id: string; state: { status: string } }) =>
        entry.id === "neutral.failure" && entry.state.status === "active"))
      return { engine, log }
    }
    async function stop(engine: ReturnType<typeof Bun.spawn>) {
      if (engine.exitCode !== null) return
      process.kill(engine.pid, "SIGTERM")
      await Promise.race([engine.exited, Bun.sleep(3000).then(() => { process.kill(engine.pid, "SIGKILL"); return engine.exited })])
    }
    try {
      if (mode !== "restart") await writeFile(faultFile, "ready")
      if (mode === "sink-failure") await writeFile(path.join(runtime, "diagnostics"), "intentionally not a directory")
      const initial = await start(false)
      await request("/api/session", { id: sessionID, title: "Neutral failure boundary", agent: "failure-fixture",
        model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })
      await request(`/api/session/${sessionID}/prompt`, { id: "msg_native_failure", text: "Return the neutral response." })
      let failedLog = initial.log
      if (mode === "restart") {
        await until(async () => Boolean((await request("/api/session/active")).data[sessionID]))
        await stop(initial.engine)
        const database = new Database(path.join(runtime, "data", "opencode", "opencode-local.db"), { readonly: true })
        try {
          const checkpoint = database.query<{ time_suspended: number | null; resume_attempts: number }, [string]>(
            "SELECT time_suspended, resume_attempts FROM session_v2 WHERE id = ?").get(sessionID)!
          expect(checkpoint.time_suspended).not.toBeNull()
          expect(checkpoint.resume_attempts).toBe(0)
        } finally { database.close() }
        expect(providerRequests).toBe(0)
        expect(await Bun.file(evidenceFile).exists()).toBe(false)
        await writeFile(faultFile, "ready")
        const restored = await start(true)
        failedLog = restored.log
        expect((await request("/api/session/active")).data).toEqual({})
        process.kill(restored.engine.pid, "SIGUSR2")
      }
      await until(async () => (await request(`/api/session/${sessionID}`)).data.outcome === "failed")
      expect(providerRequests).toBe(0)
      expect((await request(`/api/session/${sessionID}`)).data.outcome).toBe("failed")
      expect((await request("/api/session/active")).data).toEqual({})
      const log = await readFile(failedLog, "utf8")
      expect(log).not.toContain(message)
      if (mode === "sink-failure") {
        expect(log).toContain("DOCHAUS_NATIVE_FAILURE_EVIDENCE_SINK_FAILED")
        expect(await Bun.file(evidenceFile).exists()).toBe(false)
      } else {
        await until(() => Bun.file(evidenceFile).exists())
        const lines = (await readFile(evidenceFile, "utf8")).trim().split("\n")
        expect(lines).toHaveLength(1)
        expect(lines[0]).not.toContain(message)
        expect(lines[0]).not.toContain(temporary)
        const record = JSON.parse(lines[0]!)
        expect(record.event).toBe("native_session_drain_failed")
        expect(record.session_id).toBe(sessionID)
        expect(record.reasons[0].name).toBe("Error")
        expect(record.reasons[0].message_fingerprint.ref).toBe(draftingTraceReference(message))
        // Frames are best effort: this staged caller is not shipped source,
        // and Bun need not retain the constructor's definition as a callsite.
        expect(record.reasons[0].stack_status).toBe("inspected")
        expect(Array.isArray(record.reasons[0].frames)).toBe(true)
        expect(log).toContain(lines[0]!)
      }
    } finally {
      for (const engine of engines) await stop(engine)
      provider.stop(true)
      await rm(temporary, { recursive: true, force: true })
    }
  }, 90_000)
}
