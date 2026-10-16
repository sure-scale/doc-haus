import { expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { cp, mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

test("sealed engine admits durable inputs without execution and scopes child creation", async () => {
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-native-runtime-"))
  const repository = path.resolve(import.meta.dir, "../../..")
  const runtime = path.join(temporary, "runtime")
  const workspace = path.join(temporary, "workspace")
  const matter = path.join(workspace, "matter")
  await mkdir(matter, { recursive: true })
  await mkdir(path.join(workspace, "other"))
  await mkdir(path.join(runtime, "data", "opencode"), { recursive: true })
  // Synthetic credentials exercise the native import without using an account.
  const auth = JSON.stringify({ "runtime-test": { type: "api", key: "synthetic-test-only" } })
  await writeFile(path.join(runtime, "data", "opencode", "auth.json"), auth)
  await mkdir(path.join(runtime, "config", "opencode"), { recursive: true })
  await writeFile(path.join(runtime, "config", "opencode", "opencode.json"), JSON.stringify({ model: "runtime-test/neutral", small_model: "runtime-test/neutral" }))
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const env = { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_RUNTIME_DIR: runtime,
    DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: `http://127.0.0.1:${port}`, OPENCODE_HOSTNAME: "127.0.0.1",
    OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" }
  const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], {
    cwd: repository, env, stdout: "ignore", stderr: "inherit",
  })
  const base = `http://127.0.0.1:${port}`
  async function request(route: string, body?: object, directory = matter) {
    return fetch(base + route, { method: body ? "POST" : "GET",
      headers: { "content-type": "application/json", "x-opencode-directory": directory },
      body: body ? JSON.stringify(body) : undefined })
  }
  try {
    const deadline = Date.now() + 30_000
    while (!(await fetch(base + "/api/info").catch(() => undefined))?.ok) {
      if (Date.now() > deadline || engine.exitCode !== null) throw new Error("Engine did not start")
      await Bun.sleep(100)
    }
    const duplicate = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], {
      cwd: repository, env: { ...env, DOC_HAUS_ENGINE_PORT: "0" }, stdout: "ignore", stderr: "ignore",
    })
    expect(await duplicate.exited).not.toBe(0)
    expect((await request("/api/session", { id: "ses_runtime_test", location: { directory: matter } })).status).toBe(200)
    const prompt = { id: "msg_runtime_test", text: "Admit without executing", delivery: "queue", resume: false }
    const admission = await (await request("/api/session/ses_runtime_test/prompt", prompt)).json()
    expect(admission.data.payload.text).toBe(prompt.text)
    expect(await (await request("/api/session/ses_runtime_test/prompt", prompt)).json()).toEqual(admission)
    // Upstream v2.0.19 returns the existing input on conflicting ID reuse. The
    // product transport must validate the returned identity AND payload itself.
    const conflicting = await (await request("/api/session/ses_runtime_test/prompt", { ...prompt, text: "Different input" })).json()
    expect(conflicting).toEqual(admission)
    expect((await (await request("/api/session/ses_runtime_test/inbox")).json()).data).toHaveLength(1)
    expect((await (await request("/api/session/active")).json()).data).toEqual({})
    const child = await (await request("/api/dochaus/session", { id: "ses_child_test", parentID: "ses_runtime_test" })).json()
    expect(child.data.parentID).toBe("ses_runtime_test")
    expect((await request("/api/dochaus/session", { parentID: "ses_runtime_test" }, path.join(workspace, "other"))).status).toBe(403)
    const entries = await (await request("/api/config")).json()
    expect(entries.map((entry: { path: string }) => entry.path)).toEqual([
      path.join(repository, "dochaus", "opencode.json"), path.join(runtime, "config", "opencode", "opencode.json"), path.join(repository, "dochaus"),
    ])
    // Native SSE heartbeats are 15 seconds apart. Bun's default idle timeout
    // must not close the socket before the next heartbeat or durable event.
    const streamController = new AbortController()
    const stream = await fetch(base + "/api/event", { signal: streamController.signal })
    const reader = stream.body!.getReader()
    const decoder = new TextDecoder()
    let observed = ""
    const observing = (async () => {
      while (true) {
        const chunk = await reader.read()
        if (chunk.done) throw new Error("Native event stream closed while idle")
        observed += decoder.decode(chunk.value, { stream: true })
        if (observed.includes("ses_after_idle")) return
      }
    })()
    // Retain rejection immediately so a premature disconnect does not become an
    // unhandled rejection while the intentional idle interval is in progress.
    const observation = observing.then(() => ({ ok: true as const }), error => ({ ok: false as const, error }))
    try {
      await Bun.sleep(12_000)
      expect((await request("/api/session", { id: "ses_after_idle", location: { directory: matter } })).status).toBe(200)
      const result = await Promise.race([observation, Bun.sleep(5000).then(() => ({ ok: false as const, error: new Error("Native event was not delivered after idle") }))])
      if (!result.ok) throw result.error
      expect(observed).toContain("server.connected")
      expect(observed).toContain("ses_after_idle")
    } finally {
      streamController.abort()
      await reader.cancel().catch(() => undefined)
      await observation
    }
    expect(await readFile(path.join(runtime, "data", "opencode", "auth.json"), "utf8")).toBe(auth)
    const database = new Database(path.join(runtime, "data", "opencode", "opencode-local.db"), { readonly: true })
    expect(database.query("SELECT count(*) AS count FROM credential WHERE integration_id = ?").get("runtime-test")).toEqual({ count: 1 })
    database.close()
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    // A stale pending file after a committed import must not resurrect a
    // credential that the user subsequently removed.
    await rm(path.join(runtime, "engine-process.lock"), { force: true })
    await writeFile(path.join(runtime, "legacy-auth-import.pending"), "pending")
    const editable = new Database(path.join(runtime, "data", "opencode", "opencode-local.db"))
    editable.run("DELETE FROM credential WHERE integration_id = 'runtime-test'")
    editable.close()
    const restarted = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], {
      cwd: repository, env, stdout: "ignore", stderr: "inherit",
    })
    try {
      const restartDeadline = Date.now() + 30_000
      while (!(await fetch(base + "/api/info").catch(() => undefined))?.ok) {
        if (Date.now() > restartDeadline || restarted.exitCode !== null) throw new Error("Engine did not restart")
        await Bun.sleep(100)
      }
      const reread = new Database(path.join(runtime, "data", "opencode", "opencode-local.db"), { readonly: true })
      expect(reread.query("SELECT count(*) AS count FROM credential WHERE integration_id = ?").get("runtime-test")).toEqual({ count: 0 })
      reread.close()
    } finally {
      restarted.kill("SIGTERM")
      await Promise.race([restarted.exited, Bun.sleep(3000).then(() => { restarted.kill("SIGKILL"); return restarted.exited })])
    }
  } finally {
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    await rm(temporary, { recursive: true, force: true })
  }
}, 60_000)

test("managed readiness precedes the same suspended native restart and release is one-shot", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-managed-restart-"))
  const repository = path.join(temporary, "repo")
  const workspace = path.join(temporary, "workspace")
  const matter = path.join(workspace, "matter")
  const runtime = path.join(temporary, "runtime")
  const engines: ReturnType<typeof Bun.spawn>[] = []
  let calls = 0
  let executionStarts = 0
  let eventObservationFailed = false
  let executionError: { type: string | null; status: number | null; message: string | null } | undefined
  const eventsAbort = new AbortController()
  let eventsReading: Promise<void> | undefined
  let eventsReader: ReadableStreamDefaultReader<Uint8Array> | undefined
  function isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === "object" && value !== null && !Array.isArray(value)
  }
  function errorFields(value: unknown) {
    const error = isRecord(value) ? value : undefined
    return { type: typeof error?.type === "string" ? error.type.slice(0, 128) : null,
      status: typeof error?.status === "number" && Number.isInteger(error.status) ? error.status : null,
      message: typeof error?.message === "string" ? error.message.replaceAll("synthetic-test-only", "[redacted]").slice(0, 1000) : null }
  }
  let finishResume = () => {}
  const resumedResponse = new Promise<void>(resolve => { finishResume = resolve })
  const provider = Bun.serve({ hostname: "127.0.0.1", port: 0, idleTimeout: 0, async fetch(request) {
    await request.json()
    calls += 1
    const chunk = (content: string, finish: string | null) => `data: ${JSON.stringify({
      id: "chatcmpl_restart", object: "chat.completion.chunk", created: 1, model: "local-fixture",
      choices: [{ index: 0, delta: { role: "assistant", content }, finish_reason: finish }],
    })}\n\n`
    if (calls === 1) return new Response(new ReadableStream({ start(controller) {
      controller.enqueue(new TextEncoder().encode(chunk("Neutral unfinished response.", null)))
    } }), { headers: { "content-type": "text/event-stream" } })
    await resumedResponse
    return new Response(chunk("Neutral completed response.", "stop") + "data: [DONE]\n\n",
      { headers: { "content-type": "text/event-stream" } })
  } })
  try {
    await mkdir(path.join(repository, "services", "engine"), { recursive: true })
    await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
    await mkdir(matter, { recursive: true })
    await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
    await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
    await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
    await cp(path.join(original, "services", "engine", "script"), path.join(repository, "services", "engine", "script"), { recursive: true })
    for (const entry of await readdir(path.join(original, "dochaus"))) {
      if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
      await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
    }
    await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
  await symlink(path.join(original, "dochaus", "plugin", "research-mcp.ts"), path.join(repository, "dochaus", "plugin", "research-mcp.ts"))
    const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
    config.mcp = {}
    config.model = config.small_model = "local-fixture/local-fixture"
    config.provider = { "local-fixture": { npm: "@ai-sdk/openai-compatible", name: "Local restart fixture",
      options: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only" },
      models: { "local-fixture": { name: "Local fixture", tool_call: true, limit: { context: 131072, output: 8192 } } } } }
    config.agent = Object.fromEntries(Object.entries(config.agent).map(([id, value]) => [id, { ...(value as object), model: config.model }]))
    config.agent["restart-fixture"] = { mode: "primary", model: config.model,
      prompt: "Return a short neutral response.", tools: { "*": false } }
    await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
    const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
    const port = probe.port
    probe.stop(true)
    const base = `http://127.0.0.1:${port}`
    const env = { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_RUNTIME_DIR: runtime,
      DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: base, OPENCODE_HOSTNAME: "127.0.0.1",
      OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" }
    const headers = { "content-type": "application/json", "x-opencode-directory": matter }
    async function request(route: string, body?: object) {
      const response = await fetch(base + route, { headers, method: body ? "POST" : "GET",
        body: body ? JSON.stringify(body) : undefined, signal: AbortSignal.timeout(5000) })
      expect(response.status).toBe(200)
      return response.json()
    }
    async function until(check: () => Promise<boolean>) {
      const deadline = Date.now() + 30_000
      while (!await check()) {
        if (Date.now() >= deadline) throw new Error("Neutral native restart did not reach its expected boundary")
        await Bun.sleep(50)
      }
    }
    async function start(managed: boolean, recoveryBlocked = false) {
      const recoveryLog = path.join(temporary, `blocked-recovery-${engines.length}.log`)
      const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository,
        env: { ...env, DOCHAUS_MANAGED_RESTART: managed ? "1" : "" }, stdout: "ignore",
        stderr: recoveryBlocked ? Bun.file(recoveryLog) : "inherit" })
      engines.push(engine)
      await until(async () => {
        if (engine.exitCode !== null) throw new Error("Neutral engine exited before readiness")
        return (await fetch(base + "/api/info").catch(() => undefined))?.ok === true
      })
      expect(Number(await readFile(path.join(runtime, "engine-process.lock"), "utf8"))).toBe(engine.pid)
      const ready = Bun.spawn([process.execPath, "run", "services/engine/script/wait-ready.ts"], {
        cwd: repository, env, stdout: "ignore", stderr: "inherit" })
      expect(await ready.exited).toBe(0)
      await until(async () => (await request("/api/plugin")).data.some((entry: { id: string; state: { status: string } }) =>
        entry.id === "dochaus.legal" && entry.state.status === "active"))
      if (recoveryBlocked) {
        if (managed) process.kill(engine.pid, "SIGUSR2")
        // Observe the actual sweep's admission decision, rather than relying
        // on a delay during which a slow native restart might not have begun.
        await until(async () => (await readFile(recoveryLog, "utf8")).includes("Suspended session recovery deferred:"))
      }
      return engine
    }
    async function stop(engine: ReturnType<typeof Bun.spawn>) {
      if (engine.exitCode !== null) return
      engine.kill("SIGTERM")
      await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    }
    const nativeState = () => {
      const db = new Database(path.join(runtime, "data", "opencode", "opencode-local.db"), { readonly: true })
      try {
        return { session: db.query("SELECT time_suspended, resume_attempts FROM session_v2 WHERE id = ?").get("ses_restart_boundary"),
          messages: db.query("SELECT id, type, data FROM session_message WHERE session_id = ? ORDER BY seq").all("ses_restart_boundary") }
      } finally { db.close() }
    }
    const initial = await start(false)
    await request("/api/session", { id: "ses_restart_boundary", title: "Neutral restart boundary", agent: "restart-fixture",
      model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })
    await request("/api/session/ses_restart_boundary/prompt", { id: "msg_restart_boundary", text: "Return the neutral response." })
    await until(async () => calls === 1 && Boolean((await request("/api/session/active")).data.ses_restart_boundary))
    await stop(initial)
    const checkpoint = nativeState()
    expect(checkpoint.session).toMatchObject({ resume_attempts: 0 })
    expect(checkpoint.session).not.toMatchObject({ time_suspended: null })
    const unreleased = await start(true)
    expect(calls).toBe(1)
    expect((await request("/api/session/active")).data).toEqual({})
    expect(nativeState()).toEqual(checkpoint)
    await stop(unreleased)
    expect(calls).toBe(1)
    expect(nativeState()).toEqual(checkpoint)
    const configFile = path.join(repository, "dochaus", "opencode.json")
    const setupFile = path.join(runtime, "config", "opencode", "setup.json")
    await mkdir(path.dirname(setupFile), { recursive: true })
    for (const blockedBy of ["missing-models", "incomplete-setup"]) {
      if (blockedBy === "missing-models") {
        const { model, small_model, ...withoutModels } = config
        await writeFile(configFile, JSON.stringify(withoutModels))
      } else {
        await writeFile(configFile, JSON.stringify(config))
        await writeFile(setupFile, JSON.stringify({ version: 1, completed: ["models"], skipped: [] }))
      }
      for (const managed of [false, true]) {
        const blocked = await start(managed, true)
        expect(calls).toBe(1)
        expect((await request("/api/session/active")).data).toEqual({})
        expect(nativeState()).toEqual(checkpoint)
        await stop(blocked)
        expect(calls).toBe(1)
        expect(nativeState()).toEqual(checkpoint)
      }
    }
    // Completing setup permits an explicit later restart to resume the same
    // durable work, with its saved model and input retained.
    await writeFile(setupFile, JSON.stringify({ version: 1, completed: ["models"], skipped: ["details", "research"] }))
    const released = await start(true)
    expect(calls).toBe(1)
    expect(nativeState()).toEqual(checkpoint)
    const events = await fetch(base + "/api/event", { signal: eventsAbort.signal })
    if (!events.body) throw new Error("Neutral native lifecycle stream is missing")
    eventsReader = events.body.getReader()
    eventsReading = (async () => {
      const decoder = new TextDecoder()
      let buffered = ""
      while (true) {
        const chunk = await eventsReader!.read()
        if (chunk.done) return
        buffered += decoder.decode(chunk.value, { stream: true })
        if (buffered.length > 65536) throw new Error("Neutral native lifecycle frame exceeded its observation bound")
        for (;;) {
          const end = buffered.indexOf("\n\n")
          if (end < 0) break
          const frame = buffered.slice(0, end)
          buffered = buffered.slice(end + 2)
          if (!frame.startsWith("data: ")) continue
          const event: unknown = JSON.parse(frame.slice(6))
          if (!isRecord(event) || !isRecord(event.data) || event.data.sessionID !== "ses_restart_boundary") continue
          if (event.type === "session.execution.started") executionStarts += 1
          if (event.type === "session.execution.failed") executionError = errorFields(event.data.error)
        }
      }
    })().catch(() => { if (!eventsAbort.signal.aborted) eventObservationFailed = true })
    process.kill(released.pid, "SIGUSR2")
    await until(async () => calls === 2).catch(async error => {
      const activity = await fetch(base + "/api/session/active", { headers, signal: AbortSignal.timeout(5000) })
        .then(async response => response.ok ? await response.json() as unknown : undefined).catch(() => undefined)
      const activityData = isRecord(activity) && isRecord(activity.data) ? activity.data : undefined
      const active = activityData ? Object.values(activityData).slice(0, 16).map(value =>
        isRecord(value) && value.type === "running" ? "running" : "other") : null
      let native: unknown = null
      let messages: unknown = null
      let nativeObservationFailed = false
      try {
        const db = new Database(path.join(runtime, "data", "opencode", "opencode-local.db"), { readonly: true })
        try {
          native = db.query("SELECT time_suspended, resume_attempts, idle_outcome FROM session_v2 WHERE id = ?").get("ses_restart_boundary")
        } finally { db.close() }
        messages = nativeState().messages.slice(-16).map(row => {
          if (!isRecord(row) || typeof row.data !== "string") return { type: "unknown" }
          const parsed: unknown = JSON.parse(row.data)
          const data = isRecord(parsed) ? parsed : undefined
          const model = isRecord(data?.model) ? data.model : undefined
          const time = isRecord(data?.time) ? data.time : undefined
          return { type: typeof row.type === "string" ? row.type : "unknown",
            model: model ? { id: typeof model.id === "string" ? model.id.slice(0, 128) : null,
              providerID: typeof model.providerID === "string" ? model.providerID.slice(0, 128) : null } : null,
            error: errorFields(data?.error), outcome: typeof data?.outcome === "string" ? data.outcome : null,
            status: typeof data?.status === "string" ? data.status : null,
            unfinished: row.type === "assistant" && time?.completed === undefined }
        })
      } catch { nativeObservationFailed = true }
      console.log("DOCHAUS_MANAGED_RESTART_DIAGNOSTIC " + JSON.stringify({ phase: "after_release_provider_wait",
        provider_calls: calls, engine_pid: released.pid, engine_exit: released.exitCode, execution_starts: executionStarts,
        event_observation_failed: eventObservationFailed, execution_error: executionError ?? null,
        native_observation_failed: nativeObservationFailed, native, active, messages }))
      throw error
    })
    process.kill(released.pid, "SIGUSR2")
    process.kill(released.pid, "SIGUSR2")
    expect((await request("/api/session/active")).data.ses_restart_boundary).toBeDefined()
    expect(nativeState().session).toMatchObject({ resume_attempts: 1 })
    expect(calls).toBe(2)
    finishResume()
    await until(async () => !(await request("/api/session/active")).data.ses_restart_boundary)
    const history = (await request("/api/session/ses_restart_boundary/message")).data
    expect(history.filter((message: { id: string }) => message.id === "msg_restart_boundary")).toHaveLength(1)
    expect(history.some((message: { type: string; finish?: string }) => message.type === "assistant" && message.finish === "stop")).toBe(true)
    expect(nativeState().session).toEqual({ time_suspended: null, resume_attempts: 0 })
    expect(nativeState().messages.slice(0, checkpoint.messages.length)).toEqual(checkpoint.messages)
    expect(calls).toBe(2)
  } finally {
    eventsAbort.abort()
    await eventsReader?.cancel().catch(() => undefined)
    await eventsReading
    finishResume()
    for (const engine of engines) {
      if (engine.exitCode !== null) continue
      engine.kill("SIGTERM")
      await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    }
    provider.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)
