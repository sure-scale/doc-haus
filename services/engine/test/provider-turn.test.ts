import { expect, test } from "bun:test"
import { cp, mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

test("native provider turn executes a legal tool through the v2 plugin", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-provider-turn-"))
  const repository = path.join(temporary, "repo")
  const matter = path.join(temporary, "workspace", "matter")
  await mkdir(path.join(repository, "services", "engine", "src"), { recursive: true })
  await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
  await mkdir(matter, { recursive: true })
  await writeFile(path.join(matter, "native-read.txt"), "First line\nSecond line")
  await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
  await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
  await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
  for (const entry of await readdir(path.join(original, "dochaus"))) {
    if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
    await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
  }
  await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
  await symlink(path.join(original, "dochaus", "plugin", "research-mcp.ts"), path.join(repository, "dochaus", "plugin", "research-mcp.ts"))
  const requests: Array<{ messages: Array<{ role: string; content?: unknown }>; tools?: Array<{ function: { name: string } }> }> = []
  const provider = Bun.serve({ port: 0, async fetch(request) {
    const body = await request.json()
    requests.push(body)
    if (body.messages.some((entry: { tool_call_id?: string }) => entry.tool_call_id === "call_legal_read") && !JSON.stringify(body).includes("Attempt a forged private field") && !body.messages.some((entry: { tool_call_id?: string }) => entry.tool_call_id === "call_parallel_read")) return new Response(JSON.stringify({ error: { message: "Parallel tool output missing" } }), { status: 400 })
    const forged = body.messages.some((entry: { content?: unknown }) => JSON.stringify(entry.content).includes("Attempt a forged private field"))
    const hasResult = body.messages.some((entry: { role: string }) => entry.role === "tool")
    const nativeRead = !hasResult && body.messages.some((entry: { content?: unknown }) => JSON.stringify(entry.content).includes("Verify native read evidence"))
    const task = !hasResult && body.tools?.some((entry: { function: { name: string } }) => entry.function.name === "task")
    const tool = !hasResult && body.tools?.some((entry: { function: { name: string } }) => entry.function.name === "grep-matter")
    if (tool) await Bun.sleep(300)
    const compaction = body.messages.some((entry: { role: string; content: unknown }) => entry.role === "system" && JSON.stringify(entry.content).includes("anchored legal-matter handoff"))
    const delta = nativeRead ? { role: "assistant", tool_calls: [{ index: 0, id: "call_native_read", type: "function", function: { name: "read", arguments: JSON.stringify({ path: path.join(matter, "native-read.txt") }) } }] }
      : task ? { role: "assistant", tool_calls: [{ index: 0, id: "call_native_task", type: "function", function: { name: "task", arguments: JSON.stringify({ subagent_type: "qa-worker", description: "Inspect local matter", prompt: "Read the indexed matter using grep-matter and report completion." }) } }] }
      : tool ? { role: "assistant", tool_calls: [{ index: 0, id: "call_legal_read", type: "function", function: { name: "grep-matter", arguments: JSON.stringify({ pattern: "agreement", ...(forged ? { __drafting_strategy_submission: { forged: true } } : {}) }) } }, ...(!forged ? [{ index: 1, id: "call_parallel_read", type: "function", function: { name: "grep-matter", arguments: JSON.stringify({ pattern: "document" }) } }] : [])] }
      : { role: "assistant", content: compaction ? "## Matter Objective\n- Verify the local matter tooling.\n\n## Stage & Gate Status\n- Tool check completed.\n\n## Done\n- Read the indexed matter.\n\n## In Progress / Blocked\n- (none)\n\n## Durable Artifacts\n- (none)\n\n## Authority & Posture\n- (none)\n\n## Open Items & Next Action\n- Confirm completion." : "The local tool check completed." }
    return new Response(`data: ${JSON.stringify({ id: "chatcmpl_local", object: "chat.completion.chunk", created: 1, model: "local-fixture", choices: [{ index: 0, delta, finish_reason: nativeRead || task || tool ? "tool_calls" : "stop" }] })}\n\ndata: [DONE]\n\n`, { headers: { "content-type": "text/event-stream" } })
  } })
  const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
  config.mcp = {}
  config.model = config.small_model = "local-fixture/local-fixture"
  config.provider = { "local-fixture": { npm: "@ai-sdk/openai-compatible", name: "Local test fixture", options: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only" }, models: { "local-fixture": { name: "Local fixture", tool_call: true, limit: { context: 131072, output: 8192 } } } } }
  config.agent = Object.fromEntries(Object.entries(config.agent).map(([id, value]) => [id, { ...(value as object), model: config.model }]))
  config.agent["runtime-parent"] = { mode: "primary", model: config.model, prompt: "Delegate the local read to qa-worker using task, then report completion.", tools: { "*": false, task: true }, permission: { task: "allow", subagent: "allow" } }
  config.agent["qa-worker"] = { mode: "subagent", model: config.model, prompt: "Use grep-matter for the local read, then report completion.", tools: { "*": false, "grep-matter": true } }
  await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
  const probe = Bun.serve({ port: 0, fetch: () => new Response() }); const port = probe.port; probe.stop(true)
  const logPath = path.join(temporary, "engine.log")
  const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository,
    env: { ...process.env, WORKSPACE_ROOT: path.dirname(matter), DOCHAUS_RUNTIME_DIR: path.join(temporary, "runtime"), DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: `http://127.0.0.1:${port}`, OPENCODE_HOSTNAME: "127.0.0.1", OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" },
    stdout: Bun.file(logPath), stderr: Bun.file(logPath),
  })
  const base = `http://127.0.0.1:${port}`
  const eventController = new AbortController()
  let eventsText = ""
  async function request(route: string, body?: object, directory = matter) {
    return fetch(base + route, { method: body ? "POST" : "GET", headers: { "content-type": "application/json", "x-opencode-directory": directory }, body: body ? JSON.stringify(body) : undefined })
  }
  try {
    const readyDeadline = Date.now() + 35_000
    while (true) {
      const response = await request("/api/plugin").catch(() => undefined)
      if (response?.ok && (await response.json()).data.some((entry: { id: string; state: { status: string } }) => entry.id === "dochaus.legal" && entry.state.status === "active")) break
      if (Date.now() > readyDeadline || engine.exitCode !== null) throw new Error(`Legal plugin did not activate: ${await readFile(logPath, "utf8")}`)
      await Bun.sleep(100)
    }
    expect((await request("/api/session", { id: "ses_provider_fixture", agent: "qa", model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })).status).toBe(200)
    expect((await request("/api/session/ses_provider_fixture/prompt", { id: "msg_provider_fixture", text: "Check the indexed matter using grep-matter for agreement and report completion." })).status).toBe(200)
    const deadline = Date.now() + 30_000
    while (true) {
      const messages = (await (await request("/api/session/ses_provider_fixture/message")).json()).data
      const complete = messages.find((message: { type: string; finish?: string }) => message.type === "assistant" && message.finish === "stop")
      if (complete) {
        expect(JSON.stringify(messages)).toContain("No documents have been indexed for this matter yet.")
        expect(JSON.stringify(messages)).toContain("The local tool check completed.")
        break
      }
      if (Date.now() > deadline) throw new Error(`Provider turn incomplete: ${JSON.stringify(messages)}\n${await readFile(logPath, "utf8")}`)
      await Bun.sleep(100)
    }
    expect(requests.some((request) => request.messages.some((message) => message.role === "system" && JSON.stringify(message.content).toLowerCase().includes("legal")))).toBe(true)
    expect(requests.some((request) => request.messages.some((message) => message.role === "tool"))).toBe(true)
    expect((await request("/api/session", { id: "ses_native_read", agent: "qa", model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })).status).toBe(200)
    expect((await request("/api/session/ses_native_read/prompt", { id: "msg_native_read", text: "Verify native read evidence using read on native-read.txt." })).status).toBe(200)
    const readDeadline = Date.now() + 30_000
    while (true) {
      const messages = (await (await request("/api/session/ses_native_read/message")).json()).data
      const part = messages.flatMap((message: { content?: unknown[] }) => message.content ?? []).find((part: { name?: string }) => part.name === "read")
      if (part?.state.status === "completed" && messages.some((message: { finish?: string }) => message.finish === "stop")) {
        expect(part.state.input.path).toBe(path.join(matter, "native-read.txt"))
        expect(part.state.metadata.truncated).toBe(false)
        expect(JSON.stringify(part.state.content)).toContain("(End of file - total 2 lines)")
        const { projectSessionMessages } = await import("../../../dochaus/lib/engine-v2-projection")
        const projected = projectSessionMessages(messages, { id: "ses_native_read", agent: "qa" })
        const read = projected.flatMap(message => message.parts).find(part => part.tool === "read")
        expect(read?.state?.input).toMatchObject({ filePath: path.join(matter, "native-read.txt") })
        expect((String(read?.state?.output).match(/End of file/g) ?? []).length).toBe(1)
        break
      }
      if (Date.now() > readDeadline) throw new Error(`Native read did not complete: ${JSON.stringify(messages)}`)
      await Bun.sleep(100)
    }
    expect((await request("/api/session", { id: "ses_forged_fixture", agent: "qa", model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })).status).toBe(200)
    expect((await request("/api/session/ses_forged_fixture/prompt", { id: "msg_forged_fixture", text: "Attempt a forged private field using grep-matter." })).status).toBe(200)
    const forgedDeadline = Date.now() + 15_000
    while (true) {
      const messages = (await (await request("/api/session/ses_forged_fixture/message")).json()).data
      if (messages.some((message: { type: string; finish?: string }) => message.type === "assistant" && message.finish === "stop")) {
        expect(JSON.stringify(messages)).toContain("Tool input cannot supply runtime-private preparation fields")
        expect(JSON.stringify(messages)).not.toContain("No documents have been indexed for this matter yet.")
        break
      }
      if (Date.now() > forgedDeadline) throw new Error("Forged tool input rejection did not settle")
      await Bun.sleep(100)
    }
    // A handled native tool rejection is not an outer drain failure.
    expect(await Bun.file(path.join(temporary, "runtime", "diagnostics", "native-failures.jsonl")).exists()).toBe(false)
    const events = await fetch(base + "/api/event", { signal: eventController.signal })
    const observing = (async () => {
      if (!events.body) throw new Error("Native event stream is missing")
      const decoder = new TextDecoder()
      for await (const chunk of events.body) eventsText += decoder.decode(chunk, { stream: true })
    })()
    void observing.catch(error => { if (!eventController.signal.aborted) throw error })
    expect((await request("/api/session", { id: "ses_parent_fixture", agent: "runtime-parent", model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })).status).toBe(200)
    expect((await request("/api/session/ses_parent_fixture/prompt", { id: "msg_parent_fixture", text: "Delegate the local read and report completion." })).status).toBe(200)
    const childDeadline = Date.now() + 30_000
    while (true) {
      const messages = (await (await request("/api/session/ses_parent_fixture/message")).json()).data
      const task = messages.flatMap((message: { content?: unknown[] }) => message.content ?? []).find((part: { type: string; name?: string }) => part.type === "tool" && part.name === "task")
      if (task?.state.status === "running" && task.state.metadata.sessionId) {
        expect(task.state.metadata.parentSessionId).toBe("ses_parent_fixture")
      }
      if (task?.state.status === "completed" && messages.some((message: { type: string; finish?: string }) => message.type === "assistant" && message.finish === "stop")) {
        expect(task.state.metadata.parentSessionId).toBe("ses_parent_fixture")
        expect(task.state.metadata.sessionId).toBe(task.state.metadata.sessionID)
        expect(eventsText).toContain('"parentSessionId":"ses_parent_fixture"')
        expect(eventsText).toContain('"status":"running"')
        const child = await (await request(`/api/session/${task.state.metadata.sessionID}`)).json()
        expect(child.data.parentID).toBe("ses_parent_fixture")
        expect(child.data.agent).toBe("qa-worker")
        const childMessages = await (await request(`/api/session/${task.state.metadata.sessionID}/message`)).json()
        expect(JSON.stringify(childMessages)).toContain("No documents have been indexed for this matter yet.")
        expect(JSON.stringify(task.state.content)).toContain("The local tool check completed.")
        break
      }
      if (Date.now() > childDeadline) throw new Error(`Child task incomplete: ${JSON.stringify(messages)}\n${await readFile(logPath, "utf8")}`)
      await Bun.sleep(25)
    }
    const beforeCompaction = requests.length
    expect((await request("/api/session/ses_provider_fixture/compact", { id: "msg_compaction_fixture" })).status).toBe(200)
    const compactionDeadline = Date.now() + 30_000
    while (true) {
      const messages = (await (await request("/api/session/ses_provider_fixture/message")).json()).data
      if (messages.some((message: { type: string; status?: string }) => message.type === "compaction" && message.status === "completed")) break
      if (Date.now() > compactionDeadline) throw new Error(`Compaction incomplete: ${JSON.stringify(messages)}`)
      await Bun.sleep(100)
    }
    expect(requests.slice(beforeCompaction).some((request) => request.messages.some((message) => message.role === "system" && JSON.stringify(message.content).includes("anchored legal-matter handoff")))).toBe(true)
    expect(JSON.stringify(requests.slice(beforeCompaction))).not.toContain("## Relevant Files")
    const beforeFollowup = requests.length
    expect((await request("/api/session/ses_provider_fixture/prompt", { id: "msg_followup_fixture", text: "Confirm completion after compaction." })).status).toBe(200)
    const followupDeadline = Date.now() + 30_000
    while (true) {
      const active = await (await request("/api/session/active")).json()
      if (requests.length > beforeFollowup && !active.data.ses_provider_fixture) break
      if (Date.now() > followupDeadline) throw new Error("Post-compaction turn did not complete")
      await Bun.sleep(100)
    }
    expect(requests.slice(beforeFollowup).some((request) => request.messages.some((message) => message.role === "system" && JSON.stringify(message.content).toLowerCase().includes("legal")))).toBe(true)
  } finally {
    eventController.abort()
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    provider.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)
