import { expect, test } from "bun:test"
import { chmod, cp, mkdir, mkdtemp, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

type Mode = "on" | "off" | "failing"
type ToolPart = { type: string; id?: string; name?: string; state?: { status?: string; input?: unknown; content?: unknown; error?: unknown } }

// The same provider turn, successful tool call and failing tool call must look
// identical to the model and the session whether tracing is on, off, or its sink
// fails. Only the trace root's contents may differ.
test("tracing on, off and failing yields identical model-visible requests, tool results and tool errors", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const base = await mkdtemp(path.join(os.tmpdir(), "dochaus-trace-equivalence-"))
  const repository = path.join(base, "repo")
  const matter = path.join(base, "workspace", "matter")
  const runtime = path.join(base, "runtime")
  const failingRoot = path.join(base, "failing-captures")
  await mkdir(path.join(repository, "services", "engine", "src"), { recursive: true })
  await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
  await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
  await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
  await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
  for (const entry of await readdir(path.join(original, "dochaus"))) {
    if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
    await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
  }
  await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
  await symlink(path.join(original, "dochaus", "plugin", "research-mcp.ts"), path.join(repository, "dochaus", "plugin", "research-mcp.ts"))
  const missing = path.join(matter, "missing-note.txt")
  const received: Record<Mode, string[]> = { on: [], off: [], failing: [] }
  const current: { mode: Mode } = { mode: "on" }
  const provider = Bun.serve({ port: 0, async fetch(request) {
    const text = await request.text()
    received[current.mode].push(text)
    const body = JSON.parse(text)
    const hasResult = body.messages.some((entry: { role: string }) => entry.role === "tool")
    const toolsOffered = body.tools?.some((entry: { function: { name: string } }) => entry.function.name === "grep-matter") === true
    const delta = !hasResult && toolsOffered
      ? { role: "assistant", tool_calls: [
          { index: 0, id: "call_equivalence_grep", type: "function", function: { name: "grep-matter", arguments: JSON.stringify({ pattern: "agreement" }) } },
          { index: 1, id: "call_equivalence_read", type: "function", function: { name: "read", arguments: JSON.stringify({ path: missing }) } },
        ] }
      : { role: "assistant", content: "The neutral equivalence check completed." }
    const chunk = { id: "chatcmpl_equivalence", object: "chat.completion.chunk", created: 1, model: "local-fixture", choices: [{ index: 0, delta, finish_reason: "tool_calls" in delta ? "tool_calls" : "stop" }] }
    return new Response(`data: ${JSON.stringify(chunk)}\n\ndata: [DONE]\n\n`, { headers: { "content-type": "text/event-stream" } })
  } })
  const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
  config.mcp = {}
  config.model = config.small_model = "local-fixture/local-fixture"
  config.provider = { "local-fixture": { npm: "@ai-sdk/openai-compatible", name: "Local test fixture", options: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only" }, models: { "local-fixture": { name: "Local fixture", tool_call: true, limit: { context: 131072, output: 8192 } } } } }
  config.agent = Object.fromEntries(Object.entries(config.agent).map(([id, value]) => [id, { ...(value as object), model: config.model }]))
  await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))

  const run = async (mode: Mode) => {
    current.mode = mode
    await rm(path.dirname(matter), { recursive: true, force: true })
    await rm(runtime, { recursive: true, force: true })
    await rm(failingRoot, { recursive: true, force: true })
    await mkdir(matter, { recursive: true })
    if (mode === "failing") await mkdir(failingRoot, { mode: 0o700 })
    const probe = Bun.serve({ port: 0, fetch: () => new Response() }); const port = probe.port; probe.stop(true)
    const logPath = path.join(base, `engine-${mode}.log`)
    const environment = { ...process.env, WORKSPACE_ROOT: path.dirname(matter), DOCHAUS_RUNTIME_DIR: runtime, DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: `http://127.0.0.1:${port}`, OPENCODE_HOSTNAME: "127.0.0.1", OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1",
      DOCHAUS_PROVIDER_CAPTURE_DIR: mode === "off" ? "off" : mode === "failing" ? failingRoot : path.join(runtime, "captures") }
    const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository, env: environment, stdout: Bun.file(logPath), stderr: Bun.file(`${logPath}.stderr`) })
    const request = (route: string, body?: object) => fetch(`http://127.0.0.1:${port}${route}`, { method: body ? "POST" : "GET", headers: { "content-type": "application/json", "x-opencode-directory": matter }, body: body ? JSON.stringify(body) : undefined })
    try {
      const readyDeadline = Date.now() + 35_000
      while (true) {
        const response = await request("/api/plugin").catch(() => undefined)
        if (response?.ok && (await response.json()).data.some((entry: { id: string; state: { status: string } }) => entry.id === "dochaus.legal" && entry.state.status === "active")) break
        if (Date.now() > readyDeadline || engine.exitCode !== null) throw new Error(`Legal plugin did not activate (${mode}): ${await readFile(logPath, "utf8")}`)
        await Bun.sleep(100)
      }
      // The sink fails for every write after activation: no folder or stream can be created.
      if (mode === "failing") await chmod(failingRoot, 0o500)
      expect((await request("/api/session", { id: "ses_trace_equivalence", agent: "qa", model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })).status).toBe(200)
      expect((await request("/api/session/ses_trace_equivalence/prompt", { id: "msg_trace_equivalence", text: "Check the matter for agreement and read the missing note, then report completion." })).status).toBe(200)
      const deadline = Date.now() + 40_000
      while (true) {
        const messages = (await (await request("/api/session/ses_trace_equivalence/message")).json()).data as Array<{ type: string; finish?: string; content?: ToolPart[] }>
        if (messages.some(message => message.type === "assistant" && message.finish === "stop")) {
          // Let fire-and-forget trace work settle before the engine stops.
          await Bun.sleep(300)
          return {
            tools: messages.flatMap(message => message.content ?? []).filter(part => part.type === "tool")
              .map(part => ({ id: part.id, name: part.name, status: part.state?.status, input: part.state?.input, content: part.state?.content, error: part.state?.error })),
            log: await readFile(`${logPath}.stderr`, "utf8"),
          }
        }
        if (Date.now() > deadline) throw new Error(`Turn incomplete (${mode}): ${JSON.stringify(messages)}\n${await readFile(logPath, "utf8")}`)
        await Bun.sleep(100)
      }
    } finally {
      engine.kill("SIGTERM")
      await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
      if (mode === "failing") await chmod(failingRoot, 0o700)
    }
  }

  try {
    const on = await run("on")
    const onTraceEntries = await readdir(path.join(runtime, "captures"))
    const onToolStream = await Bun.file(path.join(runtime, "captures", "streams", "tool-calls", "ses_trace_equivalence.jsonl")).text()
    const off = await run("off")
    const offCaptureExists = await Bun.file(path.join(runtime, "captures")).exists() || await readdir(path.join(runtime, "captures")).then(() => true, () => false)
    const failing = await run("failing")

    // Model-visible request bytes are identical across modes.
    expect(received.on.length).toBeGreaterThanOrEqual(2)
    expect([...received.off].sort()).toEqual([...received.on].sort())
    expect([...received.failing].sort()).toEqual([...received.on].sort())
    // Tool results and tool errors reach the session identically.
    expect(off.tools).toEqual(on.tools)
    expect(failing.tools).toEqual(on.tools)
    expect(on.tools.find(tool => tool.id === "call_equivalence_grep")?.status).toBe("completed")
    const failed = on.tools.find(tool => tool.id === "call_equivalence_read")
    expect(failed?.status).toBe("error")
    expect(JSON.stringify(failed?.error)).toContain("missing-note.txt")
    // Each mode did what it claims with the trace root.
    expect(onTraceEntries.length).toBeGreaterThan(1)
    expect(onToolStream).toContain("call_equivalence_read")
    expect(offCaptureExists).toBe(false)
    expect(off.log).not.toContain("TRACE CAPTURE DISABLED")
    expect(failing.log).toMatch(/PROVIDER CAPTURE FAILED|TRACE WRITE FAILED/)
    expect(await readdir(failingRoot)).toEqual([])
  } finally {
    provider.stop(true)
    await chmod(failingRoot, 0o700).catch(() => undefined)
    await rm(base, { recursive: true, force: true })
  }
}, 240_000)
