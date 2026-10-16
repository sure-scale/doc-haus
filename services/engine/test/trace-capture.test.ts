import { expect, test } from "bun:test"
import { cp, lstat, mkdir, mkdtemp, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

type Capture = { folder: string; request: Record<string, unknown> & { headers: Record<string, string> }; receipt: Record<string, unknown>; body: string; response?: string }
type ToolLine = { phase: string; session_id: string; message_id: string; call_id: string; tool: string; input: Record<string, unknown>; result?: { content?: unknown; metadata?: Record<string, unknown> } }

// Runs the real engine entrypoint against a local OpenAI-compatible fixture and
// proves the default-on trace records what the provider and tools actually saw.
test("explicit engine trace captures system prompt, provider wire, model output, full tool I/O, subagent sessions and stateless generation", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-trace-capture-"))
  const repository = path.join(temporary, "repo")
  const matter = path.join(temporary, "workspace", "matter")
  const runtime = path.join(temporary, "runtime")
  await mkdir(path.join(repository, "services", "engine", "src"), { recursive: true })
  await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
  await mkdir(matter, { recursive: true })
  await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
  await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
  await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
  for (const entry of await readdir(path.join(original, "dochaus"))) {
    if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
    await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
  }
  await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
  await symlink(path.join(original, "dochaus", "plugin", "research-mcp.ts"), path.join(repository, "dochaus", "plugin", "research-mcp.ts"))
  const parentPrompt = "Delegate the neutral local read to qa-worker using task, then report completion."
  const workerPrompt = "Use grep-matter for the neutral local read, then report completion."
  const received: string[] = []
  const provider = Bun.serve({ port: 0, async fetch(request) {
    const text = await request.text()
    received.push(text)
    const body = JSON.parse(text)
    const answer = "The neutral trace check completed."
    if (!Array.isArray(body.messages)) return Response.json({ error: { message: "unexpected request" } }, { status: 400 })
    const hasResult = body.messages.some((entry: { role: string }) => entry.role === "tool")
    const offered = (name: string) => body.tools?.some((entry: { function: { name: string } }) => entry.function.name === name) === true
    const delta = !hasResult && offered("task")
      ? { role: "assistant", tool_calls: [{ index: 0, id: "call_trace_task", type: "function", function: { name: "task", arguments: JSON.stringify({ subagent_type: "qa-worker", description: "Inspect local matter", prompt: "Read the matter using grep-matter and report completion." }) } }] }
      : !hasResult && offered("grep-matter")
        ? { role: "assistant", tool_calls: [{ index: 0, id: "call_trace_grep", type: "function", function: { name: "grep-matter", arguments: JSON.stringify({ pattern: "agreement" }) } }] }
        : { role: "assistant", content: answer }
    const chunk = { id: "chatcmpl_trace", object: "chat.completion.chunk", created: 1, model: "local-fixture", choices: [{ index: 0, delta, finish_reason: "tool_calls" in delta ? "tool_calls" : "stop" }] }
    if (body.stream === false) return Response.json({ id: "chatcmpl_trace", object: "chat.completion", created: 1, model: "local-fixture", choices: [{ index: 0, message: { role: "assistant", content: answer }, finish_reason: "stop" }] })
    return new Response(`data: ${JSON.stringify(chunk)}\n\ndata: [DONE]\n\n`, { headers: { "content-type": "text/event-stream" } })
  } })
  const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
  config.mcp = {}
  config.model = config.small_model = "local-fixture/local-fixture"
  config.provider = { "local-fixture": { npm: "@ai-sdk/openai-compatible", name: "Local test fixture", options: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-test-only" }, models: { "local-fixture": { name: "Local fixture", tool_call: true, limit: { context: 131072, output: 8192 } } } } }
  config.agent = Object.fromEntries(Object.entries(config.agent).map(([id, value]) => [id, { ...(value as object), model: config.model }]))
  config.agent["runtime-parent"] = { mode: "primary", model: config.model, prompt: parentPrompt, tools: { "*": false, task: true }, permission: { task: "allow", subagent: "allow" } }
  config.agent["qa-worker"] = { mode: "subagent", model: config.model, prompt: workerPrompt, tools: { "*": false, "grep-matter": true } }
  await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
  const probe = Bun.serve({ port: 0, fetch: () => new Response() }); const port = probe.port; probe.stop(true)
  const logPath = path.join(temporary, "engine.log")
  const environment = { ...process.env, WORKSPACE_ROOT: path.dirname(matter), DOCHAUS_RUNTIME_DIR: runtime, DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: `http://127.0.0.1:${port}`, OPENCODE_HOSTNAME: "127.0.0.1", OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" }
  Object.assign(environment, { DOCHAUS_PROVIDER_CAPTURE_DIR: path.join(runtime, "captures") })
  const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository, env: environment, stdout: Bun.file(logPath), stderr: Bun.file(`${logPath}.stderr`) })
  const base = `http://127.0.0.1:${port}`
  const request = (route: string, body?: object) => fetch(base + route, { method: body ? "POST" : "GET", headers: { "content-type": "application/json", "x-opencode-directory": matter }, body: body ? JSON.stringify(body) : undefined })
  const root = path.join(runtime, "captures")
  try {
    const readyDeadline = Date.now() + 35_000
    while (true) {
      const response = await request("/api/plugin").catch(() => undefined)
      if (response?.ok && (await response.json()).data.some((entry: { id: string; state: { status: string } }) => entry.id === "dochaus.legal" && entry.state.status === "active")) break
      if (Date.now() > readyDeadline || engine.exitCode !== null) throw new Error(`Legal plugin did not activate: ${await readFile(logPath, "utf8")}`)
      await Bun.sleep(100)
    }
    expect((await lstat(root)).mode & 0o777).toBe(0o700)
    expect((await request("/api/session", { id: "ses_trace_parent", agent: "runtime-parent", model: { providerID: "local-fixture", id: "local-fixture" }, location: { directory: matter } })).status).toBe(200)
    expect((await request("/api/session/ses_trace_parent/prompt", { id: "msg_trace_parent", text: "Delegate the neutral read and report completion." })).status).toBe(200)
    const deadline = Date.now() + 40_000
    const childID = await (async () => {
      while (true) {
        const messages = (await (await request("/api/session/ses_trace_parent/message")).json()).data
        const task = messages.flatMap((message: { content?: unknown[] }) => message.content ?? []).find((part: { type: string; name?: string }) => part.type === "tool" && part.name === "task")
        if (task?.state.status === "completed" && messages.some((message: { type: string; finish?: string }) => message.type === "assistant" && message.finish === "stop")) return String(task.state.metadata.sessionID)
        if (Date.now() > deadline) throw new Error(`Parent turn incomplete: ${JSON.stringify(messages)}\n${await readFile(logPath, "utf8")}`)
        await Bun.sleep(100)
      }
    })()
    const generated = await fetch(base + "/api/dochaus/generate", { method: "POST", headers: { "content-type": "application/json", "x-opencode-directory": matter }, body: JSON.stringify({ prompt: "Summarise the neutral trace check." }) })
    expect(generated.status).toBe(200)
    const parentMessages = (await (await request("/api/session/ses_trace_parent/message")).json()).data as Array<{ id: string; type: string }>
    const childMessages = (await (await request(`/api/session/${childID}/message`)).json()).data as Array<{ id: string; type: string; content?: Array<{ type: string; id?: string; name?: string; state?: { content?: unknown } }> }>

    const captures = await settledCaptures(root)
    const finals = captures.filter(capture => capture.request.request_fidelity === "final")
    const parentPrimary = finals.filter(capture => capture.request.session_id === "ses_trace_parent" && capture.request.kind === "primary")
    const childPrimary = finals.filter(capture => capture.request.session_id === childID && capture.request.kind === "primary")
    // System prompt and user prompt reach the trace exactly as sent on the wire.
    expect(parentPrimary.length).toBeGreaterThanOrEqual(2)
    expect(parentPrimary.every(capture => capture.body.includes(parentPrompt))).toBe(true)
    expect(parentPrimary[0]!.body).toContain("Delegate the neutral read and report completion.")
    for (const capture of finals) expect(received).toContain(capture.body)
    // Each primary capture names the assistant message its provider turn produced.
    const assistantIDs = new Set([...parentMessages, ...childMessages].filter(message => message.type === "assistant").map(message => message.id))
    for (const capture of [...parentPrimary, ...childPrimary]) expect(assistantIDs.has(String(capture.request.assistant_message_id))).toBe(true)
    // Subagent session: its own captures carry the parent link and the worker prompt.
    expect(childPrimary.length).toBeGreaterThanOrEqual(2)
    expect(childPrimary.every(capture => capture.request.headers["x-parent-session-id"] === "ses_trace_parent" && capture.body.includes(workerPrompt))).toBe(true)
    // Model output: every response byte the runtime consumed is saved. The native
    // client may cancel after the final event instead of reading EOF.
    expect(parentPrimary.some(capture => capture.response?.includes("call_trace_task"))).toBe(true)
    expect(childPrimary.some(capture => capture.response?.includes("The neutral trace check completed."))).toBe(true)
    expect(finals.every(capture => capture.receipt.response_bytes === capture.receipt.observed_response_bytes && capture.response?.length === capture.receipt.response_bytes)).toBe(true)
    expect(finals.every(capture => capture.receipt.capture_error === undefined || capture.receipt.capture_error === "cancelled")).toBe(true)

    // Tool input and output, raw and final, correlated to session, message and call.
    const childTools = await jsonl<ToolLine>(path.join(root, "streams", "tool-calls", `${childID}.jsonl`))
    const grep = childTools.filter(line => line.call_id === "call_trace_grep")
    expect(grep.map(line => line.phase).sort()).toEqual(["final", "raw"])
    expect(grep.every(line => line.tool === "grep-matter" && line.input.pattern === "agreement")).toBe(true)
    const nativeGrep = childMessages.flatMap(message => (message.content ?? []).map(part => ({ message: message.id, part }))).find(entry => entry.part.type === "tool" && entry.part.id === "call_trace_grep")
    expect(grep.every(line => line.message_id === nativeGrep?.message)).toBe(true)
    expect(JSON.stringify(grep.find(line => line.phase === "final")?.result?.content)).toContain(JSON.stringify(nativeGrep?.part.state?.content).slice(1, 40))
    const parentTools = await jsonl<ToolLine>(path.join(root, "streams", "tool-calls", "ses_trace_parent.jsonl"))
    expect(parentTools.find(line => line.call_id === "call_trace_task" && line.phase === "final")?.result?.metadata?.sessionID).toBe(childID)

    // Stateless generation bypasses native provider hooks and is traced by the engine.
    const generations = await jsonl<{ source: string; prompt: string; text?: string; directory?: string }>(path.join(root, "streams", "generations", "stateless.jsonl"))
    expect(generations.find(line => line.source === "engine_generate_endpoint")).toMatchObject({ prompt: "Summarise the neutral trace check.", text: "The neutral trace check completed." })

    for (const entry of await listTree(root)) expect(entry.mode).toBe(entry.directory ? 0o700 : 0o600)
    expect(await readFile(`${logPath}.stderr`, "utf8")).not.toContain("[drafting-trace]")
  } finally {
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
    provider.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)

test("engine startup fails loudly when the configured trace root is unsafe", async () => {
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-trace-refusal-"))
  try {
    const unsafe = path.join(temporary, "captures")
    await mkdir(unsafe, { mode: 0o755 })
    const probe = Bun.serve({ port: 0, fetch: () => new Response() }); const port = probe.port; probe.stop(true)
    const engine = Bun.spawn([process.execPath, "run", "src/server.ts"], { cwd: path.resolve(import.meta.dir, ".."), stdout: "pipe", stderr: "pipe",
      env: { ...process.env, WORKSPACE_ROOT: path.join(temporary, "workspace"), DOCHAUS_RUNTIME_DIR: path.join(temporary, "runtime"), DOCHAUS_PROVIDER_CAPTURE_DIR: unsafe, DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_HOSTNAME: "127.0.0.1", OPENCODE_DISABLE_MODELS_FETCH: "1" } })
    await mkdir(path.join(temporary, "workspace"))
    const exit = await Promise.race([engine.exited, Bun.sleep(30_000).then(() => { engine.kill("SIGKILL"); return "timeout" as const })])
    expect(exit).not.toBe(0)
    expect(exit).not.toBe("timeout")
    expect(await new Response(engine.stderr).text()).toContain("expected 700")
  } finally { await rm(temporary, { recursive: true, force: true }) }
}, 60_000)

async function settledCaptures(root: string) {
  for (let attempt = 0; attempt < 200; attempt++) {
    const folders = (await readdir(root, { withFileTypes: true })).filter(entry => entry.isDirectory() && entry.name !== "streams")
    const captures = await Promise.all(folders.map(async (entry): Promise<Capture | undefined> => {
      const folder = path.join(root, entry.name)
      if (!await Bun.file(path.join(folder, "receipt.json")).exists()) return undefined
      const response = Bun.file(path.join(folder, "response-body.bin"))
      return { folder, request: await Bun.file(path.join(folder, "request.json")).json(), receipt: await Bun.file(path.join(folder, "receipt.json")).json(),
        body: await Bun.file(path.join(folder, "request-body.bin")).text(), ...(await response.exists() ? { response: await response.text() } : {}) }
    }))
    if (captures.every(capture => capture !== undefined)) return captures.filter(capture => capture !== undefined)
    await Bun.sleep(25)
  }
  throw new Error("Provider captures did not settle")
}

async function jsonl<T>(file: string) {
  for (let attempt = 0; attempt < 200 && !await Bun.file(file).exists(); attempt++) await Bun.sleep(25)
  return (await Bun.file(file).text()).split("\n").filter(Boolean).map(line => JSON.parse(line) as T)
}

async function listTree(directory: string): Promise<Array<{ directory: boolean; mode: number }>> {
  const entries = await readdir(directory, { withFileTypes: true })
  const nested = await Promise.all(entries.map(async entry => {
    const file = path.join(directory, entry.name)
    const mode = (await lstat(file)).mode & 0o777
    return entry.isDirectory() ? [{ directory: true, mode }, ...await listTree(file)] : [{ directory: false, mode }]
  }))
  return nested.flat()
}
