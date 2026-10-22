import { expect, test } from "bun:test"
import { cp, mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

// The real native engine and legal plugin use a local scripted provider and
// ingest transport. Word editing itself is exercised by ingest's draft tests.
test("a source-template native session reads its attachment and prepares a proposal without publishing", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-template-native-"))
  const repository = path.join(temporary, "repo")
  const workspace = path.join(temporary, "workspace")
  const id = "38216820-d3de-4c91-ac09-837c5d76bb05"
  const directory = path.join(workspace, ".template-drafts", id)
  await mkdir(path.join(repository, "services", "engine"), { recursive: true })
  await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
  await mkdir(directory, { recursive: true })
  await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
  await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
  await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
  for (const entry of await readdir(path.join(original, "dochaus"))) {
    if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
    await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
  }
  for (const plugin of ["legal-v2.ts", "research-mcp.ts"])
    await symlink(path.join(original, "dochaus", "plugin", plugin), path.join(repository, "dochaus", "plugin", plugin))
  const draft = { id, dir: directory, source: { name: "Synthetic source.docx", matterId: "synthetic-source" },
    revision: 0, name: "reusable-template.docx", description: "", status: "draft", changes: [] as object[], notes: [] }
  const edits: object[] = []
  const ingestCalls: string[] = []
  const ingest = Bun.serve({ port: 0, async fetch(request) {
    const url = new URL(request.url)
    ingestCalls.push(`${request.method} ${url.pathname}`)
    if (request.method === "GET" && url.pathname === `/template-drafts/${id}`) return Response.json(draft)
    if (request.method === "GET" && url.pathname === `/template-drafts/${id}/text`) return Response.json({ paragraphs: [
      { paragraph_handle: "paragraph-synthetic-party", scope: "body", text: "Dovecote Example Client receives the services.", table: null },
    ] })
    if (request.method === "POST" && url.pathname === `/template-drafts/${id}/edits`) {
      const input = await request.json()
      edits.push(input)
      draft.revision++
      draft.name = input.name
      draft.changes = input.edits
      return Response.json(draft)
    }
    return Response.json({ error: "No library publication or matter access is expected" }, { status: 404 })
  } })
  const toolCatalogs: string[][] = []
  const provider = Bun.serve({ port: 0, async fetch(request) {
    const body = await request.json()
    const tools = (body.tools ?? []).map((entry: { function: { name: string } }) => entry.function.name)
    toolCatalogs.push(tools)
    const messages = body.messages as { role: string; tool_call_id?: string; content?: unknown }[]
    const prepared = messages.some(message => message.tool_call_id === "call_prepare_template")
    const read = messages.some(message => message.tool_call_id === "call_read_template")
    const call = !tools.includes("read-template-source") || prepared ? undefined : read
      ? { id: "call_prepare_template", name: "prepare-template", args: { revision: 0, name: "services-template.docx", description: "Synthetic services drafting base",
        edits: [{ paragraph_handle: "paragraph-synthetic-party", occurrence: 0, before_quote: "Dovecote Example Client", replacement: "[insert client name]" }] } }
      : { id: "call_read_template", name: "read-template-source", args: {} }
    const delta = call ? { role: "assistant", tool_calls: [{ index: 0, id: call.id, type: "function", function: { name: call.name, arguments: JSON.stringify(call.args) } }] }
      : { role: "assistant", content: "The proposed template is ready for your review. It has not been saved to the library." }
    return new Response(`data: ${JSON.stringify({ id: "chatcmpl_template_local", object: "chat.completion.chunk", created: 1, model: "template-local", choices: [{ index: 0, delta, finish_reason: call ? "tool_calls" : "stop" }] })}\n\ndata: [DONE]\n\n`, { headers: { "content-type": "text/event-stream" } })
  } })
  const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
  config.mcp = {}
  config.model = config.small_model = "template-local/template-local"
  config.provider = { "template-local": { npm: "@ai-sdk/openai-compatible", name: "Local template fixture",
    options: { baseURL: `http://127.0.0.1:${provider.port}/v1`, apiKey: "synthetic-only" },
    models: { "template-local": { name: "Local template", tool_call: true, limit: { context: 131072, output: 8192 } } } } }
  config.agent = Object.fromEntries(Object.entries(config.agent).map(([name, value]) => [name, { ...(value as object), model: config.model }]))
  await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const base = `http://127.0.0.1:${port}`
  const log = path.join(temporary, "engine.log")
  const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], { cwd: repository,
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_RUNTIME_DIR: path.join(temporary, "runtime"), INGEST_URL: ingest.url.href,
      DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: base, OPENCODE_HOSTNAME: "127.0.0.1", OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" },
    stdout: Bun.file(log), stderr: Bun.file(log) })
  const request = (route: string, body?: object) => fetch(base + route, { method: body ? "POST" : "GET",
    headers: { "content-type": "application/json", "x-opencode-directory": directory }, body: body ? JSON.stringify(body) : undefined })
  try {
    const deadline = Date.now() + 35_000
    while (true) {
      const response = await request("/api/plugin").catch(() => undefined)
      if (response?.ok && (await response.json()).data.some((entry: { id: string; state: { status: string } }) => entry.id === "dochaus.legal" && entry.state.status === "active")) break
      if (Date.now() > deadline || engine.exitCode !== null) throw new Error(`Template plugin did not activate: ${(await readFile(log, "utf8")).slice(-4000)}`)
      await Bun.sleep(100)
    }
    expect((await request("/api/session", { id: "ses_template_fixture", agent: "template-builder", model: { providerID: "template-local", id: "template-local" }, location: { directory } })).status).toBe(200)
    const response = await request("/api/session/ses_template_fixture/prompt", { id: "msg_template_fixture", text: "Prepare a reusable template from the attached source for my review.", agent: "template-builder", model: { providerID: "template-local", id: "template-local" } })
    expect(response.ok).toBe(true)
    const finish = Date.now() + 30_000
    while (edits.length === 0 && Date.now() < finish) await Bun.sleep(50)
    expect(edits).toEqual([{ revision: 0, name: "services-template.docx", description: "Synthetic services drafting base",
      edits: [{ id: "edit-1", paragraph_handle: "paragraph-synthetic-party", occurrence: 0, before_quote: "Dovecote Example Client", replacement: "[insert client name]" }] }])
    expect(ingestCalls.every(call => call.includes(`/template-drafts/${id}`))).toBe(true)
    expect(ingestCalls.some(call => /publish|\/templates$|\/matters\//.test(call))).toBe(false)
    const sourceTools = toolCatalogs.find(catalog => catalog.includes("read-template-source"))!
    expect(sourceTools).toContain("prepare-template")
    expect(sourceTools).toContain("question")
    for (const forbidden of ["read", "glob", "grep", "list", "read-document", "draft-document"])
      expect(sourceTools).not.toContain(forbidden)
  } finally {
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { if (engine.exitCode === null) engine.kill("SIGKILL"); return engine.exited })])
    provider.stop(true)
    ingest.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 90_000)
