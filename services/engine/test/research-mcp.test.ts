import { expect, test } from "bun:test"
import { cp, mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

test("Research settings connect and disconnect the native MCP in a warm matter", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-research-runtime-"))
  const repository = path.join(temporary, "repo")
  const workspace = path.join(temporary, "workspace")
  const matter = path.join(workspace, "matter")
  const headers: string[] = []
  const requestPaths: string[] = []
  const upstream = Bun.serve({ port: 0, async fetch(request) {
    headers.push(request.headers.get("Authorization") ?? "")
    requestPaths.push(new URL(request.url).pathname)
    if (request.method !== "POST") return new Response(null, { status: 405 })
    const body = await request.json()
    if (body.id === undefined) return new Response(null, { status: 202 })
    const result = body.method === "initialize"
      ? { protocolVersion: "2025-06-18", capabilities: { tools: {} }, serverInfo: { name: "neutral-research-fixture", version: "1" } }
      : body.method === "tools/list" ? { tools: [{ name: "lookup", description: "Neutral lookup", inputSchema: { type: "object" } }] }
      : {}
    return Response.json({ jsonrpc: "2.0", id: body.id, result })
  } })
  let engine: ReturnType<typeof Bun.spawn> | undefined
  try {
    await mkdir(path.join(repository, "services", "engine"), { recursive: true })
    await mkdir(path.join(repository, "dochaus", "plugin"), { recursive: true })
    await mkdir(path.join(workspace, ".preferences"), { recursive: true })
    await mkdir(matter, { recursive: true })
    await symlink(path.join(original, "node_modules"), path.join(repository, "node_modules"))
    await symlink(path.join(original, "services", "engine", "node_modules"), path.join(repository, "services", "engine", "node_modules"))
    await cp(path.join(original, "services", "engine", "src"), path.join(repository, "services", "engine", "src"), { recursive: true })
    for (const entry of await readdir(path.join(original, "dochaus"))) {
      if (["opencode.json", "plugin", ".env.local", ".env"].includes(entry)) continue
      await symlink(path.join(original, "dochaus", entry), path.join(repository, "dochaus", entry))
    }
    await symlink(path.join(original, "dochaus", "plugin", "legal-v2.ts"), path.join(repository, "dochaus", "plugin", "legal-v2.ts"))
    // Run the real connector plugin against a neutral local MCP, never an account.
    const plugin = await readFile(path.join(original, "dochaus", "plugin", "research-mcp.ts"), "utf8")
    const approvedUrl = `http://127.0.0.1:${upstream.port}/approved-research/`
    await writeFile(path.join(repository, "dochaus", "plugin", "research-mcp.ts"), plugin.replace("https://mcp.courtlistener.com/", approvedUrl))
    const config = JSON.parse(await readFile(path.join(original, "dochaus", "opencode.json"), "utf8"))
    config.mcp = { servers: { "unapproved-account": { type: "remote", url: `http://127.0.0.1:${upstream.port}/unapproved-account/`, oauth: false } } }
    await writeFile(path.join(repository, "dochaus", "opencode.json"), JSON.stringify(config))
    const probe = Bun.serve({ port: 0, fetch: () => new Response() })
    const port = probe.port
    probe.stop(true)
    const base = `http://127.0.0.1:${port}`
    engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], {
      cwd: repository, stdout: "ignore", stderr: Bun.file(path.join(temporary, "engine.log")),
      env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_RUNTIME_DIR: path.join(temporary, "runtime"),
        DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: base, OPENCODE_HOSTNAME: "127.0.0.1",
        OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" },
    })
    await until(async () => (await fetch(base + "/api/info").catch(() => undefined))?.ok === true)
    const list = async () => {
      let servers: { name: string; status: { status: string } }[] = []
      await until(async () => {
        // Preference writes deliberately make managed consent readiness pending
        // while native reconciliation applies the new token/opt-out. Observe
        // that boundary before inspecting the warm matter's MCP inventory.
        const ready = await fetch(base + "/api/dochaus/ready", { headers: { "x-opencode-directory": matter }, signal: AbortSignal.timeout(30_000) })
        const state = await ready.json()
        expect(["active", "loading"]).toContain(state.legalState)
        expect(["active", "pending"]).toContain(state.consentState)
        if (ready.status === 503) return false
        expect(ready.status).toBe(200)
        expect(state.ready).toBe(true)
        const response = await fetch(base + "/api/mcp", { headers: { "x-opencode-directory": matter }, signal: AbortSignal.timeout(30_000) })
        // A filesystem notification can start another reload between the two
        // reads; its 503 must settle rather than become a false test failure.
        if (response.status === 503) return false
        expect(response.status).toBe(200)
        servers = (await response.json()).data
        return true
      })
      return servers
    }
    const save = (enabled: boolean, token: string) => writeFile(path.join(workspace, ".preferences", "preferences.json"),
      JSON.stringify({ courtlistenerMcpEnabled: enabled, caseLawApiKey: token }))
    expect((await list()).some(server => server.name === "courtlistener")).toBe(false)
    expect((await list()).some(server => server.name === "unapproved-account")).toBe(false)
    expect(headers).toEqual([])
    await save(true, "")
    await Bun.sleep(300)
    expect((await list()).some(server => server.name === "courtlistener")).toBe(false)
    expect(headers).toEqual([])
    await save(true, "first-synthetic-token")
    await until(async () => (await list()).some(server => server.name === "courtlistener" && server.status.status === "connected"))
    expect(headers.length).toBeGreaterThan(0)
    expect(headers.every(header => header === "Token first-synthetic-token")).toBe(true)
    const accountRefresh = await fetch(base + "/api/experimental/mcp/account-refresh", { method: "PUT",
      headers: { "x-opencode-directory": matter, "Content-Type": "application/json" },
      body: JSON.stringify({ config: { type: "remote", url: `http://127.0.0.1:${upstream.port}/unapproved-refresh/`, oauth: false } }) })
    expect(accountRefresh.status).toBe(204)
    expect((await list()).some(server => server.name === "account-refresh")).toBe(false)
    // Account reconciliation may reconnect the approved server. Distinct paths
    // prove that neither unapproved endpoint received a request.
    expect(requestPaths.filter(requestPath => requestPath !== "/approved-research/")).toEqual([])
    await save(true, "second-synthetic-token")
    await until(async () => (await list()).some(server => server.name === "courtlistener" && server.status.status === "connected")
      && headers.includes("Token second-synthetic-token"))
    await save(false, "second-synthetic-token")
    await until(async () => !(await list()).some(server => server.name === "courtlistener"))
    const count = headers.length
    await Bun.sleep(300)
    expect(headers.length).toBe(count)
    // A native/operator denial remains authoritative on subsequent preference
    // and token changes; consent cannot re-add a server explicitly removed.
    await save(true, "second-synthetic-token")
    await until(async () => (await list()).some(server => server.name === "courtlistener" && server.status.status === "connected"))
    const disabledCount = headers.length
    const disabled = await fetch(base + "/api/experimental/mcp/courtlistener", { method: "PUT",
      headers: { "x-opencode-directory": matter, "Content-Type": "application/json" },
      body: JSON.stringify({ config: { type: "remote", url: approvedUrl, oauth: false, disabled: true } }) })
    expect(disabled.status).toBe(204)
    expect((await list()).some(server => server.name === "courtlistener")).toBe(false)
    await Bun.sleep(300)
    expect(headers.length).toBe(disabledCount)
    const allowed = await fetch(base + "/api/experimental/mcp/courtlistener", { method: "PUT",
      headers: { "x-opencode-directory": matter, "Content-Type": "application/json" },
      body: JSON.stringify({ config: { type: "remote", url: approvedUrl, oauth: false } }) })
    expect(allowed.status).toBe(204)
    await until(async () => (await list()).some(server => server.name === "courtlistener" && server.status.status === "connected"))
    const denied = await fetch(base + "/api/experimental/mcp/courtlistener", { method: "DELETE", headers: { "x-opencode-directory": matter } })
    expect(denied.status).toBe(204)
    const deniedCount = headers.length
    await save(true, "third-synthetic-token")
    await Bun.sleep(300)
    expect((await list()).some(server => server.name === "courtlistener")).toBe(false)
    expect(headers.length).toBe(deniedCount)
    expect(requestPaths.filter(requestPath => requestPath !== "/approved-research/")).toEqual([])
  } finally {
    if (engine) {
      engine.kill("SIGTERM")
      await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine?.kill("SIGKILL"); return engine?.exited })])
    }
    upstream.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 60_000)

async function until(check: () => Promise<boolean>) {
  const deadline = Date.now() + 30_000
  while (!await check()) {
    if (Date.now() > deadline) throw new Error("Native research connector did not reach the expected state")
    await Bun.sleep(50)
  }
}
