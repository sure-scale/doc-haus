import { expect, test } from "bun:test"
import { cp, mkdtemp, mkdir, readdir, rm, symlink, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

const upstream = ["build", "general", "explore", "plan"]

test("a cold matter's inventory lists answer only with its activated legal configuration", async () => {
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-inventory-activation-"))
  const workspace = path.join(temporary, "workspace")
  await mkdir(workspace, { recursive: true })
  const engine = await start(path.resolve(import.meta.dir, "../../.."), workspace, path.join(temporary, "runtime"))
  try {
    // Created only now, so no request has booted this matter's location yet.
    const matter = path.join(workspace, "cold-matter")
    await mkdir(matter)
    const query = `?location%5Bdirectory%5D=${encodeURIComponent(matter)}`
    const observed = await Promise.all([
      observe(engine.base + "/api/agent" + query, {}),
      observe(engine.base + "/api/agent", { "x-opencode-directory": encodeURIComponent(matter) }),
      observe(engine.base + "/api/skill" + query, {}),
      observe(engine.base + "/api/command", { "x-opencode-directory": matter }),
    ])
    for (const answers of observed) {
      const final = answers.at(-1)!.ids
      expect(final.length).toBeGreaterThan(0)
      expect(answers).toEqual(answers.map(() => ({ status: 200, ids: final })))
    }
    const agents = observed[0]!.at(-1)!.ids
    expect(agents[0]).toBe("draft")
    expect(agents.filter((id) => upstream.includes(id))).toEqual([])
    expect(observed[1]!.at(-1)!.ids).toEqual(agents)
  } finally {
    await engine.stop()
    await rm(temporary, { recursive: true, force: true })
  }
}, 60_000)

test("a warm matter's agent list waits through a plugin re-activation", async () => {
  const original = path.resolve(import.meta.dir, "../../..")
  const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-inventory-reactivation-"))
  const repository = path.join(temporary, "repo")
  const workspace = path.join(temporary, "workspace")
  const matter = path.join(workspace, "matter")
  await mkdir(path.join(repository, "services", "engine"), { recursive: true })
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
  await cp(path.join(original, "dochaus", "opencode.json"), path.join(repository, "dochaus", "opencode.json"))
  const engine = await start(repository, workspace, path.join(temporary, "runtime"))
  try {
    const query = `?location%5Bdirectory%5D=${encodeURIComponent(matter)}`
    const deadline = Date.now() + 30_000
    while ((await observe(engine.base + "/api/agent" + query, {}, 0))[0]!.ids[0] !== "draft") {
      if (Date.now() > deadline) throw new Error("Matter did not activate")
      await Bun.sleep(50)
    }
    // A new plugin sorted before the legal plugin re-activates every later
    // plugin, including the one that replaces the upstream agents.
    const observing = observe(engine.base + "/api/agent" + query, {})
    await Bun.sleep(300)
    await writeFile(path.join(repository, "dochaus", "plugin", "aa-reactivation-probe.ts"),
      `import { Plugin } from "@opencode/plugin"\nexport default Plugin.define({ id: "test.reactivation", setup: async () => {} })\n`)
    const answers = await observing
    const plugins: { data: { id: string; state: { status: string } }[] } = await (await fetch(engine.base + "/api/plugin" + query)).json()
    expect(plugins.data.find((plugin) => plugin.id === "test.reactivation")?.state.status).toBe("active")
    const final = answers.at(-1)!.ids
    expect(final[0]).toBe("draft")
    expect(final.filter((id) => upstream.includes(id))).toEqual([])
    expect(answers).toEqual(answers.map(() => ({ status: 200, ids: final })))
  } finally {
    await engine.stop()
    await rm(temporary, { recursive: true, force: true })
  }
}, 60_000)

async function start(repository: string, workspace: string, runtime: string) {
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const base = `http://127.0.0.1:${port}`
  const engine = Bun.spawn([process.execPath, "run", "services/engine/src/server.ts"], {
    cwd: repository, stdout: "ignore", stderr: "inherit",
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_RUNTIME_DIR: runtime,
      DOC_HAUS_ENGINE_PORT: String(port), OPENCODE_URL: base, OPENCODE_HOSTNAME: "127.0.0.1",
      OPENCODE_SERVER_PASSWORD: "", OPENCODE_DISABLE_MODELS_FETCH: "1" },
  })
  const stop = async () => {
    engine.kill("SIGTERM")
    await Promise.race([engine.exited, Bun.sleep(3000).then(() => { engine.kill("SIGKILL"); return engine.exited })])
  }
  const deadline = Date.now() + 30_000
  while (!(await fetch(base + "/api/info").catch(() => undefined))?.ok) {
    if (Date.now() > deadline || engine.exitCode !== null) {
      await stop()
      throw new Error("Engine did not start")
    }
    await Bun.sleep(100)
  }
  return { base, stop }
}

// Polls one list from now for the window, keeping every answer in order.
async function observe(url: string, headers: Record<string, string>, window = 3000) {
  const started = performance.now()
  const answers: { status: number; ids: string[] }[] = []
  while (performance.now() - started < window || answers.length < 1) {
    const response = await fetch(url, { headers, signal: AbortSignal.timeout(40_000) })
    const body: { data?: { id?: string; name?: string }[] } = await response.json()
    answers.push({ status: response.status, ids: (body.data ?? []).map((item) => item.id ?? item.name ?? "") })
    await Bun.sleep(5)
  }
  return answers
}
