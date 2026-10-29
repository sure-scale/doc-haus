import { expect, test } from "bun:test"
import { mkdtemp, mkdir, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

test("grid HTTP rejects stale writes and invalid bodies while retaining the accepted grid", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-grid-http-"))
  const config = path.join(root, "config")
  await mkdir(config)
  await Bun.write(path.join(config, "opencode.json"), "{}")
  const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => Response.json({}) })
  const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const request = (endpoint: string, input?: unknown, method = "GET") => fetch(`http://127.0.0.1:${port}${endpoint}`, {
    method,
    headers: { "Content-Type": "application/json" },
    ...(input === undefined ? {} : { body: JSON.stringify(input) }),
    signal: AbortSignal.timeout(15000),
  })
  const child = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: {
      ...process.env, DOCHAUS_DATA_ROOT: root, WORKSPACE_ROOT: undefined, DOCHAUS_RUNTIME_DIR: undefined,
      DOCHAUS_DIR: config, INGEST_HOST: "127.0.0.1", INGEST_PORT: String(port),
      OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "",
    },
    stdout: Bun.file(path.join(root, "ingest.log")), stderr: Bun.file(path.join(root, "ingest.log")),
  })
  try {
    const deadline = Date.now() + 20000
    while (!(await request("/matters").catch(() => undefined))?.ok) {
      if (Date.now() > deadline || child.exitCode !== null) throw new Error("Isolated ingest did not start")
      await Bun.sleep(50)
    }
    const created = await request("/matters", { title: "Grid HTTP check" }, "POST")
    expect(created.status).toBe(200)
    const { id } = await created.json() as { id: string }
    const url = `/matters/${id}/grid`
    const initial = await (await request(url)).json() as { version: number; columns: Array<{ id: string; question: string }>; cells: Record<string, unknown> }
    expect(initial).toEqual({ version: 0, columns: [], cells: {} })
    const accepted = await request(url, { ...initial, columns: [{ id: "q", question: "Question?" }] }, "PUT")
    expect(accepted.status).toBe(200)
    const saved = await accepted.json()
    expect(saved).toEqual({ version: 1, columns: [{ id: "q", question: "Question?" }], cells: {} })
    const stale = await request(url, { ...initial, columns: [{ id: "other", question: "Other?" }] }, "PUT")
    expect(stale.status).toBe(409)
    expect((await stale.json()).grid).toEqual(saved)
    expect((await request(url, { columns: [], cells: {} }, "PUT")).status).toBe(400)
    expect((await request(url, { ...saved, cells: [] }, "PUT")).status).toBe(400)
    expect(await (await request(url)).json()).toEqual(saved)
  } finally {
    child.kill("SIGTERM")
    await child.exited
    engine.stop(true)
    await rm(root, { recursive: true, force: true })
  }
}, 30000)
