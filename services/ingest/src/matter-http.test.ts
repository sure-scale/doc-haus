import { expect, test } from "bun:test"
import { mkdtemp, mkdir, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

// Runs the checks against an isolated ingest with its own data root.
type Ingest = {
  root: string
  request: (route: string, options?: RequestInit) => Promise<Response>
  send: (method: string, route: string, body: unknown) => Promise<Response>
}

async function withIsolatedIngest(check: (ingest: Ingest) => Promise<void>) {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-matter-http-"))
  const config = path.join(root, "config")
  await mkdir(config)
  await Bun.write(path.join(config, "opencode.json"), "{}")
  const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const child = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: {
      ...process.env,
      DOCHAUS_DATA_ROOT: root,
      WORKSPACE_ROOT: undefined,
      DOCHAUS_RUNTIME_DIR: undefined,
      DOCHAUS_DIR: config,
      INGEST_HOST: "127.0.0.1",
      INGEST_PORT: String(port),
    },
    stdout: Bun.file(path.join(root, "ingest.log")),
    stderr: Bun.file(path.join(root, "ingest.log")),
  })
  const request = (route: string, options?: RequestInit) =>
    fetch(`http://127.0.0.1:${port}${route}`, { ...options, signal: AbortSignal.timeout(15000) })
  const send = (method: string, route: string, body: unknown) =>
    request(route, { method, headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
  try {
    const deadline = Date.now() + 20000
    while (!(await request("/matters").catch(() => undefined))?.ok) {
      if (Date.now() > deadline || child.exitCode !== null) throw new Error("Isolated ingest did not start")
      await Bun.sleep(50)
    }
    await check({ root, request, send })
  } finally {
    child.kill("SIGTERM")
    await child.exited
    await rm(root, { recursive: true, force: true })
  }
}

// The matter routes on an isolated ingest: a PATCH changes the posture only
// when its body carries the key, and a value outside the three is refused.
test("matter routes round-trip the posture and refuse one outside the three", () => withIsolatedIngest(async ({ root, send }) => {
  const created = await send("POST", "/matters", { title: "Supply", client: "Northwind Traders Ltd", posture: "conservative" })
  expect(created.status).toBe(200)
  const matter = (await created.json()) as { id: string; posture?: string }
  expect(matter.posture).toBe("conservative")
  const stored = async () => JSON.parse(await readFile(path.join(root, "workspace", matter.id, "matter.json"), "utf8")) as Record<string, unknown>

  expect((await send("PATCH", `/matters/${matter.id}`, { title: "Supply renamed" })).status).toBe(200)
  expect(await stored()).toMatchObject({ title: "Supply renamed", client: "Northwind Traders Ltd", posture: "conservative" })

  const refused = await send("PATCH", `/matters/${matter.id}`, { title: "Supply renamed", posture: "aggressive" })
  expect(refused.status).toBe(400)
  expect(((await refused.json()) as { error: string }).error).toBe("Posture must be Client-favorable, Balanced or Conservative.")
  expect((await send("POST", "/matters", { title: "Refused", posture: 42 })).status).toBe(400)

  expect((await send("PATCH", `/matters/${matter.id}`, { title: "Supply renamed", posture: "balanced" })).status).toBe(200)
  expect((await stored()).posture).toBe("balanced")
  expect((await send("PATCH", `/matters/${matter.id}`, { title: "Supply renamed", posture: "" })).status).toBe(200)
  expect("posture" in (await stored())).toBe(false)
  expect((await stored()).client).toBe("Northwind Traders Ltd")
}), 60000)

// A matter id that names no matter (a deleted matter, a wrong link, or a
// caller that sent none and got "undefined") is a 404 the web shows as
// "Matter not found", with no server path in the reply.
test("an unknown matter answers 404 on read, rename and conflicts", () => withIsolatedIngest(async ({ request, send }) => {
  const created = await send("POST", "/matters", { title: "Supply" })
  const matter = (await created.json()) as { id: string }
  const existing = await request(`/matters/${matter.id}`)
  expect(existing.status).toBe(200)
  expect(await existing.json()).toMatchObject({ id: matter.id, title: "Supply", documents: [] })

  for (const id of ["undefined", "deleted-matter-a1b2c3"]) {
    for (const response of [
      await request(`/matters/${id}`),
      await send("PATCH", `/matters/${id}`, { title: "Renamed" }),
      await send("POST", `/matters/${id}/conflicts`, { status: "cleared", by: "A. Partner" }),
    ]) {
      expect(response.status).toBe(404)
      expect(await response.json()).toEqual({ error: "Matter not found." })
    }
  }
}), 60000)
