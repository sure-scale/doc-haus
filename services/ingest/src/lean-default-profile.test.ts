import { expect, test } from "bun:test"
import { existsSync, readdirSync } from "node:fs"
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import refusedOpen from "../../../dochaus/workflows/lean/fixtures/start-refused-open.json"
import { initLeanState, writeLeanState } from "../../../dochaus/workflows/lean/state"

// The drafting start route through the real ingest server: a start that names no
// profile runs the lean profile; any other named profile is refused.
// The stand-in engine only creates sessions, so no model runs.
test("drafting start without a profile runs lean; a legacy profile is refused; a finished lean run gives way to a new one", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-lean-default-"))
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  await mkdir(workspace)
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  const sessionCreates: Array<{ directory: string | null; body: { title?: string; metadata?: { dochaus_role?: string } } }> = []
  const engine = Bun.serve({ port: 0, async fetch(request) {
    const url = new URL(request.url)
    if (request.method === "POST" && ["/api/session", "/api/dochaus/session"].includes(url.pathname)) {
      const directory = request.headers.get("x-opencode-directory")
      sessionCreates.push({ directory, body: await request.json() })
      return Response.json({ data: { id: `ses_stand_in_${sessionCreates.length}`, location: { directory } } })
    }
    return new Response("stand-in engine", { status: 503 })
  } })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const serviceLog = path.join(temporary, "ingest.log")
  const service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config,
      INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}` },
    stdout: Bun.file(serviceLog), stderr: Bun.file(serviceLog),
  })
  const base = `http://127.0.0.1:${port}`
  const request = (route: string, options?: RequestInit) => fetch(base + route, { ...options, signal: AbortSignal.timeout(60_000) })
  const json = (body: unknown): RequestInit => ({ method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) })
  const createMatter = async (title: string) => (await (await request("/matters", json({ title }))).json()) as { id: string; dir: string }
  const start = { request: "Draft a short supply note.", agent: "draft", model: { providerID: "stand-in", modelID: "none" } }
  try {
    const deadline = Date.now() + 20_000
    while (!(await request("/matters").then(response => response.ok).catch(() => false))) {
      if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(serviceLog, "utf8")}`)
      await Bun.sleep(50)
    }

    const lean = await createMatter("Default profile")
    const started = await request(`/matters/${lean.id}/drafting-runs/start`, json(start))
    expect(started.status).toBe(200)
    const body = await started.json() as { profile: string; mode: string; ownerSessionID: string; status: { schema: string; profile: string } }
    expect(body).toMatchObject({ profile: "lean", mode: "started", ownerSessionID: "ses_stand_in_1", status: { schema: "dochaus.lean.state/2", profile: "lean" } })
    expect(sessionCreates[0]).toMatchObject({ directory: lean.dir, body: { metadata: { dochaus_role: "lean-drafting-owner" } } })
    expect(existsSync(path.join(lean.dir, ".dochaus", "lean", "state.json"))).toBe(true)
    expect(await (await request(`/matters/${lean.id}/drafting-status`)).json()).toMatchObject({ profile: "lean", owner_session_id: "ses_stand_in_1" })

    const other = await createMatter("Named profiles")
    expect(await (await request(`/matters/${other.id}/drafting-status`)).json()).toBeNull()
    const creates = sessionCreates.length
    for (const profile of ["legacy", "fast"]) {
      const refused = await request(`/matters/${other.id}/drafting-runs/start`, json({ ...start, profile }))
      expect(refused.status).toBe(400)
      expect(await refused.json()).toEqual({ error: "Drafting profile must be \"lean\".", code: "invalid_drafting_profile" })
    }
    expect(sessionCreates.length).toBe(creates)
    expect(existsSync(path.join(other.dir, ".dochaus", "lean"))).toBe(false)

    // A finished lean run does not hold the matter: the next request starts a
    // fresh run and the finished one is archived. An open run points to Continue.
    const previousRun = (directory: string, status: "completed" | "blocked") => {
      const state = initLeanState(directory, { request: "Draft the first note.", ownerSessionID: "ses_previous", model: start.model })
      writeLeanState(directory, { ...state, status, stop_reason: status === "blocked" ? "blocked in an earlier pass" : null })
    }
    const finished = await createMatter("Finished lean run")
    previousRun(finished.dir, "completed")
    const fresh = await request(`/matters/${finished.id}/drafting-runs/start`, json(start))
    expect(fresh.status).toBe(200)
    const freshBody = await fresh.json() as { ownerSessionID: string }
    expect(freshBody).toMatchObject({ profile: "lean", mode: "started", status: { status: "running", owner_session_id: freshBody.ownerSessionID } })
    expect(freshBody.ownerSessionID).not.toBe("ses_previous")
    expect(await readFile(path.join(finished.dir, ".dochaus", "lean", "request.md"), "utf8")).toBe(`${start.request}\n`)
    const archived = readdirSync(path.join(finished.dir, ".dochaus", "lean-archive"))
    expect(archived).toHaveLength(1)
    expect(await readFile(path.join(finished.dir, ".dochaus", "lean-archive", archived[0]!, "request.md"), "utf8")).toBe("Draft the first note.\n")

    const open = await createMatter("Blocked lean run")
    previousRun(open.dir, "blocked")
    const refused = await request(`/matters/${open.id}/drafting-runs/start`, json(start))
    expect(refused.status).toBe(409)
    expect(await refused.json()).toEqual(refusedOpen)
    expect(existsSync(path.join(open.dir, ".dochaus", "lean-archive"))).toBe(false)
    const resumed = await request(`/matters/${open.id}/lean/resume`, { method: "POST" })
    expect(resumed.status).toBe(202)
    expect(await resumed.json()).toMatchObject({ profile: "lean", mode: "queued", status: { owner_session_id: "ses_previous" } })
  } finally {
    service.kill()
    await service.exited
    engine.stop(true)
    if (process.env.DOCHAUS_LEAN_DEFAULT_TEST_LOG) await writeFile(process.env.DOCHAUS_LEAN_DEFAULT_TEST_LOG, await readFile(serviceLog))
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)
