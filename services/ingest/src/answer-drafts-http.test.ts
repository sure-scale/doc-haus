import { expect, test } from "bun:test"
import { mkdtemp, mkdir, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"

test("full ingest API saves both draft types, survives service restart and cleans up deleted sessions and matters", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-draft-http-"))
  const config = path.join(root, "config")
  await mkdir(config)
  await Bun.write(path.join(config, "opencode.json"), "{}")
  const engine = Bun.serve({
    hostname: "127.0.0.1",
    port: 0,
    fetch: (request) =>
      request.method === "DELETE"
        ? new Response(null, { status: 204 })
        : Response.json({
            data: {
              id: new URL(request.url).pathname.split("/").at(-1),
              location: { directory: request.headers.get("x-opencode-directory") },
            },
          }),
  })
  const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  let child: ReturnType<typeof Bun.spawn> | undefined
  const request = (endpoint: string, input?: unknown, method = "GET", headers: Record<string, string> = {}) =>
    fetch(`http://127.0.0.1:${port}${endpoint}`, {
      method,
      headers: { "Content-Type": "application/json", ...headers },
      ...(input === undefined ? {} : { body: JSON.stringify(input) }),
      signal: AbortSignal.timeout(15000),
    })
  const start = async () => {
    child = Bun.spawn([process.execPath, "run", "src/server.ts"], {
      cwd: path.resolve(import.meta.dir, ".."),
      env: {
        ...process.env,
        DOCHAUS_DATA_ROOT: root,
        WORKSPACE_ROOT: undefined,
        DOCHAUS_RUNTIME_DIR: undefined,
        DOCHAUS_DIR: config,
        INGEST_HOST: "127.0.0.1",
        INGEST_PORT: String(port),
        OPENCODE_URL: `http://127.0.0.1:${engine.port}`,
        OPENCODE_SERVER_PASSWORD: "",
      },
      stdout: Bun.file(path.join(root, "ingest.log")),
      stderr: Bun.file(path.join(root, "ingest.log")),
    })
    const deadline = Date.now() + 20000
    while (!(await request("/matters").catch(() => undefined))?.ok) {
      if (Date.now() > deadline || child.exitCode !== null) throw new Error("Isolated ingest did not start")
      await Bun.sleep(50)
    }
  }
  const stop = async () => {
    child?.kill("SIGTERM")
    await child?.exited
    child = undefined
  }
  try {
    await start()
    const created = await request("/matters", { title: "Storage HTTP check" }, "POST")
    expect(created.status).toBe(200)
    const matter = (await created.json()) as { id: string; dir: string }
    expect(matter.dir).toBe(path.join(root, "workspace", matter.id))
    const route = `/matters/${matter.id}`
    const draft = `${route}/answer-drafts/lean.owner`
    expect(await (await request(draft)).json()).toEqual({ answers: {}, revision: 0 })
    expect((await request(draft, { revision: 0, answers: { q1: "Drafting answer" } }, "PUT")).status).toBe(200)
    expect(
      (
        await request(
          `${route}/answer-drafts/session.chat`,
          { revision: 0, answers: { form1: { text: "Chat answer", selected: ["one"] } } },
          "PUT",
        )
      ).status,
    ).toBe(200)
    expect((await request(draft, { revision: 0, answers: { q1: "Stale answer" } }, "PUT")).status).toBe(409)
    expect((await request(draft, { revision: -1, answers: {} }, "PUT")).status).toBe(400)
    expect(
      (await request(`${route}/answer-drafts/session.chat`, { revision: 1, answers: { form1: "Wrong type" } }, "PUT"))
        .status,
    ).toBe(400)
    expect((await request(`${route}/answer-drafts/bad`)).status).toBe(400)
    expect((await request("/matters/missing/answer-drafts/lean.owner")).status).toBe(404)
    const cors = await request(draft, undefined, "GET", { Origin: "http://localhost:5174" })
    expect(cors.headers.get("Access-Control-Allow-Origin")).toBe("http://localhost:5174")
    const rejectedOrigin = await request(draft, undefined, "GET", { Origin: "https://unrelated.example" })
    expect(rejectedOrigin.headers.get("Access-Control-Allow-Origin")).toBeNull()
    await stop()
    await start()
    expect(await (await request(draft)).json()).toEqual({ answers: { q1: "Drafting answer" }, revision: 1 })
    expect((await (await request(`${route}/answer-drafts/session.chat`)).json()).answers.form1.text).toBe("Chat answer")
    expect((await request(`${route}/sessions/chat`, undefined, "DELETE")).status).toBe(200)
    expect(await (await request(`${route}/answer-drafts/session.chat`)).json()).toEqual({ answers: {}, revision: 0 })
    expect((await request(route)).status).toBe(200)
    expect((await request(route, undefined, "DELETE")).status).toBe(200)
    expect((await request(draft)).status).toBe(404)
  } finally {
    await stop()
    engine.stop(true)
    await rm(root, { recursive: true, force: true })
  }
}, 60000)
