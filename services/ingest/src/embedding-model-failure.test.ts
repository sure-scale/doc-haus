import { expect, test } from "bun:test"
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { MODEL } from "./embed"

test("an embedding model that cannot download fails the upload with 503 and the service keeps serving", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-embedding-failure-"))
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  const modelCache = path.join(temporary, "models")
  await mkdir(workspace)
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  // A cached config lets the load reach the parallel weight downloads, the
  // stage where the library leaves sibling download promises unawaited.
  await mkdir(path.join(modelCache, MODEL), { recursive: true })
  await writeFile(path.join(modelCache, MODEL, "config.json"), JSON.stringify({
    model_type: "modernbert",
    "transformers.js_config": { use_external_data_format: { "model_quantized.onnx": 1 } },
  }))
  // Every outbound HTTPS request goes through this proxy, which resets the
  // connection — the transport failure seen when a first-run model download dies.
  let proxyConnections = 0
  const proxy = Bun.listen({ hostname: "127.0.0.1", port: 0, socket: {
    open(socket) {
      proxyConnections++
      socket.end()
    },
    data() {},
  } })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const serviceLog = path.join(temporary, "ingest.log")
  const proxyURL = `http://127.0.0.1:${proxy.port}`
  const service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config, DOCHAUS_MODEL_CACHE_DIR: modelCache,
      INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1",
      HTTPS_PROXY: proxyURL, https_proxy: proxyURL, HTTP_PROXY: proxyURL, http_proxy: proxyURL,
      NO_PROXY: "127.0.0.1,localhost", no_proxy: "127.0.0.1,localhost" },
    stdout: Bun.file(serviceLog), stderr: Bun.file(serviceLog),
  })
  const base = `http://127.0.0.1:${port}`
  const request = (route: string, options?: RequestInit) => fetch(base + route, {
    ...options, signal: AbortSignal.timeout(30_000),
  })
  const upload = async (matterID: string) => {
    const form = new FormData()
    form.append("file", new File([new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: [
      new Paragraph("1. Term"), new Paragraph("The agreement runs for one year."),
      new Paragraph("2. Notice"), new Paragraph("Notice must be given in writing."),
    ] }] })))], "source.docx"))
    return request(`/matters/${matterID}/documents`, { method: "POST", body: form })
  }
  try {
    const deadline = Date.now() + 20_000
    while (true) {
      if (await request("/matters").then(response => response.ok).catch(() => false)) break
      if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(serviceLog, "utf8")}`)
      await Bun.sleep(50)
    }
    const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ title: "Embedding model unavailable" }) })
    expect(created.status).toBe(200)
    const matter = await created.json() as { id: string }

    const first = await upload(matter.id)
    expect(first.status).toBe(503)
    const failure = await first.json() as { error: string; code: string }
    expect(failure.code).toBe("embedding_model_unavailable")
    expect(failure.error).toContain(MODEL)
    const connectionsAfterFirst = proxyConnections
    expect(connectionsAfterFirst).toBeGreaterThan(0)

    // Let any download the library left unawaited settle, then prove the
    // process survived it and a later upload retries the load from scratch.
    await Bun.sleep(1_000)
    expect(service.exitCode).toBeNull()
    const second = await upload(matter.id)
    expect(second.status).toBe(503)
    expect(proxyConnections).toBeGreaterThan(connectionsAfterFirst)

    // The failed uploads left nothing indexed, and the service still serves.
    const detail = await request(`/matters/${matter.id}`)
    expect(detail.status).toBe(200)
    expect((await detail.json() as { documents: unknown[] }).documents).toEqual([])
    expect(service.exitCode).toBeNull()
  } finally {
    service.kill()
    await service.exited
    proxy.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 60_000)
