import { expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { inspectParserCapabilities } from "../../../services/ingest/src/parser-capabilities"
import { parserCapabilitiesRoutes } from "../../../services/ingest/src/parser-capabilities-route"
import { getParserCapabilities } from "../src/api/parser-capabilities"

// Bun's fetch type also carries preconnect; keep it so a stub is a whole fetch.
const stubFetch = (respond: () => Promise<Response>): typeof fetch => Object.assign(respond, { preconnect: fetch.preconnect })

test("web client uses the isolated product Hono GET and preserves unknown observations", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "parser-inventory-web-"))
  const calls: Request[] = []
  try {
    // Use Hono through its owning ingest module; the web package has no Hono dependency.
    const app = parserCapabilitiesRoutes({ diskDirectory: directory })
      .route("/host", parserCapabilitiesRoutes({ diskDirectory: directory }))
    const transport = (async (input, init) => {
      const request = new Request(input, init)
      calls.push(request)
      return app.request(request)
    }) as typeof fetch
    const snapshot = await getParserCapabilities({ fetch: transport })
    expect(calls).toHaveLength(1)
    expect(calls[0]!.method).toBe("GET")
    expect(new URL(calls[0]!.url).pathname).toBe("/host/parser-capabilities")
    expect(snapshot.host.basis).toBe("ingest_process")
    expect(snapshot.accelerator).toEqual({ gpu: "unknown", dedicatedVramBytes: null, unifiedMemoryBytes: null, reason: "not_probed" })
    expect(snapshot.components).toHaveLength(6)
    expect(JSON.stringify(snapshot)).not.toContain(directory)
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})

test("web client rejects unsupported schemas, fabricated capacity and invalid component evidence", async () => {
  const snapshot = await inspectParserCapabilities({ searchPath: "", diskDirectory: "/not-an-existing-inventory-filesystem" })
  const invalid = [
    { ...snapshot, schemaVersion: 2 },
    { ...snapshot, memory: { ...snapshot.memory, free: { status: "observed", source: "os.freemem", bytes: -1 } } },
    { ...snapshot, accelerator: { ...snapshot.accelerator, unifiedMemoryBytes: snapshot.memory.total.bytes } },
    { ...snapshot, components: [snapshot.components[0], ...snapshot.components.slice(0, 5)] },
    { ...snapshot, components: snapshot.components.map(component => component.id === "uvx" ? { ...component, status: "ready" } : component) },
  ]
  for (const response of invalid)
    await expect(getParserCapabilities({ fetch: stubFetch(async () => Response.json(response)) })).rejects.toThrow("unsupported inventory response")
  await expect(getParserCapabilities({ fetch: stubFetch(async () => new Response("not JSON")) })).rejects.toThrow("unsupported inventory response")
})

test("web client reports an unavailable endpoint without presenting empty capabilities", async () => {
  await expect(getParserCapabilities({ fetch: stubFetch(async () => new Response("private error context", { status: 503 })) })).rejects.toThrow("The ingestion host inventory is unavailable.")
  await expect(getParserCapabilities({ fetch: stubFetch(async () => { throw new Error("private network detail") }) })).rejects.toThrow("The ingestion host inventory is unavailable.")
})

test("web client forwards cancellation and makes one request without retrying", async () => {
  const controller = new AbortController()
  let requests = 0
  const transport = ((_input, init) => {
    requests += 1
    return new Promise<Response>((_resolve, reject) => {
      init!.signal!.addEventListener("abort", () => reject(new Error("cancelled")), { once: true })
    })
  }) as typeof fetch
  const request = getParserCapabilities({ fetch: transport, signal: controller.signal })
  controller.abort()
  await expect(request).rejects.toThrow("The ingestion host inventory is unavailable.")
  expect(requests).toBe(1)
})
