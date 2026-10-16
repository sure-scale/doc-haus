import { expect, test } from "bun:test"
import { mkdtempSync, readdirSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { inspectParserCapabilities, type ParserCapabilitiesSnapshot } from "./parser-capabilities"
import { parserCapabilitiesRoutes } from "./parser-capabilities-route"

test("isolated Hono GET reports current passive facts without creating storage", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "parser-inventory-route-"))
  try {
    let reads = 0
    const app = parserCapabilitiesRoutes({ inspect: async () => {
      reads += 1
      return inspectParserCapabilities({ diskDirectory: directory, searchPath: directory })
    } })
    expect(reads).toBe(0)
    for (const route of ["/parser-capabilities", "/parser-capabilities?directory=/not-the-observed-filesystem"]) {
      const response = await app.request(route)
      expect(response.status).toBe(200)
      expect(response.headers.get("cache-control")).toBe("no-store")
      const snapshot = await response.json() as ParserCapabilitiesSnapshot
      expect(snapshot.schemaVersion).toBe(1)
      expect(snapshot.host.basis).toBe("ingest_process")
      expect(snapshot.memory.total.source).toBe("os.totalmem")
      expect(snapshot.disk.available.status).toBe("observed")
      expect(snapshot.components).toHaveLength(6)
      expect(snapshot.components.filter(component => component.kind === "package").every(component => component.status === "present")).toBe(true)
      expect(snapshot.accelerator.gpu).toBe("unknown")
      expect(JSON.stringify(snapshot)).not.toContain(directory)
    }
    expect(reads).toBe(2)
    for (const method of ["POST", "PUT", "DELETE", "PATCH"])
      expect((await app.request("/parser-capabilities", { method })).status).toBe(404)
    expect(reads).toBe(2)
    expect(readdirSync(directory)).toEqual([])
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})

test("inventory failures return a bounded public error and do not retry", async () => {
  let reads = 0
  const app = parserCapabilitiesRoutes({ inspect: async () => {
    reads += 1
    throw new Error("private host path and credential context")
  } })
  const response = await app.request("/parser-capabilities")
  expect(response.status).toBe(503)
  expect(await response.json()).toEqual({ error: "The ingestion host inventory is unavailable." })
  expect(reads).toBe(1)
})
