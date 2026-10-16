import { Hono } from "hono"

type ReuseCatalog = { sources: { id: string; label: string; status: string; connections: { id: string; integrationID: string; label: string; authType: "key" | "oauth" }[] }[] }
const validIntegrationID = (value: unknown): value is string => typeof value === "string" && /^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(value)

// Only metadata and opaque selection cross this proxy. Values stay inside the
// engine, which owns the native credential transaction.
export function authReuseRoutes(options: { engine?: string; password?: () => string | undefined } = {}) {
  const app = new Hono()
  app.on(["GET", "POST"], "/auth/reuse", async c => {
    const origin = c.req.header("origin")
    if (origin && !/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) return c.json({ error: "Origin is not allowed." }, 403)
    const integrationID = c.req.query("integrationID")
    if (integrationID !== undefined && !validIntegrationID(integrationID)) return c.json({ error: "Select one provider." }, 400)
    let selection: { sourceId: string; connectionId: string } | undefined
    let expectedIntegrationID: string | undefined
    if (c.req.method === "POST") {
      if (c.req.header("content-type")?.split(";")[0].trim().toLowerCase() !== "application/json") return c.json({ error: "Select one connection using JSON." }, 415)
      const text = await readSelectionBody(c.req.raw)
      const input: unknown = text === undefined ? undefined : (() => { try { return JSON.parse(text) } catch { return undefined } })()
      if (!input || typeof input !== "object" || Array.isArray(input) || !["connectionId,sourceId", "connectionId,expectedIntegrationID,sourceId"].includes(Object.keys(input).sort().join(",")) ||
          !("sourceId" in input) || !("connectionId" in input) || typeof input.sourceId !== "string" || typeof input.connectionId !== "string" ||
          input.sourceId.length > 128 || input.connectionId.length > 128) return c.json({ error: "Select one existing connection." }, 400)
      if ("expectedIntegrationID" in input) {
        if (!validIntegrationID(input.expectedIntegrationID)) return c.json({ error: "Select one provider." }, 400)
        expectedIntegrationID = input.expectedIntegrationID
      }
      selection = { sourceId: input.sourceId, connectionId: input.connectionId }
    }
    c.header("Cache-Control", "no-store")
    try {
      const engine = options.engine ?? process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`
      const password = options.password ? options.password() : process.env.OPENCODE_SERVER_PASSWORD
      const request = (selected?: { sourceId: string; connectionId: string }) => fetch(`${engine}/api/dochaus/credential-reuse`, {
        method: selected ? "POST" : "GET", signal: AbortSignal.timeout(10_000),
        headers: { ...(password ? { Authorization: `Basic ${btoa(`opencode:${password}`)}` } : {}), ...(selected ? { "Content-Type": "application/json" } : {}) },
        ...(selected ? { body: JSON.stringify(selected) } : {}),
      })
      if (selection && expectedIntegrationID) {
        const selected = selection
        const discovery = await request()
        if (!discovery.ok) return c.json({ error: "Could not check the selected sign-in. Refresh saved sign-ins or connect directly." }, 400)
        const catalog = await discovery.json() as ReuseCatalog
        const matches = catalog.sources.some(source => source.id === selected.sourceId && source.status === "ready" && source.connections.some(connection => connection.id === selected.connectionId && connection.integrationID === expectedIntegrationID))
        if (!matches) return c.json({ error: "The selected sign-in does not match this provider. Refresh saved sign-ins or connect directly." }, 400)
      }
      const response = await request(selection)
      if (!response.ok) return c.json({ error: "Could not reuse existing sign-ins. Refresh the list or connect again." }, 400)
      const result = await response.json()
      if (!selection && integrationID !== undefined) {
        const catalog = result as ReuseCatalog
        return c.json({ sources: catalog.sources.flatMap(source => {
          if (source.status !== "ready") return []
          const connections = source.connections.filter(connection => connection.integrationID === integrationID).map(connection => ({ id: connection.id, integrationID: connection.integrationID, label: connection.label, authType: connection.authType }))
          return connections.length ? [{ id: source.id, label: source.label, status: "ready", unavailableCount: 0, connections }] : []
        }) })
      }
      if (selection && expectedIntegrationID && result.connection?.integrationID !== expectedIntegrationID) return c.json({ error: "The copied sign-in could not be checked for this provider. Refresh the provider catalog before continuing." }, 400)
      return c.json(result)
    } catch { return c.json({ error: "Could not reach the engine. Retry when it is running." }, 503) }
  })
  return app
}

async function readSelectionBody(request: Request) {
  const reader = request.body?.getReader()
  if (!reader) return
  const timer = setTimeout(() => { void reader.cancel().catch(() => {}) }, 5000)
  const chunks: Uint8Array[] = []
  let length = 0
  try {
    while (true) {
      const chunk = await reader.read()
      if (chunk.done) return new TextDecoder().decode(Buffer.concat(chunks))
      length += chunk.value.length
      if (length > 1024) { await reader.cancel(); return }
      chunks.push(chunk.value)
    }
  } catch { return }
  finally { clearTimeout(timer); reader.releaseLock() }
}
