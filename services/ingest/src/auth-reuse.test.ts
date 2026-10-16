import { expect, test } from "bun:test"
import { authReuseRoutes } from "./auth-reuse"

test("proxy forwards metadata reads and one selected connection with engine authentication", async () => {
  const received: { method: string; authorization: string | null; body: unknown }[] = []
  const engine = Bun.serve({ port: 0, hostname: "127.0.0.1", async fetch(request) {
    received.push({ method: request.method, authorization: request.headers.get("authorization"), body: request.method === "POST" ? await request.json() : undefined })
    return Response.json(request.method === "GET" ? { sources: [{ id: "source", label: "OpenCode", connections: [{ id: "connection", integrationID: "opencode-go", label: "Saved", authType: "key" }] }] }
      : { connection: { id: "imported", integrationID: "opencode-go", label: "Saved", authType: "key" } })
  } })
  try {
    const app = authReuseRoutes({ engine: `http://127.0.0.1:${engine.port}`, password: () => "synthetic-engine-password" })
    const response = await app.request("/auth/reuse")
    expect(response.status).toBe(200)
    expect(response.headers.get("cache-control")).toBe("no-store")
    expect((await response.json()).sources[0].connections[0].integrationID).toBe("opencode-go")
    expect(received).toHaveLength(1)
    const selected = await app.request("/auth/reuse", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ sourceId: "source", connectionId: "connection" }) })
    expect((await selected.json()).connection.id).toBe("imported")
    expect(received[1].body).toEqual({ sourceId: "source", connectionId: "connection" })
    expect(received.every(request => request.authorization === `Basic ${btoa("opencode:synthetic-engine-password")}`)).toBe(true)
  } finally { engine.stop(true) }
})

test("proxy rejects unsafe origin, content type, oversized body and browser-supplied paths before fetching", async () => {
  let requests = 0
  const engine = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch() { requests++; return Response.json({}) } })
  try {
    const app = authReuseRoutes({ engine: `http://127.0.0.1:${engine.port}`, password: () => undefined })
    for (const [headers, body, status] of [
      [{ Origin: "https://foreign.invalid", "Content-Type": "application/json" }, "{}", 403],
      [{ "Content-Type": "text/plain" }, "{}", 415],
      [{ "Content-Type": "application/jsonp" }, "{}", 415],
      [{ "Content-Type": "application/json" }, " ".repeat(2048), 400],
      [{ "Content-Type": "application/json" }, JSON.stringify({ sourceId: "source", connectionId: "connection", path: "/synthetic/source" }), 400],
    ] as const) expect((await app.request("/auth/reuse", { method: "POST", headers, body })).status).toBe(status)
    expect(requests).toBe(0)
  } finally { engine.stop(true) }
})

test("engine failure responses are replaced by a fixed safe message", async () => {
  const engine = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch() { return Response.json({ error: "synthetic-error-secret /synthetic/source" }, { status: 500 }) } })
  try {
    const app = authReuseRoutes({ engine: `http://127.0.0.1:${engine.port}`, password: () => undefined })
    const response = await app.request("/auth/reuse")
    expect(response.status).toBe(400)
    expect(await response.text()).not.toContain("synthetic-error-secret")
  } finally { engine.stop(true) }
})

test("provider-scoped discovery returns exact matches without unrelated accounts or store errors", async () => {
  const engine = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch() {
    return Response.json({ sources: [
      { id: "mixed", label: "OpenCode", status: "ready", unavailableCount: 3, message: "unrelated-store-message", connections: [
        { id: "openai", integrationID: "openai", label: "Work account", authType: "oauth", extra: "not-public-metadata" },
        { id: "go", integrationID: "opencode-go", label: "Unrelated account", authType: "key" },
      ] },
      { id: "other", label: "Unrelated source", status: "ready", unavailableCount: 0, connections: [{ id: "openai-like", integrationID: "openai-compatible", label: "Other provider", authType: "key" }] },
      { id: "unavailable", label: "Unavailable source", status: "unavailable", unavailableCount: 5, connections: [], message: "unrelated-source-error" },
    ] })
  } })
  try {
    const app = authReuseRoutes({ engine: `http://127.0.0.1:${engine.port}`, password: () => undefined })
    const response = await app.request("/auth/reuse?integrationID=openai")
    expect(response.status).toBe(200)
    expect(response.headers.get("cache-control")).toBe("no-store")
    expect(await response.json()).toEqual({ sources: [{ id: "mixed", label: "OpenCode", status: "ready", unavailableCount: 0, connections: [{ id: "openai", integrationID: "openai", label: "Work account", authType: "oauth" }] }] })
    expect(await (await app.request("/auth/reuse?integrationID=google")).json()).toEqual({ sources: [] })
    const unscoped = await (await app.request("/auth/reuse")).json()
    expect(unscoped.sources).toHaveLength(3)
    expect(unscoped.sources[0].unavailableCount).toBe(3)
    expect(unscoped.sources[2].message).toBe("unrelated-source-error")
  } finally { engine.stop(true) }
})

test("provider-scoped copy checks the selected source and integration before copying", async () => {
  const received: { method: string; authorization: string | null; body?: unknown }[] = []
  const engine = Bun.serve({ port: 0, hostname: "127.0.0.1", async fetch(request) {
    received.push({ method: request.method, authorization: request.headers.get("authorization"), body: request.method === "POST" ? await request.json() : undefined })
    return Response.json(request.method === "GET" ? { sources: [{ id: "source", label: "OpenCode", status: "ready", connections: [
      { id: "go", integrationID: "opencode-go", label: "Go account", authType: "key" },
      { id: "openai", integrationID: "openai", label: "OpenAI account", authType: "oauth" },
    ] }] } : { connection: { id: "imported", integrationID: "openai", label: "OpenAI account", authType: "oauth" } })
  } })
  try {
    const app = authReuseRoutes({ engine: `http://127.0.0.1:${engine.port}`, password: () => "synthetic-engine-password" })
    const copy = (sourceId: string, connectionId: string, expectedIntegrationID: string) => app.request("/auth/reuse", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ sourceId, connectionId, expectedIntegrationID }) })
    expect((await copy("source", "go", "openai")).status).toBe(400)
    expect((await copy("other-source", "openai", "openai")).status).toBe(400)
    expect(received.map(request => request.method)).toEqual(["GET", "GET"])
    const matched = await copy("source", "openai", "openai")
    expect(matched.status).toBe(200)
    expect((await matched.json()).connection.integrationID).toBe("openai")
    expect(received.map(request => request.method)).toEqual(["GET", "GET", "GET", "POST"])
    expect(received[3].body).toEqual({ sourceId: "source", connectionId: "openai" })
    expect(received.every(request => request.authorization === `Basic ${btoa("opencode:synthetic-engine-password")}`)).toBe(true)
  } finally { engine.stop(true) }
})

test("provider-scoped copy rejects a mismatched engine response with a safe message", async () => {
  const engine = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch(request) {
    return Response.json(request.method === "GET" ? { sources: [{ id: "source", label: "OpenCode", status: "ready", connections: [{ id: "connection", integrationID: "openai", label: "Account", authType: "oauth" }] }] }
      : { connection: { id: "unexpected", integrationID: "opencode-go", label: "unrelated-account-label", authType: "key" } })
  } })
  try {
    const app = authReuseRoutes({ engine: `http://127.0.0.1:${engine.port}`, password: () => undefined })
    const response = await app.request("/auth/reuse", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ sourceId: "source", connectionId: "connection", expectedIntegrationID: "openai" }) })
    expect(response.status).toBe(400)
    expect(await response.text()).not.toContain("unrelated-account-label")
  } finally { engine.stop(true) }
})

test("provider filters must be exact integration identifiers", async () => {
  let requests = 0
  const engine = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch() { requests++; return Response.json({}) } })
  try {
    const app = authReuseRoutes({ engine: `http://127.0.0.1:${engine.port}`, password: () => undefined })
    for (const integrationID of ["", "openai*", "openai/path", "x".repeat(129)]) {
      expect((await app.request(`/auth/reuse?integrationID=${encodeURIComponent(integrationID)}`)).status).toBe(400)
      expect((await app.request("/auth/reuse", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ sourceId: "source", connectionId: "connection", expectedIntegrationID: integrationID }) })).status).toBe(400)
    }
    expect(requests).toBe(0)
  } finally { engine.stop(true) }
})
