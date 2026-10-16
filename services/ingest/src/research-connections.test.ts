import { expect, test } from "bun:test"
import { DEFAULT_DRAFTING } from "./preferences"
import { testResearchConnection } from "./research-connections"

test("connection checks require saved credentials and separate MCP consent", async () => {
  const never = () => { throw new Error("No network call expected") }
  expect((await testResearchConnection("exa", DEFAULT_DRAFTING, never)).status).toBe("not-configured")
  expect((await testResearchConnection("courtlistener-mcp", { ...DEFAULT_DRAFTING, caseLawApiKey: "synthetic-token" }, never)).status).toBe("not-configured")
})

test("explicit connection tests use fixed read-only requests and return no credentials", async () => {
  const calls: { url: string; options: RequestInit | undefined }[] = []
  const transport = (async (input: string | URL | Request, options?: RequestInit) => {
    calls.push({ url: String(input), options })
    return Response.json(String(input).includes("exa.ai") ? { results: [] } : String(input).includes("mcp.courtlistener")
      ? { jsonrpc: "2.0", id: 1, result: { protocolVersion: "2025-06-18", serverInfo: { name: "neutral-fixture" } } } : { current_usage: [] })
  })
  const prefs = { ...DEFAULT_DRAFTING, searchApiKey: "synthetic-exa-key", caseLawApiKey: "synthetic-case-token", courtlistenerMcpEnabled: true }
  for (const service of ["exa", "courtlistener", "courtlistener-mcp"] as const) {
    const result = await testResearchConnection(service, prefs, transport)
    expect(result.status).toBe("verified")
    expect(JSON.stringify(result)).not.toContain("synthetic")
  }
  expect(calls.map(call => [call.url, call.options?.method])).toEqual([
    ["https://api.exa.ai/search", "POST"], ["https://www.courtlistener.com/api/rest/v4/api-usage/", "GET"], ["https://mcp.courtlistener.com/", "POST"],
  ])
  expect(calls[0].options?.headers).toMatchObject({ "x-api-key": "synthetic-exa-key" })
  expect(JSON.parse(calls[0].options?.body as string)).toMatchObject({ numResults: 1, contents: { text: false } })
  expect(calls[1].options?.headers).toMatchObject({ Authorization: "Token synthetic-case-token" })
  expect(JSON.parse(calls[2].options?.body as string).method).toBe("initialize")
  expect(calls.every(call => call.options?.signal && call.options.redirect === "error")).toBe(true)
})

test("remote denial, timeout and malformed responses never echo secret-bearing errors or bodies", async () => {
  const prefs = { ...DEFAULT_DRAFTING, searchApiKey: "synthetic-secret" }
  for (const transport of [
    async () => new Response("synthetic-secret", { status: 403 }),
    async () => { throw new Error("synthetic-secret") },
    async () => new Response("synthetic-secret"),
    async () => Response.json({ wrong: true }),
    async () => new Response("a".repeat(70_000)),
  ]) {
    const result = await testResearchConnection("exa", prefs, transport)
    expect(result.status).toBe("failed")
    expect(result.message).not.toContain("synthetic-secret")
  }
})

test("connector test accepts a bounded SSE initialize response without executing a tool", async () => {
  const result = await testResearchConnection("courtlistener-mcp", { ...DEFAULT_DRAFTING, caseLawApiKey: "synthetic-token", courtlistenerMcpEnabled: true },
    async () => new Response('event: message\ndata: {"jsonrpc":"2.0","id":1,"result":{"protocolVersion":"2025-06-18","serverInfo":{"name":"neutral-fixture"}}}\n\n', { headers: { "Content-Type": "text/event-stream" } }))
  expect(result.status).toBe("verified")
})
