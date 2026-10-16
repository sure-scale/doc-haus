import { Hono } from "hono"
import { readDraftingPreferences } from "./preferences"
import type { DraftingPreferences } from "../../../dochaus/lib/drafting-preferences"

export type ResearchConnection = "exa" | "courtlistener" | "courtlistener-mcp"
export type ResearchConnectionResult = { status: "verified" | "failed" | "not-configured"; message: string }

// Fixed endpoints only; tests cannot be repurposed as an arbitrary URL proxy.
// Each call is explicit, read-only and bounded. No test runs when settings save.
export async function testResearchConnection(
  service: ResearchConnection,
  prefs: DraftingPreferences,
  request: (url: string, options: RequestInit) => Promise<Response> = fetch,
): Promise<ResearchConnectionResult> {
  const key = service === "exa" ? prefs.searchApiKey : prefs.caseLawApiKey
  if (!key) return { status: "not-configured", message: "Save a credential before testing this connection." }
  if (service === "courtlistener-mcp" && !prefs.courtlistenerMcpEnabled)
    return { status: "not-configured", message: "Save explicit CourtListener connector consent before testing it." }
  const name = service === "exa" ? "Exa" : service === "courtlistener" ? "CourtListener API" : "CourtListener connector"
  try {
    const response = await request(service === "exa" ? "https://api.exa.ai/search" : service === "courtlistener"
      ? "https://www.courtlistener.com/api/rest/v4/api-usage/" : "https://mcp.courtlistener.com/", {
      method: service === "courtlistener" ? "GET" : "POST",
      redirect: "error",
      signal: AbortSignal.timeout(12_000),
      headers: service === "exa"
        ? { "x-api-key": key, "Content-Type": "application/json" }
        : { Authorization: `Token ${key}`, "Content-Type": "application/json", Accept: "application/json, text/event-stream" },
      body: service === "exa" ? JSON.stringify({ query: "official legal sources", numResults: 1, contents: { text: false, highlights: false, summary: false } })
        : service === "courtlistener" ? undefined
        : JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize", params: {
          protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "doc.haus-connection-test", version: "1" },
        } }),
    })
    if (!response.ok) {
      await response.body?.cancel()
      return { status: "failed", message: `${name} rejected the test (HTTP ${response.status}). Check the saved credential, account access and service allowance.` }
    }
    const body = await readTestResponse(response)
    const valid = service === "exa" ? Array.isArray(body.results) : service === "courtlistener" ? Array.isArray(body.current_usage)
      : typeof body.result?.protocolVersion === "string" && typeof body.result?.serverInfo?.name === "string"
    return valid
      ? { status: "verified", message: `${name} responded successfully to this test. This does not confirm an active matter's native connector status.` }
      : { status: "failed", message: `${name} returned an unexpected response. The connection is unverified.` }
  } catch {
    // Do not echo exception messages or remote bodies: either may contain keys.
    return { status: "failed", message: `${name} could not be verified within 12 seconds. Check the service and network, then test again.` }
  }
}

async function readTestResponse(response: Response) {
  const reader = response.body?.getReader()
  if (!reader) throw new Error("Missing response")
  const decoder = new TextDecoder()
  let text = ""
  let bytes = 0
  try {
    while (true) {
      const chunk = await reader.read()
      if (chunk.done) break
      bytes += chunk.value.byteLength
      if (bytes > 65_536) throw new Error("Response too large")
      text += decoder.decode(chunk.value, { stream: true })
      if (response.headers.get("content-type")?.includes("text/event-stream")) {
        const event = text.split(/\r?\n\r?\n/).slice(0, -1).find(event => event.split(/\r?\n/).some(line => line.startsWith("data:")))
        if (event) return JSON.parse(event.split(/\r?\n/).filter(line => line.startsWith("data:")).map(line => line.slice(5).trim()).join("\n"))
      }
    }
    return JSON.parse(text)
  } finally {
    await reader.cancel()
  }
}

export const researchConnectionRoutes = new Hono().post("/:service/test", async c => {
  const service = c.req.param("service")
  if (service !== "exa" && service !== "courtlistener" && service !== "courtlistener-mcp")
    return c.json({ error: "Unknown research service." }, 400)
  try {
    return c.json(await testResearchConnection(service, readDraftingPreferences()))
  } catch {
    return c.json({ status: "failed", message: "Saved research preferences could not be read. No connection was tested." }, 503)
  }
})
