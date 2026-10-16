import { afterAll, expect, test } from "bun:test"

const requests: Array<{ method: string; url: URL; body?: Record<string, unknown> }> = []
let configAvailable = true
let executionFailed = false
// Serve the engine and ingest routes in-process through a fetch stub. The shared
// config module keeps the origins of whichever test file first evaluates it, so
// this file must not depend on owning that evaluation.
const originalFetch = globalThis.fetch
globalThis.fetch = (async (input, init) => handle(new Request(input, init))) as typeof fetch
async function handle(request: Request) {
  const url = new URL(request.url)
  const body = request.method === "POST" ? await request.json().catch(() => undefined) as Record<string, unknown> | undefined : undefined
  requests.push({ method: request.method, url, body })
  if (url.pathname === "/engine/config") return configAvailable ? Response.json({ model: "provider/selected" }) : new Response(null, { status: 503 })
  expect(request.headers.get("x-opencode-directory")).toBe("/active-matter")
  if (url.pathname === "/api/dochaus/provider-catalog") return Response.json({ data: [
    { id: "provider", name: "Provider", activation: "enabled", integrationID: "provider-account" },
    { id: "amazon-bedrock", name: "Bedrock", activation: "auto", integrationID: "amazon-bedrock", host: { type: "aws-chain", options: {}, sources: [] } },
  ] })
  if (url.pathname === "/api/provider") return Response.json({ data: [{ id: "provider", activation: "enabled", integrationID: "provider-account" }] })
  if (url.pathname === "/api/model") return Response.json({ data: [{ id: "selected", modelID: "wire-id", name: "Selected", providerID: "provider", enabled: true }] })
  if (url.pathname === "/api/integration") return Response.json({ data: [{ id: "provider-account", connections: [{ type: "env", name: "TEST_MODEL_KEY" }] }] })
  if (url.pathname === "/api/session" && request.method === "POST") return Response.json({ data: { id: "ses_created", location: body?.location } })
  if (url.pathname === "/api/session") return Response.json({ data: [], cursor: url.searchParams.has("cursor") ? {} : { next: "next-page" } })
  if (/\/api\/session\/[^/]+\/(agent|model)$/.test(url.pathname)) return new Response(null, { status: 204 })
  if (url.pathname.endsWith("/wait")) return new Response(null, { status: 204 })
  if (url.pathname.endsWith("/message")) return Response.json({ data: [
    { type: "idle", id: "msg_idle", time: { created: 3 }, outcome: executionFailed ? "failed" : "succeeded" },
    { type: "assistant", id: "msg_result", time: { created: 2 }, agent: "extract", content: [], ...(executionFailed ? { error: { type: "provider.invalid-request", message: "Provider rejected the request" } } : {}) },
  ], cursor: {} })
  if (request.method === "POST") return Response.json({ data: { id: "msg_admitted", sessionID: "ses_created" } })
  return new Response(null, { status: 404 })
}
const clientApi = await import("../src/api/opencode/client")
const sessionsApi = await import("../src/api/opencode/sessions")
const providersApi = await import("../src/api/opencode/providers")
afterAll(() => {
  globalThis.fetch = originalFetch
})

test("web creation passes the native matter location in the body", async () => {
  const session = await sessionsApi.createSession(clientApi.matterClient("/active-matter"), "Matter conversation")
  expect(session.location.directory).toBe("/active-matter")
  expect(requests.at(-1)?.body).toMatchObject({ title: "Matter conversation", location: { directory: "/active-matter" } })
})
test("web catalog uses selection IDs and list pagination follows native cursors", async () => {
  const client = clientApi.matterClient("/active-matter")
  const catalog = await providersApi.listProviders(client)
  expect(Object.keys(catalog.all[0]!.models)).toEqual(["selected"])
  expect(catalog.all[0]!.integrationID).toBe("provider-account")
  expect(catalog.all[1]).toMatchObject({ id: "amazon-bedrock", activation: "auto", models: {}, host: { type: "aws-chain" } })
  expect(catalog.connected).toEqual(["provider"])
  expect(requests.filter((request) => request.url.pathname === "/api/dochaus/provider-catalog")).toHaveLength(1)
  expect(requests.filter((request) => request.url.pathname === "/api/integration")).toHaveLength(1)
  await sessionsApi.listSessions(client)
  expect(requests.at(-1)?.url.searchParams.get("cursor")).toBe("next-page")
  expect(requests.at(-1)?.url.searchParams.has("order")).toBe(false)
})
test("web prompt refuses a failed model settings read before engine mutation", async () => {
  configAvailable = false
  const before = requests.filter(request => request.method === "POST").length
  await expect(sessionsApi.sendPrompt(clientApi.matterClient("/active-matter"), "ses_created", "draft", "Prepare the requested document.")).rejects.toThrow()
  expect(requests.filter(request => request.method === "POST")).toHaveLength(before)
})


test("web prompt distinguishes a successful drain from failed native execution", async () => {
  configAvailable = true
  const client = clientApi.matterClient("/active-matter")
  executionFailed = true
  await expect(sessionsApi.sendPrompt(client, "ses_created", "extract", "Extract the requested term.")).rejects.toThrow("Provider rejected the request")
  executionFailed = false
  expect(await sessionsApi.sendPrompt(client, "ses_created", "extract", "Extract the requested term.")).toMatchObject({ id: "msg_admitted" })
})
