import { afterEach, expect, test } from "bun:test"
import { engineV2Fetch } from "./engine-v2"
import { projectSessionMessages } from "./engine-v2-projection"
const servers: ReturnType<typeof Bun.serve>[] = []
afterEach(() => { servers.splice(0).forEach(server => server.stop(true)) })
const session = { id: "ses_test", location: { directory: "/matter" }, agent: "draft", model: { providerID: "selected", id: "model" } }
function serve(handler: (request: Request) => Response | Promise<Response>) {
  const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: handler })
  servers.push(server)
  return server.url.origin
}
const headers = { "x-opencode-directory": "/matter", "content-type": "application/json" }

test("requires exact durable admission and sends queue delivery", async () => {
  let admitted: Record<string, unknown> | undefined
  const base = serve(async request => {
    if (request.method === "GET") return Response.json({ data: session })
    admitted = await request.json() as Record<string, unknown>
    return Response.json({ data: { id: admitted.id, sessionID: session.id, type: "user", delivery: admitted.delivery, payload: { text: admitted.text } } })
  })
  const response = await engineV2Fetch()(`${base}/api/session/ses_test/prompt`, { method: "POST", headers, body: JSON.stringify({ messageID: "msg_exact", agent: "draft", model: { providerID: "selected", modelID: "model" }, parts: [{ type: "text", text: "Continue from the checkpoint." }] }) })
  expect(response.ok).toBe(true)
  expect(admitted).toEqual({ id: "msg_exact", text: "Continue from the checkpoint.", delivery: "queue", metadata: {} })
})
test("rejects mismatched model before sending a prompt", async () => {
  let posts = 0
  const base = serve(request => { if (request.method === "POST") posts++; return Response.json({ data: session }) })
  const response = await engineV2Fetch()(`${base}/api/session/ses_test/prompt`, { method: "POST", headers, body: JSON.stringify({ messageID: "msg_exact", model: { providerID: "other", modelID: "model" }, parts: [{ type: "text", text: "Continue." }] }) })
  expect(response.status).toBe(409)
  expect(posts).toBe(0)
})
test("rejects success without matching admission identity", async () => {
  const base = serve(request => Response.json({ data: request.method === "GET" ? session : { id: "msg_other" } }))
  await expect(engineV2Fetch()(`${base}/api/session/ses_test/prompt`, { method: "POST", headers, body: JSON.stringify({ messageID: "msg_exact", text: "Continue." }) })).rejects.toThrow("exact durable")
})
test("keeps failed reads failed and rejects malformed success", async () => {
  const base = serve(request => new URL(request.url).searchParams.has("bad") ? Response.json({}) : new Response("Unavailable", { status: 503 }))
  expect((await engineV2Fetch()(`${base}/api/session/active`)).status).toBe(503)
  await expect(engineV2Fetch()(`${base}/api/session/active?bad`)).rejects.toThrow("envelope")
})
test("rejects a session from another matter", async () => {
  const base = serve(() => Response.json({ data: { ...session, location: { directory: "/other" } } }))
  await expect(engineV2Fetch()(`${base}/api/session/ses_test`, { headers })).rejects.toThrow("outside")
})
test("projects active state without inventing session data", async () => {
  const base = serve(request => Response.json({ data: new URL(request.url).pathname === "/api/session/active" ? { ses_test: { type: "running" } } : session }))
  expect(await (await engineV2Fetch()(`${base}/api/session/active`, { headers })).json()).toEqual({ ses_test: { type: "busy" } })
})
test("projects tool identity, completion and restrictions", () => {
  const messages = projectSessionMessages([
    { id: "msg_select", type: "agent-switched", agent: "review", time: { created: 1 } },
    { id: "msg_user", type: "user", text: "Review", metadata: { dochaus_tools: { write: false } }, time: { created: 2 } },
    { id: "msg_result", type: "assistant", agent: "review", model: session.model, time: { created: 3, completed: 5 }, finish: "stop", content: [{ type: "tool", id: "call_same", name: "read-document", state: { status: "completed", input: { path: "document.docx" }, content: [{ type: "text", text: "Evidence" }] }, time: { created: 3, completed: 4 } }] },
    { id: "msg_idle", type: "idle", time: { created: 6 }, outcome: "succeeded" },
  ], session)
  expect(messages).toHaveLength(2)
  expect(messages[0]!.info.tools).toEqual({ write: false })
  expect(messages[1]!.info).toMatchObject({ parentID: "msg_user", time: { completed: 5 } })
  expect(messages[1]!.parts[0]).toMatchObject({ callID: "call_same", tool: "read-document", state: { output: "Evidence", time: { end: 4 } } })
})

test("never admits or interrupts a session outside the requested matter", async () => {
  let mutations = 0
  const base = serve(request => {
    if (request.method !== "GET") mutations++
    return Response.json({ data: { ...session, location: { directory: "/other" } } })
  })
  for (const suffix of ["prompt", "interrupt"]) {
    await expect(engineV2Fetch()(`${base}/api/session/ses_test/${suffix}`, { method: "POST", headers,
      body: JSON.stringify({ messageID: "msg_exact", text: "Continue." }),
    })).rejects.toThrow("outside")
  }
  expect(mutations).toBe(0)
})
test("keeps native selection IDs and excludes unavailable models", async () => {
  const base = serve(request => Response.json({ data: new URL(request.url).pathname === "/api/provider"
    ? [{ id: "selected", activation: "auto" }, { id: "disabled", activation: "disabled" }]
    : [{ id: "alias", modelID: "wire-model", providerID: "selected", enabled: true }, { id: "off", providerID: "selected", enabled: false }] }))
  const catalog = await (await engineV2Fetch()(`${base}/api/provider`, { headers })).json()
  expect(catalog.connected).toEqual(["selected"])
  expect(Object.keys(catalog.all[0].models)).toEqual(["alias"])
  expect(catalog.all[0].models.alias.modelID).toBe("wire-model")
})
test("traverses history cursors for unbounded reads before projecting", async () => {
  const base = serve(request => {
    const url = new URL(request.url)
    if (!url.pathname.endsWith("/message")) return Response.json({ data: session })
    if (url.searchParams.has("cursor")) {
      expect(url.searchParams.has("order")).toBe(false)
      return Response.json({ data: [{ id: "msg_original", type: "user", text: "Original", time: { created: 1 } }], cursor: {} })
    }
    return Response.json({ data: [{ id: "msg_recent", type: "user", text: "Recent", time: { created: 2 } }], cursor: { next: "older" } })
  })
  const response = await engineV2Fetch()(`${base}/api/session/ses_test/message`, { headers })
  expect((await response.json()).map((row: { info: { id: string } }) => row.info.id)).toEqual(["msg_original", "msg_recent"])
})

test("rejects a bare no-content response as proof of prompt admission", async () => {
  const base = serve(request => request.method === "GET" ? Response.json({ data: session }) : new Response(null, { status: 204 }))
  await expect(engineV2Fetch()(`${base}/api/session/ses_test/prompt`, { method: "POST", headers, body: JSON.stringify({ messageID: "msg_exact", text: "Continue." }) })).rejects.toThrow("exact durable")
})
test("collects every child session page for durable worker discovery", async () => {
  const base = serve(request => {
    const url = new URL(request.url)
    expect(url.searchParams.get("directory")).toBe("/matter")
    return url.searchParams.has("cursor")
      ? Response.json({ data: [{ ...session, id: "ses_older" }], cursor: {} })
      : Response.json({ data: [session], cursor: { next: "older" } })
  })
  const rows = await (await engineV2Fetch()(`${base}/api/session?parentID=ses_parent`, { headers })).json()
  expect(rows.map((row: { id: string }) => row.id)).toEqual(["ses_test", "ses_older"])
})

test.each([["google", "google"], ["google-vertex", "vertex"]])("preserves %s signed tool state in the legal policy projection", (providerID, key) => {
  const state = { thoughtSignature: "opaque-provider-signature" }
  const [message] = projectSessionMessages([{ id: "msg_signed", type: "assistant", agent: "draft", model: { providerID, id: "gemini-3" }, time: { created: 1 }, content: [{
    type: "tool", id: "call_signed", name: "task", providerState: state, state: { status: "running", input: {}, metadata: { sessionId: "ses_child", progress: "Working" } }, time: { created: 1 },
  }] }], session)
  expect(message!.parts[0]).toMatchObject({ metadata: { [key!]: state }, providerState: state, state: { metadata: { sessionId: "ses_child", progress: "Working" } } })
})

test("scopes process-wide native activity to the requested matter", async () => {
  const base = serve(request => {
    const route = new URL(request.url).pathname
    return Response.json({ data: route === "/api/session/active" ? { ses_test: { type: "running" }, ses_foreign: { type: "running" } }
      : route.endsWith("ses_foreign") ? { ...session, id: "ses_foreign", location: { directory: "/other" } } : session })
  })
  expect(await (await engineV2Fetch()(`${base}/api/session/active`, { headers })).json()).toEqual({ ses_test: { type: "busy" } })
})

test.each([["default", 200], ["high", 409]] as const)("native variant %s is compared against an omitted default without switching", async (variant, status) => {
  let posts = 0
  const base = serve(async request => {
    if (request.method === "GET") return Response.json({ data: { ...session, model: { ...session.model, variant } } })
    posts++
    const body = await request.json() as { id: string; text: string; delivery: string }
    return Response.json({ data: { id: body.id, sessionID: session.id, type: "user", delivery: body.delivery, payload: { text: body.text } } })
  })
  const response = await engineV2Fetch()(`${base}/api/session/ses_test/prompt`, { method: "POST", headers, body: JSON.stringify({ messageID: "msg_variant", model: { providerID: "selected", modelID: "model" }, text: "Continue." }) })
  expect(response.status).toBe(status)
  expect(posts).toBe(status === 200 ? 1 : 0)
})
