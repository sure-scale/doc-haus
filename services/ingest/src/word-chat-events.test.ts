import { expect, spyOn, test } from "bun:test"
import { Hono } from "hono"
import { installWordChatEvents } from "./word-chat-events"
import type { WordChatAuth } from "./word-chat-context"

type NativeSession = { id: string; parentID?: string; location: { directory: string }; metadata: { type: string; document: string } }
type Event = { id: string; created: number; type: string; data: Record<string, unknown> }
const directory = "/private/scoped-matter"
const document = "synthetic-document"

function fixture() {
  let authorized = true, disconnected = false
  let controller: ReadableStreamDefaultController<Uint8Array> | undefined
  let connections = 0
  const connected = Promise.withResolvers<void>()
  const requests: string[] = []
  const session = (id: string, parentID?: string, type = "word-chat-turn", doc = document, dir = directory): NativeSession => ({ id, parentID, location: { directory: dir }, metadata: { type, document: doc } })
  const sessions = new Map([
    ["ses_root", session("ses_root", undefined, "word-chat")],
    ["ses_turn", session("ses_turn", "ses_root")],
    ["ses_batch", session("ses_batch", "ses_turn")],
    ["ses_review", session("ses_review", "ses_batch", "word-chat-review")],
    ["ses_foreign", session("ses_foreign", "ses_root", "word-chat-turn", document, "/private/another-matter")],
    ["ses_other_document", session("ses_other_document", "ses_root", "word-chat-turn", "another-document")],
    ["ses_other_root", session("ses_other_root", undefined, "word-chat")],
    ["ses_other_turn", session("ses_other_turn", "ses_other_root")],
    ["ses_cycle", session("ses_cycle", "ses_cycle")],
  ])
  const encoder = new TextEncoder()
  const push = (type: string, sessionID: string, extra: object = {}) => controller!.enqueue(encoder.encode(`data: ${JSON.stringify({ id: "private-native-event-id", created: 1, type, location: { directory }, data: { sessionID, ...extra } })}\n\n`))
  const transport = Object.assign(async (input: RequestInfo | URL, init?: RequestInit) => {
    const request = new Request(input, init)
    const url = new URL(request.url)
    expect(request.headers.get("x-opencode-directory")).toBe(directory)
    requests.push(url.pathname)
    if (url.pathname === "/api/event") {
      connections++
      const body = new ReadableStream<Uint8Array>({
        start(value) {
          controller = value
          connected.resolve()
          request.signal.addEventListener("abort", () => { disconnected = true; try { controller!.close() } catch {} }, { once: true })
          value.enqueue(encoder.encode('data: {"type":"server.connected","data":{"directory":"private-hidden"}}\n\n'))
        },
        cancel() { disconnected = true },
      })
      return new Response(body, { headers: { "Content-Type": "text/event-stream" } })
    }
    const id = /^\/api\/session\/([^/]+)$/.exec(url.pathname)?.[1]
    if (id && sessions.has(id)) return Response.json({ data: sessions.get(id) })
    return Response.json({}, { status: 404 })
  }, { preconnect() {} }) as typeof fetch
  const auth: WordChatAuth = { matter: { id: "synthetic", title: "Synthetic matter", dir: directory }, target: { name: "Draft.docx", docPath: `${directory}/Draft.docx` },
    grant: { token_hash: "synthetic-capability", document_id: document, issued_version: "synthetic-version", origin: "https://localhost", revoked_at: null, overlap: 0 } }
  const app = new Hono()
  installWordChatEvents(app, () => authorized ? auth : undefined, { fetch: transport, engineURL: "http://synthetic.invalid" })
  const open = (conversationId = "ses_root") => app.request("/word/chat/events", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ conversationId }) })
  return { app, open, push, sessions, requests, ready: connected.promise, revoke: () => { authorized = false }, connections: () => connections, disconnected: () => disconnected }
}

async function receive(response: Response) {
  const reader = response.body!.getReader()
  const decoder = new TextDecoder()
  let buffer = ""
  return {
    async next(): Promise<Event> {
      for (;;) {
        const at = buffer.indexOf("\n\n")
        if (at >= 0) {
          const block = buffer.slice(0, at); buffer = buffer.slice(at + 2)
          const data = block.split("\n").find(line => line.startsWith("data: "))?.slice(6)
          if (data) return JSON.parse(data) as Event
          continue
        }
        const result = await reader.read()
        if (result.done) throw new Error("Stream ended")
        buffer += decoder.decode(result.value, { stream: true })
      }
    },
    closed: async () => { for (;;) { if ((await reader.read()).done) return } },
    cancel: () => reader.cancel(),
  }
}

test("Word events map only authorized conversation descendants to public turns and omit private payloads", async () => {
  const f = fixture()
  const response = await f.open()
  expect(response.status).toBe(200)
  expect(response.headers.get("content-type")).toContain("text/event-stream")
  expect(response.headers.get("cache-control")).toBe("no-store")
  const stream = await receive(response)
  const events = [await stream.next(), await stream.next()]
  expect(events[0]?.type).toBe("server.connected")
  expect(events[0]?.data).toEqual({})
  expect(events[1]?.data).toEqual({ sessionID: "ses_root", status: { type: "idle" } })
  await f.ready
  try {
    // Starting authorized input delivers a status. Changing only ownership
    // prevents a notification, and the next valid status serves as a fence.
    for (const id of ["ses_foreign", "ses_other_document", "ses_other_turn", "ses_cycle"])
      f.push("session.status", id, { status: { type: "idle" }, private: "must-not-cross" })
    f.push("session.status", "ses_turn", { status: { type: "busy" }, private: "must-not-cross" })
    events.push(await stream.next())
    expect(events.at(-1)?.data).toEqual({ sessionID: "ses_turn", status: { type: "busy" } })
    f.push("session.text.delta", "ses_review", { delta: "PRIVATE_REVIEW_REASONING" })
    f.push("session.tool.success", "ses_batch", { output: "PRIVATE_TOOL_CONTENT" })
    f.push("session.status", "ses_review", { status: { type: "retry", message: "PRIVATE_PROVIDER_DIAGNOSTIC" } })
    events.push(await stream.next())
    expect(events.at(-1)?.data).toEqual({ sessionID: "ses_turn", status: { type: "busy" } })
    f.push("session.idle", "ses_batch")
    events.push(await stream.next())
    expect(events.at(-1)?.data).toEqual({ sessionID: "ses_turn", status: { type: "idle" } })
    const serialized = JSON.stringify(events)
    for (const hidden of [directory, "must-not-cross", "PRIVATE_", "ses_review", "ses_batch", "private-native-event-id", "synthetic-capability"])
      expect(serialized).not.toContain(hidden)
    expect(events.every(event => Object.keys(event).toSorted().join(",") === "created,data,id,type")).toBe(true)
  } finally { await stream.cancel() }
  for (let attempt = 0; attempt < 20 && !f.disconnected(); attempt++) await Bun.sleep(1)
  expect(f.disconnected()).toBe(true)
})

test("Word events refuse other roots, malformed bodies and revoked links before subscribing", async () => {
  const f = fixture()
  expect((await f.open("ses_turn")).status).toBe(404)
  const root = f.sessions.get("ses_root")!
  root.location.directory = "/private/another-matter"
  expect((await f.open()).status).toBe(404)
  root.location.directory = directory
  root.metadata.document = "another-document"
  expect((await f.open()).status).toBe(404)
  expect((await f.open("../invalid")).status).toBe(400)
  let cancelled = false
  const body = new ReadableStream<Uint8Array>({ start(controller) { controller.enqueue(new Uint8Array(2049)) }, cancel() { cancelled = true } })
  expect((await f.app.request("/word/chat/events", { method: "POST", body, duplex: "half" } as RequestInit)).status).toBe(400)
  expect(cancelled).toBe(true)
  f.revoke()
  expect((await f.open()).status).toBe(401)
  expect(f.connections()).toBe(0)
})

test("revocation closes the Word event bridge and aborts the native subscription before another notification", async () => {
  const f = fixture()
  const stream = await receive(await f.open())
  await stream.next(); await stream.next()
  await f.ready
  f.revoke()
  f.push("session.status", "ses_turn", { status: { type: "busy" } })
  await stream.closed()
  expect(f.disconnected()).toBe(true)
})

test("the idle five-second heartbeat, inside Bun's ten-second idle limit, rechecks authorization without needing a model event", async () => {
  const original = globalThis.setInterval
  let heartbeat: (() => void) | undefined
  const timer = spyOn(globalThis, "setInterval").mockImplementation(((callback: () => void, delay: number) => {
    if (delay === 5_000) heartbeat = callback
    return original(callback, delay)
  }) as typeof setInterval)
  const f = fixture()
  let stream: Awaited<ReturnType<typeof receive>> | undefined
  try {
    stream = await receive(await f.open())
    await stream.next(); await stream.next(); await f.ready
    timer.mockRestore()
    expect(heartbeat).toBeDefined()
    heartbeat!()
    expect(f.disconnected()).toBe(false)
    f.revoke()
    heartbeat!()
    await stream.closed()
    expect(f.disconnected()).toBe(true)
  } finally { timer.mockRestore(); await stream?.cancel() }
})
