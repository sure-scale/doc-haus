import { expect, test } from "bun:test"
import { chatClient } from "./chat-client"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"

const link = { version: 1 as const, apiOrigin: "https://localhost:4612", origin: "http://localhost:4444", matter: "synthetic", matterName: "Synthetic matter", document: "draft.docx", documentName: "Draft", token: "synthetic-document-capability", baseVersion: "original" }

test("native SSE reader uses document-granted POST and strips all content-bearing frames and metadata", async () => {
  let init: RequestInit | undefined, path = ""
  const frames = [
    { id: "connected", type: "server.connected", data: {}, metadata: { secret: "private" } },
    { id: "raw", type: "session.text.delta", data: { sessionID: "turn", delta: "unreviewed JSON" } },
    { id: "thinking", type: "session.reasoning.delta", data: { sessionID: "turn", delta: "private reviewer" } },
    { id: "status", created: 1, type: "session.status", location: { directory: "/private" }, metadata: { quote: "unchecked" }, data: { sessionID: "turn", status: { type: "idle", text: "unchecked proposal" } } },
  ]
  const encoded = new TextEncoder().encode(frames.map(frame => `data: ${JSON.stringify(frame)}\r\n\r\n`).join(""))
  const api = chatClient(link, async (url, options) => {
    path = String(url); init = options
    return new Response(new ReadableStream({ start(controller) { controller.enqueue(encoded.slice(0, 37)); controller.enqueue(encoded.slice(37)); controller.close() } }), { headers: { "Content-Type": "text/event-stream" } })
  })
  const controller = new AbortController(), received = []
  for await (const frame of api.events!("conversation", controller.signal)) received.push(frame)
  expect(path).toBe(`${link.apiOrigin}/word/chat/events`); expect(path).not.toContain(link.token)
  expect(init?.method).toBe("POST"); expect(init?.credentials).toBe("omit"); expect(init?.redirect).toBe("error"); expect(init?.referrerPolicy).toBe("no-referrer")
  expect(new Headers(init?.headers).get("Authorization")).toBe(`Bearer ${link.token}`)
  expect(new Headers(init?.headers).get("x-dochaus-word-version")).toBe(WORD_ADDIN_VERSION)
  expect(JSON.parse(init?.body as string)).toEqual({ conversationId: "conversation" })
  expect(received).toEqual([{ id: "connected", type: "server.connected", data: {} }, { id: "status", created: 1, type: "session.status", data: { sessionID: "turn", status: { type: "idle" } } }])
})

test("aborting a quiet Word stream cancels the actual native reader and transport", async () => {
  let requested: AbortSignal | null | undefined, cancelled = false, opened = false
  const server = Bun.serve({ port: 0, fetch() {
    return new Response(new ReadableStream({ start(controller) { opened = true; controller.enqueue(new TextEncoder().encode(": connected\n\n")) }, cancel() { cancelled = true } }), { headers: { "Content-Type": "text/event-stream" } })
  } })
  const api = chatClient(link, async (_url, init) => {
    requested = init?.signal
    return fetch(server.url, init)
  })
  const controller = new AbortController()
  const listening = (async () => { for await (const _event of api.events!("conversation", controller.signal)) { /* quiet stream */ } })()
  try {
    const deadline = Date.now() + 1_000
    while (!opened && Date.now() < deadline) await Bun.sleep(5)
    expect(opened).toBe(true); expect(requested?.aborted).toBe(false)
    controller.abort(); await listening
    while (!cancelled && Date.now() < deadline) await Bun.sleep(5)
    expect(requested?.aborted).toBe(true); expect(cancelled).toBe(true)
  } finally { controller.abort(); await listening; server.stop(true) }
})
