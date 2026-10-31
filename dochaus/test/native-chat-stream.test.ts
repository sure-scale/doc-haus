import { expect, test } from "bun:test"
import type { V2Event } from "@opencode/client"
import { subscribeNativeChatStream } from "../lib/native-chat-stream"

function status(sessionID: string, directory?: string): V2Event {
  return { id: `evt_${sessionID}`, created: 1, type: "session.status", ...(directory ? { location: { directory } } : {}), data: { sessionID, status: { type: "busy" } } }
}
function feed() {
  let output!: ReadableStreamDefaultController<V2Event>, cancelled = false
  const stream = new ReadableStream<V2Event>({ start(controller) { output = controller }, cancel() { cancelled = true } })
  return {
    send: (event: V2Event) => output.enqueue(event), close: () => output.close(), cancelled: () => cancelled,
    async *connect(signal: AbortSignal) {
      const reader = stream.getReader(), abort = () => { void reader.cancel() }
      signal.addEventListener("abort", abort, { once: true })
      try { while (!signal.aborted) { const item = await reader.read(); if (item.done) return; yield item.value } }
      finally { signal.removeEventListener("abort", abort); await reader.cancel(); reader.releaseLock() }
    },
  }
}
async function until(check: () => boolean) {
  const end = Date.now() + 1_000
  while (!check() && Date.now() < end) await Bun.sleep(5)
  expect(check()).toBe(true)
}

test("shared stream scopes before hydration, coalesces bursts, and discards a retired in-flight snapshot", async () => {
  const events = feed(), controller = new AbortController(), reads: string[] = [], snapshots: string[] = []
  let release!: () => void, deleted = false
  const held = new Promise<void>(resolve => { release = resolve })
  const listening = subscribeNativeChatStream({
    connect: signal => events.connect(signal), signal: controller.signal, directory: "/matter",
    acceptsSession: id => id === "ses_chat", needsRefresh: event => event.type === "session.status", refreshDelayMs: 5,
    read: async id => { reads.push(id); if (reads.length === 2) await held; return id },
    onSnapshot: (_id, value) => snapshots.push(value),
    onEvent: event => { if (event.type === "session.deleted") deleted = true },
  })
  try {
    events.send(status("ses_foreign", "/matter")); events.send(status("ses_chat", "/other"))
    events.send(status("ses_chat", "/matter")); events.send(status("ses_chat", "/matter")); events.send(status("ses_chat", "/matter"))
    await until(() => snapshots.length === 1)
    expect(reads).toEqual(["ses_chat"])
    events.send(status("ses_chat")); await until(() => reads.length === 2)
    events.send({ id: "deleted", created: 2, type: "session.deleted", durable: { aggregateID: "ses_chat", seq: 2, version: 2 }, data: { sessionID: "ses_chat" } })
    events.send(status("ses_chat")); await until(() => deleted); release(); await Bun.sleep(20)
    expect(snapshots).toEqual(["ses_chat"]); expect(reads).toEqual(["ses_chat", "ses_chat"])
  } finally { release(); controller.abort(); await listening }
  expect(events.cancelled()).toBe(true)
})

test("a refresh invalidated during its read drains once more; abort suppresses late output", async () => {
  const events = feed(), controller = new AbortController(), snapshots: number[] = []
  let release!: () => void, releaseLast!: () => void, reads = 0, readSignal: AbortSignal | undefined
  const held = new Promise<void>(resolve => { release = resolve })
  const last = new Promise<void>(resolve => { releaseLast = resolve })
  const listening = subscribeNativeChatStream({
    connect: signal => events.connect(signal), signal: controller.signal, acceptsSession: id => id === "ses_chat",
    needsRefresh: () => true, refreshDelayMs: 5,
    read: async (_id, signal) => { readSignal = signal; const current = ++reads; if (current === 1) await held; if (current === 3) await last; return current },
    onSnapshot: (_id, value) => snapshots.push(value),
  })
  try {
    events.send(status("ses_chat")); await until(() => reads === 1)
    events.send(status("ses_chat")); events.send(status("ses_chat")); await Bun.sleep(0)
    release(); await until(() => snapshots.length === 2)
    expect(snapshots).toEqual([1, 2]); expect(reads).toBe(2)
    events.send(status("ses_chat")); await until(() => reads === 3)
    controller.abort(); expect(readSignal?.aborted).toBe(true)
    releaseLast(); await Bun.sleep(0); expect(snapshots).toEqual([1, 2])
  } finally { release(); releaseLast(); controller.abort(); await listening }
})

test("shared lifecycle reconnects and abort interrupts the reconnect wait", async () => {
  const first = feed(), second = feed(), controller = new AbortController()
  let connections = 0, reconnects = 0, disconnected = 0
  const listening = subscribeNativeChatStream({
    connect: signal => (++connections === 1 ? first : second).connect(signal), signal: controller.signal,
    acceptsSession: () => false, read: async () => undefined, onSnapshot() {}, reconnectDelayMs: 5,
    onReconnect: () => reconnects++, onDisconnect: () => disconnected++,
  })
  first.send({ id: "first", type: "server.connected", data: {} }); first.close()
  await until(() => connections === 2)
  second.send({ id: "second", type: "server.connected", data: {} }); await until(() => reconnects === 1)
  expect(disconnected).toBe(1)
  second.close(); await until(() => disconnected === 2)
  controller.abort(); await listening
  expect(connections).toBe(2)
})
