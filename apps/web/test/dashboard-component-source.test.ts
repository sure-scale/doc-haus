import { afterEach, expect, test } from "bun:test"
import { createDashboardSource, createPollingDashboardSource, isJson, readBinding } from "../src/components/dashboard/visual/source"
const originalFetch = globalThis.fetch
afterEach(() => { globalThis.fetch = originalFetch })
const tick = () => new Promise((resolve) => setTimeout(resolve, 0))

test("live sources fan out updates and isolate published data", () => {
  const source = createDashboardSource()
  let first = 0, second = 0
  const stopFirst = source.subscribe(() => first++)
  const stopSecond = source.subscribe(() => second++)
  const value = { rows: [{ value: 5 }] }
  source.publish(value)
  value.rows[0]!.value = 100
  expect(source.getSnapshot().status).toBe("live")
  expect(readBinding(source.getSnapshot().data, "/rows/0/value")).toBe(5)
  expect(Object.isFrozen(source.getSnapshot().data)).toBe(true)
  expect([first, second]).toEqual([1, 1])
  expect(source.getSnapshot()).toBe(source.getSnapshot())
  stopFirst()
  source.publish({ rows: [{ value: 8 }] })
  expect([first, second]).toEqual([1, 2])
  source.fail()
  expect(source.getSnapshot().status).toBe("error")
  expect(source.getSnapshot().updatedAt).toBeDefined()
  source.publish({ rows: [] })
  expect(source.getSnapshot().status).toBe("live")
  stopSecond()
})

test("JSON pointers support nested and escaped keys without inherited access", () => {
  const data = { "a/b": { "~value": 12 }, rows: [0, null, false] }
  expect(readBinding(data, "/a~1b/~0value")).toBe(12)
  expect(readBinding(data, "/rows/0")).toBe(0)
  expect(readBinding(data, "/rows/1")).toBeNull()
  expect(readBinding(data, "/rows/2")).toBe(false)
  expect(readBinding(data, "/rows/9")).toBeUndefined()
  for (const path of ["rows", "/bad~2key", "/__proto__", "/constructor", "/rows/map"]) expect(readBinding(data, path)).toBeUndefined()
})

test("invalid updates leave the last valid snapshot intact", () => {
  const source = createDashboardSource({ count: 2 })
  const before = source.getSnapshot()
  for (const value of [NaN, Infinity, undefined, () => 1, new Date(), { count: -Infinity }, JSON.parse('{"__proto__":{"bad":true}}')]) {
    expect(isJson(value)).toBe(false)
    expect(() => source.publish(value as never)).toThrow()
    expect(source.getSnapshot()).toBe(before)
  }
})

test("polling shares requests, reports failures and aborts after the last unsubscribe", async () => {
  const signals: AbortSignal[] = []
  let calls = 0
  let resolve: ((response: Response) => void) | undefined
  globalThis.fetch = ((_: unknown, init: RequestInit) => {
    calls++
    signals.push(init.signal as AbortSignal)
    return new Promise<Response>((done) => { resolve = done })
  }) as typeof fetch
  const source = createPollingDashboardSource({ url: "/ingest/dynamic-practice", intervalMs: 60000 })
  expect(calls).toBe(0)
  const stop = source.subscribe(() => {})
  const stop2 = source.subscribe(() => {})
  expect(calls).toBe(1)
  resolve!(Response.json({ rows: [{ value: 3 }] }))
  await tick()
  expect(source.getSnapshot().status).toBe("live")
  source.refresh!()
  expect(calls).toBe(2)
  resolve!(new Response("failed", { status: 503 }))
  await tick()
  expect(source.getSnapshot().status).toBe("error")
  source.refresh!()
  expect(calls).toBe(3)
  stop()
  expect(signals[2]!.aborted).toBe(false)
  stop2()
  expect(signals[2]!.aborted).toBe(true)
  resolve!(Response.json({ rows: [{ value: 100 }] }))
  await tick()
  expect(source.getSnapshot().status).toBe("error")
})

test("polling endpoints and refresh intervals are explicit", () => {
  for (const url of ["https://example.com/feed", "//example.com/feed", "/\\example.com", "javascript:alert(1)", "/\n/example.com", "/\t/example.com"]) expect(() => createPollingDashboardSource({ url })).toThrow()
  for (const intervalMs of [0, -1, 999, NaN, Infinity]) expect(() => createPollingDashboardSource({ url: "/api/view", intervalMs })).toThrow()
})
