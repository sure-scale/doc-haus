import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { chartSourceURL, createChartEndpointSource } from "../src/api/ingest/dashboard"

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0))
test("endpoint source coalesces refreshes, reports failure and stops after unsubscribe", async () => {
  const browser = new Window({ url: "http://localhost" })
  const saved = ["window", "document", "fetch"].map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  const pending: Array<(value: Response) => void> = []
  Object.defineProperties(globalThis, {
    window: { configurable: true, value: browser }, document: { configurable: true, value: browser.document },
    fetch: { configurable: true, value: () => new Promise<Response>((resolve) => pending.push(resolve)) },
  })
  try {
    const source = createChartEndpointSource("/dashboard/sources/chart", 60_000)
    const stop = source.subscribe(() => {})
    source.refresh?.(); source.refresh?.()
    expect(pending).toHaveLength(1)
    pending.shift()!(Response.json({ data: { total: 1 }, updatedAt: 10 }))
    await tick(); await tick()
    expect(source.getSnapshot().status).toBe("loading")
    expect(pending).toHaveLength(1)
    pending.shift()!(Response.json({ data: { total: 2 }, updatedAt: 20 }))
    await tick()
    expect(source.getSnapshot()).toMatchObject({ status: "live", data: { total: 2 }, updatedAt: 20 })
    source.refresh?.()
    pending.shift()!(new Response(null, { status: 503 }))
    await tick()
    expect(source.getSnapshot()).toMatchObject({ status: "error", data: { total: 2 }, updatedAt: 20 })
    stop()
    browser.dispatchEvent(new browser.Event("focus"))
    expect(pending).toHaveLength(0)
  } finally {
    for (const [name, descriptor] of saved) if (descriptor) Object.defineProperty(globalThis, name, descriptor); else delete (globalThis as Record<string, unknown>)[name]
    browser.happyDOM.abort()
  }
})


test("legacy chart requests use the local source without changing saved settings", () => {
  const chart = { type: "chart", title: "Spend", metric: "spend", groupBy: "none", form: "stat", period: "all" } as const
  const params = new URL(chartSourceURL(chart, "month", "Europe/Lisbon")).searchParams
  expect(JSON.parse(params.get("chart")!)).toEqual({ ...chart, source: { kind: "local-query", query: "practice-chart", version: 1 } })
  expect(chart).not.toHaveProperty("source")
  expect(params.get("period")).toBe("month")
  expect(params.get("timeZone")).toBe("Europe/Lisbon")
})
