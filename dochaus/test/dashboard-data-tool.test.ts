import { expect, test } from "bun:test"
import dashboardData from "../tool/dashboard-data"

const context = { abort: AbortSignal.timeout(5000) } as never
const chart = { type: "chart", title: "Documents", metric: "documents", groupBy: "client", form: "bars", period: "all" }
test("native discovery sends only checked app-owned queries and records evidence", async () => {
  const original = globalThis.fetch
  const calls: Array<{ url: string; body: Record<string, unknown> }> = []
  globalThis.fetch = (async (url, init) => {
    calls.push({ url: String(url), body: JSON.parse(String(init?.body)) })
    return Response.json({ updatedAt: 1, results: [] })
  }) as typeof fetch
  try {
    await dashboardData.execute({ action: "inventory", period: "month", timeZone: "UTC" }, context)
    const preview = await dashboardData.execute({ action: "preview", period: "month", timeZone: "UTC", widgets: [chart] }, context)
    expect(calls.map((call) => new URL(call.url).pathname)).toEqual(["/dashboard/discovery", "/dashboard/discovery"])
    expect(calls[1]?.body).toMatchObject({ action: "preview", widgets: [chart], period: "month", timeZone: "UTC" })
    expect(preview.metadata).toMatchObject({ action: "preview", widgets: [chart] })
    await expect(dashboardData.execute({ action: "preview", period: "month", timeZone: "UTC", widgets: [{ ...chart, period: "month" }] }, context)).rejects.toThrow("no reliable dates")
    expect(calls).toHaveLength(2)
  } finally { globalThis.fetch = original }
})
