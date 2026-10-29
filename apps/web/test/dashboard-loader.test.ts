import { expect, test } from "bun:test"
import { loadDashboard } from "../src/components/dashboard/data"

test("multi-day message failure marks cost timing incomplete and recovery restores it", async () => {
  const original = globalThis.fetch
  const created = Date.parse("2026-09-30T12:00:00Z")
  const later = Date.parse("2026-10-01T12:00:00Z")
  let failing = false
  const messageRequests: string[] = []
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = new URL(input instanceof Request ? input.url : String(input))
    const path = url.pathname
    const json = (value: unknown, status = 200) => Response.json(value, { status })
    if (path === "/matters") return json([{ id: "m1", title: "Matter", dir: "/fixture", created_at: created }])
    if (path === "/matters/m1") return json({ documents: [] })
    if (path === "/matters/m1/drafting-status") return json(null)
    if (path === "/preferences") return json({})
    if (path === "/word/setup") return json({ status: "unsupported" })
    if (path === "/templates") return json({ templates: [] })
    if (path === "/playbooks/library") return json({ playbooks: [] })
    if (path === "/clients" || path === "/jurisdictions") return json([])
    if (path === "/api/session") return json({ data: [
      { id: "long", time: { created, updated: later }, cost: 3, tokens: {} },
      { id: "short", time: { created, updated: created }, cost: 1, tokens: {} },
    ], cursor: { next: null } })
    if (path.startsWith("/api/session/") && path.endsWith("/message")) {
      messageRequests.push(path)
      if (failing) return json({ error: "unavailable" }, 500)
      return json({ data: [{ type: "assistant", time: { created }, cost: 1 }, { type: "assistant", time: { created: later }, cost: 2 }], cursor: { next: null } })
    }
    throw new Error(`Unexpected loader request: ${path}`)
  }) as typeof fetch
  try {
    const good = (await loadDashboard()).matters[0]!
    expect(good.unread).toBe(false)
    expect(good.costs.map(({ time, cost }) => [time, cost])).toEqual([[created, 1], [later, 2], [created, 1]])
    failing = true
    const partial = (await loadDashboard()).matters[0]!
    expect(partial.unread).toBe(true)
    expect(partial.costs.map(({ time, cost }) => [time, cost])).toEqual([[created, 3], [created, 1]])
    failing = false
    const recovered = (await loadDashboard()).matters[0]!
    expect(recovered.unread).toBe(false)
    expect(recovered.costs).toEqual(good.costs)
    expect(messageRequests).toEqual(Array(3).fill("/api/session/long/message"))
  } finally { globalThis.fetch = original }
})
