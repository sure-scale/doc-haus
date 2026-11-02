import { expect, test } from "bun:test"
import { loadDashboard } from "../src/components/dashboard/data"

test("spend comes from ingest's ledger; a failed read is partial, never zero", async () => {
  const original = globalThis.fetch
  const created = Date.parse("2026-09-30T12:00:00Z")
  const later = Date.parse("2026-10-01T12:00:00Z")
  let usage: "ok" | "failing" | "missing" = "ok"
  const usageBodies: unknown[] = []
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
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
    if (path === "/api/session") return json({ data: [{ id: "long", time: { created, updated: later }, cost: 3, tokens: {} }], cursor: { next: null } })
    if (path === "/dashboard/discovery") {
      usageBodies.push(JSON.parse(String(init?.body)))
      if (usage === "failing") return json({ error: "Could not inspect dashboard data." }, 503)
      // Two priced calls of a live session plus a stateless generation job.
      const matters = usage === "missing" ? {} : { m1: { costs: [{ time: created, cost: 1, agent: "qa" }, { time: later, cost: 2, agent: "qa" }, { time: later, cost: 0.5, agent: "prepare-changes" }], unread: false, unpriced: 0 } }
      return json({ updatedAt: later, matters, workspace: { costs: [{ time: later, cost: 0.25, agent: "dashboard" }], unread: false, unpriced: 0 } })
    }
    throw new Error(`Unexpected loader request: ${path}`)
  }) as typeof fetch
  try {
    const good = await loadDashboard()
    expect(good.matters[0]!.unread).toBe(false)
    expect(good.matters[0]!.costs.map(({ time, cost }) => [time, cost])).toEqual([[created, 1], [later, 2], [later, 0.5]])
    expect(good.workspace).toEqual({ costs: [{ time: later, cost: 0.25, agent: "dashboard" }], unread: false, unpriced: 0 })
    expect(good.matters).toHaveLength(1)
    usage = "failing"
    const partial = await loadDashboard()
    expect(partial.matters[0]!.unread).toBe(true)
    expect(partial.matters[0]!.costs).toEqual([])
    expect(partial.matters[0]!.sessions).toHaveLength(1)
    expect(partial.workspace?.unread).toBe(true)
    // A matter ingest has not read yet is unknown, not free.
    usage = "missing"
    expect((await loadDashboard()).matters[0]!.unread).toBe(true)
    usage = "ok"
    expect(await loadDashboard()).toMatchObject({ matters: [{ unread: false, costs: good.matters[0]!.costs }] })
    expect(usageBodies).toEqual(Array(4).fill({ action: "usage" }))
  } finally { globalThis.fetch = original }
})
