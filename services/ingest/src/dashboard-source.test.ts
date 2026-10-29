import { test, expect } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { createMatter } from "./matter"
import { openDb } from "./db"
import { dashboardChartSource, dashboardInventory, validateDashboardSources, templateAvailability } from "./dashboard-source"
import { dashboardRoutes } from "./dashboard"

test("local source reads current matters, coalesces concurrent reads, and rejects unknown mappings", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dashboard-source-"))
  const previous = process.env.WORKSPACE_ROOT, originalFetch = globalThis.fetch
  process.env.WORKSPACE_ROOT = root
  let reads = 0
  globalThis.fetch = Object.assign(async () => { reads++; await new Promise(resolve => setTimeout(resolve, 10)); return Response.json({ data: [], cursor: {} }) }, { preconnect: originalFetch.preconnect }) as typeof fetch
  const chart = { type: "chart", title: "Matters", metric: "matters", groupBy: "none", form: "stat", period: "all", source: { kind: "local-query", query: "practice-chart", version: 1 } }
  try {
    createMatter("One")
    const [first, second] = await Promise.all([dashboardChartSource(JSON.stringify(chart), "all", "UTC"), dashboardChartSource(JSON.stringify(chart), "all", "UTC")])
    expect(first.data.total).toBe(1); expect(second.data.total).toBe(1); expect(reads).toBe(1)
    createMatter("Two")
    expect((await dashboardChartSource(JSON.stringify(chart), "all", "UTC")).data.total).toBe(2)
    const before = reads
    await expect(dashboardChartSource(JSON.stringify({ ...chart, source: { kind: "local-query", query: "other", version: 1 } }), "all", "UTC")).rejects.toThrow("Invalid dashboard source")
    expect(reads).toBe(before)
    const response = await dashboardRoutes().request("/dashboard/discovery", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ action: "inventory", period: "all", timeZone: "UTC" }) })
    expect(response.status).toBe(200)
    expect((await response.json()).metrics.matters.total).toBe(2)
    const inventory = await dashboardInventory("all", "UTC")
    expect(inventory.metrics.matters.total).toBe(2)
    expect(inventory.metrics.spend.status).toBe("empty")
    expect(inventory.metrics.documents.periods.month.status).toBe("unavailable")
    expect(inventory.metrics.drafts.periods.quarter.status).toBe("unavailable")
    expect(inventory.filters.matters).toContain("One")
    const checks = await validateDashboardSources({ widgets: [chart, { ...chart, title: "Unknown", matter: "Missing" }, { type: "actions" }], period: "all", timeZone: "UTC" })
    expect(checks.results.map(result => result.usable)).toEqual([true, false, true])
    expect(checks.results[1]?.status).toBe("empty")
    await expect(validateDashboardSources({ widgets: [chart], period: "all", timeZone: "bad-zone" })).rejects.toThrow("Invalid dashboard source")
  } finally {
    globalThis.fetch = originalFetch
    if (previous === undefined) delete process.env.WORKSPACE_ROOT; else process.env.WORKSPACE_ROOT = previous
    rmSync(root, { recursive: true, force: true })
  }
})

test("native paginated sessions and messages use per-message spend, with failed reads disclosed", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dashboard-native-"))
  const previous = process.env.WORKSPACE_ROOT, originalFetch = globalThis.fetch
  process.env.WORKSPACE_ROOT = root
  const matter = createMatter("One")
  const db = openDb(matter.dir)
  try { db.run("INSERT INTO documents (doc_path, name, created_at, role) VALUES (?, ?, ?, ?)", ["source.docx", "source.docx", Date.now(), "source"]) } finally { db.close() }
  // São Paulo skipped midnight on 4 November 2018. The real source route
  // must advance through that date while attributing each message's spend.
  const first = Date.parse("2018-11-03T12:00:00Z"), second = Date.parse("2018-11-04T12:00:00Z")
  const session = (id: string, cost: number) => ({ id, location: { directory: matter.dir }, title: id, cost, tokens: { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: first, updated: second } })
  let failMessages = false
  globalThis.fetch = Object.assign(async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(String(input)), cursor = url.searchParams.get("cursor")
    if (url.pathname === "/api/session") return Response.json({ data: [cursor ? session("two", 1) : session("one", 3)], cursor: { next: cursor ? null : "next" } })
    if (url.pathname === "/api/session/one" || url.pathname === "/api/session/two") return Response.json({ data: session(url.pathname.endsWith("one") ? "one" : "two", url.pathname.endsWith("one") ? 3 : 1) })
    if (url.pathname.endsWith("/message")) {
      if (failMessages) return new Response("failed", { status: 503 })
      if (cursor && url.searchParams.has("order")) return new Response("cursor cannot have order", { status: 400 })
      if (cursor) return Response.json({ data: [], cursor: {} })
      const one = url.pathname.includes("/one/")
      return Response.json({ data: [{ id: one ? "m1" : "m2", type: "assistant", time: { created: one ? first : second }, cost: one ? 3 : 1, content: [] }], cursor: { next: "next-message" } })
    }
    return new Response("missing", { status: 404 })
  }, { preconnect: originalFetch.preconnect }) as typeof fetch
  const chart = { type: "chart", title: "Spend", metric: "spend", groupBy: "day", form: "line", period: "all", source: { kind: "local-query", query: "practice-chart", version: 1 } }
  try {
    const query = new URLSearchParams({ chart: JSON.stringify(chart), period: "all", timeZone: "America/Sao_Paulo" })
    const response = await dashboardRoutes().request(`/dashboard/sources/chart?${query}`)
    expect(response.status).toBe(200)
    const result = await response.json()
    expect(result.data.total).toBe(4)
    expect(result.data.basis).not.toContain("partial")
    const docs = { ...chart, title: "Documents", metric: "documents", groupBy: "matter", form: "bars" }
    const tokens = { ...chart, title: "Tokens", metric: "tokens", groupBy: "none", form: "stat" }
    expect((await dashboardChartSource(JSON.stringify(docs), "all", "UTC")).data.total).toBe(1)
    expect((await dashboardChartSource(JSON.stringify(tokens), "all", "UTC")).data.total).toBe(4)
    const availability = await validateDashboardSources({ widgets: [docs, tokens, { type: "actions" }, { type: "library" }], period: "all", timeZone: "UTC" })
    expect(availability.results.map(row => row.dataBacked)).toEqual([true, true, false, true])
    expect(result.data.points.find((point: { time: number }) => point.time === Date.parse("2018-11-03T03:00:00Z"))?.value).toBe(3)
    expect(result.data.points.find((point: { time: number }) => point.time === Date.parse("2018-11-04T03:00:00Z"))?.value).toBe(1)
    failMessages = true
    const partial = await dashboardChartSource(JSON.stringify(chart), "all", "UTC")
    expect(partial.data.basis).toContain("partial")
    expect(partial.data.notes.join(" ")).toContain("incomplete")
    const unaffected = await validateDashboardSources({ widgets: [docs], period: "all", timeZone: "UTC" })
    expect(unaffected.results[0]?.status).toBe("available")
    await expect(validateDashboardSources({ widgets: [{ ...docs, period: "month" }], period: "month", timeZone: "UTC" })).rejects.toThrow("Invalid dashboard source")
  } finally {
    globalThis.fetch = originalFetch
    if (previous === undefined) delete process.env.WORKSPACE_ROOT; else process.env.WORKSPACE_ROOT = previous
    rmSync(root, { recursive: true, force: true })
  }
})


test("template availability follows actual time windows and correlates draft cost to its own run", () => {
  const now = Date.now(), old = now - 400 * 86400000
  const session = (id: string, time: number, cost: number) => ({ id, title: id, cost, tokens: { input: 1, output: 1, reasoning: 0 }, time: { created: time, updated: time } })
  const row = { matter: { id: "a", title: "A", created_at: old }, sessions: [session("unrelated", old, 4)], costs: [{ time: old, cost: 4 }], unread: false, sessionUnread: false, draftingUnread: false, documentsUnread: false, unpriced: 0, drafting: { status: "completed", owner_session_id: "owner", stages: {}, documents: { status: "completed", documents: [{ kind: "body", deliverable: "body.docx" }] } } } as any
  expect(templateAvailability("week", { matters: [row] }, "all", "UTC").usable).toBe(false)
  expect(templateAvailability("activity", { matters: [row] }, "all", "UTC").usable).toBe(false)
  expect(templateAvailability("draft-cost", { matters: [row] }, "all", "UTC").usable).toBe(false)
  row.sessions.push(session("owner", now, 2))
  expect(templateAvailability("draft-cost", { matters: [row] }, "all", "UTC").usable).toBe(true)
  row.draftingUnread = true
  expect(templateAvailability("tokens", { matters: [row] }, "all", "UTC").status).toBe("available")
  expect(templateAvailability("draft-cost", { matters: [row] }, "all", "UTC").status).toBe("partial")
})


test("queued answers do not make a completed run need attention", () => {
  const row = { matter: { id: "a", title: "A", created_at: 1 }, sessions: [], costs: [], unread: false, sessionUnread: false, draftingUnread: false, documentsUnread: false, unpriced: 0, drafting: { status: "completed", stopped: false, in_flight: false, questions: [{ id: "q1", status: "open" }], queued_answer_ids: ["q1"] } } as any
  expect(templateAvailability("attention", { matters: [row] }, "all", "UTC").usable).toBe(false)
  row.drafting.queued_answer_ids = []
  expect(templateAvailability("attention", { matters: [row] }, "all", "UTC").usable).toBe(true)
})
