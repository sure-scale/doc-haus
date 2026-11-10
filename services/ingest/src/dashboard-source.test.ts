import { test, expect } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { mkdirSync } from "node:fs"
import { createMatter } from "./matter"
import { openDb } from "./db"
import { openTelemetryDb, recordGeneration } from "../../../dochaus/lib/telemetry"
import { dashboardChartSource, dashboardInventory, dashboardUsage, validateDashboardSources, templateAvailability } from "./dashboard-source"
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

test("spend splits by plain agent name and background agents, with Word suggestions counted as background", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dashboard-agents-"))
  const previous = process.env.WORKSPACE_ROOT, originalFetch = globalThis.fetch
  process.env.WORKSPACE_ROOT = root
  const one = createMatter("One"), two = createMatter("Two")
  const time = Date.parse("2026-10-20T12:00:00Z")
  const session = (id: string, dir: string, agent: string, cost: number, metadata?: Record<string, unknown>) => ({ id, location: { directory: dir }, title: id, agent, cost, ...(metadata ? { metadata } : {}), tokens: { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: time, updated: time } })
  // "ask" carries 1 of title upkeep beyond its priced step; "ideas" is a Word suggestions run on the Q&A agent.
  const sessions = [session("ask", one.dir, "qa", 3), session("ideas", one.dir, "qa", 1.5, { kind: "suggestions" }), session("chron", one.dir, "chronology-assess", .5), session("narr", two.dir, "narrator", 2), session("chron2", two.dir, "chronology-assess", .25)]
  const steps: Record<string, number> = { ask: 2, ideas: 1.5, chron: .5, narr: 2, chron2: .25 }
  globalThis.fetch = Object.assign(async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(String(input))
    if (url.pathname === "/api/session") return Response.json({ data: sessions.filter(row => row.location.directory === url.searchParams.get("directory")), cursor: {} })
    const id = /^\/api\/session\/([^/]+)/.exec(url.pathname)?.[1]
    const found = sessions.find(row => row.id === id)
    if (found && url.pathname.endsWith("/message")) return Response.json({ data: [{ id: `m-${id}`, type: "assistant", time: { created: time }, cost: steps[id!], content: [] }], cursor: {} })
    if (found) return Response.json({ data: found })
    return new Response("missing", { status: 404 })
  }, { preconnect: originalFetch.preconnect }) as typeof fetch
  const base = { type: "chart", title: "Agents", metric: "spend", groupBy: "agent", form: "list", period: "all", source: { kind: "local-query", query: "practice-chart", version: 1 } }
  const read = async (patch: Record<string, unknown>) => (await dashboardChartSource(JSON.stringify({ ...base, ...patch }), "all", "UTC")).data
  const rows = (data: { points: { label: string; value: number }[] }) => data.points.map(point => [point.label, point.value])
  try {
    const agents = await read({})
    expect(agents.total).toBe(7.25)
    expect(rows(agents)).toEqual([["Narrate drafting progress", 2], ["Questions", 2], ["Suggest ideas in Word", 1.5], ["Conversation titles and compaction", 1], ["Suggest a chronology", .75]])
    expect(agents.badge).toBe("5 agents")
    const split = await read({ groupBy: "agent-group", form: "bars" })
    expect(rows(split)).toEqual([["Background agents", 4.25], ["Your work", 3]])
    expect(split.badge).toBe("")
    const background = await read({ groupBy: "matter", agentGroup: "background" })
    expect(background.total).toBe(4.25)
    expect(rows(background)).toEqual([["Two", 2.25], ["One", 2]])
    expect(background.badge).toBe("2 matters")
    expect(background.period).toBe("All time · Background agents")
    const word = await read({ matter: "One", agentGroup: "background" })
    expect(rows(word)).toEqual([["Suggest ideas in Word", 1.5], ["Suggest a chronology", .5]])
    expect(word.period).toBe("All time · One · Background agents")
    expect((await read({ agentGroup: "work" })).total).toBe(3)
    const inventory = await dashboardInventory("all", "UTC")
    expect(inventory.filters.agentGroups.background).toMatchObject({ label: "Background agents", status: "available", total: 4.25 })
    expect(inventory.filters.agentGroups.work.total).toBe(3)
    const checks = await validateDashboardSources({ widgets: [{ ...base, groupBy: "matter", agentGroup: "background" }], period: "all", timeZone: "UTC" })
    expect(checks.results[0]?.usable).toBe(true)
    await expect(read({ metric: "tokens" })).rejects.toThrow("Invalid dashboard source")
  } finally {
    globalThis.fetch = originalFetch
    if (previous === undefined) delete process.env.WORKSPACE_ROOT; else process.env.WORKSPACE_ROOT = previous
    rmSync(root, { recursive: true, force: true })
  }
})

test("built-in usage and charts share one ledger: generation rows count, dashboard building is outside every matter", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dashboard-ledger-"))
  const previous = process.env.WORKSPACE_ROOT, originalFetch = globalThis.fetch
  process.env.WORKSPACE_ROOT = root
  const one = createMatter("One"), two = createMatter("Two")
  const sample = createMatter("Sample", undefined, undefined, undefined, undefined, undefined, true)
  const dashboards = path.join(root, ".dashboard")
  mkdirSync(dashboards, { recursive: true })
  const time = Date.parse("2026-10-20T12:00:00Z")
  const tokens = { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }
  const session = (id: string, dir: string, agent: string, cost: number) => ({ id, location: { directory: dir }, title: id, agent, cost, tokens, time: { created: time, updated: time } })
  const sessions = [session("ask", one.dir, "qa", 2), session("try", sample.dir, "qa", 9), session("plan", dashboards, "dashboard", .5)]
  // A stateless job in matter Two and a deleted chart session in the dashboard folder.
  for (const [dir, agent, cost] of [[two.dir, "prepare-changes", 1], [dashboards, "dashboard-chart", .25]] as const) {
    const db = openTelemetryDb(dir)
    try { recordGeneration(db, { agent, tokens, cost, ts: time, ...(agent === "dashboard-chart" ? { session: "ses_deleted" } : {}) }) } finally { db.close() }
  }
  globalThis.fetch = Object.assign(async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(String(input))
    if (url.pathname === "/api/session") return Response.json({ data: sessions.filter(row => row.location.directory === url.searchParams.get("directory")), cursor: {} })
    const id = /^\/api\/session\/([^/]+)/.exec(url.pathname)?.[1]
    const found = sessions.find(row => row.id === id)
    if (found && url.pathname.endsWith("/message")) return Response.json({ data: [{ id: `m-${id}`, type: "assistant", time: { created: time }, cost: found.cost, content: [] }], cursor: {} })
    if (found) return Response.json({ data: found })
    return new Response("missing", { status: 404 })
  }, { preconnect: originalFetch.preconnect }) as typeof fetch
  const base = { type: "chart", title: "Spend", metric: "spend", groupBy: "none", form: "stat", period: "all", source: { kind: "local-query", query: "practice-chart", version: 1 } }
  const read = async (patch: Record<string, unknown>) => (await dashboardChartSource(JSON.stringify({ ...base, ...patch }), "all", "UTC")).data
  const rows = (data: { points: { label: string; value: number }[] }) => data.points.map(point => [point.label, point.value])
  try {
    const usage = await dashboardUsage()
    expect(usage.matters[one.id]?.costs.map(cost => [cost.cost, cost.agent])).toEqual([[2, "qa"]])
    expect(usage.matters[two.id]?.costs.map(cost => [cost.cost, cost.agent])).toEqual([[1, "prepare-changes"]])
    // The sample matter's own spend reaches its row in Recent, never a total.
    expect(usage.matters[sample.id]?.costs.map(cost => cost.cost)).toEqual([9])
    expect(usage.workspace.costs.map(cost => [cost.cost, cost.agent])).toEqual([[.5, "dashboard"], [.25, "dashboard-chart"]])
    expect(usage.workspace.unread).toBe(false)
    expect((await read({})).total).toBe(3.75)
    expect(rows(await read({ groupBy: "matter", form: "bars" }))).toEqual([["One", 2], ["Two", 1], ["Not in a matter", .75]])
    expect((await read({ groupBy: "matter", form: "bars" })).badge).toBe("2 matters")
    expect(rows(await read({ groupBy: "client", form: "bars" }))).toEqual([["No client recorded", 3], ["Not in a matter", .75]])
    expect(rows(await read({ groupBy: "agent", form: "bars" }))).toEqual([["Questions", 2], ["Prepare changes", 1], ["Build dashboards", .75]])
    expect((await read({ matter: "One" })).total).toBe(2)
    expect((await read({ matter: "Two", groupBy: "agent", form: "bars" })).total).toBe(1)
    const inventory = await dashboardInventory("all", "UTC")
    expect(inventory.metrics.matters.total).toBe(2)
    expect(inventory.metrics.spend.total).toBe(3.75)
    expect((await read({ metric: "matters" })).total).toBe(2)
    expect((await read({ metric: "conversations" })).total).toBe(1)
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


test("human titles stay in authoritative conversation charts, availability and inventory", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dashboard-title-"))
  const previous = process.env.WORKSPACE_ROOT, originalFetch = globalThis.fetch
  process.env.WORKSPACE_ROOT = root
  const matter = createMatter("Title authority")
  const now = Date.now()
  const base = { location: { directory: matter.dir }, title: "doc.haus contract questions", cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: now - 1000, updated: now } }
  const sessions = [{ ...base, id: "human" }, { ...base, id: "owner", metadata: { dochaus_role: "lean-drafting-owner" } }, { ...base, id: "job", title: "ordinary title", metadata: { dochaus_role: "review-grid" } }, { ...base, id: "child", parentID: "human" }]
  globalThis.fetch = Object.assign(async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(String(input))
    if (url.pathname === "/api/session") return Response.json({ data: sessions, cursor: {} })
    if (url.pathname.endsWith("/message")) return Response.json({ data: [], cursor: {} })
    const found = sessions.find(session => url.pathname === "/api/session/" + session.id)
    return found ? Response.json({ data: found }) : new Response(null, { status: 404 })
  }, { preconnect: originalFetch.preconnect }) as typeof fetch
  try {
    const chart = { type: "chart", title: "Conversations", metric: "conversations", groupBy: "none", form: "stat", period: "all", source: { kind: "local-query", query: "practice-chart", version: 1 } }
    expect((await dashboardChartSource(JSON.stringify(chart), "all", "UTC")).data.total).toBe(2)
    const inventory = await dashboardInventory("all", "UTC")
    expect(inventory.metrics.conversations.total).toBe(2)
    const check = await validateDashboardSources({ widgets: [{ type: "activity" }], period: "all", timeZone: "UTC" })
    expect(check.results[0]).toMatchObject({ usable: true, total: 2 })
  } finally {
    globalThis.fetch = originalFetch
    if (previous === undefined) delete process.env.WORKSPACE_ROOT; else process.env.WORKSPACE_ROOT = previous
    rmSync(root, { recursive: true, force: true })
  }
})
