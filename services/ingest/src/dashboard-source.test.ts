import { test, expect } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { createMatter } from "./matter"
import { dashboardChartSource } from "./dashboard-source"

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
  const first = Date.parse("2026-01-01T12:00:00Z"), second = Date.parse("2026-01-02T12:00:00Z")
  const session = (id: string, cost: number) => ({ id, location: { directory: matter.dir }, title: id, cost, tokens: { input: 1, output: 1, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: first, updated: second } })
  let failMessages = false
  globalThis.fetch = Object.assign(async (input: Parameters<typeof fetch>[0]) => {
    const url = new URL(String(input)), cursor = url.searchParams.get("cursor")
    if (url.pathname === "/api/session") return Response.json({ data: [cursor ? session("two", 1) : session("one", 3)], cursor: { next: cursor ? null : "next" } })
    if (url.pathname === "/api/session/one" || url.pathname === "/api/session/two") return Response.json({ data: session(url.pathname.endsWith("one") ? "one" : "two", url.pathname.endsWith("one") ? 3 : 1) })
    if (url.pathname.endsWith("/message")) {
      if (failMessages) return new Response("failed", { status: 503 })
      const one = url.pathname.includes("/one/")
      return Response.json({ data: [{ id: one ? "m1" : "m2", type: "assistant", time: { created: one ? first : second }, cost: one ? 3 : 1, content: [] }], cursor: {} })
    }
    return new Response("missing", { status: 404 })
  }, { preconnect: originalFetch.preconnect }) as typeof fetch
  const chart = { type: "chart", title: "Spend", metric: "spend", groupBy: "day", form: "line", period: "all", source: { kind: "local-query", query: "practice-chart", version: 1 } }
  try {
    const result = await dashboardChartSource(JSON.stringify(chart), "all", "UTC")
    expect(result.data.total).toBe(4)
    expect(result.data.points.find(point => point.time === Date.parse("2026-01-01T00:00:00Z"))?.value).toBe(3)
    expect(result.data.points.find(point => point.time === Date.parse("2026-01-02T00:00:00Z"))?.value).toBe(1)
    failMessages = true
    const partial = await dashboardChartSource(JSON.stringify(chart), "all", "UTC")
    expect(partial.data.basis).toContain("partial")
    expect(partial.data.notes.join(" ")).toContain("incomplete")
  } finally {
    globalThis.fetch = originalFetch
    if (previous === undefined) delete process.env.WORKSPACE_ROOT; else process.env.WORKSPACE_ROOT = previous
    rmSync(root, { recursive: true, force: true })
  }
})
