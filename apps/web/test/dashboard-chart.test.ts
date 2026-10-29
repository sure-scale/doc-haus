process.env.TZ = "Europe/London"
import { expect, test } from "bun:test"
import type { Chart } from "../../../dochaus/lib/dashboard-plan"
import type { Matter } from "../src/api/ingest/matters"
import type { LeanDraftingStatus } from "../src/api/ingest/lean"
import { chartData } from "../src/components/dashboard/chart-data"
import type { DashboardData, MatterActivity, Session } from "../src/components/dashboard/data"

const at = (iso: string) => new Date(iso).getTime()
const now = at("2026-10-29T12:00:00Z")
const chart = (patch: Partial<Chart> = {}): Chart => ({ type: "chart", title: "Spend", metric: "spend", groupBy: "client", form: "bars", period: "dashboard", ...patch })
const session = (id: string, patch: Partial<Session> = {}): Session => ({ id, title: "Payment terms", cost: 4, time: { created: at("2026-10-28T09:00:00Z"), updated: now }, tokens: { input: 10, output: 5, reasoning: 0, cache: { read: 0, write: 0 } }, ...patch }) as Session
const row = (id: string, patch: Partial<MatterActivity> = {}): MatterActivity => ({ matter: { id, title: id, dir: `/fictional/${id}`, client: "Alder House", created_at: at("2026-09-25T12:00:00Z") } as Matter, sessions: [], costs: [], unread: false, unpriced: 0, ...patch })
const data = (...matters: MatterActivity[]): DashboardData => ({ matters })

test("AI charts count priced calls when spent, respect filters, omit sample data, and disclose unpriced activity", () => {
  const a = row("a", { costs: [{ session: "call", cost: 7, time: at("2026-10-25T23:30:00Z") }, { session: "call", cost: 5, time: at("2026-09-30T12:00:00Z") }], unpriced: 1 })
  const b = row("b", { matter: { ...row("b").matter, client: "Birch Studio" }, costs: [{ session: "b", cost: 3, time: now }] })
  const sample = row("sample", { matter: { ...row("sample").matter, sample: true }, costs: [{ session: "sample", cost: 100, time: now }] })
  const view = chartData(chart({ client: "Alder House" }), data(a, b, sample), "month", now)
  expect(view.total).toBe(7)
  expect(view.points).toEqual([{ key: "Alder House", label: "Alder House", value: 7 }])
  expect(view.notes.join(" ")).toContain("no price")
  expect(chartData(chart({ period: "last-month" }), data(a, b), "month", now).total).toBe(5)
  expect(chartData(chart({ client: "Unknown" }), data(a), "month", now).empty).toBe("No matters found for Unknown.")
})

test("time charts fill local calendar buckets across clock changes, and conversations exclude runtime sessions", () => {
  const entries = [session("first", { time: { created: at("2026-10-24T23:30:00+01:00"), updated: now } }), session("second", { time: { created: at("2026-10-25T23:30:00Z"), updated: now } }), session("hidden", { metadata: { dochaus_role: "dashboard" } }), session("child", { parentID: "first" })]
  const view = chartData(chart({ metric: "conversations", groupBy: "day", form: "line" }), data(row("a", { sessions: entries })), "month", now)
  expect(view.total).toBe(2)
  expect(view.points).toHaveLength(29)
  expect(view.points.filter((point) => point.value).map(({ label, value }) => [label, value])).toEqual([["24 Oct", 1], ["25 Oct", 1]])
  const weeks = chartData(chart({ metric: "conversations", groupBy: "week", form: "line" }), data(row("a", { sessions: entries })), "month", now)
  expect(weeks.points.filter((point) => point.value).map(({ label, value }) => [label, value])).toEqual([["Week of 19 Oct", 2]])
})

test("all-time matter charts start at the first matter, even when it predates spend or chat", () => {
  const view = chartData(chart({ metric: "matters", period: "all", groupBy: "month", form: "line" }), data(row("old")), "month", now)
  expect(view.total).toBe(1)
  expect(view.points.map(({ label, value }) => [label, value])).toEqual([["Sep 2026", 1], ["Oct 2026", 0]])
  expect(chartData(chart({ metric: "matters", groupBy: "none", form: "stat" }), data(row("old")), "month", now).total).toBe(0)
})

test("draft totals count only delivered bodies from finished latest runs and never invent dates", () => {
  const drafting = { status: "completed", updated_at: "2026-10-28T12:00:00Z", documents: { status: "completed", documents: [{ kind: "body", deliverable: "body.docx" }, { kind: "redline", deliverable: "redline.docx" }, { kind: "body", deliverable: null }] } } as unknown as LeanDraftingStatus
  const rows = data(row("done", { drafting }), row("running", { drafting: { ...drafting, status: "running" } }))
  const view = chartData(chart({ metric: "drafts", groupBy: "none", form: "stat", period: "last-month" }), rows, "month", now)
  expect(view.total).toBe(1)
  expect(view.period).toContain("all time")
  expect(view.notes.join(" ")).toContain("Draft dates are not tracked")
  const invalid = chartData(chart({ metric: "drafts", groupBy: "day", form: "line" }), rows, "month", now)
  expect(invalid.unavailable).toBe(true)
  expect(invalid.points).toEqual([])
})

test("rank limits preserve the full total and missing activity never looks like a complete empty record", () => {
  const rows = [row("a", { sessions: [session("1"), session("2")] }), row("b", { sessions: [session("3")], unread: true })]
  const view = chartData(chart({ metric: "conversations", groupBy: "matter", limit: 1 }), data(...rows), "month", now)
  expect(view.points.map(({ label, value }) => [label, value])).toEqual([["a", 2]])
  expect(view.total).toBe(3)
  expect(view.notes.join(" ")).toContain("top 1 of 2 matters")
  expect(view.incomplete).toBe(true)
  const failed = chartData(chart(), { matters: [], error: "Could not load matters" }, "month", now)
  expect(failed.unavailable).toBe(true)
  expect(failed.empty).toBe("These figures could not be loaded.")
  const repeated = chartData(chart({ matter: "a", groupBy: "none", form: "stat" }), data(row("a"), row("b", { matter: { ...row("b").matter, title: "a" } })), "month", now)
  expect(repeated.unavailable).toBe(true)
})

test("calendar buckets follow the requested time zone across daylight saving", () => {
  const view = chartData(chart({ metric: "matters", groupBy: "day", form: "line", period: "month" }), data(row("dst", { matter: { ...row("dst").matter, created_at: Date.parse("2026-03-29T23:30:00Z") } })), "month", Date.parse("2026-03-31T12:00:00Z"), "Europe/London")
  expect(view.points.find(point => point.label === "30 Mar")?.value).toBe(1)
  const before = view.points.find(point => point.label === "29 Mar")!
  const after = view.points.find(point => point.label === "30 Mar")!
  expect(after.time! - before.time!).toBe(23 * 60 * 60 * 1000)
})

test("civil day buckets progress through midnight gaps, repeats and skipped dates", () => {
  const cases = [
    { zone: "America/Sao_Paulo", created: "2018-11-02T12:00:00Z", end: "2018-11-06T12:00:00Z", labels: ["2 Nov", "3 Nov", "4 Nov", "5 Nov", "6 Nov"], transition: "4 Nov", boundary: "2018-11-04T03:00:00Z" },
    { zone: "America/Havana", created: "2020-10-31T12:00:00Z", end: "2020-11-02T12:00:00Z", labels: ["31 Oct", "1 Nov", "2 Nov"], transition: "1 Nov", boundary: "2020-11-01T04:00:00Z" },
    { zone: "Pacific/Apia", created: "2011-12-29T12:00:00Z", end: "2011-12-31T12:00:00Z", labels: ["29 Dec", "31 Dec", "1 Jan"], transition: "31 Dec", boundary: "2011-12-30T10:00:00Z" },
  ]
  for (const item of cases) {
    const view = chartData(chart({ metric: "matters", groupBy: "day", form: "line", period: "all" }), data(row("civil", { matter: { ...row("civil").matter, created_at: at(item.created) } })), "all", at(item.end), item.zone)
    expect(view.unavailable).toBe(false)
    expect(view.points.map(point => point.label)).toEqual(item.labels)
    expect(view.points.find(point => point.label === item.transition)?.time).toBe(at(item.boundary))
    expect(view.points.reduce((sum, point) => sum + point.value, 0)).toBe(view.total)
    for (const groupBy of ["week", "month"] as const) {
      const grouped = chartData(chart({ metric: "matters", groupBy, form: "line", period: "all" }), data(row("civil", { matter: { ...row("civil").matter, created_at: at(item.created) } })), "all", at(item.end), item.zone)
      expect(grouped.unavailable).toBe(false)
      expect(grouped.points.reduce((sum, point) => sum + point.value, 0)).toBe(grouped.total)
    }
  }
})


test("both occurrences of a repeated midnight belong to the same earliest civil-day bucket", () => {
  const records = ["2020-11-01T04:30:00Z", "2020-11-01T05:30:00Z"].map((time, i) => row(String(i), { matter: { ...row(String(i)).matter, created_at: at(time) } }))
  const view = chartData(chart({ metric: "matters", groupBy: "day", form: "line", period: "all" }), data(...records), "all", at("2020-11-02T12:00:00Z"), "America/Havana")
  expect(view.points[0]).toMatchObject({ label: "1 Nov", time: at("2020-11-01T04:00:00Z"), value: 2 })
  expect(view.points.reduce((sum, point) => sum + point.value, 0)).toBe(2)
})
