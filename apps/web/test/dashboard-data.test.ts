process.env.TZ = "Europe/London"
import { expect, test } from "bun:test"
import type { Matter } from "../src/api/ingest/matters"
import type { LeanDraftingStatus } from "../src/api/ingest/lean"
import {
  attention, chartStart, costByMatter, costEntries, dailyConversations, dailySpend, drafts, lastDays, periodRange, previousRange, progress, spendCsv, summarize,
  isConversation, type CostEntry, type MatterActivity, type Session,
} from "../src/components/dashboard/data"

const matter = { id: "m1", title: "Alder House · Services agreement", dir: "/fixture/alder", created_at: Date.UTC(2026, 0, 5) } as Matter
const at = (iso: string) => new Date(iso).getTime()
const iso = (time: number) => new Date(time).toISOString()
const session = (id: string, created: string, patch: Partial<Session> = {}) =>
  ({ id, title: "Payment terms", time: { created: at(created), updated: at(created) }, cost: 1, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, ...patch }) as Session
// Each session's spend on the day it started, as loading gives a single-day session.
const startCosts = (sessions: Session[]): CostEntry[] => sessions.map((row) => ({ time: row.time.created, cost: row.cost }))
const row = (sessions: Session[], patch: Partial<MatterActivity> = {}): MatterActivity =>
  ({ matter, sessions, costs: startCosts(sessions), unread: false, unpriced: 0, ...patch })
const activity = (sessions: Session[], drafting?: Partial<LeanDraftingStatus>): MatterActivity[] =>
  [row(sessions, { drafting: drafting as LeanDraftingStatus | undefined })]

// British clocks went back on 25 October 2026, so that day is 25 hours long.
const now = at("2026-10-29T10:00:00Z")

test("each day's spend lands on its own calendar day across a clock change", () => {
  const range = lastDays(7, now)
  const days = dailySpend(activity([
    session("a", "2026-10-24T12:00:00+01:00", { cost: 3 }),
    session("b", "2026-10-25T23:30:00Z", { cost: 5 }),
    session("c", "2026-10-26T00:30:00Z", { cost: 7 }),
  ]), range)
  expect(days.map((day) => new Date(day.day).toLocaleDateString("en-GB", { day: "numeric" }))).toEqual(["23", "24", "25", "26", "27", "28", "29"])
  expect(days.map((day) => day.cost)).toEqual([0, 3, 5, 7, 0, 0, 0])
  // Local midnight on 23 October, still summer time, to midnight after today in winter time.
  expect([iso(range.start), iso(range.end)]).toEqual(["2026-10-22T23:00:00.000Z", "2026-10-30T00:00:00.000Z"])
})

test("ledger spend lands on each entry's own day and month", () => {
  const long = session("long", "2026-09-30T16:00:00+01:00", { cost: 10, time: { created: at("2026-09-30T16:00:00+01:00"), updated: at("2026-10-02T09:00:00+01:00") } })
  // As ingest's ledger reports it: each priced call, then the leftover upkeep.
  const costs = [
    { time: at("2026-09-30T16:01:00+01:00"), cost: 3, agent: "qa" },
    { time: at("2026-10-01T10:00:00+01:00"), cost: 4, agent: "qa" },
    { time: at("2026-10-02T09:00:00+01:00"), cost: 3, agent: "conversation-upkeep" },
  ]
  const matters = [row([long], { costs })]
  const october = at("2026-10-15T12:00:00+01:00")
  expect(summarize(matters, "month", october)).toMatchObject({ spend: 7, previous: 3 })
  expect(summarize(matters, "last-month", october).spend).toBe(3)
  const days = dailySpend(matters, periodRange("month", october))
  expect(days.slice(0, 3).map((day) => day.cost)).toEqual([4, 3, 0])
  expect(days[0].matters).toEqual([{ title: matter.title, cost: 4 }])
})

test("spend outside any matter counts in totals, trends and the export, never as a matter", () => {
  const matters = [row([session("o", "2026-10-28T09:00:00Z", { cost: 2 })])]
  // A deleted chart session's generation row and a live planner session.
  const workspace = { costs: [{ time: at("2026-10-27T09:00:00Z"), cost: 0.5, agent: "dashboard" }, { time: at("2026-09-02T09:00:00Z"), cost: 0.25, agent: "dashboard-chart" }], unread: false, unpriced: 0 }
  expect(summarize(matters, "month", now, workspace)).toMatchObject({ spend: 2.5, previous: 0.25, conversations: 1, activeMatters: 1 })
  expect(costByMatter(matters, periodRange("month", now)).map((item) => item.cost)).toEqual([2])
  const days = dailySpend(matters, lastDays(3, now), workspace)
  expect(days.map((day) => day.cost)).toEqual([0.5, 2, 0])
  expect(days[0].matters).toEqual([{ title: "Not in a matter", cost: 0.5 }])
  expect(costEntries(matters, workspace).reduce((sum, entry) => sum + entry.value, 0)).toBe(2.75)
  expect(iso(chartStart(matters, periodRange("all", now), workspace))).toBe("2026-09-01T23:00:00.000Z")
  expect(spendCsv(costByMatter(matters, periodRange("month", now)), 0.5).split("\r\n").slice(-2)).toEqual([",Not in a matter,,,0.50", ""])
  expect(summarize(matters, "month", now, { ...workspace, unpriced: 1 }).unpriced).toBe(1)
})

test("the fictional sample matter adds nothing to spend", () => {
  const matters = [
    { ...row([session("s", "2026-10-28T09:00:00Z", { cost: 50 })]), matter: { ...matter, id: "sample", sample: true } as Matter },
    row([session("o", "2026-10-28T09:00:00Z", { cost: 2 })]),
  ]
  expect(summarize(matters, "month", now)).toMatchObject({ spend: 2, conversations: 1, activeMatters: 1 })
  expect(costByMatter(matters, periodRange("month", now)).map((item) => item.matter.id)).toEqual(["m1"])
})

test("only the lawyer's own conversations count; runtime and subagent sessions add only spend", () => {
  const sessions = [
    session("chat", "2026-10-28T09:00:00Z"),
    session("run", "2026-10-28T09:05:00Z", { metadata: { dochaus_role: "lean-drafting-owner" } }),
    session("stage", "2026-10-28T09:06:00Z", { metadata: { dochaus_role: "lean-drafting" } }),
    session("cell", "2026-10-28T09:07:00Z", { title: "doc.haus grid: payment", metadata: { dochaus_role: "review-grid" } }),
    session("sub", "2026-10-28T09:08:00Z", { parentID: "chat" }),
  ]
  const summary = summarize(activity(sessions), "month", now)
  expect(summary.conversations).toBe(2)
  expect(summary.spend).toBe(5)
  const first = at("2026-10-28T00:00:00Z")
  expect(dailyConversations(activity(sessions), first, 1)).toEqual([2])
})

test("periods follow calendar months, and January looks back into December", () => {
  const range = (value: { start: number; end: number } | undefined) => value && [iso(value.start), iso(value.end)]
  const january = at("2027-01-20T15:00:00Z")
  expect(range(periodRange("month", january))).toEqual(["2027-01-01T00:00:00.000Z", "2027-01-21T00:00:00.000Z"])
  expect(range(previousRange("month", january))).toEqual(["2026-12-01T00:00:00.000Z", "2027-01-01T00:00:00.000Z"])
  expect(range(periodRange("last-month", january))).toEqual(["2026-12-01T00:00:00.000Z", "2027-01-01T00:00:00.000Z"])
  expect(range(previousRange("last-month", january))).toEqual(["2026-11-01T00:00:00.000Z", "2026-12-01T00:00:00.000Z"])
  expect(range(periodRange("quarter", january))).toEqual(["2026-11-01T00:00:00.000Z", "2027-01-21T00:00:00.000Z"])
  // Local midnight on 1 August is in summer time.
  expect(range(previousRange("quarter", january))).toEqual(["2026-07-31T23:00:00.000Z", "2026-11-01T00:00:00.000Z"])
  expect(periodRange("all", january).start).toBe(-Infinity)
  expect(previousRange("all", january)).toBeUndefined()
  // Local midnight on 1 October is still summer time.
  expect(iso(periodRange("month", now).start)).toBe("2026-09-30T23:00:00.000Z")
})

test("a draft changed after the final check is never shown as ready", () => {
  const completed = (patch: Partial<LeanDraftingStatus>) => ({
    status: "completed", updated_at: "2026-10-28T12:00:00Z", owner_session_id: "run", stages: {}, open_items: [],
    gate: { status: "pass", delivery: "execution_ready", blocking: [] },
    documents: { status: "completed", documents: [
      { kind: "body", document: "d1", title: "Services agreement", deliverable: "Services agreement.docx" },
      { kind: "redline", document: "d1", title: null, deliverable: "Services agreement - redline.docx" },
      { kind: "body", document: "d2", title: null, deliverable: null },
    ] },
    ...patch,
  }) as unknown as Partial<LeanDraftingStatus>
  const state = (patch: Partial<LeanDraftingStatus>) => drafts(activity([session("run", "2026-10-28T09:00:00Z", { cost: 12 })], completed(patch)))
  expect(state({})).toMatchObject([{ state: "ready", titles: ["Services agreement"], cost: 12 }])
  expect(state({ changed_after_check: { since: "2026-10-28T12:00:00Z", documents: ["Services agreement.docx"], missing: [] } })[0].state).toBe("changed")
})

test("first steps count only the lawyer's own matters and point at the newest one with documents", () => {
  const own = (id: string, created: string) => ({ ...matter, id, title: id, created_at: at(created) }) as Matter
  const rows: MatterActivity[] = [
    { ...row([session("s", "2026-10-28T10:00:00Z")]), matter: { ...own("sample", "2026-10-28T09:00:00Z"), sample: true } as Matter, documents: 4 },
    { ...row([session("o", "2026-10-21T09:00:00Z")]), matter: own("older", "2026-10-20T09:00:00Z"), documents: 2 },
    { ...row([]), matter: own("newer", "2026-10-27T09:00:00Z"), documents: 0,
      drafting: { status: "running", owner_session_id: "run", questions: [{ status: "open" }] } as unknown as LeanDraftingStatus },
  ]
  const steps = progress(rows)
  expect([steps.matters, steps.documents, steps.conversations]).toEqual([2, 2, 1])
  expect(steps.target?.id).toBe("older")
  expect(steps.run).toEqual({ matter: rows[2].matter, session: "run", waiting: true })
})

test("a stopped run asks for nothing; a running one nothing is working on asks to be continued", () => {
  const base = { owner_session_id: "run", questions: [] }
  expect(attention(activity([], { ...base, status: "failed", stopped: true }))).toEqual([])
  expect(attention(activity([], { ...base, status: "failed", stopped: false })).map((item) => item.kind)).toEqual(["failed"])
  expect(attention(activity([], { ...base, status: "running", in_flight: true }))).toEqual([])
  expect(attention(activity([], { ...base, status: "running", in_flight: false })).map((item) => item.kind)).toEqual(["stalled"])
})

test("a run stopped on credit asks for credit, not a reconnect", () => {
  const base = { owner_session_id: "run", questions: [] }
  expect(attention(activity([], { ...base, status: "failed", provider_stop: "credit" })).map((item) => item.kind)).toEqual(["credit"])
  expect(attention(activity([], { ...base, status: "failed", provider_stop: "auth" })).map((item) => item.kind)).toEqual(["provider"])
})

test("the spend export quotes commas and quotes and defuses formula cells", () => {
  const titled = { ...matter, title: 'Alder, "Phase 2"', client: "Alder Ltd", reference: "AL-7" } as Matter
  const csv = spendCsv([
    { matter: { ...matter, id: "none", title: "No client" }, cost: 9, conversations: 1 },
    { matter: titled, cost: 1.5, conversations: 2 },
    { matter: { ...matter, id: "formula", title: "=1+1", client: "@Birch" }, cost: 0.5, conversations: 1 },
  ])
  expect(csv.split("\r\n")).toEqual([
    "Client,Matter,Reference,Conversations,Spend (USD)",
    "'@Birch,'=1+1,,1,0.50",
    'Alder Ltd,"Alder, ""Phase 2""",AL-7,2,1.50',
    ",No client,,1,9.00",
    "",
  ])
})


test("conversation titles never classify user content as runtime ownership", () => {
  for (const title of ["doc.haus contract review", "doc.haus grid: payment", "Ordinary conversation"]) {
    expect(isConversation(session("human", "2026-10-28T09:00:00Z", { title }))).toBe(true)
    expect(isConversation(session("system", "2026-10-28T09:00:00Z", { title, metadata: { dochaus_role: "review-grid" } }))).toBe(false)
  }
})
