process.env.TZ = "Europe/London"
import { expect, test } from "bun:test"
import type { Matter } from "../src/api/ingest/matters"
import type { LeanDraftingStatus } from "../src/api/ingest/lean"
import { attention, dailyConversations, dailySpend, periodStart, progress, summarize, type MatterActivity, type Session } from "../src/components/dashboard/data"

const matter = { id: "m1", title: "Alder House · Services agreement", dir: "/fixture/alder", created_at: Date.UTC(2026, 0, 5) } as Matter
const at = (iso: string) => new Date(iso).getTime()
const session = (id: string, created: string, patch: Partial<Session> = {}) =>
  ({ id, title: "Payment terms", time: { created: at(created), updated: at(created) }, cost: 1, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, ...patch }) as Session
const activity = (sessions: Session[], drafting?: Partial<LeanDraftingStatus>): MatterActivity[] =>
  [{ matter, sessions, drafting: drafting as LeanDraftingStatus | undefined, failed: false }]

// British clocks went back on 25 October 2026, so that day is 25 hours long.
const now = at("2026-10-29T10:00:00Z")

test("each day's spend lands on its own calendar day across a clock change", () => {
  const days = dailySpend(activity([
    session("a", "2026-10-24T12:00:00+01:00", { cost: 3 }),
    session("b", "2026-10-25T23:30:00Z", { cost: 5 }),
    session("c", "2026-10-26T00:30:00Z", { cost: 7 }),
  ]), 7, now)
  expect(days.map((day) => new Date(day.day).toLocaleDateString("en-GB", { day: "numeric" }))).toEqual(["23", "24", "25", "26", "27", "28", "29"])
  expect(days.map((day) => day.cost)).toEqual([0, 3, 5, 7, 0, 0, 0])
  // Local midnight on 23 October, still summer time.
  expect(new Date(periodStart(7, now)).toISOString()).toBe("2026-10-22T23:00:00.000Z")
})

test("only the lawyer's own conversations count; runtime and subagent sessions add only spend", () => {
  const sessions = [
    session("chat", "2026-10-28T09:00:00Z"),
    session("run", "2026-10-28T09:05:00Z", { metadata: { dochaus_role: "lean-drafting-owner" } }),
    session("stage", "2026-10-28T09:06:00Z", { metadata: { dochaus_role: "lean-drafting" } }),
    session("cell", "2026-10-28T09:07:00Z", { title: "doc.haus grid: payment" }),
    session("sub", "2026-10-28T09:08:00Z", { parentID: "chat" }),
  ]
  const summary = summarize(activity(sessions), 7, now)
  expect(summary.conversations).toBe(2)
  expect(summary.spend).toBe(5)
  const first = at("2026-10-28T00:00:00Z")
  expect(dailyConversations(activity(sessions), first, 1)).toEqual([2])
})

test("a drafted document is one delivered body, however many files deliver it", () => {
  const drafting = {
    status: "completed", updated_at: "2026-10-28T12:00:00Z", owner_session_id: "run", stages: {},
    documents: { status: "completed", documents: [
      { kind: "body", document: "d1", deliverable: "Agreement.docx" },
      { kind: "redline", document: "d1", deliverable: "Agreement (redline).docx" },
      { kind: "drafting_notes", document: "d1", deliverable: "Notes.docx" },
      { kind: "combined", document: "d1", deliverable: "Agreement and notes.docx" },
      { kind: "body", document: "d2", deliverable: null },
    ] },
  } as unknown as Partial<LeanDraftingStatus>
  const summary = summarize(activity([session("run", "2026-10-28T09:00:00Z", { cost: 12 })], drafting), 30, now)
  expect(summary.documents).toBe(1)
  expect(summary.averageDocument).toBe(12)
})

test("first steps count only the lawyer's own matters and point at the newest one with documents", () => {
  const own = (id: string, created: string) => ({ ...matter, id, title: id, created_at: at(created) }) as Matter
  const rows: MatterActivity[] = [
    { matter: { ...own("sample", "2026-10-28T09:00:00Z"), sample: true } as Matter, sessions: [session("s", "2026-10-28T10:00:00Z")], documents: 4, failed: false },
    { matter: own("older", "2026-10-20T09:00:00Z"), sessions: [session("o", "2026-10-21T09:00:00Z")], documents: 2, failed: false },
    { matter: own("newer", "2026-10-27T09:00:00Z"), sessions: [], documents: 0, failed: false,
      drafting: { status: "running", owner_session_id: "run", questions: [{ status: "open" }] } as unknown as LeanDraftingStatus },
  ]
  const steps = progress(rows)
  expect([steps.matters, steps.documents, steps.conversations]).toEqual([2, 2, 1])
  expect(steps.target?.id).toBe("older")
  expect(steps.run).toEqual({ matter: rows[2].matter, session: "run", waiting: true })
})

test("a run the lawyer stopped does not ask for attention", () => {
  const base = { owner_session_id: "run", questions: [] }
  expect(attention(activity([], { ...base, status: "failed", stopped: true }))).toEqual([])
  expect(attention(activity([], { ...base, status: "failed", stopped: false })).map((item) => item.kind)).toEqual(["failed"])
})
