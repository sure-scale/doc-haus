import { afterEach, expect, test } from "bun:test"
import { conflictsChanged, conflictsChip, conflictsProblems, recordConflicts } from "../src/api/ingest/conflicts"

const realFetch = globalThis.fetch
afterEach(() => { globalThis.fetch = realFetch })

function mockFetch(status: number, body: unknown) {
  const calls: { url: string; body: unknown }[] = []
  globalThis.fetch = (async (url: string, init: RequestInit) => {
    calls.push({ url, body: JSON.parse(String(init.body)) })
    return new Response(JSON.stringify(body), { status })
  }) as unknown as typeof fetch
  return calls
}

test("each state sends only the fields it needs", async () => {
  const calls = mockFetch(200, { id: "m" })
  await recordConflicts("m", { status: "cleared", by: " A. Lawyer ", owner: "stale", reason: "stale" })
  await recordConflicts("m", { status: "pending", owner: "B. Partner", due: "", by: "stale" })
  await recordConflicts("m", { status: "bypassed", by: "C", reason: "Urgent filing", owner: "stale" })
  expect(calls.map((c) => c.body)).toEqual([
    { status: "cleared", by: "A. Lawyer" },
    { status: "pending", owner: "B. Partner" },
    { status: "bypassed", by: "C", reason: "Urgent filing" },
  ])
  expect(calls[0]!.url).toEndWith("/matters/m/conflicts")
})

test("the endpoint's validation message is surfaced as returned", async () => {
  mockFetch(400, { error: 'conflicts "bypassed" requires "by" and "reason" — a documented bypass names who decided and why' })
  await expect(recordConflicts("m", { status: "bypassed", by: "C" })).rejects.toThrow('requires "by" and "reason"')
})

test("a blank required name is a field message, before any request", () => {
  expect(Object.keys(conflictsProblems({ status: "pending", owner: "  ", by: "stale" }))).toEqual(["owner"])
  expect(Object.keys(conflictsProblems({ status: "bypassed", by: "C" }))).toEqual(["reason"])
  expect(conflictsProblems({ status: "cleared", by: "A" })).toEqual({})
})

test("an edit records the check only when it changes", () => {
  const current = { status: "pending" as const, owner: "B", due: "2026-11-02", recorded_at: 1 }
  expect(conflictsChanged({ status: "pending", owner: " B ", due: "2026-11-02", by: "ignored" }, current)).toBe(false)
  expect(conflictsChanged({ status: "pending", owner: "B", due: "" }, current)).toBe(true)
  expect(conflictsChanged({ status: "cleared", by: "B" }, current)).toBe(true)
  expect(conflictsChanged({ status: "cleared", by: "B" }, undefined)).toBe(true)
})

test("the chip tone and date metadata follow the saved record", () => {
  expect(conflictsChip(undefined).tone).toBe("none")
  expect(conflictsChip({ status: "cleared", by: "A", recorded_at: Date.UTC(2026, 9, 23, 12) })).toMatchObject({ tone: "ok" })
  expect(conflictsChip({ status: "cleared", by: "A", recorded_at: Date.UTC(2026, 9, 23, 12) }).text).toContain("23 Oct 2026")
  expect(conflictsChip({ status: "pending", owner: "B", due: "2026-11-02", recorded_at: 0 }).text).toContain("2 Nov 2026")
  expect(conflictsChip({ status: "bypassed", by: "C", reason: "r", recorded_at: 0 }).tone).toBe("bypassed")
})
