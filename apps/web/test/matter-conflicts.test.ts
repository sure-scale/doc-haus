import { afterEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { MemoryRouter } from "react-router-dom"
import { conflictsChanged, conflictsChip, conflictsProblems, recordConflicts } from "../src/api/ingest/conflicts"
import type { MatterDetail } from "../src/api/ingest/matters"
import MatterHeader from "../src/components/matter/MatterHeader"

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
  expect(conflictsProblems({ status: "pending", owner: "  ", by: "stale" })).toEqual({ owner: "Name who will complete the check." })
  expect(conflictsProblems({ status: "bypassed", by: "C" })).toEqual({ reason: "Give the reason for the bypass." })
  expect(conflictsProblems({ status: "cleared", by: "A" })).toEqual({})
})

test("an edit records the check only when it changes", () => {
  const current = { status: "pending" as const, owner: "B", due: "2026-11-02", recorded_at: 1 }
  expect(conflictsChanged({ status: "pending", owner: " B ", due: "2026-11-02", by: "ignored" }, current)).toBe(false)
  expect(conflictsChanged({ status: "pending", owner: "B", due: "" }, current)).toBe(true)
  expect(conflictsChanged({ status: "cleared", by: "B" }, current)).toBe(true)
  expect(conflictsChanged({ status: "cleared", by: "B" }, undefined)).toBe(true)
})

test("the chip words follow the saved record", () => {
  expect(conflictsChip(undefined)).toEqual({ text: "Conflicts: not recorded", tone: "none" })
  expect(conflictsChip({ status: "cleared", by: "A", recorded_at: Date.UTC(2026, 9, 23, 12) })).toEqual({ text: "Conflicts: cleared by A on 23 Oct 2026", tone: "ok" })
  expect(conflictsChip({ status: "pending", owner: "B", due: "2026-11-02", recorded_at: 0 }).text).toBe("Conflicts: pending — owner B, due 2 Nov 2026")
  expect(conflictsChip({ status: "bypassed", by: "C", reason: "r", recorded_at: 0 }).text).toBe("Conflicts: bypassed by C")
})

test("the matter header shows the chip from the loaded matter", () => {
  const noop = () => {}
  const done = async () => {}
  const header = (fields: Partial<MatterDetail>) => renderToStaticMarkup(createElement(MemoryRouter, null, createElement(MatterHeader, {
    matter: { id: "m", title: "Supply note", dir: "/tmp/m", created_at: 0, documents: [], ...fields },
    jurisdictions: [],
    editingRef: false, setEditingRef: noop, draftRef: "", setDraftRef: noop, commitRef: done,
    editingClient: false, setEditingClient: noop, draftClient: "", setDraftClient: noop, commitClient: done,
    renaming: false, setRenaming: noop, draftTitle: "", setDraftTitle: noop, commitRename: done,
    patch: done, onEditConflicts: noop,
  })))
  expect(header({})).toContain("Conflicts: not recorded")
  expect(header({ conflicts: { status: "bypassed", by: "C", reason: "r", recorded_at: 0 } })).toContain("Conflicts: bypassed by C")
})
