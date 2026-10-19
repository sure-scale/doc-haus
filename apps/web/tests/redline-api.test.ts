import { afterAll, expect, test } from "bun:test"
import { acceptAllRedlines, acceptRedline, rejectAllRedlines, rejectRedline } from "../src/api/ingest/redlines"
import { combinedNotUpdatedNotice } from "../src/components/DocumentViewer"

// Stub fetch for this file only and restore it afterwards. Mocking the shared
// config module instead would leak a dead ingest origin into every later test
// file, because module mocks are process-global.
const requests: string[] = []
let response = Response.json({ ok: true })
const originalFetch = globalThis.fetch
globalThis.fetch = (async (input, init) => {
  const url = new URL(String(input))
  requests.push(`${init?.method ?? "GET"} ${url.pathname}${url.search}`)
  return response.clone()
}) as typeof fetch
afterAll(() => {
  globalThis.fetch = originalFetch
})

const actions = [
  { call: () => acceptRedline("matter", 7), path: "/matters/matter/redlines/7/accept", resolved: {} },
  { call: () => rejectRedline("matter", 8), path: "/matters/matter/redlines/8/reject", resolved: {} },
  { call: () => acceptAllRedlines("matter", "Source file.docx"), path: "/matters/matter/redlines/accept-all?name=Source%20file.docx", resolved: {} },
  { call: () => rejectAllRedlines("matter", "Source file.docx"), path: "/matters/matter/redlines/reject-all?name=Source%20file.docx", resolved: {} },
]

test("review resolution wrappers submit exact routes and reject server failures", async () => {
  for (const action of actions) {
    response = Response.json({ ok: true })
    expect(await action.call()).toEqual(action.resolved)
    expect(requests.at(-1)).toBe(`POST ${action.path}`)
    response = Response.json({ error: "The proposal anchor is no longer current." }, { status: 409 })
    await expect(action.call()).rejects.toThrow("The proposal anchor is no longer current.")
  }
  response = new Response("Service unavailable", { status: 503 })
  await expect(actions[0]!.call()).rejects.toThrow("Could not resolve tracked changes (503)")
})

test("accept wrappers return only the drafting-notes and redline outcomes", async () => {
  response = Response.json({ ok: true, accepted: 2, redline_updated: true, notes_updated: false, notes_skip_reason: "write_failed" })
  expect(await acceptAllRedlines("matter", "Draft.docx")).toEqual({ redline_updated: true, notes_updated: false, notes_skip_reason: "write_failed" })
  response = Response.json({ ok: true, redline_updated: false, redline_skip_reason: "rebuild_failed", notes_updated: true })
  expect(await acceptRedline("matter", 7)).toEqual({ redline_updated: false, redline_skip_reason: "rebuild_failed", notes_updated: true })
  // Each outcome is read on its own.
  response = Response.json({ ok: true, redline_updated: false, redline_skip_reason: "write_failed" })
  expect(await acceptRedline("matter", 7)).toEqual({ redline_updated: false, redline_skip_reason: "write_failed" })
})

test("reject wrappers return the combined file's outcome, so the viewer says when it was not rebuilt", async () => {
  response = Response.json({ ok: true, rejected: 1, combined_updated: false, combined_skip_reason: "write_failed" })
  for (const result of [await rejectRedline("matter", 8), await rejectAllRedlines("matter", "Draft.docx")]) {
    expect(result).toEqual({ combined_updated: false, combined_skip_reason: "write_failed" })
    expect(combinedNotUpdatedNotice(result)).toBe("The combined file was not updated: the combined file could not be saved.")
  }
  // A reject that owed no rebuild says nothing about the combined file.
  response = Response.json({ ok: true, rejected: 1 })
  expect(combinedNotUpdatedNotice(await rejectRedline("matter", 8))).toBeUndefined()
})
