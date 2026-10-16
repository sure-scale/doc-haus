import { afterAll, expect, test } from "bun:test"
import { acceptAllRedlines, acceptRedline, rejectAllRedlines, rejectRedline } from "../src/api/ingest/redlines"

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
  { call: () => acceptRedline("matter", 7), path: "/matters/matter/redlines/7/accept" },
  { call: () => rejectRedline("matter", 8), path: "/matters/matter/redlines/8/reject" },
  { call: () => acceptAllRedlines("matter", "Source file.docx"), path: "/matters/matter/redlines/accept-all?name=Source%20file.docx" },
  { call: () => rejectAllRedlines("matter", "Source file.docx"), path: "/matters/matter/redlines/reject-all?name=Source%20file.docx" },
]

test("review resolution wrappers submit exact routes and reject server failures", async () => {
  for (const action of actions) {
    response = Response.json({ ok: true })
    await expect(action.call()).resolves.toBeUndefined()
    expect(requests.at(-1)).toBe(`POST ${action.path}`)
    response = Response.json({ error: "The proposal anchor is no longer current." }, { status: 409 })
    await expect(action.call()).rejects.toThrow("The proposal anchor is no longer current.")
  }
  response = new Response("Service unavailable", { status: 503 })
  await expect(actions[0]!.call()).rejects.toThrow("Could not resolve tracked changes (503)")
})
