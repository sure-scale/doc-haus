import { afterEach, expect, test } from "bun:test"
import { detectInjection } from "../lib/untrusted-text"
import caseLaw from "../tool/case-law"

// The search is answered locally: no request leaves the test.
const originalFetch = globalThis.fetch
const originalRoot = process.env.WORKSPACE_ROOT
const originalToken = process.env.COURTLISTENER_API_TOKEN
afterEach(() => {
  globalThis.fetch = originalFetch
  if (originalRoot === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = originalRoot
  if (originalToken === undefined) delete process.env.COURTLISTENER_API_TOKEN
  else process.env.COURTLISTENER_API_TOKEN = originalToken
})

async function search(status: number) {
  delete process.env.WORKSPACE_ROOT
  delete process.env.COURTLISTENER_API_TOKEN
  globalThis.fetch = (async () => new Response("{}", { status })) as unknown as typeof fetch
  return String(await caseLaw.execute({ query: "notice periods" }, {} as never))
}

test("a search refused for lack of a token says a token is needed and nothing was searched", async () => {
  for (const status of [401, 403]) {
    const reply = await search(status)
    // Was "The case-law database may be unreachable": an outage, not a missing token.
    expect(reply).toContain("needs a CourtListener API token, added under Settings > Research")
    expect(reply).toContain("No case law was searched. Disclose this as a research limitation")
    expect(reply).not.toContain("unreachable")
    expect(detectInjection(reply)).toEqual([])
  }
})

test("any other failed search is reported as unreachable, with nothing searched", async () => {
  const reply = await search(503)
  expect(reply).toContain("CourtListener search failed (HTTP 503); the service may be unreachable. No case law was searched.")
  expect(detectInjection(reply)).toEqual([])
})
