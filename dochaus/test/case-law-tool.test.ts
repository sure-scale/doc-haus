import { expect, test } from "bun:test"
import { detectInjection } from "../lib/untrusted-text"
import caseLaw from "../tool/case-law"

// The search is answered locally: no request leaves the test.
async function search(status: number) {
  const originalFetch = globalThis.fetch
  const originalRoot = process.env.WORKSPACE_ROOT
  const originalToken = process.env.COURTLISTENER_API_TOKEN
  delete process.env.WORKSPACE_ROOT
  delete process.env.COURTLISTENER_API_TOKEN
  globalThis.fetch = (async () => new Response("{}", { status })) as unknown as typeof fetch
  try {
    return String(await caseLaw.execute({ query: "notice periods" }, {} as never))
  } finally {
    globalThis.fetch = originalFetch
    if (originalRoot === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = originalRoot
    if (originalToken === undefined) delete process.env.COURTLISTENER_API_TOKEN
    else process.env.COURTLISTENER_API_TOKEN = originalToken
  }
}

test("a search refused for lack of a token reports missing credentials rather than an outage", async () => {
  for (const status of [401, 403]) {
    const reply = await search(status)
    expect(reply).toMatch(/API token/i)
    expect(reply).toMatch(/no case law was searched/i)
    expect(reply).not.toMatch(/unreachable/i)
    expect(detectInjection(reply)).toEqual([])
  }
})

test("any other failed search reports the status and unreachability with nothing searched", async () => {
  const reply = await search(503)
  expect(reply).toContain("503")
  expect(reply).toMatch(/unreachable/i)
  expect(reply).toMatch(/no case law was searched/i)
  expect(reply).not.toMatch(/API token/i)
  expect(detectInjection(reply)).toEqual([])
})
