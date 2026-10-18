import { expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { fetchExaSourceText } from "../lib/research-source-exa"
import { fetchSourceText } from "../lib/research-source-fetch"
import { preserveResearchSourceSnapshot, readResearchSourceSnapshot } from "../lib/research-source-snapshot"
import { displayedSourceSnippets } from "../lib/research-source-snippets"

const url = "https://records.gov/record"
const text = "The public record contains the following statement."
const payload = () => ({ statuses: [{ id: url, status: "success", source: "crawled" }], results: [{ url, text }] })

test("Exa fallback rereads the app's saved research key on every call", async () => {
  const previous = process.env.WORKSPACE_ROOT
  const directory = mkdtempSync(path.join(tmpdir(), "research-exa-settings-"))
  process.env.WORKSPACE_ROOT = directory
  mkdirSync(path.join(directory, ".preferences"))
  try {
    for (const key of ["first-saved-key", "updated-saved-key"]) {
      writeFileSync(path.join(directory, ".preferences", "preferences.json"), JSON.stringify({ searchApiKey: key }))
      const result = await fetchExaSourceText(url, { fetch: async (_endpoint, init) => {
        expect(new Headers(init?.headers).get("x-api-key")).toBe(key)
        return Response.json(payload())
      } })
      expect(result?.text).toBe(text)
    }
  } finally {
    if (previous === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previous
    rmSync(directory, { recursive: true, force: true })
  }
})

test("Exa retrieval requests only exact URL text with fresh crawl and retains transport provenance", async () => {
  const result = await fetchExaSourceText(`${url}#section`, { key: "test-only", fetch: async (endpoint, init) => {
    expect(endpoint).toBe("https://api.exa.ai/contents")
    expect(init?.redirect).toBe("error")
    expect(JSON.parse(String(init?.body))).toEqual({ urls: [url], text: true, maxAgeHours: 0, livecrawlTimeout: 15000 })
    expect(new Headers(init?.headers).get("x-api-key")).toBe("test-only")
    return Response.json(payload())
  } })
  expect(result?.text).toBe(text)
  expect(result?.url).toBe(url)
  expect(result?.provenance).toMatchObject({ provider: "exa", requestedFreshness: "live", statusSource: "crawled" })
  expect(result?.provenance.responseHash).toMatch(/^[a-f0-9]{64}$/)
})

test("Exa cannot replace missing source text with cached, mismatched, multiple or generated results", async () => {
  const cases = [
    { ...payload(), statuses: [{ id: url, status: "success", source: "cached" }] },
    { ...payload(), statuses: [{ id: url, status: "error", source: "crawled" }] },
    { ...payload(), statuses: [{ id: `${url}/other`, status: "success", source: "crawled" }] },
    { ...payload(), results: [{ url: "https://other.gov/record", text }] },
    { ...payload(), results: [{ url, summary: text, highlights: [text] }] },
    { ...payload(), results: [{ url, text: " " }] },
    { ...payload(), results: [payload().results[0], payload().results[0]] },
    {},
  ]
  for (const value of cases)
    expect(await fetchExaSourceText(url, { key: "test-only", fetch: async () => Response.json(value) })).toBeNull()
  for (const target of ["https://example.com/record", "https://user:password@records.gov/record", "http://records.gov/record", "https://127.0.0.1/record"])
    expect(await fetchExaSourceText(target, { key: "test-only", fetch: async () => { throw new Error("must not request") } })).toBeNull()
})

test("Exa failures and response limits fail closed without exposing provider errors", async () => {
  for (const response of [new Response("provider details", { status: 429 }), new Response("bad json"),
    new Response("small", { headers: { "content-length": String(6 * 1024 * 1024) } }), new Response("x".repeat(5 * 1024 * 1024 + 1))])
    expect(await fetchExaSourceText(url, { key: "test-only", fetch: async () => response })).toBeNull()
  expect(await fetchExaSourceText(url, { key: "test-only", fetch: async () => { throw new Error("secret provider detail") } })).toBeNull()
})

test("shared fallback preserves extracted text without claiming original bytes or full-source verification", async () => {
  const previous = process.env.EXA_API_KEY
  const previousRoot = process.env.WORKSPACE_ROOT
  const directory = mkdtempSync(path.join(tmpdir(), "research-exa-"))
  process.env.EXA_API_KEY = "test-only"
  // Web research defaults to the open web; the redirect fence below needs the official-sources scope.
  process.env.WORKSPACE_ROOT = directory
  mkdirSync(path.join(directory, ".preferences"))
  writeFileSync(path.join(directory, ".preferences", "preferences.json"), JSON.stringify({ webResearch: "official" }))
  try {
    const calls: string[] = []
    const result = await fetchSourceText(url, { browserFallback: false, fetch: async endpoint => {
      calls.push(String(endpoint))
      return String(endpoint) === url ? new Response("unavailable", { status: 503 }) : Response.json(payload())
    } })
    expect(calls).toEqual([url, "https://api.exa.ai/contents"])
    expect(result.status).toBe("ok")
    if (result.status === "blocked") throw new Error("Expected retrieved extraction")
    expect(result.sourceBytes).toBeUndefined()
    expect(result.descriptor.bodyHash).toBe("")
    const reference = preserveResearchSourceSnapshot(directory, result, displayedSourceSnippets(result.text, [], 4))
    const snapshot = readResearchSourceSnapshot(directory, reference)
    expect(snapshot?.text).toBe(text)
    expect(snapshot?.extraction).toBe("exa-contents-text/v1")
    expect(snapshot?.descriptor.retrieval?.provider).toBe("exa")
    expect(snapshot?.original_base64).toBeNull()
    expect(snapshot?.completeness).toBe("unknown")
    calls.length = 0
    await fetchSourceText(url, { browserFallback: false, exaFallback: false, fetch: async endpoint => {
      calls.push(String(endpoint))
      return new Response("unavailable", { status: 503 })
    } })
    expect(calls).toEqual([url])
    // A successful direct fetch or an explicit fence failure never contacts Exa.
    for (const response of [new Response(text, { headers: { "content-type": "text/plain" } }),
      new Response(null, { status: 302, headers: { location: "https://example.com/record" } })]) {
      calls.length = 0
      await fetchSourceText(url, { browserFallback: false, fetch: async endpoint => {
        calls.push(String(endpoint))
        return response
      } })
      expect(calls).toEqual([url])
    }
  } finally {
    if (previous === undefined) delete process.env.EXA_API_KEY
    else process.env.EXA_API_KEY = previous
    if (previousRoot === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previousRoot
    rmSync(directory, { recursive: true, force: true })
  }
})
