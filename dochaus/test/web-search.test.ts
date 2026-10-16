import { describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import webSearch from "../tool/web-search"
import { parseResponse } from "../tool/web-search"

describe("Exa MCP response parsing", () => {
  test("treats an embedded MCP tool error as an error", () => {
    const body =
      'event: message\ndata: {"result":{"content":[{"type":"text","text":"web_search_exa error (401): Invalid API key"}],"isError":true},"jsonrpc":"2.0","id":1}\n'

    expect(parseResponse(body)).toEqual({ error: "web_search_exa error (401): Invalid API key" })
  })

  test("returns successful MCP text", () => {
    const body = '{"result":{"content":[{"type":"text","text":"Search results"}]},"jsonrpc":"2.0","id":1}'

    expect(parseResponse(body)).toEqual({ text: "Search results" })
  })
})

// Redirect only the existing tool's HTTP request to a real local server.
// No public endpoint, provider response mock, or firm preferences are used.
async function localSearch(
  handler: (request: Request) => Response | Promise<Response>,
  run: (directory: string, requests: URL[], signals: AbortSignal[]) => Promise<void>,
) {
  const directory = mkdtempSync(path.join(tmpdir(), "web-search-http-"))
  const previousRoot = process.env.WORKSPACE_ROOT
  const previousKey = process.env.EXA_API_KEY
  const nativeFetch = globalThis.fetch
  const requests: URL[] = []
  const signals: AbortSignal[] = []
  const server = Bun.serve({ hostname: "127.0.0.1", port: 0, idleTimeout: 60, fetch: handler })
  process.env.WORKSPACE_ROOT = directory
  delete process.env.EXA_API_KEY
  mkdirSync(path.join(directory, ".preferences"))
  globalThis.fetch = Object.assign((input: RequestInfo | URL, init?: RequestInit) => {
    const requested = new URL(String(input))
    expect(requested.origin).toBe("https://mcp.exa.ai")
    expect(requested.pathname).toBe("/mcp")
    requests.push(requested)
    if (init?.signal) signals.push(init.signal)
    return nativeFetch(new URL(`${requested.pathname}${requested.search}`, server.url), init)
  }, nativeFetch)
  try {
    await run(directory, requests, signals)
  } finally {
    globalThis.fetch = nativeFetch
    if (previousRoot === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previousRoot
    if (previousKey === undefined) delete process.env.EXA_API_KEY
    else process.env.EXA_API_KEY = previousKey
    server.stop(true)
    rmSync(directory, { recursive: true, force: true })
  }
}

const context = {
  directory: "unused", worktree: "unused", agent: "research",
  sessionID: "ses_local_search", messageID: "msg_local_search", callID: "call_local_search",
  abort: new AbortController().signal, metadata() {}, async ask() {},
}
const resultText = "Title: Public record\nURL: https://records.gov/record\nText: Public source lead."
const response = JSON.stringify({ jsonrpc: "2.0", id: 1, result: { content: [{ type: "text", text: resultText }] } })

test("keyless discovery uses the same MCP request and preserves the legal authority warning", async () => {
  await localSearch(async request => {
    expect(await request.json()).toEqual({ jsonrpc: "2.0", id: 1, method: "tools/call",
      params: { name: "web_search_exa", arguments: { query: "public record", type: "auto", numResults: 8, livecrawl: "fallback" } } })
    return new Response(response, { headers: { "content-type": "application/json" } })
  }, async (_directory, requests) => {
    const result = await webSearch.execute({ query: "public record" }, context)
    expect(requests).toHaveLength(1)
    expect(requests[0].search).toBe("")
    expect(result.output).toContain(resultText)
    expect(result.output).toContain("leads, not authority")
    expect(result.metadata).toMatchObject({ results: true })
  })
})

test("Settings keys are reread, take precedence over environment, and support SSE", async () => {
  await localSearch(() => new Response(`event: message\ndata: ${response}\n\n`, {
    headers: { "content-type": "text/event-stream" },
  }), async (directory, requests) => {
    process.env.EXA_API_KEY = "synthetic-env"
    for (const key of ["synthetic-first", "synthetic-updated", ""]) {
      writeFileSync(path.join(directory, ".preferences", "preferences.json"), JSON.stringify({ searchApiKey: key }))
      expect((await webSearch.execute({ query: "public record" }, context)).output).toContain(resultText)
    }
    expect(requests.map(url => url.searchParams.get("exaApiKey"))).toEqual([
      "synthetic-first", "synthetic-updated", "synthetic-env",
    ])
  })
})

test.each([401, 403, 429])("HTTP %i fails once without authentication downgrade or retry", async status => {
  await localSearch(() => new Response("unavailable", { status }), async (directory, requests) => {
    writeFileSync(path.join(directory, ".preferences", "preferences.json"), JSON.stringify({ searchApiKey: "synthetic-auth" }))
    await expect(webSearch.execute({ query: "public record" }, context)).rejects.toThrow(String(status))
    expect(requests).toHaveLength(1)
    expect(requests[0].searchParams.get("exaApiKey")).toBe("synthetic-auth")
  })
})

test("embedded auth rejection does not retry without the configured key", async () => {
  await localSearch(() => Response.json({ result: { isError: true, content: [{ type: "text", text: "401 Invalid API key" }] } }),
    async (directory, requests) => {
      writeFileSync(path.join(directory, ".preferences", "preferences.json"), JSON.stringify({ searchApiKey: "synthetic-auth" }))
      expect((await webSearch.execute({ query: "public record" }, context)).metadata).toMatchObject({ authenticationFailed: true, results: false })
      expect(requests).toHaveLength(1)
    })
})

test("oversized streamed body aborts without publishing partial search results", async () => {
  await localSearch(() => new Response(new ReadableStream({
    start(controller) {
      for (let index = 0; index < 9; index++) controller.enqueue(new TextEncoder().encode("x".repeat(32 * 1024)))
      controller.close()
    },
  })), async (_directory, requests, signals) => {
    await expect(webSearch.execute({ query: "public record" }, context)).rejects.toThrow("exceeded 262144 bytes")
    expect(requests).toHaveLength(1)
    expect(signals[0].aborted).toBe(true)
  })
})

test("the existing 25 second deadline covers a stalled body after headers", async () => {
  await localSearch(() => new Response(new ReadableStream({
    start(controller) { controller.enqueue(new TextEncoder().encode("event: message\n")) },
  }), { headers: { "content-type": "text/event-stream" } }), async (_directory, requests, signals) => {
    const started = Date.now()
    await expect(webSearch.execute({ query: "public record" }, context)).rejects.toThrow()
    expect(Date.now() - started).toBeGreaterThanOrEqual(24_000)
    expect(Date.now() - started).toBeLessThan(29_000)
    expect(requests).toHaveLength(1)
    expect(signals[0].aborted).toBe(true)
  })
}, 35_000)
