import { tool } from "./legacy-definition"
import { readSearchKey } from "../lib/research"

// doc.haus web-search tool — citation DISCOVERY for legal research. The agent
// uses this to FIND where a law lives (which act, which section, which official
// site) when it does not already know the citation; it then READS the law with
// webfetch against the official source, which stays fenced (lib/research.ts).
// Search results are leads, never authority.
//
// Retain the legal discovery boundary and per-call firm credential settings.
// Exa's hosted MCP supports rate-limited keyless search; a configured key is
// sent on the initial request and is never retried without authentication.

const EXA_URL = "https://mcp.exa.ai/mcp"
const MAX_RESPONSE_BYTES = 256 * 1024

export default tool({
  description:
    "Search the web to FIND a legal citation or source you do not already know — which statute governs, the act name and section number, or the official page that carries it. Returns search results as leads, not authority: once you have the citation, retrieve the actual text from the official source with webfetch (legal-research skill) before relying on it. Uses rate-limited keyless search unless the firm supplies an Exa key in Settings.",
  args: {
    query: tool.schema.string().describe("Search query identifying the legal question and the official source or citation to discover"),
    numResults: tool.schema.number().min(1).max(20).optional().describe("Number of results (default 8, max 20)"),
  },
  async execute(args) {
    const apiKey = readSearchKey()
    const result = await callExa(apiKey, {
      query: args.query,
      type: "auto",
      numResults: args.numResults ?? 8,
      livecrawl: "fallback",
    })

    if (result.error) {
      const authenticationFailed = Boolean(apiKey) && /401|403/.test(result.error)
      return {
        title: authenticationFailed ? "Exa API key rejected" : "Web search failed",
        output:
          (authenticationFailed
            ? `Exa rejected the API key saved in Settings. The key is reaching the search service, but Exa reports that it is invalid. ` +
              `Replace it with an active key from the Exa dashboard, save Research settings, and retry; no engine restart is required.`
            : `Exa could not complete the search: ${result.error}`),
        metadata: { configured: true, results: false, authenticationFailed },
      }
    }
    if (!result.text)
      return {
        title: `No results: ${args.query}`,
        output: "No search results found. Try a different query.",
        metadata: { configured: true, results: false },
      }
    return {
      title: `Web search: ${args.query}`,
      output:
        result.text +
        `\n\n[legal-research] These results are leads, not authority. Before relying on any law they mention, ` +
        `retrieve its current text from the official source with webfetch and cite what you retrieved.`,
      metadata: { configured: true, results: true },
    }
  },
})

async function callExa(apiKey: string | undefined, exaArgs: Record<string, unknown>) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), 25_000)
  const url = new URL(EXA_URL)
  if (apiKey) url.searchParams.set("exaApiKey", apiKey)
  try {
    const res = await fetch(url, {
      method: "POST",
      signal: controller.signal,
      headers: { "Content-Type": "application/json", Accept: "application/json, text/event-stream" },
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "tools/call",
        params: { name: "web_search_exa", arguments: exaArgs },
      }),
    })
    if (!res.ok) {
      controller.abort()
      throw new Error(`Web search failed: ${res.status} ${res.statusText}`)
    }
    if (Number(res.headers.get("content-length")) > MAX_RESPONSE_BYTES) {
      controller.abort()
      throw new Error(`Web search response exceeded ${MAX_RESPONSE_BYTES} bytes`)
    }
    if (!res.body) return {}
    const reader = res.body.getReader()
    const chunks: Uint8Array[] = []
    let size = 0
    try {
      while (true) {
        const chunk = await reader.read()
        if (chunk.done) break
        size += chunk.value.byteLength
        if (size > MAX_RESPONSE_BYTES) {
          controller.abort()
          throw new Error(`Web search response exceeded ${MAX_RESPONSE_BYTES} bytes`)
        }
        chunks.push(chunk.value)
      }
      return parseResponse(Buffer.concat(chunks).toString("utf8"))
    } finally {
      await reader.cancel().catch(() => {})
      reader.releaseLock()
    }
  } finally {
    clearTimeout(timer)
  }
}

// The MCP response arrives either as a plain JSON-RPC body or as SSE "data:"
// lines; the result text lives at result.content[].text in both cases.
export function parseResponse(body: string): { text?: string; error?: string } {
  const direct = parsePayload(body.trim())
  if (direct) return direct
  for (const line of body.split("\n")) {
    if (!line.startsWith("data: ")) continue
    const data = parsePayload(line.substring(6).trim())
    if (data) return data
  }
  return {}
}

function parsePayload(payload: string) {
  if (!payload.startsWith("{")) return undefined
  const parsed = JSON.parse(payload) as { result?: { content?: { text?: string }[]; isError?: boolean }; error?: { message?: string } }
  const text = parsed.result?.content?.find((item) => item.text)?.text
  if (parsed.error?.message) return { error: parsed.error.message }
  if (parsed.result?.isError && text) return { error: text }
  if (text) return { text }
  return undefined
}
