import { createHash } from "node:crypto"
import { isIP } from "node:net"
import { isOfficialSource, readSearchKey } from "./research"

const ENDPOINT = "https://api.exa.ai/contents"
const MAX_RESPONSE_BYTES = 5 * 1024 * 1024

// A third-party text extraction is not the origin's original bytes. Accept only
// an exact requested public URL and a successful live crawl, never a search
// snippet, generated summary, cached fallback or another result's text.
export async function fetchExaSourceText(url: string, options: { fetch?: typeof fetch; key?: string } = {}) {
  try {
    const target = new URL(url)
    if (target.protocol !== "https:" || target.username || target.password || isIP(target.hostname) ||
      target.hostname === "localhost" || /\.(local|internal|localhost)$/.test(target.hostname) || !isOfficialSource(url)) return null
    target.hash = ""
    const key = options.key ?? readSearchKey()
    if (!key) return null
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), 25_000)
    try {
      const response = await (options.fetch ?? globalThis.fetch)(ENDPOINT, {
        method: "POST", redirect: "error", signal: controller.signal,
        headers: { "Content-Type": "application/json", "x-api-key": key },
        body: JSON.stringify({ urls: [target.href], text: true, maxAgeHours: 0, livecrawlTimeout: 15000 }),
      })
      if (!response.ok || !response.body) {
        await response.body?.cancel()
        return null
      }
      if (Number(response.headers.get("content-length")) > MAX_RESPONSE_BYTES) {
        await response.body.cancel()
        return null
      }
      const reader = response.body.getReader()
      const chunks: Uint8Array[] = []
      let size = 0
      try {
        while (true) {
          const chunk = await reader.read()
          if (chunk.done) break
          size += chunk.value.byteLength
          if (size > MAX_RESPONSE_BYTES) return null
          chunks.push(chunk.value)
        }
      } finally {
        await reader.cancel().catch(() => {})
        reader.releaseLock()
      }
      const body = Buffer.concat(chunks)
      const value: unknown = JSON.parse(body.toString("utf8"))
      if (!record(value) || !Array.isArray(value.statuses) || value.statuses.length !== 1 ||
        !Array.isArray(value.results) || value.results.length !== 1) return null
      const status: unknown = value.statuses[0]
      const result: unknown = value.results[0]
      if (!record(status) || status.id !== target.href || status.status !== "success" || status.source !== "crawled" ||
        !record(result) || result.url !== target.href || typeof result.text !== "string" || !result.text.trim()) return null
      return {
        url: target.href, text: result.text,
        provenance: { provider: "exa" as const, requestedFreshness: "live" as const, statusSource: "crawled" as const,
          responseHash: createHash("sha256").update(body).digest("hex"), receivedAt: new Date().toISOString() },
      }
    } finally {
      clearTimeout(timer)
    }
  } catch {
    // Retain the original source failure. Never leak a provider response or key.
    return null
  }
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value)
}
