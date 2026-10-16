import { createHash } from "node:crypto"
import { allowRedirect, isOfficialSource } from "./research"
import { projectUntrustedHTML } from "./untrusted-html"
import { fetchExaSourceText } from "./research-source-exa"

export const MAX_FETCH_BYTES = 5 * 1024 * 1024
export const MAX_TEXT_CHARS = 500_000
const MAX_REDIRECTS = 5
const FETCH_TIMEOUT_MS = 20_000
const RENDER_TIMEOUT_MS = 15_000

export function hashBytes(buf: ArrayBuffer): string {
  if (buf.byteLength === 0) return ""
  return createHash("sha256").update(Buffer.from(buf)).digest("hex")
}

export function hashText(text: string): string {
  if (!text) return ""
  return createHash("sha256").update(text, "utf8").digest("hex")
}

export function resolveLocation(fromUrl: string, locationHeader: string | null): string | null {
  if (!locationHeader) return null
  try {
    return new URL(locationHeader, fromUrl).toString()
  } catch {
    return null
  }
}

export type FetchStatus = "ok" | "blocked" | "partial"

export type FetchReason =
  | "redirect-off-fence"
  | "oversize"
  | "http-error"
  | "tls-error"
  | "network-error"
  | "empty-body"
  | "extract-error"
  | "text-truncated"

export type FetchDescriptor = {
  originalUrl: string
  finalUrl: string
  host: string
  contentType: string
  byteLength: number
  textChars: number
  textHash: string
  bodyHash: string
  fetchedAt: string
  projection?: ReturnType<typeof projectUntrustedHTML>["projection"]
  retrieval?: NonNullable<Awaited<ReturnType<typeof fetchExaSourceText>>>["provenance"]
}

export type FetchResult =
  | { status: "ok"; descriptor: FetchDescriptor; text: string; sourceBytes?: Uint8Array; extraction?: string }
  | { status: "partial"; descriptor: FetchDescriptor; text: string; reason: FetchReason; sourceBytes?: Uint8Array; extraction?: string }
  | { status: "blocked"; descriptor: FetchDescriptor; reason: FetchReason }

export type RenderedSource = {
  finalUrl: string
  contentType: string
  text: string
  byteLength: number
}

export type FetchSourceTextOptions = {
  fetch?: typeof fetch
  renderer?: (url: string) => Promise<RenderedSource | null>
  browserFallback?: boolean
  exaFallback?: boolean
}

export function buildDescriptor(args: {
  originalUrl: string
  finalUrl: string
  contentType: string
  byteLength: number
  text: string
  body?: ArrayBuffer
  fetchedAt: string
}): FetchDescriptor {
  return {
    originalUrl: args.originalUrl,
    finalUrl: args.finalUrl,
    host: new URL(args.finalUrl).hostname,
    contentType: args.contentType,
    byteLength: args.byteLength,
    textChars: args.text.length,
    textHash: hashText(args.text),
    bodyHash: args.body ? hashBytes(args.body) : "",
    fetchedAt: args.fetchedAt,
  }
}

export async function fetchSourceText(originalUrl: string, options: FetchSourceTextOptions = {}): Promise<FetchResult> {
  const fetchedAt = new Date().toISOString()
  const result = await fetchPlainSourceText(originalUrl, fetchedAt, options.fetch ?? globalThis.fetch)

  if (result.status === "ok" || !isOfficialSource(originalUrl) ||
    ["redirect-off-fence", "oversize", "text-truncated"].includes(result.reason)) return result
  if (shouldTryRenderedFallback(originalUrl, result, options)) {
    const rendered = await safeRenderSource(originalUrl, options.renderer ?? renderSourceWithBrowser)
    const recovered = rendered && renderedFetchResult(originalUrl, rendered, fetchedAt)
    if (recovered) return recovered
  }
  if (options.exaFallback === false) return result
  const extracted = await fetchExaSourceText(originalUrl, { fetch: options.fetch })
  if (!extracted) return result
  const text = extracted.text.slice(0, MAX_TEXT_CHARS)
  const descriptor = { ...buildDescriptor({ originalUrl, finalUrl: extracted.url, fetchedAt,
    contentType: "text/plain", byteLength: Buffer.byteLength(text), text }), retrieval: extracted.provenance }
  // No sourceBytes/bodyHash: the original document and extraction completeness
  // are unknown. Exact text/locator checks still operate on the retained text.
  return extracted.text.length > MAX_TEXT_CHARS
    ? { status: "partial", descriptor, text, extraction: "exa-contents-text/v1", reason: "text-truncated" }
    : { status: "ok", descriptor, text, extraction: "exa-contents-text/v1" }
}

async function fetchPlainSourceText(originalUrl: string, fetchedAt: string, fetcher: typeof fetch): Promise<FetchResult> {
  let currentUrl = originalUrl
  for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
    const attempted = await safeFetch(currentUrl, fetcher)
    if (attempted.kind !== "ok")
      return blocked({ originalUrl, finalUrl: currentUrl, fetchedAt, reason: attempted.kind })

    const response = attempted.response
    if (response.status >= 300 && response.status < 400) {
      const location = resolveLocation(currentUrl, response.headers.get("location"))
      if (!location || hop === MAX_REDIRECTS)
        return blocked({ originalUrl, finalUrl: currentUrl, fetchedAt, reason: "http-error" })
      if (!allowRedirect(currentUrl, location))
        return blocked({ originalUrl, finalUrl: location, fetchedAt, reason: "redirect-off-fence" })
      currentUrl = location
      continue
    }
    if (response.status >= 400)
      return blocked({ originalUrl, finalUrl: currentUrl, fetchedAt, reason: "http-error" })

    return await readBody(response, originalUrl, currentUrl, fetchedAt)
  }
  return blocked({ originalUrl, finalUrl: currentUrl, fetchedAt, reason: "http-error" })
}

function shouldTryRenderedFallback(originalUrl: string, result: FetchResult, options: FetchSourceTextOptions) {
  if (options.browserFallback === false) return false
  if (!isOfficialSource(originalUrl)) return false
  if (result.status === "ok") return false
  return !["redirect-off-fence", "oversize", "text-truncated"].includes(result.reason)
}

async function safeRenderSource(
  originalUrl: string,
  renderer: (url: string) => Promise<RenderedSource | null>,
): Promise<RenderedSource | null> {
  try {
    return await renderer(originalUrl)
  } catch {
    return null
  }
}

function renderedFetchResult(originalUrl: string, rendered: RenderedSource, fetchedAt: string): FetchResult | null {
  try {
    if (!allowRedirect(originalUrl, rendered.finalUrl)) return null
    // innerText is already visible text; do not parse literal angle brackets as
    // HTML or silently claim that a capped projection is a complete source.
    const text = rendered.text.slice(0, MAX_TEXT_CHARS)
    if (!text) return null
    const body = new TextEncoder().encode(text).buffer
    if (rendered.byteLength > MAX_FETCH_BYTES || body.byteLength > MAX_FETCH_BYTES) return null
    const source = {
      descriptor: buildDescriptor({
        originalUrl,
        finalUrl: rendered.finalUrl,
        contentType: rendered.contentType,
        byteLength: rendered.byteLength,
        text,
        fetchedAt,
      }),
      text,
      extraction: "browser-visible-text/v1",
    }
    return rendered.text.length > MAX_TEXT_CHARS
      ? { status: "partial", ...source, reason: "text-truncated" }
      : { status: "ok", ...source }
  } catch {
    return null
  }
}

function blocked(args: {
  originalUrl: string
  finalUrl: string
  fetchedAt: string
  reason: FetchReason
  contentType?: string
  byteLength?: number
  text?: string
}): FetchResult {
  return {
    status: "blocked",
    descriptor: buildDescriptor({
      originalUrl: args.originalUrl,
      finalUrl: args.finalUrl,
      contentType: args.contentType ?? "",
      byteLength: args.byteLength ?? 0,
      text: args.text ?? "",
      fetchedAt: args.fetchedAt,
    }),
    reason: args.reason,
  }
}

async function readBody(
  response: Response,
  originalUrl: string,
  finalUrl: string,
  fetchedAt: string,
): Promise<FetchResult> {
  const declaredLength = Number(response.headers.get("content-length") ?? "0")
  if (declaredLength > MAX_FETCH_BYTES)
    return blocked({
      originalUrl,
      finalUrl,
      fetchedAt,
      contentType: response.headers.get("content-type") ?? "",
      reason: "oversize",
    })

  let buffer: ArrayBuffer
  try {
    buffer = await response.arrayBuffer()
  } catch {
    return blocked({ originalUrl, finalUrl, fetchedAt, reason: "network-error" })
  }
  if (buffer.byteLength > MAX_FETCH_BYTES)
    return blocked({
      originalUrl,
      finalUrl,
      fetchedAt,
      contentType: response.headers.get("content-type") ?? "",
      reason: "oversize",
    })

  const contentType = response.headers.get("content-type") ?? ""
  let html: ReturnType<typeof projectUntrustedHTML> | undefined
  let text: string
  try {
    html = contentType.toLowerCase().includes("html")
      ? projectUntrustedHTML(new TextDecoder().decode(buffer))
      : undefined
    text = isPdf(finalUrl, contentType)
      ? await pdfText(buffer)
      : html
        ? html.text
        : new TextDecoder().decode(buffer)
  } catch {
    return blocked({ originalUrl, finalUrl, fetchedAt, contentType, reason: "extract-error" })
  }
  const descriptor = buildDescriptor({
    originalUrl,
    finalUrl,
    contentType,
    byteLength: buffer.byteLength,
    text: text.slice(0, MAX_TEXT_CHARS),
    body: buffer,
    fetchedAt,
  })

  const source = {
    descriptor: { ...descriptor, ...(html ? { projection: html.projection } : {}) },
    text: text.slice(0, MAX_TEXT_CHARS),
    sourceBytes: new Uint8Array(buffer),
    extraction: isPdf(finalUrl, contentType) ? "pdf-text/v1" : contentType.toLowerCase().includes("html") ? "html-visible-text/v1" : "plain-text/v1",
  }
  if (!text) return { status: "partial", ...source, reason: "empty-body" }
  if (text.length > MAX_TEXT_CHARS) return { status: "partial", ...source, reason: "text-truncated" }
  return { status: "ok", ...source }
}

type SafeFetchResult =
  | { kind: "ok"; response: Response }
  | { kind: "tls-error" }
  | { kind: "network-error" }

async function safeFetch(url: string, fetcher: typeof fetch): Promise<SafeFetchResult> {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS)
  try {
    const response = await fetcher(url, {
      redirect: "manual",
      signal: controller.signal,
      headers: {
        accept: "text/html,application/xhtml+xml,application/pdf,text/plain,*/*",
        "user-agent": "ResearchSourceFetch/1.0",
      },
    })
    return { kind: "ok", response }
  } catch (error) {
    if (error instanceof Error && /certificate|verify|tls|ssl/i.test(error.message)) return { kind: "tls-error" }
    return { kind: "network-error" }
  } finally {
    clearTimeout(timer)
  }
}

export function readableText(text: string) {
  return projectUntrustedHTML(text).text
}

async function pdfText(buffer: ArrayBuffer) {
  const { extractText, getDocumentProxy } = await import("unpdf")
  // PDF workers transfer their input buffer. Retain the original source bytes
  // for the descriptor hash and the immutable research source snapshot.
  const { text } = await extractText(await getDocumentProxy(new Uint8Array(buffer.slice(0))), { mergePages: true })
  return text
}

function isPdf(url: string, contentType: string) {
  return contentType.toLowerCase().includes("application/pdf") || new URL(url).pathname.toLowerCase().endsWith(".pdf")
}

async function renderSourceWithBrowser(url: string): Promise<RenderedSource | null> {
  try {
    const { chromium } = await import("playwright")
    const browser = await chromium.launch({ headless: true })
    try {
      const page = await browser.newPage()
      page.setDefaultTimeout(RENDER_TIMEOUT_MS)
      page.setDefaultNavigationTimeout(RENDER_TIMEOUT_MS)
      const response = await page.goto(url, { waitUntil: "domcontentloaded", timeout: RENDER_TIMEOUT_MS })
      if (response && response.status() >= 400) return null
      // A failed visible-text extraction must remain unavailable. Raw DOM
      // textContent includes hidden/script content and cannot carry this proof.
      const text = await page.locator("body").innerText()
      if (typeof text !== "string") return null
      const byteLength = Buffer.byteLength(text, "utf8")
      if (byteLength > MAX_FETCH_BYTES) return null
      return {
        finalUrl: page.url(),
        contentType: response?.headers()["content-type"] ?? "text/html",
        text,
        byteLength,
      }
    } finally {
      await browser.close().catch(() => {})
    }
  } catch {
    return null
  }
}
