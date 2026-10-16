import { describe, expect, test, beforeAll, afterAll } from "bun:test"
import { tmpdir } from "node:os"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import { hashBytes, hashText } from "../lib/research-source-fetch"

describe("hashBytes / hashText", () => {
  test("hashBytes returns sha256 hex of raw bytes", () => {
    const buf = new TextEncoder().encode("hello").buffer
    expect(hashBytes(buf)).toBe("2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824")
  })

  test("hashBytes returns empty string for empty input", () => {
    expect(hashBytes(new ArrayBuffer(0))).toBe("")
  })

  test("hashText returns sha256 hex of UTF-8 bytes", () => {
    expect(hashText("hello")).toBe("2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824")
  })

  test("hashText returns empty string for empty input", () => {
    expect(hashText("")).toBe("")
  })

  test("hashText is deterministic across calls", () => {
    expect(hashText("legal source text")).toBe(hashText("legal source text"))
  })
})

import { resolveLocation } from "../lib/research-source-fetch"

describe("resolveLocation", () => {
  test("absolute URL in Location passes through", () => {
    expect(resolveLocation("https://a.gov/x", "https://b.gov/y")).toBe("https://b.gov/y")
  })

  test("relative path resolved against base URL", () => {
    expect(resolveLocation("https://a.gov/dir/page", "/final")).toBe("https://a.gov/final")
  })

  test("relative path without leading slash resolved against base directory", () => {
    expect(resolveLocation("https://a.gov/dir/page", "next")).toBe("https://a.gov/dir/next")
  })

  test("missing Location header returns null", () => {
    expect(resolveLocation("https://a.gov/x", null)).toBe(null)
  })

  test("empty Location header returns null", () => {
    expect(resolveLocation("https://a.gov/x", "")).toBe(null)
  })
})

import { allowRedirect } from "../lib/research"

describe("allowRedirect", () => {
  let savedRoot: string | undefined
  let tempRoot: string

  beforeAll(async () => {
    savedRoot = process.env.WORKSPACE_ROOT
    tempRoot = `${tmpdir()}/allowRedirect-${process.pid}-${Date.now()}`
    await mkdir(tempRoot, { recursive: true })
    process.env.WORKSPACE_ROOT = tempRoot
    await Bun.write(`${tempRoot}/.preferences/preferences.json`, JSON.stringify({ webResearch: "official" }))
  })

  afterAll(async () => {
    if (savedRoot === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = savedRoot
    await rm(tempRoot, { recursive: true, force: true })
  })

  test("same host passes", async () => {
    expect(allowRedirect("https://law.gov/x", "https://law.gov/y")).toBe(true)
  })

  test("subdomain of original passes", async () => {
    expect(allowRedirect("https://law.gov/x", "https://sub.law.gov/y")).toBe(true)
  })

  test("parent of original passes (subdomain to apex)", async () => {
    expect(allowRedirect("https://sub.law.gov/x", "https://law.gov/y")).toBe(true)
  })

  test("official-source landing passes (gov to gov)", async () => {
    expect(allowRedirect("https://law.gov/x", "https://other.gov/y")).toBe(true)
  })

  test("off-fence landing blocked (gov to example.com)", async () => {
    expect(allowRedirect("https://law.gov/x", "https://example.com/y")).toBe(false)
  })

  test("open scope allows everything", async () => {
    await Bun.write(`${tempRoot}/.preferences/preferences.json`, JSON.stringify({ webResearch: "open" }))
    expect(allowRedirect("https://law.gov/x", "https://example.com/y")).toBe(true)
    await Bun.write(`${tempRoot}/.preferences/preferences.json`, JSON.stringify({ webResearch: "official" }))
    expect(allowRedirect("https://law.gov/x", "https://example.com/y")).toBe(false)
  })
})

import { buildDescriptor } from "../lib/research-source-fetch"

describe("buildDescriptor", () => {
  test("assembles descriptor with host derived from finalUrl", () => {
    const d = buildDescriptor({
      originalUrl: "https://a.gov/x",
      finalUrl: "https://sub.b.gov/y",
      contentType: "text/html",
      byteLength: 100,
      text: "hello",
      fetchedAt: "2026-07-04T00:00:00.000Z",
    })
    expect(d.host).toBe("sub.b.gov")
    expect(d.textChars).toBe(5)
    expect(d.textHash).toBe(hashText("hello"))
    expect(d.bodyHash).toBe("")
  })

  test("computes bodyHash when body provided", () => {
    const body = new TextEncoder().encode("abc").buffer
    const d = buildDescriptor({
      originalUrl: "https://a.gov/x",
      finalUrl: "https://a.gov/y",
      contentType: "application/pdf",
      byteLength: 3,
      text: "",
      body,
      fetchedAt: "2026-07-04T00:00:00.000Z",
    })
    expect(d.bodyHash).toBe(hashBytes(body))
    expect(d.textHash).toBe("")
  })
})

import { fetchSourceText, MAX_FETCH_BYTES } from "../lib/research-source-fetch"

describe("fetchSourceText integration", () => {
  const previousRoot = process.env.WORKSPACE_ROOT
  let root: string
  beforeAll(async () => {
    root = await mkdtemp(`${tmpdir()}/source-fetch-official-`)
    process.env.WORKSPACE_ROOT = root
    await Bun.write(`${root}/.preferences/preferences.json`, JSON.stringify({ webResearch: "official" }))
  })
  afterAll(async () => {
    if (previousRoot === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previousRoot
    await rm(root, { recursive: true, force: true })
  })
  const baseUrl = "https://records.gov"
  const fetchFixture = async (input: RequestInfo | URL) => {
    const url = new URL(input instanceof Request ? input.url : input.toString())
    switch (url.pathname) {
      case "/on-fence-redirect":
        return new Response(null, { status: 302, headers: { Location: "/final" } })
      case "/off-fence-redirect":
        return new Response(null, { status: 302, headers: { Location: "https://example.invalid/evil" } })
      case "/loop":
        return new Response(null, { status: 302, headers: { Location: "/loop" } })
      case "/oversize-body":
        return new Response("x".repeat(MAX_FETCH_BYTES + 1), { status: 200 })
      case "/empty-body":
        return new Response("<html><body></body></html>", { status: 200, headers: { "Content-Type": "text/html" } })
      case "/final":
        return new Response("<html><body>Retrieved official source text</body></html>", {
          status: 200,
          headers: { "Content-Type": "text/html" },
        })
      case "/missing":
        return new Response("not found", { status: 404 })
      case "/corrupt-pdf":
        return new Response("%-bad-pdf", { status: 200, headers: { "Content-Type": "application/pdf" } })
      default:
        return new Response("not found", { status: 404 })
    }
  }
  const fetchFixtureSource = (path: string) =>
    fetchSourceText(`${baseUrl}${path}`, { fetch: fetchFixture, browserFallback: false })

  test("follows on-fence redirect and returns ok with finalUrl set", async () => {
    const result = await fetchFixtureSource("/on-fence-redirect")
    expect(result.status).toBe("ok")
    if (result.status !== "ok") return
    expect(result.descriptor.finalUrl).toBe(`${baseUrl}/final`)
    expect(result.descriptor.originalUrl).toBe(`${baseUrl}/on-fence-redirect`)
    expect(result.text).toContain("Retrieved official source text")
  })

  test("blocks off-fence redirect with reason redirect-off-fence", async () => {
    const result = await fetchFixtureSource("/off-fence-redirect")
    expect(result.status).toBe("blocked")
    if (result.status !== "blocked") return
    expect(result.reason).toBe("redirect-off-fence")
    expect(result.descriptor.finalUrl).toBe("https://example.invalid/evil")
  })

  test("blocks redirect loop at MAX_REDIRECTS with reason http-error", async () => {
    const result = await fetchFixtureSource("/loop")
    expect(result.status).toBe("blocked")
    if (result.status !== "blocked") return
    expect(result.reason).toBe("http-error")
  })

  test("blocks oversize by actual byte count", async () => {
    const result = await fetchFixtureSource("/oversize-body")
    expect(result.status).toBe("blocked")
    if (result.status !== "blocked") return
    expect(result.reason).toBe("oversize")
    expect(result.descriptor.bodyHash).toBe("")
  })

  test("returns partial for empty body after extraction", async () => {
    const result = await fetchFixtureSource("/empty-body")
    expect(result.status).toBe("partial")
    if (result.status !== "partial") return
    expect(result.reason).toBe("empty-body")
    expect(result.descriptor.textHash).toBe("")
    expect(result.descriptor.host).toBeTruthy()
    expect(result.descriptor.fetchedAt).toBeTruthy()
  })

  test("blocks 404 final with reason http-error", async () => {
    const result = await fetchFixtureSource("/missing")
    expect(result.status).toBe("blocked")
    if (result.status !== "blocked") return
    expect(result.reason).toBe("http-error")
  })

  test("ok fetch: textHash matches sha256 of extracted text", async () => {
    const result = await fetchFixtureSource("/final")
    expect(result.status).toBe("ok")
    if (result.status !== "ok") return
    expect(result.descriptor.textHash).toBe(hashText(result.text))
    expect(new Date(result.descriptor.fetchedAt).toISOString()).toBe(result.descriptor.fetchedAt)
  })

  test("blocks corrupt PDF with reason extract-error instead of throwing", async () => {
    const result = await fetchFixtureSource("/corrupt-pdf")
    expect(result.status).toBe("blocked")
    if (result.status !== "blocked") return
    expect(result.reason).toBe("extract-error")
  })

  test("retains original PDF bytes, length and hash after worker text extraction", async () => {
    const original = await Bun.file(new URL("../../services/ingest/src/fixtures/synthetic-text-layer.pdf", import.meta.url)).arrayBuffer()
    const result = await fetchSourceText(`${baseUrl}/source.pdf`, {
      fetch: async () => new Response(original.slice(0), { headers: { "Content-Type": "application/pdf" } }),
      browserFallback: false,
      exaFallback: false,
    })
    expect(result.status).toBe("ok")
    if (result.status !== "ok") return
    expect(result.text.trim().length).toBeGreaterThan(0)
    expect(result.descriptor.byteLength).toBe(original.byteLength)
    expect(result.descriptor.bodyHash).toBe(hashBytes(original))
    expect(result.descriptor.textHash).toBe(hashText(result.text))
    expect(result.sourceBytes).toEqual(new Uint8Array(original))
    expect(result.extraction).toBe("pdf-text/v1")
  })
})

describe("fetchSourceText rendered official-source fallback", () => {
  test("retries an official source through rendered transport when plain extraction is empty", async () => {
    const url = "https://records.gov/source"
    const renderedUrls: string[] = []
    const result = await fetchSourceText(url, {
      fetch: async () =>
        new Response("<html><body></body></html>", {
          status: 200,
          headers: { "Content-Type": "text/html" },
        }),
      renderer: async (renderUrl) => {
        renderedUrls.push(renderUrl)
        return {
          finalUrl: renderUrl,
          contentType: "text/html",
          text: "Rendered official source text",
          byteLength: new TextEncoder().encode("Rendered official source text").byteLength,
        }
      },
    })

    expect(renderedUrls).toEqual([url])
    expect(result.status).toBe("ok")
    if (result.status !== "ok") return
    expect(result.text).toBe("Rendered official source text")
    expect(result.descriptor.originalUrl).toBe(url)
    expect(result.descriptor.finalUrl).toBe(url)
  })

  test("does not use rendered transport for a non-official source", async () => {
    const result = await fetchSourceText("https://example.com/source", {
      fetch: async () =>
        new Response("<html><body></body></html>", {
          status: 200,
          headers: { "Content-Type": "text/html" },
        }),
      renderer: async () => {
        throw new Error("renderer should not be called")
      },
    })

    expect(result.status).toBe("partial")
    if (result.status !== "partial") return
    expect(result.reason).toBe("empty-body")
  })

  test("keeps the original unresolved result when rendered transport fails", async () => {
    const result = await fetchSourceText("https://records.gov/missing", {
      fetch: async () => new Response("not found", { status: 404 }),
      renderer: async () => null,
    })

    expect(result.status).toBe("blocked")
    if (result.status !== "blocked") return
    expect(result.reason).toBe("http-error")
  })
})
