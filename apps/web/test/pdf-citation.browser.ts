// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/pdf-citation.browser.ts
// Mounts the real DocumentViewer on a PDF; only ingest HTTP is mocked. Set PLAYWRIGHT_CHROMIUM_PATH=<chromium> when Playwright's own browser is missing.
// When the citation changes while the viewer stays open on the same PDF, the page found for the previous
// excerpt must never be shown, even for one commit, beside the new excerpt.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-pdf-citation-"))
const first = { section: "3. Liability", excerpt: "The supplier's total liability is capped at the fees paid." }
const second = { section: "6. Returns", excerpt: "Returned samples are wrapped in the original covers." }
const pages: Record<string, number> = { [first.excerpt]: 3, [second.excerpt]: 6 }

const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentViewer from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentViewer.tsx"))}
const root = createRoot(document.getElementById("root"))
window.showCitation = (citation) => root.render(createElement(DocumentViewer, { matterId: "matter-viewer", name: "Supply terms.pdf", citation, onClose: () => {} }))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://viewer.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response(`<html><meta charset="utf-8"><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(10_000)
  // Each excerpt's page; the second answer arrives late, so the change is seen before it.
  await page.route("http://viewer.invalid/**", async route => {
    const request = route.request()
    const { pathname } = new URL(request.url())
    if (request.method() === "POST" && pathname === "/matters/matter-viewer/documents/pdf-page") {
      const { excerpt } = request.postDataJSON() as { excerpt: string }
      if (excerpt === second.excerpt) await new Promise(resolve => setTimeout(resolve, 300))
      return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ status: "matched", page: pages[excerpt] }) })
    }
    if (pathname === "/matters/matter-viewer/documents/content") return route.fulfill({ status: 200, contentType: "application/pdf", body: "%PDF-1.4\n%%EOF\n" })
    return route.fulfill({ status: 404, body: "" })
  })
  await page.goto(server.url.href)
  await page.waitForFunction(() => typeof (window as unknown as { showCitation?: unknown }).showCitation === "function")
  await page.evaluate(citation => (window as unknown as { showCitation: (c: unknown) => void }).showCitation(citation), first)
  await page.locator('iframe.pdf-render[src$="#page=3"]').waitFor({ state: "attached" })
  // Record what every commit shows from here: the frame's page and the status line.
  await page.evaluate(() => {
    const seen: { src: string | null; status: string | null }[] = []
    Object.assign(window, { seen })
    new MutationObserver(() => seen.push({
      src: document.querySelector("iframe.pdf-render")?.getAttribute("src") ?? null,
      status: document.querySelector('[role="status"]')?.textContent ?? null,
    })).observe(document.body, { subtree: true, childList: true, attributes: true, characterData: true })
  })
  await page.evaluate(citation => (window as unknown as { showCitation: (c: unknown) => void }).showCitation(citation), second)
  await page.locator('iframe.pdf-render[src$="#page=6"]').waitFor({ state: "attached" })
  const seen = await page.evaluate(() => (window as unknown as { seen: { src: string | null; status: string | null }[] }).seen)
  const stale = seen.filter(state => state.status?.includes("Returned samples") && (state.src?.endsWith("#page=3") || state.status.includes("page 3")))
  console.log(`${stale.length ? "FAIL" : "PASS"} citation change: ${seen.length} states seen, ${stale.length} showing page 3 beside the new excerpt`)
  assert.deepEqual(stale, [], "The new excerpt is never shown with the previous excerpt's page")
  assert.ok(seen.some(state => state.status?.includes("page 6") && state.status.includes("Returned samples")), "The new excerpt's page is shown")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
