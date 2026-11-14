// Run from apps/web: bun test/document-upload-name-check.browser.ts
// Mounts the real DocumentUpload with the product's stylesheets. A same-name
// upload is answered from the matter's own document list, so no request fails
// in the browser console; the server's 409 stays as the race guard. #365 2.8.
import assert from "node:assert/strict"
import { mkdirSync } from "node:fs"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Document } from "../src/api/ingest/documents"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const shots = process.env.REVIEW_SHOTS_DIR ?? path.join(os.tmpdir(), "dochaus-review-help")
mkdirSync(shots, { recursive: true })
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-upload-name-check-"))
const documents: Document[] = [
  { id: 1, name: "Agreement.docx", doc_path: "/m/Agreement.docx", created_at: 1, role: "source" },
  { id: 2, name: "Terms.pdf", doc_path: "/m/Terms.pdf", created_at: 2, role: "source" },
]
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/primitives.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-documents.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-review.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/upload.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/motion.css"))}
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentUpload from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentUpload.tsx"))}
import { ToastProvider } from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/Toast.tsx"))}
createRoot(document.getElementById("root")).render(createElement(ToastProvider, null,
  createElement(DocumentUpload, {
    matterId: "matter-upload", documents: ${JSON.stringify(documents)},
    onUploaded: () => { window.uploaded = (window.uploaded ?? 0) + 1 },
    onUploadBatch: () => { window.batches = (window.batches ?? 0) + 1 },
    onView: () => {},
  }),
))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://ingest.invalid", VITE_OPENCODE_URL: "http://engine.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = (await Promise.all(build.outputs.filter(output => output.kind === "asset" && output.path.endsWith(".css")).map(output => output.text()))).join("\n")
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => {
  const { pathname } = new URL(request.url)
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname === "/harness.css") return new Response(css, { headers: { "content-type": "text/css" } })
  return new Response(`<html><meta charset="utf-8"><link rel="stylesheet" href="/harness.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  await page.setViewportSize({ width: 1440, height: 900 })
  const uploads: string[] = []
  const consoleErrors: string[] = []
  page.on("console", message => { if (message.type() === "error") consoleErrors.push(message.text()) })
  page.on("pageerror", error => consoleErrors.push(String(error)))
  await page.route("http://ingest.invalid/**", route => {
    const request = route.request()
    const url = new URL(request.url())
    if (url.pathname === "/matters/matter-upload/documents" && request.method() === "POST") {
      uploads.push(url.pathname)
      return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ name: "New terms.pdf", quality: { status: "ok" } }) })
    }
    return route.fulfill({ status: 200, contentType: "application/json", body: "{}" })
  })
  await page.goto(server.url.href)
  const picker = page.locator('input[type="file"]').first()
  await picker.setInputFiles({ name: "terms.PDF", mimeType: "application/pdf", buffer: Buffer.from("%PDF-1.4 synthetic") })
  await page.getByText("Already in this matter. Save this as a new version of it, or keep both?").waitFor()
  assert.deepEqual(uploads, [], "a same-name file never leaves the browser")
  assert.equal(consoleErrors.filter(line => line.includes("409")).length, 0, "no failed request is logged")
  await page.screenshot({ path: path.join(shots, "upload-same-name-1440.png") })
  await page.setViewportSize({ width: 400, height: 900 })
  await page.screenshot({ path: path.join(shots, "upload-same-name-400.png") })
  await page.setViewportSize({ width: 1440, height: 900 })

  // A name the matter does not list still uploads; a batch of two reports once,
  // after both have settled, so work that reads the matter starts then.
  await picker.setInputFiles([
    { name: "New terms.pdf", mimeType: "application/pdf", buffer: Buffer.from("%PDF-1.4 synthetic") },
    { name: "Schedule.md", mimeType: "text/markdown", buffer: Buffer.from("# Schedule") },
  ])
  await page.waitForFunction(() => document.body.textContent?.includes("New terms.pdf"))
  await page.waitForFunction(() => (window as unknown as { batches?: number }).batches === 1)
  assert.deepEqual(uploads, ["/matters/matter-upload/documents", "/matters/matter-upload/documents"], "new names upload as before")
  assert.equal(await page.evaluate(() => (window as unknown as { batches?: number }).batches), 1, "one upload batch reports once")
  console.log("PASS a same-name upload is answered from the document list, with no failed request, new names still upload, and a batch reports once")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
