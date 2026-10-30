// Run from apps/web: bun test/concurrency-settings.browser.ts
// Uses a real browser with mocked local APIs; no model requests are made.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-concurrency-ui-"))
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ConcurrencySettings from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/settings/ConcurrencySettings.tsx"))}
function Harness() {
  const [busy, setBusy] = useState(false)
  return createElement('div', null, createElement('span', {id:'busy'}, String(busy)), createElement(ConcurrencySettings, {onBusy:setBusy, onDirty:()=>{}}))
}
createRoot(document.getElementById('root')).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://ingest.invalid", VITE_OPENCODE_URL: "http://engine.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response('<div id="root"></div><script type="module" src="/harness.js"></script>', { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage({ viewport: { width: 375, height: 812 } })
  let saved = 4, tests = 0, patches = 0
  let held: import("@playwright/test").Route | undefined
  await page.route("http://ingest.invalid/**", async route => {
    const request = route.request()
    if (request.method() === "PATCH") { patches++; saved = request.postDataJSON().model_concurrency }
    await route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ model_concurrency: saved }) })
  })
  await page.route("http://engine.invalid/**", route => { tests++; assert.equal(route.request().postDataJSON().count, saved); held = route })
  await page.goto(server.url.href)
  const field = page.getByLabel("Request limit")
  await field.waitFor()
  await page.getByRole("button", { name: "Test request capacity" }).waitFor()
  assert.equal(await field.getAttribute("aria-valuetext"), "5 requests")
  assert.equal(tests, 0, "test is never automatic")
  await field.focus()
  await page.keyboard.press("ArrowRight")
  await page.getByRole("button", { name: "Test request capacity" }).click()
  await page.getByText("Testing 10 requests using your saved primary model…").waitFor()
  assert.equal(patches, 1, "the selected limit is saved before testing")
  assert.equal(await page.locator("#busy").textContent(), "true")
  await held!.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ model: "fixture/model", requested: 10, succeeded: 8, failed: 2, peak: 9, elapsed_ms: 3400, errors: [{ kind: "unavailable", count: 2 }] }) })
  await page.getByText("8 of 10 requests succeeded; 2 failed.").waitFor()
  assert.ok(await page.getByText("Peak overlapping requests in doc.haus: 9.", { exact: false }).isVisible())
  assert.ok(await page.getByText("Connection unavailable: 2").isVisible())
  assert.equal(await page.locator("#busy").textContent(), "false")
  await page.getByRole("button", { name: "Test request capacity" }).click()
  await page.getByText("Testing 10 requests using your saved primary model…").waitFor()
  await held!.fulfill({ status: 409, contentType: "application/json", body: JSON.stringify({ error: "A concurrency test is already running." }) })
  await page.getByRole("alert").getByText("A concurrency test is already running.").waitFor()
  console.log("PASS saved limit, explicit test, progress, errors and observed overlap")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
