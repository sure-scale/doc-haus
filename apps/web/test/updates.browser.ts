// Run with `bun test/updates.browser.ts` from apps/web; requires Playwright Chromium.
// Every update request is intercepted. No live service or repository is changed.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Page, Route } from "@playwright/test"
import type { useUpdates } from "../src/components/Updates"

declare global { interface Window { updateHarness: ReturnType<typeof useUpdates> } }

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE || "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-update-browser-"))
const component = path.resolve(import.meta.dir, "../src/components/Updates.tsx")
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { useUpdates, UpdateButton, UpdateModal } from ${JSON.stringify(component)}
function Harness() {
  const updates = useUpdates(true)
  window.updateHarness = updates
  return createElement("main", null,
    createElement("textarea", { "aria-label": "Draft", defaultValue: "Work in progress" }),
    createElement(UpdateButton, { updates }),
    updates.open && createElement(UpdateModal, { updates }))
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://updates.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs[0].text()
const css = await Bun.file(path.resolve(import.meta.dir, "../src/styles/launch-notice.css")).text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response(`<html><style>:root { --panel:white; --text:#13271c; --text-secondary:#4e6156; --accent:#26734d; --border:#aaa; --border-soft:#ddd } body { font-family:sans-serif } ${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH })
const current = { version: 1, name: "Current", notes: "Current notes" }
const latest = { version: 2, name: "Next", notes: "Next notes" }
const offer = { current, latest, target: "a".repeat(40), available: true }
const results: string[] = []

async function scenario(name: string, run: (page: Page, calls: { method: string; url: string; body: string | null }[], respond: (value: unknown) => void, accept: (value: "accepted" | "lost") => void) => Promise<void>) {
  if (process.env.UPDATE_SCENARIO && !name.includes(process.env.UPDATE_SCENARIO)) return
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } })
  const page = await context.newPage()
  page.setDefaultTimeout(5000)
  const calls: { method: string; url: string; body: string | null }[] = []
  let value: unknown = { current, available: false }
  let post = "accepted"
  await page.route("http://updates.invalid/**", async (route: Route) => {
    const request = route.request()
    calls.push({ method: request.method(), url: request.url(), body: request.postData() })
    if (request.method() === "POST") {
      if (post === "lost") return route.abort("connectionclosed")
      return route.fulfill({ status: 202, contentType: "application/json", body: JSON.stringify({ phase: "restarting" }) })
    }
    if (value === "offline") return route.abort("connectionfailed")
    return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(value) })
  })
  try {
    await run(page, calls, next => { value = next }, next => { post = next })
    results.push(name)
    console.log(`PASS ${name}`)
  } finally { await context.close() }
}

async function visible(page: Page, name: string) {
  await page.getByRole("button", { name, exact: true }).waitFor({ state: "visible" })
}

try {
  await scenario("HTTP 200 error is visible, non-blocking, retryable, and never prompts repeatedly", async (page, calls, respond) => {
    respond({ current, available: false, error: "Git could not complete the update. Check your repository access and connection." })
    await page.goto(server.url.href)
    await visible(page, "Update check failed")
    assert.equal(await page.getByRole("dialog").count(), 0)
    await page.getByRole("textbox", { name: "Draft" }).fill("Drafting remains available")
    await page.getByRole("button", { name: "Update check failed", exact: true }).click()
    await visible(page, "Check again")
    assert.match(await page.getByRole("alert").innerText(), /repository access and connection/)
    assert.equal(await page.getByRole("button", { name: "Update now", exact: true }).count(), 0)
    await page.evaluate(() => window.updateHarness.apply())
    assert.equal(calls.filter(call => call.method === "POST").length, 0)
    await page.getByRole("button", { name: "Close", exact: true }).click()
    const before = calls.length
    await page.evaluate(() => window.dispatchEvent(new Event("focus")))
    await page.waitForFunction(() => !window.updateHarness.checking)
    assert.ok(calls.length > before)
    assert.equal(await page.getByRole("dialog").count(), 0)
    await page.getByRole("button", { name: "Update check failed", exact: true }).click()
    if (process.env.UPDATE_SCREENSHOT) await page.screenshot({ path: process.env.UPDATE_SCREENSHOT })
    respond({ current, available: false })
    let releaseRetry = () => {}
    const retryWait = new Promise<void>(resolve => { releaseRetry = resolve })
    let refreshes = 0
    await page.route("http://updates.invalid/updates?refresh=1", async route => {
      refreshes++
      await retryWait
      await route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ current, available: false }) })
    })
    await page.getByRole("button", { name: "Check again", exact: true }).click()
    assert.equal(await page.getByRole("button", { name: "Checking…", exact: true }).isDisabled(), true)
    const checks = calls.length
    await page.evaluate(() => { window.updateHarness.retry(); window.dispatchEvent(new Event("focus")) })
    assert.equal(refreshes, 1)
    assert.equal(calls.length, checks)
    releaseRetry()
    await page.getByText("Doc.Haus is up to date.", { exact: true }).waitFor()
    assert.equal(await page.getByRole("alert").count(), 0)
    assert.equal(refreshes, 1)
    assert.equal(calls.filter(call => call.method === "POST").length, 0)
  })

  await scenario("network failure clears on retry and a valid offer can be installed", async (page, calls, respond) => {
    respond("offline")
    await page.goto(server.url.href)
    await visible(page, "Update check failed")
    assert.equal(await page.getByRole("dialog").count(), 0)
    await page.getByRole("button", { name: "Update check failed", exact: true }).click()
    assert.match(await page.getByRole("alert").innerText(), /Could not reach/)
    respond(offer)
    await page.getByRole("button", { name: "Check again", exact: true }).click()
    await visible(page, "Update now")
    assert.equal(await page.getByRole("alert").count(), 0)
    await page.getByRole("button", { name: "Update now", exact: true }).click()
    await page.getByRole("heading", { name: "Updating Doc.Haus" }).waitFor()
    assert.deepEqual(calls.filter(call => call.method === "POST").map(call => JSON.parse(call.body!)), [{ target: offer.target }])
  })

  await scenario("invalid target cannot create an offer or POST", async (page, calls, respond) => {
    respond({ ...offer, target: "invalid" })
    await page.goto(server.url.href)
    await visible(page, "Update check failed")
    assert.equal(await page.getByRole("dialog").count(), 0)
    await page.evaluate(() => window.updateHarness.apply())
    await page.getByRole("button", { name: "Update check failed", exact: true }).click()
    assert.match(await page.getByRole("alert").innerText(), /information is incomplete/)
    assert.equal(await page.getByRole("button", { name: "Update now", exact: true }).count(), 0)
    assert.equal(calls.filter(call => call.method === "POST").length, 0)
  })

  for (const response of ["accepted", "lost"] as const) await scenario(`${response} apply response reconnects through disconnect without a second POST`, async (page, calls, respond, accept) => {
    respond(offer)
    accept(response)
    await page.goto(server.url.href)
    await visible(page, "Update now")
    // The heading also appears while the POST is still requesting. Keep the
    // fixture online until acceptance (or a lost response) has established the
    // pending marker and the updating effect has completed its first check.
    const reconnectCheck = page.waitForResponse(result => result.request().method() === "GET" && result.url() === "http://updates.invalid/updates")
    await page.getByRole("button", { name: "Update now", exact: true }).click()
    await page.getByRole("heading", { name: "Updating Doc.Haus" }).waitFor()
    await reconnectCheck
    await page.waitForFunction(() => sessionStorage.getItem("dh.update-pending") === "2" && !window.updateHarness.checking)
    respond("offline")
    const checks = calls.length
    await page.evaluate(() => window.dispatchEvent(new Event("focus")))
    await page.waitForTimeout(2200)
    assert.ok(calls.length > checks)
    assert.equal(await page.getByRole("alert").count(), 0)
    assert.equal(await page.getByRole("button", { name: "Check again", exact: true }).count(), 0)
    assert.equal(await page.getByRole("button", { name: "Updating…", exact: true }).last().isDisabled(), true)
    await page.evaluate(() => window.updateHarness.apply())
    assert.equal(calls.filter(call => call.method === "POST").length, 1)
    assert.equal(await page.evaluate(() => sessionStorage.getItem("dh.update-pending")), "2")
    respond({ current: latest, available: false })
    const reload = page.waitForEvent("load")
    await page.evaluate(() => window.dispatchEvent(new Event("focus")))
    await reload
    await page.waitForFunction(() => window.updateHarness && !window.updateHarness.updating)
    assert.equal(await page.evaluate(() => sessionStorage.getItem("dh.update-pending")), "")
    assert.equal(calls.filter(call => call.method === "POST").length, 1)
  })
  console.log(`${results.length} update browser scenarios passed`)
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
