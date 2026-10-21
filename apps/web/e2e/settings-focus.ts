// Desktop keyboard regression against a separately started Vite/preview server.
// All API traffic is intercepted; this never reaches a live service or account.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, reducedMotion: "reduce" })
const page = await context.newPage()
page.setDefaultTimeout(10000)
const errors: string[] = []
const unexpected: string[] = []
const requests: { method: string; path: string }[] = []
const results: { name: string; status: string; error?: string }[] = []
let config = { model: "fixture/primary", small_model: "fixture/fast" }
page.on("pageerror", (error: Error) => errors.push(error.message))
await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
await context.route("**/*", async route => {
  const request = route.request()
  const url = new URL(request.url())
  const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
  const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || pathname.startsWith("/api/") || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
  if (!api && url.origin === new URL(baseURL).origin) return route.continue()
  if (!api) { unexpected.push(request.url()); return route.abort() }
  const headers = { "access-control-allow-origin": new URL(baseURL).origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
  if (request.method() === "OPTIONS") return route.fulfill({ status: 204, headers })
  requests.push({ method: request.method(), path: pathname })
  const reply = (body: unknown) => route.fulfill({ status: 200, headers, contentType: "application/json", body: JSON.stringify(body) })
  if (request.method() === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
  if (request.method() === "PATCH" && pathname === "/engine/config") {
    config = { ...config, ...request.postDataJSON() as typeof config }
    return reply(config)
  }
  if (request.method() !== "GET") { unexpected.push(`${request.method()} ${pathname}`); return route.abort() }
  if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
  if (pathname === "/setup") return reply({ version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" })
  if (pathname === "/matters") return reply([])
  if (pathname === "/jurisdictions") return reply([{ code: "england-wales", name: "England and Wales" }, { code: "scotland", name: "Scotland" }])
  if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
  if (pathname === "/engine/config") return reply(config)
  if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/fixture" }, data: [{ id: "fixture", name: "Fixture Models", integrationID: "fixture", activation: "enabled", package: "@opencode/ai/providers/openai-compatible" }] })
  if (pathname === "/api/model") return reply({ location: { directory: "/fixture" }, data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
  if (pathname === "/api/integration") return reply({ location: { directory: "/fixture" }, data: [{ id: "fixture", name: "Fixture", methods: [], connections: [{ type: "credential", id: "fixture-account", label: "Fixture", method: "key" }] }] })
  unexpected.push(`${request.method()} ${pathname}`)
  return route.abort()
})

async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}
const opener = page.getByRole("button", { name: "Settings", exact: true })
const modal = page.getByRole("dialog", { name: "Settings", exact: true })
async function open() {
  await opener.focus()
  await page.keyboard.press("Enter")
  await expect(modal).toBeVisible()
  await expect(modal.getByRole("button", { name: "Close", exact: true })).toBeFocused()
}
async function expectContained() {
  expect(await modal.evaluate((node: HTMLDialogElement) => node.matches(":modal"))).toBe(true)
  // Chromium can place sequential navigation's wrap point on BODY. It must
  // never put focus on any interactive control behind the modal.
  expect(await page.evaluate(() => document.activeElement === document.body || Boolean(document.activeElement?.closest("dialog.settings-overlay")))).toBe(true)
}

try {
  await page.goto(baseURL)
  await expect(opener).toBeVisible()
  await scenario("Enter opens an accessible native modal; forward/backward Tab cannot reach the page; Escape restores opener", async () => {
    await open()
    await expect(modal.getByLabel("Primary model", { exact: true })).toBeEnabled()
    for (const key of ["Tab", "Shift+Tab"]) {
      for (let index = 0; index < 25; index++) { await page.keyboard.press(key); await expectContained() }
    }
    await page.keyboard.press("Escape")
    await expect(modal).toHaveCount(0)
    await expect(opener).toBeFocused()
  })
  await scenario("Nested jurisdiction picker remains usable; Escape dismisses it before Settings; unsaved model state survives", async () => {
    await open()
    await modal.getByLabel("Primary model", { exact: true }).selectOption("fixture/fast")
    await modal.getByRole("button", { name: "Matter defaults", exact: true }).click()
    const jurisdiction = modal.getByLabel("Jurisdictions", { exact: true })
    await jurisdiction.focus(); await page.keyboard.press("Enter")
    const search = modal.getByPlaceholder("Search jurisdictions")
    await expect(search).toBeFocused()
    await search.fill("Scot")
    await expect(modal.locator(".jx-option")).toHaveText(["Scotlandscotland"])
    await page.keyboard.press("Tab")
    await expect(modal.getByRole("checkbox", { name: "Scotland scotland", exact: true })).toBeFocused()
    await page.keyboard.press("Shift+Tab")
    await expect(search).toBeFocused()
    await page.screenshot({ path: path.join(evidence, "settings-jurisdictions.png") })
    await page.keyboard.press("Escape")
    await expect(search).toHaveCount(0)
    await expect(jurisdiction).toBeFocused()
    await expect(modal).toBeVisible()
    await modal.getByRole("button", { name: "Models", exact: true }).click()
    await expect(modal.getByLabel("Primary model", { exact: true })).toHaveValue("fixture/fast")
    expect(requests.filter(request => request.method !== "GET" && request.path !== "/api/dochaus/provider-test")).toEqual([])
    await modal.getByRole("button", { name: "Close", exact: true }).focus()
    await page.keyboard.press("Enter")
    await expect(opener).toBeFocused()
  })
  await scenario("Existing Save stays open and persists only the fixture; backdrop dismissal restores opener", async () => {
    await open()
    await expect(modal.getByLabel("Primary model", { exact: true })).toHaveValue("fixture/primary")
    await modal.getByLabel("Primary model", { exact: true }).selectOption("fixture/fast")
    const save = modal.getByRole("button", { name: "Save", exact: true })
    await save.focus(); await page.keyboard.press("Enter")
    await expect(page.getByRole("status")).toContainText("Models saved.")
    await expect(modal).toBeVisible()
    expect(config).toEqual({ model: "fixture/fast", small_model: "fixture/fast" })
    await expectContained()
    await page.screenshot({ path: path.join(evidence, "settings-modal.png") })
    await modal.click({ position: { x: 8, y: 8 } })
    await expect(modal).toHaveCount(0)
    await expect(opener).toBeFocused()
  })
  await scenario("Closing after an opener is detached does not focus a stale element", async () => {
    await open()
    await opener.evaluate((node: HTMLElement) => node.remove())
    await page.keyboard.press("Escape")
    await expect(modal).toHaveCount(0)
    expect(await page.evaluate(() => document.activeElement?.isConnected)).toBe(true)
  })
  expect(unexpected).toEqual([])
  expect(errors).toEqual([])
} finally {
  await writeFile(path.join(evidence, "settings-focus-results.json"), JSON.stringify({ results, errors, unexpected, requests }, null, 2) + "\n")
  await browser.close()
}
