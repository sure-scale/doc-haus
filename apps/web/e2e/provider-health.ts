// Synthetic browser acceptance: all API traffic is intercepted; no provider is contacted.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import type { Locator, Page } from "@playwright/test"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
const appOrigin = new URL(baseURL).origin
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: { name: string; status: string; error?: string }[] = []
const errors: string[] = []
const unexpected: string[] = []
const requests: { scenario: string; method: string; path: string; body?: unknown }[] = []
type Outcome = { ok: true } | { ok: false; kind: "auth" | "model" | "unavailable"; expired?: boolean }
type State = { outcome: Outcome; setup: "done" | "models"; config: { model: string; small_model: string }; events: unknown[] }

async function fixture(name: string, options: { accepted?: boolean; width?: number; outcome?: Outcome; setup?: State["setup"]; distinct?: boolean } = {}) {
  const context = await browser.newContext({ viewport: { width: options.width ?? 1440, height: 1000 }, reducedMotion: "reduce" })
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  page.on("pageerror", error => errors.push(`${name}: ${error.message}`))
  if (options.accepted !== false) await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  const state: State = { outcome: options.outcome ?? { ok: true }, setup: options.setup ?? "done", config: { model: "opencode-go/primary", small_model: `opencode-go/${options.distinct ? "fast" : "primary"}` }, events: [] }
  await context.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || pathname.startsWith("/api/") || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if (!api && url.origin === appOrigin) return route.continue()
    if (!api) { unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": appOrigin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (request.method() === "OPTIONS") return route.fulfill({ status: 204, headers })
    requests.push({ scenario: name, method: request.method(), path: pathname, ...(request.postData() ? { body: request.postDataJSON() } : {}) })
    const reply = (body: unknown) => route.fulfill({ status: 200, headers, contentType: "application/json", body: JSON.stringify(body) })
    if (request.method() === "POST" && pathname === "/api/dochaus/provider-test") return reply(state.outcome)
    if (request.method() === "POST" && pathname === "/api/integration/opencode-go/connect/key") {
      expect(request.postDataJSON()).toEqual({ key: "synthetic-replacement-key", answer: {} })
      state.outcome = { ok: true }
      return route.fulfill({ status: 204, headers })
    }
    if (request.method() !== "GET") { unexpected.push(`${name}: ${request.method()} ${pathname}`); return route.abort() }
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply({ version: 1, ready: true, completed: state.setup === "done" ? ["models", "research", "details"] : [], skipped: [], next: state.setup })
    if (pathname === "/matters") return reply([])
    if (pathname === "/jurisdictions") return reply([{ code: "england-wales", name: "England and Wales" }])
    if (pathname === "/preferences") return reply({})
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/engine/config") return reply(state.config)
    if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/synthetic" }, data: [{ id: "opencode-go", name: "OpenCode Go", integrationID: "opencode-go", activation: "enabled" }] })
    if (pathname === "/api/model") return reply({ location: { directory: "/synthetic" }, data: ["primary", "fast"].map(id => ({ providerID: "opencode-go", id, name: id, enabled: true })) })
    if (pathname === "/api/integration") return reply({ location: { directory: "/synthetic" }, data: [{ id: "opencode-go", name: "OpenCode Go", methods: [{ type: "key", label: "API key" }], connections: [{ type: "credential", id: "synthetic-account", label: "Synthetic", method: "key" }] }] })
    if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: state.events.map(event => `data: ${JSON.stringify(event)}\n\n`).join("") })
    if (pathname === "/api/session/owned-session") return reply({ data: { id: "owned-session", model: { providerID: "opencode-go", id: "primary" } } })
    if (pathname === "/api/session/owned-session/message") return reply({ data: [] })
    unexpected.push(`${name}: ${request.method()} ${pathname}`)
    return route.abort()
  })
  await page.goto(baseURL!)
  return { page, context, state, probes: () => requests.filter(request => request.scenario === name && request.path === "/api/dochaus/provider-test") }
}

async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}
async function hitTarget(locator: Locator, minimum = 40) {
  await locator.scrollIntoViewIfNeeded()
  const target = await locator.evaluate((node: HTMLElement) => {
    const rect = node.getBoundingClientRect()
    return { width: rect.width, height: rect.height, uncovered: node.contains(document.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2)) }
  })
  expect(target).toMatchObject({ uncovered: true })
  expect(target.width).toBeGreaterThanOrEqual(minimum)
  expect(target.height).toBeGreaterThanOrEqual(minimum)
}
async function noOverflow(page: Page) {
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true)
}

try {
  await scenario("No provider check before launch acknowledgement; equal model selections are checked once", async () => {
    const f = await fixture("notice-gate", { accepted: false })
    try {
      await expect(f.page.getByRole("dialog", { name: "Before you begin" })).toBeVisible()
      await f.page.waitForTimeout(250)
      expect(requests.filter(request => request.scenario === "notice-gate")).toEqual([])
      await f.page.getByRole("button", { name: "I accept", exact: true }).click()
      await expect.poll(() => f.probes().length).toBe(1)
      await expect(f.page.getByRole("button", { name: "Settings", exact: true })).toBeVisible()
      expect(f.probes()[0].body).toEqual({ providerID: "opencode-go", modelID: "primary" })
      await expect(f.page.locator(".provider-connection-banner")).toHaveCount(0)
    } finally { await f.context.close() }
  })
  await scenario("Incomplete setup does not run a provider probe", async () => {
    const f = await fixture("setup-gate", { setup: "models" })
    try {
      await expect(f.page.getByRole("heading", { name: "Choose your models", exact: true })).toBeVisible()
      await f.page.waitForTimeout(250)
      expect(f.probes()).toEqual([])
    } finally { await f.context.close() }
  })
  await scenario("Distinct primary and fast models both verify; Settings shows verified access", async () => {
    const f = await fixture("healthy", { distinct: true })
    try {
      await expect.poll(() => f.probes().length).toBe(2)
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      const modal = f.page.getByRole("dialog", { name: "Settings", exact: true })
      await modal.getByRole("button", { name: "Providers", exact: true }).click()
      await expect(modal.locator(".settings-provider-health")).toHaveText("Connection verified")
      await expect(f.page.locator(".provider-connection-banner")).toHaveCount(0)
      await f.page.screenshot({ path: path.join(evidence, "healthy-settings.png") })
    } finally { await f.context.close() }
  })
  for (const width of [1440, 390]) await scenario(`Expired key recovery works at ${width}px with visible pointer targets and no overflow`, async () => {
    const f = await fixture(`expired-${width}`, { width, outcome: { ok: false, kind: "auth", expired: true } })
    try {
      const banner = f.page.locator(".provider-connection-banner")
      await expect(banner).toHaveAttribute("role", "alert")
      await expect(banner).toContainText("OpenCode Go needs attention")
      await expect(banner).toContainText("The saved API key has expired")
      await expect(banner).toContainText("Signing in again will not renew this key")
      const reconnect = banner.getByRole("button", { name: "Reconnect provider", exact: true })
      await hitTarget(reconnect)
      await hitTarget(banner.getByRole("button", { name: "Check again", exact: true }))
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, `expired-${width}.png`) })
      await reconnect.click()
      const modal = f.page.getByRole("dialog", { name: "Settings", exact: true })
      await expect(modal.getByRole("heading", { name: "Connect a provider", exact: true })).toBeVisible()
      await expect(modal.getByLabel(/^Provider/)).toHaveValue("opencode-go")
      await expect(modal.locator(".settings-provider-health")).toContainText("API key has expired")
      await modal.getByLabel(/^Sign-in method/).selectOption("0")
      await modal.getByLabel("API key", { exact: true }).fill("synthetic-replacement-key")
      const connect = modal.getByRole("button", { name: "Connect", exact: true })
      await hitTarget(connect, 24)
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, `reconnect-${width}.png`) })
      await connect.click()
      await expect(modal.locator(".settings-provider-health")).toHaveText("Connection verified")
      await expect(banner).toHaveCount(0)
      expect(f.probes().length).toBe(2)
      expect(f.probes()[1].body).toEqual({ providerID: "opencode-go", modelID: "primary", force: true })
      await modal.getByRole("button", { name: "Close", exact: true }).click()
      await noOverflow(f.page)
    } finally { await f.context.close() }
  })
  await scenario("Transient provider failure asks to check availability and Check again recovers", async () => {
    const f = await fixture("transient", { outcome: { ok: false, kind: "unavailable" } })
    try {
      const banner = f.page.locator(".provider-connection-banner")
      await expect(banner).toContainText("Check your network and provider availability")
      await expect(banner.getByRole("button", { name: "Reconnect provider", exact: true })).toHaveCount(0)
      f.state.outcome = { ok: true }
      await banner.getByRole("button", { name: "Check again", exact: true }).click()
      await expect.poll(() => f.probes().length).toBe(2)
      await expect(banner).toHaveCount(0)
    } finally { await f.context.close() }
  })
  await scenario("Model access failure asks for model settings without claiming expired credentials", async () => {
    const f = await fixture("model-access", { outcome: { ok: false, kind: "model" } })
    try {
      const banner = f.page.locator(".provider-connection-banner")
      await expect(banner).toContainText("The selected model is unavailable")
      await expect(banner.getByRole("button", { name: "Reconnect provider", exact: true })).toHaveCount(0)
      await expect(banner.getByRole("button", { name: "Open provider settings", exact: true })).toBeVisible()
    } finally { await f.context.close() }
  })
  await scenario("Native event authentication rejection affects only the owned session and directory", async () => {
    const f = await fixture("native-events")
    try {
      await expect.poll(() => f.probes().length).toBe(1)
      const failure = (sessionID: string, directory: string) => ({ id: `synthetic-${sessionID}-${directory}`, type: "session.step.failed", location: { directory }, data: { sessionID, error: { type: "provider.auth", message: "The synthetic API key expired" } } })
      async function receive(events: unknown[]) {
        f.state.events = events
        // Exercise the same native subscription used by the matter UI, while the
        // visible application banner observes its shared health monitor.
        await f.page.evaluate(async () => {
          const eventsPath = "/src/api/opencode/events.ts"
          const clientPath = "/src/api/opencode/client.ts"
          const [{ subscribeEvents }, { matterClient }] = await Promise.all([import(eventsPath), import(clientPath)])
          const controller = new AbortController()
          const timer = setTimeout(() => controller.abort(), 250)
          try { await subscribeEvents(matterClient("/synthetic/matter"), () => {}, controller.signal, { acceptsSession: (id: string) => id === "owned-session" }) }
          finally { clearTimeout(timer) }
        })
      }
      await receive([failure("owned-session", "/synthetic/other-matter"), failure("foreign-session", "/synthetic/matter")])
      await expect(f.page.locator(".provider-connection-banner")).toHaveCount(0)
      expect(requests.filter(request => request.scenario === "native-events" && request.path.startsWith("/api/session/"))).toEqual([])
      await receive([failure("owned-session", "/synthetic/matter")])
      await expect(f.page.locator(".provider-connection-banner")).toContainText("The saved API key has expired")
      expect(requests.some(request => request.scenario === "native-events" && request.path === "/api/session/owned-session")).toBe(true)
      await f.page.screenshot({ path: path.join(evidence, "native-auth-rejection.png") })
    } finally { await f.context.close() }
  })
  expect(unexpected).toEqual([])
  expect(errors).toEqual([])
} finally {
  await writeFile(path.join(evidence, "provider-health-results.json"), JSON.stringify({ results, errors, unexpected, requests }, null, 2) + "\n")
  await browser.close()
}
