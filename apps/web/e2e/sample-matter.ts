// Sample matter regression against a separately started Vite/preview server.
// Setup's last step offers the sample matter, pre-selected only for an empty
// workspace; finishing with it adds the sample once, finishing without it adds
// nothing, and a failed add shows the error with a retry without undoing setup.
// Settings → Matter defaults adds it later and is then disabled. The matters
// list labels it. All API traffic is intercepted; this never reaches a live
// service or account.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
type Page = import("@playwright/test").Page
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: { name: string; status: string; error?: string }[] = []

type Matter = { id: string; title: string; reference?: string; jurisdictions?: string[]; sample?: boolean; dir: string; created_at: number }
const existing: Matter = { id: "existing-1", title: "Existing synthetic matter", dir: "/fixture/existing-1", created_at: 1 }
const sample: Matter = { id: "aldgate-mills-engagement-1", title: "Aldgate Mills — Engagement", reference: "A&C/2026-0042", jurisdictions: ["EW"], sample: true, dir: "/fixture/sample", created_at: 2 }

function fixture(setup: "done" | "details", matters: Matter[], sampleFailures = 0) {
  return { setup, matters: [...matters], sampleFailures, sampleRequests: 0, unexpected: [] as string[], errors: [] as string[] }
}

async function open(state: ReturnType<typeof fixture>, width = 1440, height = 1000) {
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 600, reducedMotion: "reduce" })
  await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  await context.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if (!api && url.origin === new URL(baseURL!).origin) return route.continue()
    if (!api) { state.unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": new URL(baseURL!).origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (request.method() === "OPTIONS") return route.fulfill({ status: 204, headers })
    const reply = (body: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(body) })
    const method = request.method()
    if (method === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    // Settings opens on AI setup, which e2e/ai-setup.ts covers; refuse its traffic.
    if (pathname === "/auth/reuse" || pathname.startsWith("/api/dochaus/provider-candidate")) return route.abort()
    if (method === "POST" && pathname === "/sample-matter") {
      state.sampleRequests++
      if (state.sampleFailures > 0) { state.sampleFailures--; return reply({ error: "Could not add the sample matter." }, 500) }
      const found = state.matters.find(matter => matter.sample)
      if (found) return reply({ matter: found, created: false })
      state.matters.push(sample)
      return reply({ matter: sample, created: true }, 201)
    }
    if (pathname === "/preferences" && method === "PUT") return reply({ ...DEFAULT_DRAFTING, ...(request.postDataJSON() as object) })
    if (pathname === "/setup" && method === "PATCH") {
      state.setup = "done"
      return reply({ version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" })
    }
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    const provider = { id: "fixture", name: "Fixture Models", integrationID: "fixture", activation: "enabled", package: "@opencode/ai/providers/openai-compatible" }
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply(state.setup === "done"
      ? { version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" }
      : { version: 1, ready: true, completed: ["models", "research"], skipped: [], next: "details" })
    if (pathname === "/preferences") return reply({ ...DEFAULT_DRAFTING, attorney: "Fixture Lawyer" })
    if (pathname === "/jurisdictions") return reply([{ code: "EW", name: "England & Wales", citationStyle: "oscola" }])
    if (pathname === "/matters") return reply(state.matters)
    if (pathname === "/playbooks") return reply([])
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/engine/config") return reply({ model: "fixture/primary", small_model: "fixture/fast" })
    if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/fixture" }, data: [provider] })
    if (pathname === "/api/model") return reply({ location: { directory: "/fixture" }, data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
    if (pathname === "/api/integration") return reply({ location: { directory: "/fixture" }, data: [{ id: "fixture", name: "Fixture", methods: [], connections: [{ type: "credential", id: "fixture-account", label: "Fixture", method: "key" }] }] })
    state.unexpected.push(`${method} ${pathname}`)
    return route.abort()
  })
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  page.on("pageerror", (error: Error) => state.errors.push(error.message))
  await page.goto(baseURL!)
  return { page, context }
}

async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}
const shot = (page: Page, name: string) => page.screenshot({ path: path.join(evidence!, `${name}.png`) })
const setupDialog = (page: Page) => page.getByRole("dialog", { name: "Let's get you set up" })
const sampleChoice = (page: Page) => setupDialog(page).getByRole("checkbox", { name: "Add a sample matter to explore", exact: true })
const sampleRow = (page: Page) => page.locator(".matters-index-list .list-row").filter({ hasText: sample.title })
async function finish(page: Page) {
  await setupDialog(page).getByRole("button", { name: /^Save and finish/ }).click()
}

try {
  for (const [width, height] of [[1440, 1000], [400, 860]] as const) {
    await scenario(`An empty workspace finishes setup with the pre-selected sample matter (${width}x${height})`, async () => {
      const state = fixture("details", [])
      const { page, context } = await open(state, width, height)
      try {
        await expect(sampleChoice(page)).toBeChecked()
        await expect(sampleChoice(page)).toHaveAccessibleDescription(/fictional engagement letter/)
        await sampleChoice(page).scrollIntoViewIfNeeded()
        await shot(page, `onboarding-sample-choice-${width}x${height}`)
        await finish(page)
        await expect(setupDialog(page)).toHaveCount(0)
        await expect(sampleRow(page)).toHaveCount(1)
        await expect(sampleRow(page).locator(".badge")).toHaveText("Sample")
        await shot(page, `matters-sample-labelled-${width}x${height}`)
        expect(state.sampleRequests).toBe(1)
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } finally { await context.close() }
    })
  }

  await scenario("Finishing setup with the sample unchecked adds no matter", async () => {
    const state = fixture("details", [])
    const { page, context } = await open(state)
    try {
      await sampleChoice(page).uncheck()
      await finish(page)
      await expect(setupDialog(page)).toHaveCount(0)
      await expect(page.getByRole("heading", { name: "No matters yet" })).toBeVisible()
      expect(state.sampleRequests).toBe(0)
      expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
    } finally { await context.close() }
  })

  await scenario("A workspace with matters leaves the sample unselected", async () => {
    const state = fixture("details", [existing])
    const { page, context } = await open(state)
    try {
      await expect(sampleChoice(page)).not.toBeChecked()
      await finish(page)
      await expect(setupDialog(page)).toHaveCount(0)
      expect(state.sampleRequests).toBe(0)
      expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
    } finally { await context.close() }
  })

  await scenario("A failed sample add keeps setup saved and offers a retry", async () => {
    const state = fixture("details", [], 1)
    const { page, context } = await open(state)
    try {
      await finish(page)
      const alert = setupDialog(page).getByRole("alert")
      await expect(alert).toContainText("Could not add the sample matter. Your setup is saved.")
      expect(state.setup).toBe("done")
      await shot(page, "onboarding-sample-failed")
      await alert.getByRole("button", { name: "Retry", exact: true }).click()
      await expect(setupDialog(page)).toHaveCount(0)
      await expect(sampleRow(page)).toHaveCount(1)
      expect(state.sampleRequests).toBe(2)
      expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
    } finally { await context.close() }
  })

  await scenario("Settings adds the sample matter later, then disables the action", async () => {
    const state = fixture("done", [existing])
    const { page, context } = await open(state)
    try {
      await expect(page.locator(".matters-index-list .list-row")).toHaveCount(1)
      await page.getByRole("button", { name: "Settings", exact: true }).click()
      const modal = page.getByRole("dialog", { name: "Settings", exact: true })
      await modal.getByRole("tab", { name: "Matter defaults", exact: true }).click()
      const add = modal.getByRole("button", { name: "Add sample matter", exact: true })
      await expect(add).toBeEnabled()
      await shot(page, "settings-sample-add")
      await add.click()
      const added = modal.getByRole("button", { name: "Sample matter added", exact: true })
      await expect(added).toBeDisabled()
      await expect(added).toHaveAccessibleDescription(/The sample matter is in your matters list/)
      await shot(page, "settings-sample-added")
      await modal.getByRole("button", { name: "Close", exact: true }).click()
      await expect(sampleRow(page)).toHaveCount(1)
      expect(state.sampleRequests).toBe(1)

      // Reopened, Settings still knows the sample exists.
      await page.getByRole("button", { name: "Settings", exact: true }).click()
      await modal.getByRole("tab", { name: "Matter defaults", exact: true }).click()
      await expect(modal.getByRole("button", { name: "Sample matter added", exact: true })).toBeDisabled()
      expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
    } finally { await context.close() }
  })
} finally {
  await writeFile(path.join(evidence, "sample-matter-results.json"), JSON.stringify(results, null, 2) + "\n")
  await browser.close()
}
console.log(`PASS ${results.length} sample matter scenarios`)
