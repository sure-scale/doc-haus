// Page layout regression against a separately started Vite/preview server.
// Settings → Drafting sets every page layout control; onboarding's details step
// asks the two legends, marks suggestions from the default jurisdictions without
// choosing any, and previews an agreement page. Each save must send flat keys.
// All API traffic is intercepted; this never reaches a live service or account.
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
// The draft line is dated when a draft is first made; the preview shows today.
const today = new Date("2026-10-20T10:00:00Z")
const jurisdictions = [
  { code: "EW", name: "England & Wales", citationStyle: "oscola" },
  { code: "FR", name: "France", citationStyle: "fr" },
  { code: "US-NY", name: "New York", citationStyle: "bluebook" },
]
const sizes = [[1440, 1000], [400, 860]] as const

function fixture(setup: "done" | "details") {
  return {
    setup,
    preferences: { ...DEFAULT_DRAFTING, attorney: "Fixture Lawyer", firm: setup === "done" ? "Fixture Partners LLP" : "", dateFormat: "day-month-year" as const },
    puts: [] as Record<string, unknown>[],
    unexpected: [] as string[],
    errors: [] as string[],
  }
}

async function open(state: ReturnType<typeof fixture>, width: number, height: number) {
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 600, reducedMotion: "reduce", timezoneId: "Europe/London" })
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
    const reply = (body: unknown) => route.fulfill({ status: 200, headers, contentType: "application/json", body: JSON.stringify(body) })
    const method = request.method()
    if (method === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    // Settings opens on AI setup, which e2e/ai-setup.ts covers; refuse its traffic.
    if (pathname === "/auth/reuse" || pathname.startsWith("/api/dochaus/provider-candidate")) return route.abort()
    if (pathname === "/preferences" && method === "PUT") {
      const body = request.postDataJSON() as Record<string, unknown>
      state.puts.push(body)
      state.preferences = { ...state.preferences, ...body }
      return reply(state.preferences)
    }
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
    if (pathname === "/preferences") return reply(state.preferences)
    if (pathname === "/jurisdictions") return reply(jurisdictions)
    if (pathname === "/matters" || pathname === "/playbooks") return reply([])
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
  await page.clock.setFixedTime(today)
  page.on("pageerror", (error: Error) => state.errors.push(error.message))
  await page.goto(baseURL!)
  return { page, context }
}

async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}
const preview = (page: Page) => ({
  header: page.getByRole("figure", { name: "Agreement page preview" }).locator('[data-preview="header"] p'),
  footer: page.getByRole("figure", { name: "Agreement page preview" }).locator('[data-preview="footer"]'),
})
const choice = (page: Page, group: string, label: string) => page.getByRole("radiogroup", { name: group }).getByRole("radio", { name: label, exact: true })
const segment = (page: Page, group: string, label: string) => page.getByRole("group", { name: group }).getByRole("button", { name: label, exact: true })
const check = (page: Page, label: string) => page.getByRole("checkbox", { name: label, exact: true })
const shot = (page: Page, name: string) => page.screenshot({ path: path.join(evidence!, `${name}.png`) })
// The picker opens below its trigger; start with the trigger at the top so the
// list is in view at a narrow height too.
async function openJurisdictions(page: Page) {
  const trigger = page.locator("#onboarding-jurisdictions")
  await trigger.evaluate((node: HTMLElement) => node.scrollIntoView({ block: "start" }))
  await trigger.click()
}
const layoutKeys = ["pageNumbers", "pageNumberPosition", "numberFirstPage", "draftLine", "draftLineFirm", "agreementLegend", "adviceLegend", "notesMarking"]

try {
  for (const [width, height] of sizes) {
    await scenario(`Settings → Drafting sets each page layout control and saves flat keys (${width}x${height})`, async () => {
      const state = fixture("done")
      const { page, context } = await open(state, width, height)
      try {
        await page.getByRole("button", { name: "Settings", exact: true }).click()
        const modal = page.getByRole("dialog", { name: "Settings", exact: true })
        await modal.getByRole("tab", { name: "Drafting", exact: true }).click()
        const heading = modal.getByRole("heading", { name: "Page layout", exact: true })
        await heading.scrollIntoViewIfNeeded()
        await expect(segment(page, "Page numbers", "2")).toHaveAttribute("aria-pressed", "true")
        await expect(segment(page, "Page numbers", "- 2 -")).toHaveAttribute("aria-pressed", "false")
        await expect(segment(page, "Page number position", "Center")).toHaveAttribute("aria-pressed", "true")
        await expect(check(page, "Number the first page")).toBeChecked()
        await expect(check(page, "Number the first page")).toHaveAccessibleDescription("Letters always start numbering on page 2.")
        await expect(segment(page, "Draft line", "On")).toHaveAttribute("aria-pressed", "true")
        await expect(check(page, "Include the firm name")).toBeChecked()
        await expect(choice(page, "Agreement legend", "None")).toBeChecked()
        await expect(choice(page, "Advice legend", "None")).toBeChecked()
        await expect(segment(page, "Drafting notes marking", "On")).toHaveAttribute("aria-pressed", "true")
        await expect(modal.getByText('On agreements, advice and letters: "Fixture Partners LLP draft: 20 October 2026", dated when the draft is first made.')).toBeVisible()
        await expect(preview(page).header).toHaveText(["Fixture Partners LLP draft: 20 October 2026"])
        await expect(preview(page).footer).toHaveText("1")
        await expect(modal.getByText(/without prejudice/i)).toHaveCount(0)
        await shot(page, `settings-page-layout-${width}x${height}-defaults`)

        await segment(page, "Page numbers", "Off").click()
        await expect(page.getByRole("group", { name: "Page number position" })).toHaveCount(0)
        await expect(check(page, "Number the first page")).toHaveCount(0)
        await expect(preview(page).footer).toHaveText("No page number")
        await segment(page, "Page numbers", "Page 2").click()
        await segment(page, "Page number position", "Right").click()
        await expect(preview(page).footer).toHaveText("Page 1")
        await expect(preview(page).footer).toHaveClass(/is-right/)
        await segment(page, "Page numbers", "- 2 -").click()
        await expect(preview(page).footer).toHaveText("- 1 -")
        await check(page, "Number the first page").uncheck()
        await expect(preview(page).footer).toHaveText("Numbered from page 2")
        await expect(preview(page).footer).toHaveClass(/is-right/)
        // Hidden with page numbers off, and still unticked when they return.
        await segment(page, "Page numbers", "Off").click()
        await expect(check(page, "Number the first page")).toHaveCount(0)
        await segment(page, "Page numbers", "- 2 -").click()
        await expect(check(page, "Number the first page")).not.toBeChecked()
        await expect(preview(page).footer).toHaveText("Numbered from page 2")
        await modal.locator(".settings-form-row", { has: page.getByRole("group", { name: "Page numbers" }) }).evaluate((node: HTMLElement) => node.scrollIntoView({ block: "start" }))
        await shot(page, `settings-page-layout-${width}x${height}-numbering`)
        await choice(page, "Agreement legend", "Subject to contract").check()
        await expect(preview(page).header).toHaveText(["Subject to contract", "Fixture Partners LLP draft: 20 October 2026"])
        await choice(page, "Advice legend", "Custom text").check()
        const custom = modal.getByLabel("Advice legend text", { exact: true })
        await custom.fill("x".repeat(130))
        await expect(custom).toHaveValue("x".repeat(120))
        await custom.fill("Fixture advice legend")
        await choice(page, "Advice legend", "Privileged & confidential").check()
        await choice(page, "Advice legend", "Custom text").check()
        await expect(custom).toHaveValue("Fixture advice legend")
        await page.getByRole("radiogroup", { name: "Agreement legend" }).evaluate((node: HTMLElement) => node.scrollIntoView({ block: "start" }))
        await shot(page, `settings-page-layout-${width}x${height}-legends`)
        // Without a firm name there is nothing to include.
        await modal.getByLabel("Firm", { exact: true }).fill("")
        await expect(check(page, "Include the firm name")).toHaveCount(0)
        await modal.getByLabel("Firm", { exact: true }).fill("Other Fixture LLP")
        await expect(modal.getByText('On agreements, advice and letters: "Other Fixture LLP draft: 20 October 2026", dated when the draft is first made.')).toBeVisible()
        await check(page, "Include the firm name").uncheck()
        await expect(modal.getByText('On agreements, advice and letters: "Draft: 20 October 2026", dated when the draft is first made.')).toBeVisible()
        await expect(preview(page).header).toHaveText(["Subject to contract", "Draft: 20 October 2026"])
        await segment(page, "Draft line", "Off").click()
        await expect(check(page, "Include the firm name")).toHaveCount(0)
        await expect(preview(page).header).toHaveText(["Subject to contract"])
        await segment(page, "Draft line", "On").click()
        await expect(check(page, "Include the firm name")).not.toBeChecked()
        await expect(preview(page).header).toHaveText(["Subject to contract", "Draft: 20 October 2026"])
        await segment(page, "Draft line", "Off").click()
        await segment(page, "Drafting notes marking", "Off").click()
        await preview(page).footer.scrollIntoViewIfNeeded()
        await shot(page, `settings-page-layout-${width}x${height}-changed`)

        await modal.getByRole("button", { name: "Save", exact: true }).click()
        await expect(page.getByRole("status")).toContainText("Drafting preferences saved.")
        expect(state.puts).toHaveLength(1)
        const body = state.puts[0]!
        expect(body).toMatchObject({ firm: "Other Fixture LLP", pageNumbers: "dashed", pageNumberPosition: "right", numberFirstPage: false, draftLine: false, draftLineFirm: false, agreementLegend: "Subject to contract", adviceLegend: "Fixture advice legend", notesMarking: false })
        // Flat keys: the ingest save merges top-level keys only, so a nested
        // object would replace its siblings. Only the existing two stay nested.
        expect(Object.entries(body).filter(([, value]) => value !== null && typeof value === "object").map(([key]) => key).sort()).toEqual(["approvedSources", "telemetry"])
        for (const key of layoutKeys) expect(Object.hasOwn(body, key)).toBe(true)
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } finally { await context.close() }
    })

    await scenario(`Onboarding marks suggested legends from default jurisdictions, chooses none, previews and saves flat keys (${width}x${height})`, async () => {
      const state = fixture("details")
      const { page, context } = await open(state, width, height)
      try {
        const dialog = page.getByRole("dialog", { name: "Let's get you set up" })
        await expect(dialog.getByRole("heading", { name: "Make doc.haus yours", exact: true })).toBeVisible()
        await expect(choice(page, "Agreement legend", "None")).toBeChecked()
        await expect(choice(page, "Advice legend", "None")).toBeChecked()
        await expect(dialog.getByText("Suggested", { exact: true })).toHaveCount(0)
        await expect(preview(page).header).toHaveText(["Draft: 20 October 2026"])
        await dialog.getByLabel("Firm", { exact: true }).fill("Fixture Partners LLP")
        await expect(preview(page).header).toHaveText(["Fixture Partners LLP draft: 20 October 2026"])
        await expect(preview(page).footer).toHaveText("1")

        await openJurisdictions(page)
        await page.getByRole("checkbox", { name: "England & Wales EW" }).check()
        await page.keyboard.press("Escape")
        await expect(choice(page, "Agreement legend", "Subject to contract Suggested")).toBeVisible()
        await expect(choice(page, "Advice legend", "Privileged & confidential Suggested")).toBeVisible()
        await expect(dialog.getByText("Suggested", { exact: true })).toHaveCount(2)
        await expect(choice(page, "Agreement legend", "None")).toBeChecked()
        await expect(choice(page, "Advice legend", "None")).toBeChecked()
        await expect(dialog.getByText(/without prejudice/i)).toHaveCount(0)
        await openJurisdictions(page)
        await page.getByRole("checkbox", { name: "New York US-NY" }).check()
        await page.getByRole("checkbox", { name: "France FR" }).check()
        await page.keyboard.press("Escape")
        await expect(dialog.getByText("Suggested", { exact: true })).toHaveCount(4)
        await expect(choice(page, "Agreement legend", "Draft — for discussion purposes only Suggested")).toBeVisible()
        await expect(choice(page, "Advice legend", "Privileged & confidential — attorney-client communication Suggested")).toBeVisible()

        await choice(page, "Agreement legend", "Subject to contract Suggested").check()
        await expect(preview(page).header).toHaveText(["Subject to contract", "Fixture Partners LLP draft: 20 October 2026"])
        await page.getByRole("heading", { name: "Legends", exact: true }).scrollIntoViewIfNeeded()
        await shot(page, `onboarding-legends-${width}x${height}`)
        await preview(page).footer.scrollIntoViewIfNeeded()
        await shot(page, `onboarding-preview-${width}x${height}`)

        await dialog.getByRole("button", { name: "Save and finish", exact: true }).click()
        await expect(dialog).toHaveCount(0)
        expect(state.puts).toEqual([{ attorney: "Fixture Lawyer", firm: "Fixture Partners LLP", agreementLegend: "Subject to contract", adviceLegend: "", dateFormat: "day-month-year" }])
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } finally { await context.close() }
    })
  }

  await scenario("Onboarding left alone keeps both legends at none", async () => {
    const state = fixture("details")
    const { page, context } = await open(state, 1440, 1000)
    try {
      const dialog = page.getByRole("dialog", { name: "Let's get you set up" })
      await openJurisdictions(page)
      await page.getByRole("checkbox", { name: "New York US-NY" }).check()
      await page.keyboard.press("Escape")
      await expect(dialog.getByText("Suggested", { exact: true })).toHaveCount(2)
      await dialog.getByRole("button", { name: "Save and finish", exact: true }).click()
      await expect(dialog).toHaveCount(0)
      expect(state.puts).toEqual([{ attorney: "Fixture Lawyer", firm: "", agreementLegend: "", adviceLegend: "", dateFormat: "month-day-year" }])
      expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
    } finally { await context.close() }
  })
} finally {
  await writeFile(path.join(evidence, "page-layout-results.json"), JSON.stringify(results, null, 2) + "\n")
  await browser.close()
}
console.log(`PASS ${results.length} page layout scenarios`)
