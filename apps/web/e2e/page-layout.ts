// Page layout regression against a separately started Vite/preview server.
// Settings → Drafting exercises page layout controls and previews. Saves send flat keys.
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
  await context.addInitScript(date => localStorage.setItem("dh.launch-notice.v1", date), today.toDateString())
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
    if (pathname === "/matters" || pathname === "/playbooks" || pathname === "/clients") return reply([])
    if (pathname === "/templates") return reply({ dir: "/fixture/templates", templates: [] })
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
  if (process.env.DOCHAUS_E2E_SCENARIOS && !new RegExp(process.env.DOCHAUS_E2E_SCENARIOS).test(name)) return
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}
const preview = (page: Page) => ({
  header: page.getByRole("figure", { name: /page preview$/ }).locator('[data-preview="header"] p'),
  footer: page.getByRole("figure", { name: /page preview$/ }).locator('[data-preview="footer"]'),
})
const choice = (page: Page, group: string, label: string) => page.getByRole("radiogroup", { name: group }).getByRole("radio", { name: label, exact: true })
const segment = (page: Page, group: string, label: string) => page.getByRole("group", { name: group }).getByRole("button", { name: label, exact: true })
const check = (page: Page, label: string) => page.getByRole("checkbox", { name: label, exact: true })
const shot = (page: Page, name: string) => page.screenshot({ path: path.join(evidence!, `${name}.png`) })
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
        await modal.getByRole("group", { name: "Preview document" }).getByRole("button", { name: "Advice" }).click()
        await expect(preview(page).header).toHaveText(["Fixture advice legend", "Fixture Partners LLP draft: 20 October 2026"])
        await modal.getByRole("group", { name: "Preview document" }).getByRole("button", { name: "Drafting notes" }).click()
        await expect(preview(page).header).toHaveText(["Internal drafting notes — not for circulation"])
        await segment(page, "Drafting notes marking", "Off").click()
        await expect(preview(page).header).toHaveText(["No header"])
        await segment(page, "Drafting notes marking", "On").click()
        await modal.getByRole("group", { name: "Preview document" }).getByRole("button", { name: "Agreement" }).click()
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

  }

} finally {
  await writeFile(path.join(evidence, "page-layout-results.json"), JSON.stringify(results, null, 2) + "\n")
  await browser.close()
}
console.log(`PASS ${results.length} page layout scenarios`)
