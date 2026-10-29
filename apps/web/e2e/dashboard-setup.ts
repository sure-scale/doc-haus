// Default dashboard setup against a built web preview. All service traffic is
// synthetic and unknown requests fail closed; no provider or OS work is run.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = await import("@playwright/test")
type Page = import("@playwright/test").Page
type Route = import("@playwright/test").Route
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR")
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE, headless: true })
const results: { name: string; status: string; error?: string }[] = []
const screens: Record<string, unknown>[] = []
const NOW = Date.UTC(2026, 9, 29, 10)
const SAMPLE = { id: "setup-sample", title: "Kestrel Books · Engagement letter", sample: true, jurisdictions: ["EW"], dir: "/fixture/matters/sample", created_at: NOW }
const wordDefault = { available: false, origin: null, manifestUrl: null, platform: "macos", wordInstalled: true, canSetup: true, configured: false, registered: false, trusted: false, running: false, status: "not_configured" }
const packs = [{ code: "EW", name: "England & Wales", citationStyle: "fixture" }, { code: "SCT", name: "Scotland", citationStyle: "fixture" }, { code: "IE", name: "Ireland", citationStyle: "fixture" }]
function fixture() {
  return { ready: false, jurisdictions: [] as string[], preferences: { ...DEFAULT_DRAFTING }, failPreferences: false, holdPreferences: false, releasePreference: undefined as (() => void) | undefined, failMatters: false, word: { ...wordDefault }, templates: [] as { name: string; placeholders: unknown[] }[], matters: [] as (typeof SAMPLE)[], saved: [] as Record<string, unknown>[], samplePosts: 0, wordPosts: 0, unexpected: [] as string[], errors: [] as string[] }
}
type State = ReturnType<typeof fixture>
async function open(state: State, width = 1440, height = 1000) {
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", locale: "en-GB" })
  await context.clock.setFixedTime(new Date(NOW))
  await context.addInitScript((jurisdictions: string[]) => {
    localStorage.setItem("dh.launch-notice.v1", new Date().toDateString())
    localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1")
    if (jurisdictions.length && !localStorage.getItem("dochaus.prefs")) localStorage.setItem("dochaus.prefs", JSON.stringify({ defaultJurisdictions: jurisdictions }))
  }, state.jurisdictions)
  const origin = new URL(baseURL!).origin
  await context.route("**/*", async (route: Route) => {
    const request = route.request(), url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const method = request.method()
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if (!api && url.origin === origin) return route.continue()
    if (!api) { state.unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    const reply = (body: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(body) })
    if (method === "OPTIONS") return route.fulfill({ status: 204, headers })
    if (pathname === "/preferences" && method === "PUT") {
      const patch = request.postDataJSON() as Record<string, unknown>
      if (state.failPreferences) return reply({ error: `Synthetic failed save: ${patch.searchApiKey ?? "settings unavailable"}` }, 500)
      if (state.holdPreferences) await new Promise<void>(resolve => { state.releasePreference = resolve })
      state.saved.push(patch); state.preferences = { ...state.preferences, ...patch }
      return reply(state.preferences)
    }
    if (pathname === "/sample-matter" && method === "POST") {
      state.samplePosts++
      if (!state.matters.some(item => item.sample)) state.matters.push(SAMPLE)
      return reply({ matter: SAMPLE })
    }
    if (pathname === "/word/setup" && method === "POST") {
      state.wordPosts++; state.word = { ...state.word, status: "ready", available: true, configured: true, registered: true, trusted: true }
      return reply(state.word)
    }
    if (pathname === "/setup" && method === "PATCH") return reply({ version: 1, ready: state.ready, completed: state.ready ? ["models"] : [], skipped: [], next: state.ready ? "research" : "models" })
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "POST") return reply({ sessionID: "synthetic-candidate" })
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "DELETE") return route.fulfill({ status: 204, headers })
    if (pathname === "/api/dochaus/provider-candidate/models" && method === "POST") return reply({ data: [], customModels: false })
    if (pathname === "/api/dochaus/provider-test" && method === "POST") return reply({ ok: true })
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply({ version: 1, ready: state.ready, completed: state.ready ? ["models"] : [], skipped: [], next: state.ready ? "research" : "models" })
    if (pathname === "/preferences") return reply(state.preferences)
    if (pathname === "/matters") return state.failMatters ? reply({ error: "Matter list unavailable" }, 503) : reply(state.matters)
    if (/^\/matters\/[^/]+\/drafting-status$/.test(pathname)) return reply(null)
    if (pathname === `/matters/${SAMPLE.id}`) return reply({ ...SAMPLE, documents: [] })
    if (pathname === "/word/setup") return reply(state.word)
    if (pathname === "/templates") return reply({ dir: "/fixture/templates", templates: state.templates })
    if (pathname === "/playbooks/library") return reply({ dir: "/fixture/playbooks", playbooks: [] })
    if (pathname === "/playbooks" || pathname === "/clients") return reply([])
    if (pathname === "/jurisdictions") return reply(packs)
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/engine/config") return reply(state.ready ? { model: "fixture/primary", small_model: "fixture/fast" } : {})
    const provider = { id: "fixture", name: "Fixture Models", integrationID: "fixture", activation: "enabled", package: "@opencode/ai/providers/openai-compatible" }
    if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/fixture" }, data: [provider] })
    if (pathname === "/api/model") return reply({ location: { directory: "/fixture" }, data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
    if (pathname === "/api/integration") return reply({ location: { directory: "/fixture" }, data: [{ id: "fixture", name: "Fixture", methods: [], connections: [{ type: "credential", id: "fixture-account", label: "Fixture", method: "key" }] }] })
    if (pathname === "/api/session/active") return reply({ data: {} })
    if (pathname === "/api/session") return reply({ data: [], cursor: { previous: null, next: null } })
    if (pathname === "/api/permission/request") return reply({ location: { directory: "/fixture" }, data: [] })
    if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: `data: ${JSON.stringify({ id: "evt_setup", type: "server.connected", data: {} })}\n\n` })
    state.unexpected.push(`${method} ${pathname}`); return route.abort()
  })
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  page.on("pageerror", error => state.errors.push(error.message))
  return { context, page }
}
const setupTab = (page: Page) => page.getByRole("tab", { name: /^Setup/ })
const checklist = (page: Page) => page.getByRole("region", { name: "Finish setting up", exact: true })
async function badge(page: Page, count: number) { await expect(setupTab(page)).toContainText(`${count}/5`); await expect(checklist(page).getByRole("progressbar")).toHaveAttribute("value", String(count)) }
async function settled(page: Page) { await expect(checklist(page)).toBeVisible(); await expect(checklist(page).getByRole("textbox", { name: "Your full name", exact: true })).toBeEnabled() }
async function capture(page: Page, name: string) {
  await page.evaluate(async () => { await document.fonts.ready })
  const overflow = await page.evaluate(() => [document.documentElement, document.body, document.getElementById("root")!].map(node => ({ node: node.id || node.tagName, overflow: node.scrollWidth - node.clientWidth })))
  for (const item of overflow) expect(item.overflow, `${name}: ${item.node} horizontal overflow`).toBeLessThanOrEqual(1)
  await page.screenshot({ path: path.join(evidence!, `${name}.png`), fullPage: true })
  screens.push({ name, overflow })
}
const clean = (state: State) => { expect(state.unexpected, "Unexpected requests").toEqual([]); expect(state.errors, "Page errors").toEqual([]) }
async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { const message = error instanceof Error ? error.message : String(error); results.push({ name, status: "failed", error: message }); console.log(`FAIL ${name}: ${message}`) }
}
try {
  for (const [width, height] of [[320, 640], [375, 812], [1440, 1000]]) await scenario(`Fresh setup dashboard ${width}x${height}`, async () => {
    const state = fixture(), { context, page } = await open(state, width, height)
    try {
      await page.goto(`${baseURL}/`); await settled(page)
      await expect(page).toHaveURL(/\/dashboard\?tab=setup$/)
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "true"); await badge(page, 0)
      await expect(page.locator("dialog[open]")).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Ask AI", exact: true })).toHaveCount(0)
      await expect(page.getByRole("group", { name: "Period", exact: true })).toHaveCount(0)
      await expect(page.getByRole("button", { name: /Options for .* dashboard/ })).toHaveCount(0)
      await capture(page, `setup-default-${width}x${height}`)
      await page.getByRole("tab", { name: "Overview", exact: true }).click()
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "false")
      await expect(page.getByRole("button", { name: "Ask AI", exact: true })).toBeDisabled()
      await page.getByRole("button", { name: "Connect AI in Setup", exact: true }).click(); await settled(page)
      await setupTab(page).focus(); await page.keyboard.press("Home")
      await expect(page.getByRole("tab", { name: "Overview", exact: true })).toBeFocused()
      await page.keyboard.press("End"); await expect(setupTab(page)).toBeFocused()
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Saved identity and research progress survives reload; failures stay truthful", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      const form = checklist(page).getByRole("form", { name: "Practice identity & default law", exact: true })
      await form.getByRole("textbox", { name: "Your full name", exact: true }).fill("Alex Example")
      await form.getByRole("textbox", { name: "Firm name", exact: true }).fill("Example Partners LLP")
      await form.getByRole("button", { name: "England & Wales", exact: true }).click()
      await form.getByRole("button", { name: "Commercial", exact: true }).click()
      state.failPreferences = true; await form.getByRole("button", { name: "Save", exact: true }).click()
      await expect(form.getByRole("alert")).toBeVisible(); await badge(page, 0)
      state.failPreferences = false; await form.getByRole("button", { name: "Save", exact: true }).click()
      await badge(page, 1)
      expect(state.preferences.attorney).toBe("Alex Example")
      expect(await page.evaluate(() => JSON.parse(localStorage.getItem("dochaus.prefs")!).defaultJurisdictions)).toEqual(["EW"])
      const key = checklist(page).getByLabel("Exa API key", { exact: true })
      await expect(key).toHaveAttribute("type", "password"); await key.fill("synthetic-exa-test-key")
      state.failPreferences = true; await checklist(page).getByRole("button", { name: "Save key", exact: true }).click()
      await expect(checklist(page).getByRole("alert")).toBeVisible(); await badge(page, 1)
      await expect(checklist(page)).not.toContainText("synthetic-exa-test-key")
      state.failPreferences = false; await checklist(page).getByRole("button", { name: "Save key", exact: true }).click()
      await badge(page, 2); await expect(checklist(page)).not.toContainText("synthetic-exa-test-key")
      await page.reload(); await settled(page); await badge(page, 2)
      await expect(checklist(page).getByRole("textbox", { name: "Your full name", exact: true })).toHaveValue("Alex Example")
      await expect(checklist(page).getByRole("heading", { name: /Exa web research configured/ })).toBeVisible()
      await capture(page, "setup-saved-details-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Preference saves serialize and preserve unfinished identity edits", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      const form = checklist(page).getByRole("form", { name: "Practice identity & default law", exact: true })
      const name = form.getByRole("textbox", { name: "Your full name", exact: true })
      await name.fill("Unfinished Alex")
      await form.getByRole("textbox", { name: "Firm name", exact: true }).fill("Unfinished Firm")
      await expect(form.getByRole("button", { name: "Save", exact: true })).toBeEnabled()
      await checklist(page).getByLabel("Exa API key", { exact: true }).fill("synthetic-held-key")
      state.holdPreferences = true
      await checklist(page).getByRole("button", { name: "Save key", exact: true }).click()
      await expect.poll(() => typeof state.releasePreference).toBe("function")
      await expect(form.getByRole("button", { name: "Save", exact: true })).toBeDisabled()
      await expect(name).toBeDisabled(); await expect(name).toHaveValue("Unfinished Alex")
      state.holdPreferences = false; state.releasePreference?.(); state.releasePreference = undefined
      await badge(page, 1)
      await expect(name).toBeEnabled(); await expect(name).toHaveValue("Unfinished Alex")
      await expect(form.getByRole("textbox", { name: "Firm name", exact: true })).toHaveValue("Unfinished Firm")
      await expect(form.getByRole("button", { name: "Save", exact: true })).toBeEnabled()
      expect(state.saved).toHaveLength(1)
      expect(state.saved[0]).toEqual({ searchApiKey: "synthetic-held-key" })
      clean(state)
    } finally { state.releasePreference?.(); await context.close() }
  })
  await scenario("Settings actions refresh actual readiness, Word and house style", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      await checklist(page).getByRole("button", { name: "Connect", exact: true }).click()
      const dialog = page.getByRole("dialog", { name: "Settings", exact: true })
      await expect(dialog.getByRole("tab", { name: "AI setup", exact: true })).toHaveAttribute("aria-selected", "true")
      // Simulate a successful supported provider save outside this scenario's
      // scope; closing Settings must re-read readiness, not mark a click done.
      await dialog.getByRole("button", { name: "Close", exact: true }).click(); await badge(page, 0)
      await checklist(page).getByRole("button", { name: "Connect", exact: true }).click()
      state.ready = true
      await dialog.getByRole("button", { name: "Close", exact: true }).click(); await badge(page, 1)
      await checklist(page).getByRole("button", { name: "Set up Word", exact: true }).click()
      await expect(dialog.getByRole("tab", { name: "Add-ins", exact: true })).toHaveAttribute("aria-selected", "true")
      await dialog.getByRole("button", { name: "Set up Word", exact: true }).click()
      await expect(dialog.getByRole("status")).toContainText("Ready")
      await dialog.getByRole("button", { name: "Close", exact: true }).click(); await badge(page, 2)
      expect(state.wordPosts).toBe(1)
      state.templates = [{ name: "Services agreement.docx", placeholders: [] }]
      await checklist(page).getByRole("button", { name: "Configure", exact: true }).click()
      await expect(dialog.getByRole("tab", { name: "Drafting", exact: true })).toHaveAttribute("aria-selected", "true")
      await dialog.getByRole("textbox", { name: "House style notes", exact: true }).fill("Plain language; numbered clauses.")
      await dialog.getByRole("button", { name: "Save", exact: true }).click()
      await expect(page.getByText("Drafting preferences saved. They apply to new replies, resumed drafting and documents rendered from now on.", { exact: true })).toBeVisible()
      await dialog.getByRole("button", { name: "Close", exact: true }).click(); await badge(page, 3)
      await capture(page, "setup-connected-tools-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Demo population is single-shot and stays populated on reload", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      const demo = page.getByRole("complementary", { name: "Explore with demo data", exact: true })
      await demo.getByRole("button", { name: "Populate with demo data", exact: true }).click()
      await expect(demo.getByRole("button", { name: "Open sample matter", exact: true })).toBeEnabled()
      expect(state.samplePosts).toBe(1); expect(state.matters).toHaveLength(1)
      await page.reload(); await settled(page)
      await expect(demo.getByRole("button", { name: "Populate with demo data", exact: true })).toHaveCount(0)
      await expect(demo.getByRole("button", { name: "Open sample matter", exact: true })).toBeEnabled()
      expect(state.samplePosts).toBe(1)
      await capture(page, "setup-demo-populated-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Matter-list failure leaves saved setup actionable", async () => {
    const state = fixture(); state.failMatters = true
    const { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page); await badge(page, 0)
      await expect(page.getByRole("status").filter({ hasText: "Your matters could not be loaded. You can still update setup." })).toBeVisible()
      await expect(checklist(page).getByRole("textbox", { name: "Your full name", exact: true })).toBeEnabled()
      await expect(checklist(page).getByRole("button", { name: "Connect", exact: true })).toBeEnabled()
      state.failMatters = false
      await page.getByRole("button", { name: "Retry matters", exact: true }).click()
      await expect(page.getByRole("status").filter({ hasText: "Your matters could not be loaded. You can still update setup." })).toHaveCount(0)
      await settled(page); await badge(page, 0)
      clean(state)
    } finally { await context.close() }
  })
  await scenario("All five saved steps remain available on the Setup tab", async () => {
    const state = fixture()
    state.ready = true; state.jurisdictions = ["EW"]
    state.preferences = { ...state.preferences, attorney: "Alex Example", firm: "Example Partners LLP", practiceAreas: ["Commercial"], searchApiKey: "synthetic-saved-exa", houseStyle: "Plain language; numbered clauses." }
    state.word = { ...state.word, status: "ready", available: true }
    state.templates = [{ name: "Services agreement.docx", placeholders: [] }]
    const { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page); await badge(page, 5)
      await expect(checklist(page).locator(".dash-setup-item.is-done")).toHaveCount(5)
      await expect(checklist(page).getByLabel("Exa API key", { exact: true })).toHaveCount(0)
      await page.reload(); await settled(page); await badge(page, 5)
      await capture(page, "setup-complete-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
} finally {
  await writeFile(path.join(evidence, "setup-results.json"), JSON.stringify({ now: new Date(NOW).toISOString(), results, screens }, null, 2) + "\n")
  await browser.close()
}
expect(results.filter(result => result.status === "failed")).toEqual([])
console.log(`PASS ${results.length} setup scenarios; ${screens.length} screenshots.`)
