// Default dashboard setup against a built web preview. All service traffic is
// synthetic and unknown requests fail closed; no provider or OS work is run.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"
import type { Dashboards } from "../src/components/dashboard/layout"

const { chromium, expect } = await import("@playwright/test")
type Page = import("@playwright/test").Page
type Route = import("@playwright/test").Route
type Locator = import("@playwright/test").Locator
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
type PersistedTabs = { version: 1; workspace: Dashboards; setupHidden: boolean; view?: "setup" | "dashboard" }
type TabsEnvelope = { state: PersistedTabs | null; revision: number }
function fixture(fresh = false) {
  return { dashboard: { state: null, revision: 0 } as TabsEnvelope, tabWrites: [] as { state: PersistedTabs; revision: number }[], failTabReads: false, failTabWrites: false, holdTabWrites: false, releaseTabWrite: undefined as (() => void) | undefined, ready: !fresh, completed: fresh ? [] as string[] : ["models"], skipped: fresh ? [] as string[] : ["research", "details"], setupWrites: [] as { step: string; status: string }[], candidateInputs: [] as Record<string, unknown>[], tested: [] as { modelID: string; ok: boolean }[], failCandidateTest: false, providerSaves: 0, jurisdictions: [] as string[], preferences: { ...DEFAULT_DRAFTING }, failPreferences: false, holdPreferences: false, releasePreference: undefined as (() => void) | undefined, failMatters: false, word: { ...wordDefault }, templates: [] as { name: string; placeholders: unknown[] }[], matters: [] as (typeof SAMPLE)[], saved: [] as Record<string, unknown>[], samplePosts: 0, wordPosts: 0, unexpected: [] as string[], errors: [] as string[] }
}
type State = ReturnType<typeof fixture>
function seedTabs(state: State, workspace: unknown, setupHidden = false) {
  state.dashboard = { state: { version: 1, workspace: workspace as Dashboards, setupHidden }, revision: state.dashboard.revision + 1 }
}
async function tabsSaved(state: State, check: (saved: PersistedTabs) => boolean) {
  await expect.poll(() => state.dashboard.state !== null && check(state.dashboard.state)).toBe(true)
}
function progress(state: State) {
  const completed = state.completed.filter(step => step !== "models" || state.ready)
  const skipped = state.skipped.filter(step => step !== "models" && !completed.includes(step))
  return { version: 1, ready: state.ready, completed, skipped, next: ["models", "research", "details"].find(step => !completed.includes(step) && !skipped.includes(step)) ?? "done" }
}
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
    if (pathname === "/dashboard/tabs" && method === "GET") return state.failTabReads ? reply({ error: "Could not read saved dashboards." }, 503) : reply(state.dashboard)
    if (pathname === "/dashboard/tabs" && method === "PUT") {
      const input = request.postDataJSON() as { state: PersistedTabs; revision: number }
      state.tabWrites.push(structuredClone(input))
      if (state.failTabWrites) return reply({ error: "Could not save dashboards." }, 503)
      if (state.holdTabWrites && !state.releaseTabWrite) await new Promise<void>(resolve => { state.releaseTabWrite = resolve })
      if (input.revision !== state.dashboard.revision) return reply({ error: "Dashboards changed in another browser. Reload saved dashboards.", ...state.dashboard }, 409)
      state.dashboard = { state: structuredClone(input.state), revision: state.dashboard.revision + 1 }
      return reply(state.dashboard)
    }
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
    if (pathname === "/setup" && method === "PATCH") {
      const change = request.postDataJSON() as { step: string; status: string }
      if (change.step === "models" && (!state.ready || change.status !== "completed")) return reply({ error: "Test and save both models first." }, 400)
      if (change.step !== "models" && !state.completed.includes("models")) return reply({ error: "Complete model setup first." }, 400)
      state.setupWrites.push(change)
      state.completed = state.completed.filter(step => step !== change.step)
      state.skipped = state.skipped.filter(step => step !== change.step)
      ;(change.status === "completed" ? state.completed : state.skipped).push(change.step)
      return reply(progress(state))
    }
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "POST") { state.candidateInputs.push(request.postDataJSON()); return reply({ sessionID: "synthetic-candidate" }) }
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "DELETE") return route.fulfill({ status: 204, headers })
    if (pathname === "/api/dochaus/provider-candidate/models" && method === "POST") return reply({ data: ["primary", "fast"].map(id => ({ id, providerID: "fixture", name: id, capabilities: { input: ["text"], output: ["text"], tools: true } })), customModels: false })
    if (pathname === "/api/dochaus/provider-candidate/test" && method === "POST") {
      const input = request.postDataJSON() as { modelID: string }
      state.tested.push({ modelID: input.modelID, ok: !state.failCandidateTest })
      return reply(state.failCandidateTest ? { ok: false, kind: "unavailable", error: "Synthetic model connection unavailable." } : { ok: true })
    }
    if (pathname === "/engine/provider-setup" && method === "POST") {
      const pair = request.postDataJSON() as { primary: string; fast: string }
      if (![pair.primary, pair.fast].every(modelID => state.tested.some(test => test.modelID === modelID && test.ok))) return reply({ code: "untested", error: "Test both selected models first." }, 400)
      state.ready = true; state.providerSaves++; return reply({ ok: true })
    }
    if (pathname === "/api/dochaus/provider-test" && method === "POST") return reply({ ok: true })
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply(progress(state))
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
const closeSetup = (page: Page) => page.getByRole("button", { name: "Close Setup tab", exact: true })
const confirmCloseSetup = (page: Page) => page.getByRole("button", { name: "Confirm close Setup tab", exact: true })
const tabBadge = (page: Page) => setupTab(page).locator(".dash-view-tab-badge")
async function visuallyShown(locator: Locator) {
  return locator.evaluate(node => {
    const style = getComputedStyle(node)
    return style.display !== "none" && style.visibility !== "hidden" && Number(style.opacity) > 0
  })
}
async function badge(page: Page, count: number) { await expect(page.locator("#dash-content-tab-__setup")).toContainText(`${count}/5`); await expect(page.locator(".dash-setup-accordion progress")).toHaveCount(0) }
async function settled(page: Page) { await expect(checklist(page)).toBeVisible(); await expect(checklist(page).locator(".dash-setup-accordion-item")).toHaveCount(5) }
const stageNames: Record<string, string> = { models: "AI provider & models", identity: "Practice identity & default law", research: "Web research with Exa", word: "Microsoft Word add-in", style: "Drafting preferences" }
const stepToggle = (page: Page, id: string) => page.locator(`#dash-setup-toggle-${id}`)
const stageDialog = (page: Page, id: string) => page.getByRole("dialog", { name: stageNames[id], exact: true })
async function expandStep(page: Page, id: string) {
  const dialog = stageDialog(page, id)
  if (!await dialog.count()) await stepToggle(page, id).click()
  await expect(dialog).toBeVisible()
  return dialog
}
async function closeStage(page: Page, id: string) {
  await stageDialog(page, id).getByRole("button", { name: "Close", exact: true }).click()
  await expect(stageDialog(page, id)).toHaveCount(0)
}
async function capture(page: Page, name: string) {
  await page.evaluate(async () => { await document.fonts.ready })
  const overflow = await page.evaluate(() => [document.documentElement, document.body, document.getElementById("root")!].map(node => ({ node: node.id || node.tagName, overflow: node.scrollWidth - node.clientWidth })))
  for (const item of overflow) expect(item.overflow, `${name}: ${item.node} horizontal overflow`).toBeLessThanOrEqual(1)
  await page.screenshot({ path: path.join(evidence!, `${name}.png`), fullPage: true })
  screens.push({ name, overflow })
}
const clean = (state: State) => { expect(state.unexpected, "Unexpected requests").toEqual([]); expect(state.errors, "Page errors").toEqual([]) }
async function scenario(name: string, run: () => Promise<void>) {
  if (process.env.SETUP_SCENARIO && !name.includes(process.env.SETUP_SCENARIO)) return
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { const message = error instanceof Error ? error.message : String(error); results.push({ name, status: "failed", error: message }); console.log(`FAIL ${name}: ${message}`) }
}
try {
  for (const [width, height] of [[320, 640], [375, 812], [1440, 1000]]) await scenario(`Fresh setup dashboard ${width}x${height}`, async () => {
    const state = fixture(true), { context, page } = await open(state, width, height)
    try {
      await page.goto(`${baseURL}/`)
      const onboarding = page.getByRole("dialog", { name: "Let's get you set up", exact: true })
      await expect(onboarding).toBeVisible()
      await expect(onboarding.getByRole("button", { name: "Continue to connection", exact: true })).toBeDisabled()
      await expect(onboarding.getByRole("button", { name: "Step 2: Connection", exact: true })).toBeDisabled()
      await expect(onboarding.getByRole("button", { name: "Step 3: Models", exact: true })).toBeDisabled()
      await expect(onboarding.getByRole("button", { name: "Step 4: Test", exact: true })).toBeDisabled()
      await page.keyboard.press("Escape"); await expect(onboarding).toBeVisible()
      await capture(page, `setup-required-provider-${width}x${height}`)
      await onboarding.getByRole("combobox", { name: "Provider", exact: true }).selectOption("fixture")
      await onboarding.getByRole("button", { name: "Continue to connection", exact: true }).click()
      await expect(onboarding.getByRole("combobox", { name: "Connection method", exact: true })).toHaveValue("credential:fixture-account")
      await onboarding.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(onboarding.getByRole("button", { name: "Continue to test", exact: true })).toBeDisabled()
      await onboarding.getByRole("combobox", { name: "Primary model", exact: true }).selectOption("fixture/primary")
      await onboarding.getByRole("combobox", { name: "Fast model", exact: true }).selectOption("fixture/fast")
      await onboarding.getByRole("button", { name: "Continue to test", exact: true }).click()
      expect(state.providerSaves).toBe(0); expect(state.setupWrites).toEqual([])
      state.failCandidateTest = true
      await onboarding.getByRole("button", { name: "Test and continue", exact: true }).click()
      await expect(onboarding.getByText("Synthetic model connection unavailable.").first()).toBeVisible()
      expect(state.providerSaves).toBe(0); expect(state.setupWrites).toEqual([])
      state.failCandidateTest = false
      await onboarding.getByRole("button", { name: "Test and continue", exact: true }).click()
      await expect(onboarding.getByRole("heading", { name: "Set up legal research", exact: true })).toBeVisible()
      expect(state.providerSaves).toBe(1)
      expect(state.candidateInputs[0]).toMatchObject({ providerID: "fixture", connection: { type: "credential", id: "fixture-account" } })
      expect(state.tested.filter(test => test.ok).map(test => test.modelID)).toEqual(["primary", "fast"])
      await onboarding.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(onboarding.getByRole("heading", { name: "Make doc.haus yours", exact: true })).toBeVisible()
      await onboarding.getByRole("button", { name: /^Save and finish/ }).click()
      await expect(onboarding).toHaveCount(0); await settled(page)
      await expect(page).toHaveURL(/\/dashboard\?tab=setup$/)
      await expect(page.locator("main h1:not(.sr-only)")).toHaveCount(0)
      await expect(page.locator(".page-bar")).toHaveCount(0)
      const widths = await page.evaluate(() => ({ dashboard: document.querySelector(".dash")!.getBoundingClientRect().width, checklist: document.querySelector(".dash-setup-accordion")!.getBoundingClientRect().width }))
      expect(Math.abs(widths.dashboard - widths.checklist)).toBeLessThanOrEqual(1)
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "true"); await badge(page, 1)
      expect(state.setupWrites).toEqual([{ step: "models", status: "completed" }, { step: "research", status: "completed" }, { step: "details", status: "completed" }])
      await expect(page.locator("dialog[open]")).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Ask AI", exact: true })).toHaveCount(0)
      await expect(page.getByRole("group", { name: "Period", exact: true })).toHaveCount(0)
      await capture(page, `setup-default-${width}x${height}`)
      if (width === 375) {
        const identityDialog = await expandStep(page, "identity")
        await identityDialog.getByLabel("Jurisdictions", { exact: true }).click()
        await expect(page.getByRole("group", { name: "Jurisdictions", exact: true })).toBeVisible()
        await capture(page, "setup-jurisdictions-dropdown-375x812")
        await page.keyboard.press("Escape")
        await closeStage(page, "identity")
      }
      await expect(page.getByRole("tab", { name: "Overview", exact: true })).toHaveCount(0)
      await expect(page.getByRole("tablist", { name: "Dashboards", exact: true }).getByRole("tab")).toHaveCount(1)
      await setupTab(page).focus(); await page.keyboard.press("Home")
      await expect(setupTab(page)).toBeFocused()
      await page.keyboard.press("End"); await expect(setupTab(page)).toBeFocused()
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Setup hover close confirms inline, cancels safely and persists only visibility", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard`); await settled(page); await badge(page, 1)
      await expect(page.getByRole("tab", { name: "Overview", exact: true })).toHaveCount(0)
      await page.mouse.move(0, 0)
      await expect.poll(() => visuallyShown(tabBadge(page))).toBe(true)
      const before = await setupTab(page).boundingBox()
      await capture(page, "setup-tab-normal-1440x1000")
      await setupTab(page).locator("span").first().hover()
      await expect.poll(() => visuallyShown(closeSetup(page))).toBe(true)
      await expect.poll(() => visuallyShown(tabBadge(page))).toBe(false)
      const closeBox = await closeSetup(page).boundingBox()
      await capture(page, "setup-tab-hover-close-1440x1000")
      await closeSetup(page).click()
      await expect(confirmCloseSetup(page)).toHaveText("Confirm")
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "true")
      await expect(checklist(page)).toBeVisible()
      const confirmBox = await confirmCloseSetup(page).boundingBox()
      const confirming = await setupTab(page).boundingBox()
      expect(Math.abs((confirmBox!.x + confirmBox!.width) - (closeBox!.x + closeBox!.width))).toBeLessThanOrEqual(1)
      expect(Math.abs(confirming!.width - before!.width)).toBeLessThanOrEqual(1)
      await capture(page, "setup-tab-inline-confirm-1440x1000")
      await page.keyboard.press("Escape")
      await expect(confirmCloseSetup(page)).toHaveCount(0)
      await expect(setupTab(page)).toBeFocused()
      await page.getByRole("group", { name: "Demo data", exact: true }).locator("p").first().click()
      await page.mouse.move(0, 0)
      await expect.poll(() => visuallyShown(tabBadge(page))).toBe(true)
      await setupTab(page).locator("span").first().hover(); await closeSetup(page).click()
      await page.getByRole("group", { name: "Demo data", exact: true }).locator("p").first().click()
      await expect(confirmCloseSetup(page)).toHaveCount(0)
      await setupTab(page).focus(); await page.keyboard.press("Tab")
      await expect(closeSetup(page)).toBeFocused()
      await page.keyboard.press("Enter")
      await expect(confirmCloseSetup(page)).toBeFocused()
      await page.keyboard.press("Enter")
      await expect(setupTab(page)).toHaveCount(0); await expect(checklist(page)).toHaveCount(0)
      await tabsSaved(state, saved => saved.setupHidden)
      await expect(page.getByRole("button", { name: "Show setup", exact: true })).toBeVisible()
      await expect(page.getByRole("button", { name: "New dashboard", exact: true }).first()).toBeEnabled()
      await capture(page, "setup-tab-closed-1440x1000")
      await page.reload()
      await expect(setupTab(page)).toHaveCount(0)
      await page.getByRole("button", { name: "Show setup", exact: true }).click()
      await settled(page); await badge(page, 1)
      await tabsSaved(state, saved => !saved.setupHidden)
      await page.reload(); await settled(page); await badge(page, 1)
      expect(state.saved).toEqual([]); expect(state.setupWrites).toEqual([])
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Setup close is available on touch without hovering", async () => {
    const state = fixture(), { context, page } = await open(state, 375, 812)
    try {
      await page.goto(`${baseURL}/dashboard`); await settled(page)
      await expect.poll(() => visuallyShown(closeSetup(page))).toBe(true)
      await closeSetup(page).tap()
      await expect(confirmCloseSetup(page)).toBeVisible()
      await expect(checklist(page)).toBeVisible()
      await capture(page, "setup-tab-inline-confirm-375x812")
      await confirmCloseSetup(page).tap()
      await expect(setupTab(page)).toHaveCount(0)
      await page.getByRole("button", { name: "Show setup", exact: true }).tap()
      await settled(page); await badge(page, 1)
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Closing inactive Setup preserves custom dashboard and saved Overview", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard`); await settled(page)
      const saved = { active: "costs", dashboards: [
        { id: "overview", name: "Overview", period: "month", stages: {}, widgets: [] },
        { id: "drafts", name: "Drafts", period: "month", stages: {}, widgets: [] },
        { id: "costs", name: "Costs", period: "month", stages: {}, widgets: [] },
      ] }
      seedTabs(state, saved)
      await page.reload()
      const costs = page.getByRole("tab", { name: "Costs", exact: true })
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await expect(page.getByRole("tab", { name: "Overview", exact: true })).toHaveCount(0)
      await setupTab(page).locator("span").first().hover(); await closeSetup(page).click()
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "false")
      await confirmCloseSetup(page).click()
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await expect(costs).toBeFocused()
      await tabsSaved(state, stored => stored.setupHidden && stored.workspace.active === "costs")
      expect(state.dashboard.state!.workspace).toMatchObject(saved)
      await page.reload()
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await expect(setupTab(page)).toHaveCount(0)
      await page.getByRole("button", { name: "Show setup", exact: true }).click()
      await settled(page)
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "true")
      await setupTab(page).locator("span").first().hover()
      await closeSetup(page).click(); await confirmCloseSetup(page).click()
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await expect(page.getByRole("tab", { name: "Drafts", exact: true })).toHaveAttribute("aria-selected", "false")
      await tabsSaved(state, stored => stored.setupHidden && stored.workspace.active === "costs")
      expect(state.dashboard.state!.workspace).toMatchObject(saved)
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Undo deleting the last custom dashboard restores its selected view", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard`); await settled(page)
      const saved = { active: "costs", dashboards: [
        { id: "overview", name: "Overview", period: "month", stages: {}, widgets: [] },
        { id: "costs", name: "Costs", period: "month", stages: {}, widgets: [] },
      ] }
      seedTabs(state, saved)
      await page.reload()
      const costs = page.getByRole("tab", { name: "Costs", exact: true })
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await page.getByRole("button", { name: "Ask AI", exact: true }).click()
      await expect(page.locator("#dash-ai-panel")).toBeVisible()
      await page.locator('summary[aria-label="Options for Costs dashboard"]').click()
      await page.getByRole("button", { name: "Delete dashboard", exact: true }).click()
      await page.getByRole("group", { name: "Delete dashboard", exact: true }).getByRole("button", { name: "Delete dashboard", exact: true }).click()
      await expect(costs).toHaveCount(0)
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "true")
      await expect(page.locator("#dash-ai-panel")).toHaveCount(0)
      await tabsSaved(state, stored => !stored.workspace.dashboards.some(tab => tab.id === "costs"))
      await page.getByRole("button", { name: "Undo", exact: true }).click()
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await expect(setupTab(page)).toHaveAttribute("aria-selected", "false")
      await expect(checklist(page)).toHaveCount(0)
      await expect(page.locator("#dash-ai-panel")).toHaveCount(0)
      await tabsSaved(state, stored => stored.workspace.active === "costs" && stored.workspace.dashboards.some(tab => tab.id === "costs"))
      expect(state.dashboard.state!.workspace).toMatchObject(saved)
      await capture(page, "setup-tab-undo-restores-custom-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Empty and populated local databases both ignore stale browser tab storage", async () => {
    const state = fixture(), first = await open(state), second = await open(state)
    const saved = { active: "costs", dashboards: [
      { id: "overview", name: "Overview", period: "month", stages: {}, widgets: [] },
      { id: "costs", name: "Costs", period: "month", stages: {}, widgets: [{ key: "spend", widget: { type: "spend" }, pinned: true }] },
    ] }
    try {
      await first.context.addInitScript(() => {
        localStorage.setItem("dh.dashboard.v4", JSON.stringify({ active: "stale", dashboards: [{ id: "stale", name: "Stale browser view", period: "all", stages: {}, widgets: [] }] }))
        localStorage.setItem("dh.dashboard.setup-hidden.v1", "true")
      })
      await first.page.goto(`${baseURL}/dashboard`)
      await settled(first.page)
      await expect(setupTab(first.page)).toHaveAttribute("aria-selected", "true")
      await expect(first.page.getByRole("tab", { name: "Stale browser view", exact: true })).toHaveCount(0)
      await tabsSaved(state, stored => !stored.setupHidden && !stored.workspace.dashboards.some(tab => tab.id === "stale"))
      expect(state.tabWrites).toHaveLength(1)
      // Existing tabs come only from the local database, including in a second
      // browser which still has unrelated values in its old browser storage.
      seedTabs(state, saved, true)
      const writes = state.tabWrites.length
      await second.context.addInitScript(() => localStorage.setItem("dh.dashboard.v4", JSON.stringify({ active: "other-stale", dashboards: [{ id: "other-stale", name: "Another stale view", period: "all", stages: {}, widgets: [] }] })))
      await second.page.goto(`${baseURL}/dashboard`)
      await expect(second.page.getByRole("tab", { name: "Costs", exact: true })).toHaveAttribute("aria-selected", "true")
      await expect(second.page.getByRole("tab", { name: "Another stale view", exact: true })).toHaveCount(0)
      await expect(second.page.getByRole("button", { name: "Unpin Total spend", exact: true })).toBeVisible()
      await expect(setupTab(second.page)).toHaveCount(0)
      expect(state.dashboard.state!.workspace).toMatchObject(saved)
      expect(state.tabWrites.length).toBe(writes)
      await capture(second.page, "dashboard-tabs-from-local-database-1440x1000")
      clean(state)
    } finally { await first.context.close(); await second.context.close() }
  })
  await scenario("Every custom tab confirms close and Undo is saved for another browser", async () => {
    const state = fixture()
    const saved = { active: "costs", dashboards: [
      { id: "overview", name: "Overview", period: "month", stages: {}, widgets: [] },
      { id: "drafts", name: "Drafts", period: "month", stages: {}, widgets: [] },
      { id: "costs", name: "Costs", period: "month", stages: {}, widgets: [] },
    ] }
    seedTabs(state, saved)
    const first = await open(state), second = await open(state)
    try {
      await first.page.goto(`${baseURL}/dashboard`)
      const costs = first.page.getByRole("tab", { name: "Costs", exact: true })
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await costs.locator("span").first().hover()
      await first.page.getByRole("button", { name: "Close Costs tab", exact: true }).click()
      await expect(costs).toBeVisible()
      await expect(first.page.getByRole("button", { name: "Confirm close Costs tab", exact: true })).toHaveText("Confirm")
      await capture(first.page, "dashboard-custom-tab-inline-confirm-1440x1000")
      await first.page.getByRole("button", { name: "Confirm close Costs tab", exact: true }).click()
      await expect(costs).toHaveCount(0)
      await tabsSaved(state, stored => !stored.workspace.dashboards.some(tab => tab.id === "costs"))
      await first.page.getByRole("button", { name: "Undo", exact: true }).click()
      await expect(costs).toHaveAttribute("aria-selected", "true")
      await tabsSaved(state, stored => stored.workspace.active === "costs" && stored.workspace.dashboards.some(tab => tab.id === "costs"))
      await second.page.goto(`${baseURL}/dashboard`)
      await expect(second.page.getByRole("tab", { name: "Costs", exact: true })).toHaveAttribute("aria-selected", "true")
      await expect(second.page.getByRole("button", { name: "Close Drafts tab", exact: true })).toHaveCount(1)
      clean(state)
    } finally { await first.context.close(); await second.context.close() }
  })
  await scenario("A database read failure blocks tabs until an explicit retry", async () => {
    const state = fixture(); state.failTabReads = true
    const { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard`)
      await expect(page.getByRole("alert").filter({ hasText: "Could not load your dashboards" })).toBeVisible()
      await expect(setupTab(page)).toHaveCount(0)
      await expect(page.getByRole("button", { name: "New dashboard", exact: true })).toHaveCount(0)
      expect(state.tabWrites).toEqual([]); expect(state.dashboard.state).toBeNull()
      state.failTabReads = false
      await page.getByRole("button", { name: "Retry dashboards", exact: true }).click()
      await settled(page)
      await tabsSaved(state, () => true)
      clean(state)
    } finally { await context.close() }
  })
  await scenario("A failed database save keeps the edit visible and retry persists it", async () => {
    const state = fixture()
    seedTabs(state, { active: "costs", dashboards: [{ id: "overview", name: "Overview", period: "month", stages: {}, widgets: [] }, { id: "costs", name: "Costs", period: "month", stages: {}, widgets: [] }] })
    const { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard`)
      const costs = page.getByRole("tab", { name: "Costs", exact: true })
      await expect(costs).toHaveAttribute("aria-selected", "true")
      state.failTabWrites = true
      await costs.locator("span").first().hover()
      await page.getByRole("button", { name: "Close Costs tab", exact: true }).click()
      await page.getByRole("button", { name: "Confirm close Costs tab", exact: true }).click()
      await expect(costs).toHaveCount(0)
      await expect(page.getByRole("alert").filter({ hasText: "Could not save your dashboard changes. Your changes are still here." })).toBeVisible()
      expect(state.dashboard.state!.workspace.dashboards.some(tab => tab.id === "costs")).toBe(true)
      expect(await page.evaluate(() => localStorage.getItem("dh.dashboard.v4"))).toBeNull()
      state.failTabWrites = false
      await page.getByRole("button", { name: "Retry save", exact: true }).click()
      await tabsSaved(state, stored => !stored.workspace.dashboards.some(tab => tab.id === "costs"))
      await expect(page.getByRole("button", { name: "Retry save", exact: true })).toHaveCount(0)
      await page.reload(); await settled(page)
      await expect(costs).toHaveCount(0)
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Rapid tab selection serializes saves using each acknowledged database revision", async () => {
    const state = fixture()
    seedTabs(state, { active: "costs", dashboards: [{ id: "overview", name: "Overview", period: "month", stages: {}, widgets: [] }, { id: "costs", name: "Costs", period: "month", stages: {}, widgets: [] }, { id: "drafts", name: "Drafts", period: "month", stages: {}, widgets: [] }] })
    const { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard`)
      await expect(page.getByRole("tab", { name: "Costs", exact: true })).toHaveAttribute("aria-selected", "true")
      state.holdTabWrites = true
      await page.getByRole("tab", { name: "Drafts", exact: true }).click()
      await expect.poll(() => typeof state.releaseTabWrite).toBe("function")
      await page.getByRole("tab", { name: "Costs", exact: true }).click()
      expect(state.tabWrites).toHaveLength(1)
      state.holdTabWrites = false; state.releaseTabWrite?.(); state.releaseTabWrite = undefined
      await tabsSaved(state, stored => stored.workspace.active === "costs")
      await expect.poll(() => state.tabWrites.length).toBe(2)
      expect(state.tabWrites.map(write => write.revision)).toEqual([1, 2])
      clean(state)
    } finally { state.releaseTabWrite?.(); await context.close() }
  })
  await scenario("A stale database revision never overwrites another browser's saved tabs", async () => {
    const state = fixture()
    seedTabs(state, { active: "costs", dashboards: [{ id: "overview", name: "Overview", period: "month", stages: {}, widgets: [] }, { id: "costs", name: "Costs", period: "month", stages: {}, widgets: [] }] })
    const { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard`)
      const costs = page.getByRole("tab", { name: "Costs", exact: true })
      await expect(costs).toHaveAttribute("aria-selected", "true")
      const next = structuredClone(state.dashboard.state!)
      next.workspace.dashboards.find(tab => tab.id === "costs")!.name = "Costs updated elsewhere"
      state.dashboard = { state: next, revision: state.dashboard.revision + 1 }
      await costs.locator("span").first().hover()
      await page.getByRole("button", { name: "Close Costs tab", exact: true }).click()
      await page.getByRole("button", { name: "Confirm close Costs tab", exact: true }).click()
      await expect(page.getByRole("alert").filter({ hasText: "These dashboards changed in another window. Your unsaved changes are still here." })).toBeVisible()
      expect(state.dashboard.state!.workspace.dashboards.find(tab => tab.id === "costs")!.name).toBe("Costs updated elsewhere")
      await page.getByRole("button", { name: "Load saved dashboards", exact: true }).click()
      await expect(page.getByRole("tab", { name: "Costs updated elsewhere", exact: true })).toHaveAttribute("aria-selected", "true")
      await expect(page.getByRole("button", { name: "Load saved dashboards", exact: true })).toHaveCount(0)
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Focused stage dialogs trap focus, return it on Escape and preserve drafts", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      const doneRow = stepToggle(page, "models"), incompleteRow = stepToggle(page, "identity")
      const doneBounds = await doneRow.boundingBox(), incompleteBounds = await incompleteRow.boundingBox()
      expect(doneBounds!.height).toBeLessThan(incompleteBounds!.height - 10)
      await expect(doneRow.locator(".dash-setup-row-detail")).not.toBeVisible()
      await expect(page.locator(".dash-setup-panel")).toHaveCount(0)
      await stepToggle(page, "identity").focus(); await page.keyboard.press("Enter")
      const identity = stageDialog(page, "identity")
      await expect(identity).toBeVisible()
      await expect(identity.getByLabel("Exa API key", { exact: true })).toHaveCount(0)
      await identity.getByRole("textbox", { name: "Your full name", exact: true }).fill("Unfinished practice details")
      for (let index = 0; index < 8; index++) {
        await page.keyboard.press("Tab")
        const focused = await page.evaluate(() => ({ inside: !!document.activeElement?.closest("dialog[open]"), tag: document.activeElement?.tagName, html: document.activeElement?.outerHTML.slice(0, 350) }))
        if (!focused.inside) console.log("Focus left stage dialog", JSON.stringify(focused))
        expect(focused.inside, JSON.stringify(focused)).toBe(true)
      }
      await capture(page, "setup-identity-modal-1440x1000")
      await page.keyboard.press("Escape")
      await expect(identity).toHaveCount(0)
      await expect(stepToggle(page, "identity")).toBeFocused()
      await stepToggle(page, "models").focus(); await page.keyboard.press("Space")
      await expect(stageDialog(page, "models")).toBeVisible()
      await expect(stageDialog(page, "models").getByRole("textbox", { name: "Your full name", exact: true })).toHaveCount(0)
      await closeStage(page, "models")
      await expandStep(page, "identity")
      await expect(identity.getByRole("textbox", { name: "Your full name", exact: true })).toHaveValue("Unfinished practice details")
      expect(state.saved).toHaveLength(0)
      await closeStage(page, "identity")
      await capture(page, "setup-stage-list-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Saved identity and research progress survives reload; failures stay truthful", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      await expandStep(page, "identity")
      const form = stageDialog(page, "identity")
      await form.getByRole("textbox", { name: "Your full name", exact: true }).fill("Alex Example")
      await form.getByRole("textbox", { name: "Firm name", exact: true }).fill("Example Partners LLP")
      await form.getByLabel("Jurisdictions", { exact: true }).click()
      await page.getByRole("group", { name: "Jurisdictions", exact: true }).getByRole("checkbox", { name: /England & Wales/ }).check()
      await page.keyboard.press("Escape")
      state.failPreferences = true; await form.getByRole("button", { name: "Save", exact: true }).click()
      await expect(form.getByRole("alert")).toBeVisible(); await badge(page, 1)
      state.failPreferences = false; await form.getByRole("button", { name: "Save", exact: true }).click()
      await badge(page, 2)
      expect(state.preferences.attorney).toBe("Alex Example")
      expect(state.preferences.practiceAreas).toEqual([])
      // Practice areas are optional: name, firm and default law completed the
      // identity step before choosing any areas. Saved choices still work.
      await expect(stageDialog(page, "identity")).toHaveCount(0)
      await expandStep(page, "identity")
      await expect(form.getByLabel("Practice areas", { exact: true })).toContainText("Select practice areas")
      await form.getByLabel("Practice areas", { exact: true }).click()
      await page.getByRole("group", { name: "Practice areas", exact: true }).getByRole("checkbox", { name: "Commercial", exact: true }).check()
      await page.keyboard.press("Escape")
      await form.getByRole("button", { name: "Save", exact: true }).click()
      await expect(form).toHaveCount(0); await badge(page, 2)
      expect(state.preferences.practiceAreas).toEqual(["Commercial"])
      expect(await page.evaluate(() => JSON.parse(localStorage.getItem("dochaus.prefs")!).defaultJurisdictions)).toEqual(["EW"])
      await expect(stageDialog(page, "identity")).toHaveCount(0)
      await expandStep(page, "research")
      const key = stageDialog(page, "research").getByLabel("Exa API key", { exact: true })
      await expect(key).toHaveAttribute("type", "password"); await key.fill("synthetic-exa-test-key")
      state.failPreferences = true; await stageDialog(page, "research").getByRole("button", { name: "Save key", exact: true }).click()
      await expect(stageDialog(page, "research").getByRole("alert")).toBeVisible(); await badge(page, 2)
      await expect(page.locator("body")).not.toContainText("synthetic-exa-test-key")
      state.failPreferences = false; await stageDialog(page, "research").getByRole("button", { name: "Save key", exact: true }).click()
      await badge(page, 3); await expect(page.locator("body")).not.toContainText("synthetic-exa-test-key")
      await page.reload(); await settled(page); await badge(page, 3)
      await expandStep(page, "identity")
      await expect(stageDialog(page, "identity").getByRole("textbox", { name: "Your full name", exact: true })).toHaveValue("Alex Example")
      await expect(form.getByLabel("Jurisdictions", { exact: true })).toContainText("England & Wales")
      await expect(form.getByLabel("Practice areas", { exact: true })).toContainText("Commercial")
      await closeStage(page, "identity")
      await expect(checklist(page).getByRole("heading", { name: /Web research with Exa/ })).toBeVisible()
      await capture(page, "setup-saved-details-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
  await scenario("A focused research save preserves unfinished practice edits", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      const identity = await expandStep(page, "identity")
      await identity.getByRole("textbox", { name: "Your full name", exact: true }).fill("Unfinished Alex")
      await identity.getByRole("textbox", { name: "Firm name", exact: true }).fill("Unfinished Firm")
      await closeStage(page, "identity")
      const research = await expandStep(page, "research")
      await research.getByLabel("Exa API key", { exact: true }).fill("synthetic-held-key")
      state.holdPreferences = true
      await research.getByRole("button", { name: "Save key", exact: true }).click()
      await expect.poll(() => typeof state.releasePreference).toBe("function")
      await expect(research.getByRole("button", { name: "Save key", exact: true })).toBeDisabled()
      state.holdPreferences = false; state.releasePreference?.(); state.releasePreference = undefined
      await expect(research).toHaveCount(0); await badge(page, 2)
      await expandStep(page, "identity")
      await expect(identity.getByRole("textbox", { name: "Your full name", exact: true })).toHaveValue("Unfinished Alex")
      await expect(identity.getByRole("textbox", { name: "Firm name", exact: true })).toHaveValue("Unfinished Firm")
      expect(state.saved).toHaveLength(1)
      expect(state.saved[0]).toMatchObject({ searchApiKey: "synthetic-held-key" })
      await closeStage(page, "identity")
      clean(state)
    } finally { state.releasePreference?.(); await context.close() }
  })
  await scenario("Focused stage settings refresh readiness, Word and full drafting preferences", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      await expandStep(page, "models")
      await expect(stageDialog(page, "models").getByRole("tab", { name: "Add-ins", exact: true })).toHaveCount(0)
      state.ready = false
      await closeStage(page, "models"); await badge(page, 0)
      await expandStep(page, "models")
      state.ready = true
      await closeStage(page, "models"); await badge(page, 1)
      const word = await expandStep(page, "word")
      await word.getByRole("button", { name: "Set up Word", exact: true }).click()
      await expect(word.getByRole("status")).toContainText("Ready")
      await closeStage(page, "word"); await badge(page, 2)
      expect(state.wordPosts).toBe(1)
      const style = await expandStep(page, "style")
      await badge(page, 2)
      await expect(style.getByRole("button", { name: "Save", exact: true })).toBeEnabled()
      await expect(style.getByRole("group", { name: "Posture", exact: true })).toBeVisible()
      await expect(style.getByRole("group", { name: "Page layout", exact: true })).toBeVisible()
      await style.getByRole("group", { name: "Posture", exact: true }).getByRole("button", { name: "Conservative", exact: true }).click()
      await style.getByRole("group", { name: "Style", exact: true }).getByRole("button", { name: "Plain language", exact: true }).click()
      await style.getByRole("group", { name: "Page numbers", exact: true }).getByRole("button", { name: "Page 2", exact: true }).click()
      await expect(style.getByRole("textbox", { name: "House style notes", exact: true })).toHaveValue("")
      expect(state.preferences.draftingConfigured).toBe(false)
      await style.getByRole("button", { name: "Save", exact: true }).click()
      await expect(style).toHaveCount(0); await badge(page, 3)
      expect(state.preferences.draftingConfigured).toBe(true)
      expect(state.preferences.houseStyle).toBe("")
      expect(state.templates).toHaveLength(0)
      expect(state.preferences.posture).toBe("conservative")
      expect(state.preferences.formality).toBe("plain")
      expect(state.preferences.pageNumbers).toBe("page")
      await expandStep(page, "style")
      await expect(style.getByRole("group", { name: "Posture", exact: true }).getByRole("button", { name: "Conservative", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(style.getByRole("group", { name: "Page numbers", exact: true }).getByRole("button", { name: "Page 2", exact: true })).toHaveAttribute("aria-pressed", "true")
      await capture(page, "setup-drafting-preferences-modal-1440x1000")
      await closeStage(page, "style")
      await page.reload(); await settled(page); await badge(page, 3)
      expect(state.preferences.draftingConfigured).toBe(true)
      await capture(page, "setup-connected-tools-1440x1000")
      clean(state)
    } finally { await context.close() }
  })
  await scenario("Demo population is single-shot and stays populated on reload", async () => {
    const state = fixture(), { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page)
      const demo = page.getByRole("group", { name: "Demo data", exact: true })
      await demo.getByRole("button", { name: "Hide setup", exact: true }).click()
      await expect(demo).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Show setup", exact: true })).toBeVisible()
      expect(state.samplePosts).toBe(0)
      await capture(page, "setup-hidden-from-demo-banner-1440x1000")
      await page.getByRole("button", { name: "Show setup", exact: true }).click()
      await expect(demo).toBeVisible()
      await demo.getByRole("button", { name: "Add sample data", exact: true }).click()
      await expect(demo.getByRole("button", { name: "Open sample matter", exact: true })).toBeEnabled()
      expect(state.samplePosts).toBe(1); expect(state.matters).toHaveLength(1)
      await page.reload(); await settled(page)
      await expect(demo.getByRole("button", { name: "Add sample data", exact: true })).toHaveCount(0)
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
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page); await badge(page, 1)
      await expect(page.getByRole("status").filter({ hasText: "Your matters could not be loaded. You can still update setup." })).toBeVisible()
      const identity = await expandStep(page, "identity")
      await expect(identity.getByRole("textbox", { name: "Your full name", exact: true })).toBeEnabled()
      // Closing a focused stage refreshes dashboard data. Keep the service
      // failing until that specific refresh finishes, then test explicit Retry.
      const closeRefresh = page.waitForResponse(response => new URL(response.url()).pathname.endsWith("/matters") && response.request().method() === "GET" && response.status() === 503)
      await closeStage(page, "identity")
      await closeRefresh
      const retry = page.getByRole("button", { name: "Retry matters", exact: true })
      await expect(retry).toBeEnabled()
      await expect(page.getByRole("status").filter({ hasText: "Your matters could not be loaded. You can still update setup." })).toBeVisible()
      await expect(stepToggle(page, "models")).toBeEnabled()
      state.failMatters = false
      await retry.click()
      await expect(page.getByRole("status").filter({ hasText: "Your matters could not be loaded. You can still update setup." })).toHaveCount(0)
      await settled(page); await badge(page, 1)
      clean(state)
    } finally { await context.close() }
  })
  await scenario("All five saved steps remain available on the Setup tab", async () => {
    const state = fixture()
    state.ready = true; state.jurisdictions = ["EW"]
    state.preferences = { ...state.preferences, attorney: "Alex Example", firm: "Example Partners LLP", practiceAreas: ["Commercial"], searchApiKey: "synthetic-saved-exa", houseStyle: "Plain language; numbered clauses.", draftingConfigured: true }
    state.word = { ...state.word, status: "ready", available: true }
    state.templates = [{ name: "Services agreement.docx", placeholders: [] }]
    const { context, page } = await open(state)
    try {
      await page.goto(`${baseURL}/dashboard?tab=setup`); await settled(page); await badge(page, 5)
      await expect(checklist(page).locator(".dash-setup-accordion-item.is-done")).toHaveCount(5)
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
