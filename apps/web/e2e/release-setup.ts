// Run against a Vite/preview URL. Every backend/auth request is intercepted;
// unexpected network requests fail the check rather than reaching an account.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"
import { OFFICIAL_SOURCE_HOSTS, OFFICIAL_SOURCE_SUFFIXES } from "../../../dochaus/lib/research-sources"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
const appOrigin = new URL(baseURL).origin
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const steps = ["models", "research", "details"]
const results: { name: string; status: string; error?: string }[] = []
const scenarioPattern = process.env.DOCHAUS_E2E_SCENARIOS ? new RegExp(process.env.DOCHAUS_E2E_SCENARIOS) : null

function fixture() {
  return {
    connected: false, nativeEnv: false, nativeBedrock: false, secondProvider: false, ready: false, hostCatalog: false, hostConfigured: false, hostDetected: false, publicModels: false,
    config: {} as Record<string, unknown>, completed: [] as string[], skipped: [] as string[],
    preferences: { ...DEFAULT_DRAFTING, attorney: "Saved Lawyer", firm: "Saved Firm", searchApiKey: "saved-fixture-exa", approvedSources: ["saved.example"] },
    requests: [] as { method: string; path: string; query: Record<string, string>; body: Record<string, unknown> }[],
    unexpected: [] as string[], policyFailures: 0, readbackFailures: 0, detailsFailures: 0, preferenceReadbackFailures: 0, probeFailures: 0, researchTestFailures: 0, setupFailures: 0, jurisdictionsUnavailable: false,
    oauthStatus: "pending", commandStatus: "pending",
    reuseAvailable: false, reuseMixedStores: false, reuseSourceUnavailable: false, reuseDiscoveryUnavailable: false, reuseCopyFailures: 0, reusedGo: false, holdReuse: false,
    releaseReuse: undefined as (() => void) | undefined,
  }
}
async function open(state: ReturnType<typeof fixture>) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1100 } })
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  const errors: string[] = []
  page.on("pageerror", (error: Error) => errors.push(error.message))
  await page.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const api = pathname.startsWith("/api/") || ["/auth/reuse", "/setup", "/engine/config", "/preferences", "/matters", "/jurisdictions", "/updates", "/host/gcp-projects", "/host/aws-profiles", "/host/probe-vertex"].includes(pathname) || pathname.startsWith("/matters/") || pathname.startsWith("/research/")
    if (!api && url.origin === appOrigin) return route.continue()
    if (!api) { state.unexpected.push(request.url()); return route.abort() }
    if (request.method() === "OPTIONS") return route.fulfill({ status: 204, headers: { "access-control-allow-origin": appOrigin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" } })
    const method = request.method()
    const body = request.postData() ? JSON.parse(request.postData()!) : {}
    state.requests.push({ method, path: pathname, query: Object.fromEntries(url.searchParams), body })
    const reply = (json: unknown, status = 200) => route.fulfill({ status, contentType: "application/json", headers: { "access-control-allow-origin": appOrigin, "access-control-allow-credentials": "true" }, body: JSON.stringify(json) })
    const progress = () => ({ version: 1, ready: state.ready, completed: state.completed, skipped: state.skipped, next: steps.find(step => !state.completed.includes(step) && !state.skipped.includes(step)) ?? "done" })
    const provider = { id: "test-models", integrationID: "test-login", name: "Fixture Models", activation: "enabled", package: "@opencode/ai/providers/openai-compatible" }
    const vertex = { id: "google-vertex", integrationID: "google-vertex", name: "Google Vertex AI", activation: state.hostConfigured ? "enabled" : "auto", host: { type: "google-adc", options: state.hostDetected ? { project: "fixture-project", location: "us-central1" } : {}, sources: ["Google ADC file"] } }
    const bedrock = { id: "amazon-bedrock", integrationID: "amazon-bedrock", name: "Amazon Bedrock", activation: "auto", host: { type: "aws-chain", options: { profile: "work", region: "us-east-1" }, sources: ["AWS_PROFILE", "AWS shared profiles"] } }
    const go = { id: "opencode-go", integrationID: "opencode-go", name: "OpenCode Go", activation: "auto" }
    const publicProvider = { id: "opencode", integrationID: "opencode", name: "OpenCode", activation: "enabled", settings: { apiKey: "public" } }
    const zai = { id: "zai-coding-plan", integrationID: "zai-coding-plan", name: "Z.AI Coding Plan", activation: "auto" }
    const publicCatalog = state.publicModels ? [{ providerID: "opencode", id: "free-fixture", name: "Free Fixture", enabled: true, cost: { input: 0, output: 0 } }] : []
    const catalog = state.connected || state.nativeEnv ? [{ providerID: provider.id, id: "primary", name: "Primary", enabled: true }, { providerID: provider.id, id: "fast", name: "Fast", enabled: true }] : []
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") {
      if (method === "PATCH") {
        if (state.setupFailures-- > 0) return reply({ error: "Research protection is still loading. Retry in a moment." }, 400)
        state.completed = state.completed.filter(step => step !== body.step)
        state.skipped = state.skipped.filter(step => step !== body.step)
        if (body.status === "completed") state.completed.push(String(body.step))
        else state.skipped.push(String(body.step))
      }
      return reply(progress())
    }
    if (pathname === "/engine/config") {
      if (method === "PATCH") {
        if (state.policyFailures-- > 0) return reply({ error: "That model is excluded by installation policy." }, 400)
        state.config = { ...state.config, ...body }; state.ready = true
        if (body.provider && "google-vertex" in (body.provider as object)) state.hostConfigured = true
      } else if (state.readbackFailures > 0 && state.config.model) {
        state.readbackFailures--; return reply({ ...state.config, small_model: "test-models/stale" })
      }
      return reply(state.config)
    }
    if (pathname === "/auth/reuse") {
      if (method === "GET") {
        if (state.reuseDiscoveryUnavailable) return reply({ error: "Fixture discovery unavailable." }, 503)
        const sources = state.reuseMixedStores ? [
        { id: "current-source", label: "Current OpenCode", status: "ready", unavailableCount: 0, connections: [{ id: "go-account", integrationID: "opencode-go", label: "Go account", authType: "key" }] },
        { id: "zen-source", label: "OpenCode Zen", status: "ready", unavailableCount: 0, connections: [{ id: "zen-account", integrationID: "opencode", label: "Zen account", authType: "key" }] },
        { id: "legacy-source", label: "Legacy OpenCode", status: "ready", unavailableCount: 0, connections: [{ id: "zai-account", integrationID: "zai-coding-plan", label: "ZAI account", authType: "key" }] },
        { id: "local-source", label: "OpenCode local", status: "unavailable", unavailableCount: 0, message: "This OpenCode sign-in store is unavailable or uses an unsupported format.", connections: [] },
        ] : state.reuseAvailable ? [{ id: "native-source", label: "OpenCode", status: state.reuseSourceUnavailable ? "unavailable" : "ready", unavailableCount: state.reuseSourceUnavailable ? 0 : 1, connections: state.reuseSourceUnavailable ? [] : [{ id: "zen-account", integrationID: "opencode", label: "Zen account", authType: "key" }, { id: "go-account", integrationID: "opencode-go", label: "Go account", authType: "key" }] }] : []
        const integrationID = url.searchParams.get("integrationID")
        return reply({ sources: integrationID === null ? sources : sources.map(source => ({ ...source, connections: source.connections.filter(connection => connection.integrationID === integrationID) })) })
      }
      if (body.expectedIntegrationID !== "opencode-go" || !["current-source", "native-source"].includes(String(body.sourceId)) || body.connectionId !== "go-account") return reply({ error: "Fixture selected sign-in does not match provider." }, 400)
      if (state.holdReuse) await new Promise<void>(resolve => { state.releaseReuse = resolve })
      if (state.reuseCopyFailures-- > 0) return reply({ error: "Fixture selected copy unavailable." }, 503)
      state.reusedGo = true
      return reply({ connection: { id: "copied-go", integrationID: "opencode-go", label: "Go account", authType: "key" } })
    }
    if (pathname === "/api/dochaus/provider-catalog") return reply({ data: [provider, ...(state.publicModels ? [publicProvider] : []), ...(state.secondProvider ? [{ ...provider, id: "second-models", name: "Second Models" }] : []), ...(state.hostCatalog ? [vertex, bedrock] : []), ...(state.reuseAvailable || state.reuseMixedStores ? [go] : []), ...(state.reuseMixedStores ? [zai] : [])] })
    if (pathname === "/api/provider") return reply({ data: [...(state.publicModels ? [publicProvider] : []), ...(state.connected || state.nativeEnv ? [provider, ...(state.secondProvider ? [{ ...provider, id: "second-models", name: "Second Models" }] : [])] : []), ...(state.nativeBedrock ? [bedrock] : []), ...(state.hostConfigured ? [vertex] : []), ...(state.reusedGo ? [go] : [])] })
    if (pathname === "/api/model") return reply({ data: [...publicCatalog, ...catalog, ...(state.secondProvider ? [{ providerID: "second-models", id: "other", name: "Other", enabled: true }] : []), ...(state.nativeBedrock ? [{ providerID: "amazon-bedrock", id: "bedrock-fixture", name: "Bedrock Fixture", enabled: true }] : []), ...(state.hostConfigured ? [{ providerID: "google-vertex", id: "gemini-fixture", name: "Gemini Fixture", enabled: true }] : []), ...(state.reusedGo ? [{ providerID: "opencode-go", id: "go-fixture", name: "Go Fixture", enabled: true }] : [])] })
    if (pathname === "/host/probe-vertex") return reply({ ok: true })
    if (pathname === "/api/dochaus/provider-test") return state.probeFailures-- > 0 ? reply({ ok: false, kind: "model", error: "Fixture model access unavailable." }, 400) : reply({ ok: true })
    if (pathname === "/api/integration") return reply({ data: [...(state.publicModels ? [{ id: "opencode", name: "OpenCode Zen", methods: [{ type: "oauth", id: "console-account", label: "Console account" }, { type: "key", label: "API key" }], connections: [] }] : []), { id: "test-login", name: "Fixture Sign-in", methods: [
      { type: "oauth", id: "workspace-oauth", label: "Workspace sign-in", form: [{ key: "account", type: "string", required: true, options: [{ value: "workspace", label: "Workspace" }, { value: "personal", label: "Personal" }] }, { key: "tenant", type: "string", title: "Workspace ID", required: true, when: [{ key: "account", op: "eq", value: "workspace" }] }] },
      { type: "key", label: "API key" }, { type: "command", id: "host-login", label: "Host sign-in", command: ["fixture-auth", "login"] }, { type: "env", names: ["FIXTURE_MODEL_KEY"] },
    ], connections: state.connected ? [{ type: "credential", id: "fixture-account", label: "Fixture", method: "oauth" }] : state.nativeEnv ? [{ type: "env", name: "FIXTURE_MODEL_KEY" }] : [] }, ...(state.reuseAvailable || state.reuseMixedStores ? [{ id: "opencode-go", name: "OpenCode Go", methods: [{ type: "key" }], connections: state.reusedGo ? [{ type: "credential", id: "copied-go", label: "Go account", method: "key" }] : [] }] : []), ...(state.reuseMixedStores ? [{ id: "zai-coding-plan", name: "Z.AI Coding Plan", methods: [{ type: "key" }], connections: [] }] : [])] })
    if (pathname === "/api/integration/test-login/connect/oauth" && method === "POST") return reply({ data: { attemptID: "fixture-oauth-attempt", mode: "code", url: "https://signin.example/fixture", instructions: "Finish the fixture sign-in, then enter its code.", time: { created: 1, expires: Date.now() + 600000 } } })
    if (pathname.endsWith("/fixture-oauth-attempt/complete")) { state.oauthStatus = "complete"; state.connected = true; return route.fulfill({ status: 204 }) }
    if (pathname.endsWith("/fixture-oauth-attempt")) return reply({ data: { status: state.oauthStatus, time: { created: 1, expires: Date.now() + 600000 } } })
    if (pathname.endsWith("/connect/command") && method === "POST") return reply({ data: { attemptID: "fixture-command-attempt", time: { created: 1, expires: Date.now() + 600000 } } })
    if (pathname.endsWith("/fixture-command-attempt")) return method === "DELETE" ? route.fulfill({ status: 204 }) : reply({ data: { status: state.commandStatus, message: "Continue the fixture host sign-in in its terminal window.", time: { created: 1, expires: Date.now() + 600000 } } })
    if (pathname.endsWith("/connect/key")) { state.connected = true; return route.fulfill({ status: 204 }) }
    if (pathname === "/preferences") {
      if (method === "PUT") {
        if (state.detailsFailures-- > 0) return reply({ error: "Fixture save unavailable." }, 503)
        state.preferences = { ...state.preferences, ...body }
      }
      if (method === "GET" && state.preferenceReadbackFailures-- > 0) return reply({ ...state.preferences, searchApiKey: "fixture-stale-key" })
      return reply(state.preferences)
    }
    if (/^\/research\/(exa|courtlistener|courtlistener-mcp)\/test$/.test(pathname) && method === "POST") return state.researchTestFailures-- > 0 ? reply({ error: "Fixture research connection unavailable." }, 503) : reply({ status: "verified", message: "Fixture research connection verified without contacting an external service." })
    if (pathname === "/jurisdictions" && state.jurisdictionsUnavailable) return reply({ error: "Fixture jurisdictions unavailable." }, 503)
    if (["/matters", "/jurisdictions", "/host/gcp-projects", "/host/aws-profiles"].includes(pathname)) return reply(pathname === "/host/gcp-projects" ? { projects: [] } : pathname === "/host/aws-profiles" ? { profiles: [] } : [])
    if (pathname.startsWith("/matters/")) return reply({ id: "fixture", title: "Fixture matter", dir: "/fixture-matter", documents: [] })
    if (pathname === "/api/session") return reply({ data: [], cursor: { next: null, prev: null } })
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    state.unexpected.push(`${method} ${pathname}`)
    return reply({ error: "Unexpected fixture request" }, 501)
  })
  await page.goto(baseURL!)
  return { page, context, errors }
}
async function accept(page: Awaited<ReturnType<typeof open>>["page"]) { await page.getByRole("button", { name: "I accept", exact: true }).click(); await expect(page.getByRole("dialog", { name: "Let's get you set up" })).toBeVisible() }
async function connection(page: Awaited<ReturnType<typeof open>>["page"], provider: string) {
  await page.getByLabel(/^Model provider/).selectOption(provider)
  await page.getByRole("button", { name: "Continue to connection", exact: true }).click()
}
function section(page: Awaited<ReturnType<typeof open>>["page"], heading: string) { return page.locator("section").filter({ has: heading === "Primary model" || heading === "Fast model" ? page.getByLabel(heading, { exact: true }) : page.getByRole("heading", { name: heading, exact: true }) }) }
async function screenshot(page: Awaited<ReturnType<typeof open>>["page"], name: string) { await page.screenshot({ path: path.join(evidence!, `${name}.png`), fullPage: true }) }
async function scenario(name: string, run: () => Promise<void>) {
  if (scenarioPattern && !scenarioPattern.test(name)) return
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}

const stageLabels = ["Provider", "Connection", "Models", "Research", "Details"]
function stepButton(page: Awaited<ReturnType<typeof open>>["page"], stage: string) {
  return page.getByRole("navigation", { name: "Setup progress" }).locator("li").nth(stageLabels.indexOf(stage)).getByRole("button")
}
async function currentStage(page: Awaited<ReturnType<typeof open>>["page"], stage: string) {
  const nav = page.getByRole("navigation", { name: "Setup progress" })
  await expect(nav.locator('[aria-current="step"]')).toHaveCount(1)
  await expect(nav.locator('[aria-current="step"]')).toContainText(stage)
}

try {
  await scenario("Research defaults to Entire web and preserves unsaved approved sources while showing the complete built-in list", async () => {
    const state = fixture(); state.connected = true; state.ready = true
    state.config = { model: "test-models/primary", small_model: "test-models/fast" }; state.completed = ["models"]
    state.preferences = { ...DEFAULT_DRAFTING }
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await currentStage(page, "Research")
      const modes = page.getByRole("group", { name: "Web research", exact: true })
      await expect(modes.getByRole("button", { name: "Entire web", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(modes.getByRole("button", { name: "Approved sources only", exact: true })).toHaveAttribute("aria-pressed", "false")
      await expect(page.getByLabel("Additional approved sources", { exact: true })).toHaveCount(0)
      await expect(page.locator(".research-source-list")).toHaveCount(0)
      await expect(page.getByLabel("CourtListener API token", { exact: true })).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Skip research setup", exact: true })).toHaveCount(0)
      await expect(page.getByText("Connection tests run only when requested.", { exact: false })).toHaveCount(0)
      await screenshot(page, "23-research-entire-web-default-desktop")
      await modes.getByRole("button", { name: "Approved sources only", exact: true }).focus()
      await page.keyboard.press("Space")
      const additional = page.getByLabel("Additional approved sources", { exact: true })
      await expect(additional).toHaveValue("")
      await additional.fill("draft-approved.example\narchive-approved.example")
      const list = page.locator(".research-source-list")
      const summary = list.locator("summary")
      await expect(summary).toHaveText("View 12 hosts and 15 domain endings")
      await summary.focus(); await page.keyboard.press("Enter")
      await expect(list).toHaveAttribute("open", "")
      await expect(list.getByRole("heading", { name: "Hosts and subdomains", exact: true })).toBeVisible()
      await expect(list.getByRole("heading", { name: "Domain endings", exact: true })).toBeVisible()
      await expect(list.locator("li")).toHaveCount(27)
      expect(await list.locator("li").allTextContents()).toEqual([...OFFICIAL_SOURCE_HOSTS, ...OFFICIAL_SOURCE_SUFFIXES])
      await list.scrollIntoViewIfNeeded()
      await screenshot(page, "24-research-approved-list-desktop")
      const layouts = []
      for (const viewport of [{ width: 1440, height: 1000 }, { width: 390, height: 844 }]) {
        await page.setViewportSize(viewport)
        await page.getByRole("dialog").evaluate(element => element.getAnimations().forEach(animation => animation.finish()))
        await list.locator("li").first().scrollIntoViewIfNeeded()
        await page.screenshot({ path: path.join(evidence!, `25-research-approved-${viewport.width}-top.png`), fullPage: false, mask: [page.locator('input[type="password"]')] })
        await list.locator("li").last().scrollIntoViewIfNeeded()
        const layout = await page.evaluate(() => {
          const body = document.querySelector(".onboarding-body") as HTMLElement
          const panel = document.querySelector(".onboarding-panel") as HTMLElement
          const list = document.querySelector(".research-source-list") as HTMLElement
          const last = [...list.querySelectorAll("li code")].at(-1) as HTMLElement
          const bodyRect = body.getBoundingClientRect(); const lastRect = last.getBoundingClientRect()
          return { viewport: { width: innerWidth, height: innerHeight }, documentOverflow: document.documentElement.scrollWidth - innerWidth, bodyOverflow: body.scrollWidth - body.clientWidth, panelOverflow: panel.scrollWidth - panel.clientWidth, listOverflow: list.scrollWidth - list.clientWidth, lastEntry: last.textContent, lastEntryWithinScrollArea: lastRect.left >= bodyRect.left && lastRect.right <= bodyRect.right && lastRect.top >= bodyRect.top && lastRect.bottom <= bodyRect.bottom, columns: getComputedStyle(document.querySelector(".research-source-columns")!).gridTemplateColumns, outsideViewport: [...document.querySelectorAll("body *")].filter(element => element.getBoundingClientRect().right > innerWidth + 1 && element.getBoundingClientRect().width > 0).slice(0, 12).map(element => ({ tag: element.tagName, className: element.className, right: element.getBoundingClientRect().right, text: element.textContent?.slice(0, 100) })) }
        })
        layouts.push(layout)
        await writeFile(path.join(evidence!, "research-default-layouts.json"), JSON.stringify(layouts, null, 2))
        // The main screen behind the dialog can overflow independently; keep
        // that diagnostic in evidence while checking the Research surface.
        for (const overflow of [layout.bodyOverflow, layout.panelOverflow, layout.listOverflow]) expect(overflow).toBeLessThanOrEqual(1)
        expect(layout.lastEntryWithinScrollArea).toBe(true)
        await page.screenshot({ path: path.join(evidence!, `26-research-approved-${viewport.width}-bottom.png`), fullPage: false, mask: [page.locator('input[type="password"]')] })
      }
      await modes.getByRole("button", { name: "Entire web", exact: true }).click()
      await expect(additional).toHaveCount(0)
      await expect(list).toHaveCount(0)
      await page.locator(".onboarding-body").evaluate(element => { element.scrollTop = 0 })
      await page.screenshot({ path: path.join(evidence!, "28-research-entire-web-mobile.png"), fullPage: false, mask: [page.locator('input[type="password"]')] })
      await stepButton(page, "Models").click(); await stepButton(page, "Research").click()
      await expect(modes.getByRole("button", { name: "Entire web", exact: true })).toHaveAttribute("aria-pressed", "true")
      await modes.getByRole("button", { name: "Approved sources only", exact: true }).click()
      await expect(additional).toHaveValue("draft-approved.example\narchive-approved.example")
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      expect(state.preferences.approvedSources).toEqual([])
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "research-default-layouts.json"), JSON.stringify(layouts, null, 2))
      await writeFile(path.join(evidence!, "research-default-requests.json"), JSON.stringify(state.requests, null, 2))
    } finally { await context.close() }
  })
  await scenario("Research saves Entire web without clearing approved sources or unrelated preferences and reloads explicit approved mode", async () => {
    const state = fixture(); state.connected = true; state.ready = true
    state.config = { model: "test-models/primary", small_model: "test-models/fast" }; state.completed = ["models"]
    state.preferences = { ...state.preferences, webResearch: "official", approvedSources: ["saved.example", "archive.example"], caseLawApiKey: "saved-fixture-courtlistener", courtlistenerMcpEnabled: true, posture: "conservative", houseStyle: "Preserve this unrelated drafting preference.", telemetry: { retention_days: 31, retention_days_source: "user", rate_overrides: {} } }
    const original = structuredClone(state.preferences)
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await currentStage(page, "Research")
      const modes = page.getByRole("group", { name: "Web research", exact: true })
      await expect(modes.getByRole("button", { name: "Approved sources only", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(page.getByLabel("Additional approved sources", { exact: true })).toHaveValue("saved.example\narchive.example")
      await modes.getByRole("button", { name: "Entire web", exact: true }).click()
      await expect(page.getByLabel("Additional approved sources", { exact: true })).toHaveCount(0)
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await currentStage(page, "Details")
      expect(state.preferences).toEqual({ ...original, webResearch: "open" })
      const firstSave = state.requests.filter(request => request.path === "/preferences" && request.method === "PUT")
      expect(firstSave).toHaveLength(1)
      expect(firstSave[0].body).toEqual({ webResearch: "open", approvedSources: original.approvedSources, searchApiKey: original.searchApiKey, caseLawApiKey: original.caseLawApiKey, courtlistenerMcpEnabled: true })
      await page.reload(); await stepButton(page, "Research").click()
      await expect(modes.getByRole("button", { name: "Entire web", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(page.getByLabel("Additional approved sources", { exact: true })).toHaveCount(0)
      await modes.getByRole("button", { name: "Approved sources only", exact: true }).click()
      await expect(page.getByLabel("Additional approved sources", { exact: true })).toHaveValue("saved.example\narchive.example")
      await page.getByLabel("Additional approved sources", { exact: true }).fill("saved.example\narchive.example\nnew-approved.example")
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await currentStage(page, "Details")
      await page.reload(); await stepButton(page, "Research").click()
      await expect(modes.getByRole("button", { name: "Approved sources only", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(page.getByLabel("Additional approved sources", { exact: true })).toHaveValue("saved.example\narchive.example\nnew-approved.example")
      await expect(page.getByLabel("Exa API key", { exact: true })).toHaveValue(original.searchApiKey)
      await expect(page.getByLabel("Exa API key", { exact: true })).toHaveAttribute("type", "password")
      await expect(page.getByLabel("CourtListener API token", { exact: true })).toHaveValue(original.caseLawApiKey)
      await expect(page.getByLabel("CourtListener API token", { exact: true })).toHaveAttribute("type", "password")
      expect(state.preferences).toEqual({ ...original, approvedSources: ["saved.example", "archive.example", "new-approved.example"] })
      expect(state.requests.filter(request => request.path === "/preferences" && request.method === "PUT")).toHaveLength(2)
      expect(state.requests.filter(request => request.method !== "GET" && !(request.path === "/preferences" && request.method === "PUT") && !(request.path === "/setup" && request.method === "PATCH"))).toEqual([])
      expect(state.requests.filter(request => /probe|provider-test|research-test|connect\//.test(request.path))).toEqual([])
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await page.screenshot({ path: path.join(evidence!, "27-research-explicit-approved-reloaded.png"), fullPage: true, mask: [page.locator('input[type="password"]')] })
      await writeFile(path.join(evidence!, "research-save-requests.json"), JSON.stringify(state.requests, (key, value) => ["searchApiKey", "caseLawApiKey"].includes(key) ? "[redacted synthetic fixture]" : value, 2))
    } finally { await context.close() }
  })
  await scenario("Research preserves a hidden CourtListener token and saves only the Exa key before an explicit mocked test", async () => {
    const state = fixture(); state.connected = true; state.ready = true
    state.config = { model: "test-models/primary", small_model: "test-models/fast" }; state.completed = ["models"]
    state.preferences = { ...state.preferences, searchApiKey: "", caseLawApiKey: "saved-fixture-token", courtlistenerMcpEnabled: false }
    const original = structuredClone(state.preferences)
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await currentStage(page, "Research")
      const connector = page.getByRole("group", { name: "CourtListener connector", exact: true })
      const token = page.getByLabel("CourtListener API token", { exact: true })
      await expect(token).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Test Exa key", exact: true })).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Save key and test", exact: true })).toHaveCount(0)
      await connector.getByRole("button", { name: "On", exact: true }).click()
      await expect(token).toHaveValue("saved-fixture-token")
      await token.fill("draft-fixture-token")
      await connector.getByRole("button", { name: "Off", exact: true }).click()
      await expect(token).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Test token", exact: true })).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Test connector", exact: true })).toHaveCount(0)
      await connector.getByRole("button", { name: "On", exact: true }).click()
      await expect(token).toHaveValue("draft-fixture-token")
      await connector.getByRole("button", { name: "Off", exact: true }).click()
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await currentStage(page, "Details")
      expect(state.preferences).toEqual({ ...original, caseLawApiKey: "draft-fixture-token" })
      await page.reload(); await stepButton(page, "Research").click()
      await expect(token).toHaveCount(0)
      await connector.getByRole("button", { name: "On", exact: true }).click()
      await expect(token).toHaveValue("draft-fixture-token")
      await token.fill("unsaved-side-token")
      await page.getByRole("group", { name: "Web research", exact: true }).getByRole("button", { name: "Approved sources only", exact: true }).click()
      await page.getByLabel("Additional approved sources", { exact: true }).fill("unsaved-side.example")
      await page.getByLabel("Exa API key", { exact: true }).fill("  changed-fixture-exa  ")
      const saveAndTest = page.getByRole("button", { name: "Save key and test", exact: true })
      await expect(saveAndTest).toBeEnabled()
      const start = state.requests.length
      const failed = page.getByText("Could not save or test this connection. Try again.", { exact: true })
      state.detailsFailures = 1
      await saveAndTest.click()
      await expect(failed).toBeVisible(); await expect(saveAndTest).toBeEnabled()
      expect(state.requests.filter(request => request.path.startsWith("/research/"))).toHaveLength(0)
      expect(state.preferences.searchApiKey).toBe("")
      state.preferenceReadbackFailures = 1
      await saveAndTest.click()
      await expect.poll(() => state.requests.slice(start).filter(request => request.path === "/preferences" && request.method === "GET").length).toBe(1)
      await expect(saveAndTest).toBeEnabled()
      expect(state.requests.filter(request => request.path.startsWith("/research/"))).toHaveLength(0)
      state.researchTestFailures = 1
      await saveAndTest.click()
      await expect(page.getByRole("button", { name: "Test Exa key", exact: true })).toBeEnabled()
      await expect(failed).toBeVisible()
      expect(state.requests.filter(request => request.path === "/research/exa/test" && request.method === "POST")).toHaveLength(1)
      await page.getByRole("button", { name: "Test Exa key", exact: true }).click()
      await expect(page.getByText("Fixture research connection verified without contacting an external service.", { exact: true })).toBeVisible()
      await currentStage(page, "Research")
      expect(state.requests.slice(start).map(request => `${request.method} ${request.path}`)).toEqual(["PUT /preferences", "PUT /preferences", "GET /preferences", "PUT /preferences", "GET /preferences", "POST /research/exa/test", "POST /research/exa/test"])
      expect(state.requests.slice(start).filter(request => request.method === "PUT").map(request => request.body)).toEqual(Array(3).fill({ searchApiKey: "changed-fixture-exa" }))
      expect(state.preferences).toEqual({ ...original, caseLawApiKey: "draft-fixture-token", searchApiKey: "changed-fixture-exa" })
      await expect(token).toHaveValue("unsaved-side-token")
      await expect(page.getByLabel("Additional approved sources", { exact: true })).toHaveValue("unsaved-side.example")
      await expect(page.getByRole("group", { name: "Web research", exact: true }).getByRole("button", { name: "Approved sources only", exact: true })).toHaveAttribute("aria-pressed", "true")
      await expect(connector.getByRole("button", { name: "On", exact: true })).toHaveAttribute("aria-pressed", "true")
      expect(state.requests.filter(request => request.path === "/setup" && request.method === "PATCH")).toHaveLength(1)
      expect(state.requests.filter(request => request.method === "POST" && request.path !== "/research/exa/test")).toHaveLength(0)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await page.getByLabel("Exa API key", { exact: true }).scrollIntoViewIfNeeded()
      await page.screenshot({ path: path.join(evidence!, "29-research-exa-mocked-success-desktop.png"), fullPage: false, mask: [page.locator('input[type="password"]')] })
      await page.setViewportSize({ width: 390, height: 844 })
      await token.scrollIntoViewIfNeeded()
      await page.screenshot({ path: path.join(evidence!, "30-research-connector-on-mobile.png"), fullPage: false, mask: [page.locator('input[type="password"]')] })
      await writeFile(path.join(evidence!, "research-key-test-requests.json"), JSON.stringify(state.requests, (key, value) => ["searchApiKey", "caseLawApiKey"].includes(key) ? "[redacted synthetic fixture]" : value, 2))
    } finally { await context.close() }
  })
  await scenario("clickable steps explain prerequisites, preserve an unsaved key and never save or authenticate on navigation", async () => {
    const state = fixture(); state.secondProvider = true; const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await expect(page.getByRole("button", { name: "Advanced settings", exact: true })).toHaveCount(0)
      await expect(page.getByRole("navigation", { name: "Setup progress" }).getByRole("button")).toHaveCount(5)
      for (const stage of stageLabels) await expect(stepButton(page, stage)).toHaveAccessibleName(new RegExp(stage))
      await expect(stepButton(page, "Provider")).toBeEnabled()
      for (const stage of stageLabels.slice(1)) {
        const button = stepButton(page, stage)
        await expect(button).toBeDisabled()
        const reason = await button.evaluate(element => [element.getAttribute("title"), ...(element.getAttribute("aria-describedby") ?? "").split(/\s+/).filter(Boolean).map(id => document.getElementById(id)?.textContent)].filter(Boolean).join(" "))
        expect(reason.length, `${stage} has a visible or accessible prerequisite explanation`).toBeGreaterThan(0)
        await button.evaluate(element => (element as HTMLButtonElement).click())
        await currentStage(page, "Provider")
      }
      await page.getByLabel(/^Model provider/).selectOption("test-models")
      await expect(stepButton(page, "Connection")).toBeEnabled()
      await stepButton(page, "Connection").focus()
      await page.keyboard.press("Enter")
      await currentStage(page, "Connection")
      await expect(page.getByRole("heading", { name: "Connect Fixture Models", exact: true })).toBeFocused()
      await page.getByLabel(/^Connection method/).selectOption("1")
      await page.getByLabel("API key", { exact: true }).fill("unsaved-stepper-key")
      await stepButton(page, "Provider").focus()
      await page.keyboard.press("Space")
      await expect(page.getByRole("heading", { name: "Choose your model provider", exact: true })).toBeFocused()
      await expect(page.getByLabel(/^Model provider/)).toHaveValue("test-models")
      await stepButton(page, "Connection").click()
      await expect(page.getByLabel("API key", { exact: true })).toHaveValue("unsaved-stepper-key")
      for (const stage of ["Models", "Research", "Details"]) await expect(stepButton(page, stage)).toBeDisabled()
      await screenshot(page, "21-stepper-unsaved-key")
      await stepButton(page, "Provider").click()
      await page.getByLabel(/^Model provider/).selectOption("second-models")
      await stepButton(page, "Connection").click()
      await page.getByLabel(/^Connection method/).selectOption("1")
      await expect(page.getByLabel("API key", { exact: true })).toHaveValue("")
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      expect(state.requests.filter(request => request.path.includes("probe") || request.path.includes("provider-test"))).toHaveLength(0)
      expect(state.completed).toEqual([]); expect(state.skipped).toEqual([])
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "stepper-prerequisite-requests.json"), JSON.stringify(state.requests, null, 2))
    } finally { await context.close() }
  })
  await scenario("all enabled steps preserve unsaved model, Research and Details values without marking visits complete", async () => {
    const state = fixture(); state.connected = true; state.ready = true; state.config = { model: "test-models/primary", small_model: "test-models/fast" }; state.completed = ["models"]
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await stepButton(page, "Models").click()
      await currentStage(page, "Models")
      for (const stage of stageLabels) await expect(stepButton(page, stage)).toBeEnabled()
      await page.getByLabel("Primary model", { exact: true }).selectOption("test-models/fast")
      await page.getByLabel("Fast model", { exact: true }).selectOption("test-models/primary")
      await stepButton(page, "Research").click()
      await page.getByLabel("Exa API key", { exact: true }).fill("unsaved-stepper-research")
      await stepButton(page, "Details").click()
      await page.getByPlaceholder("Jane Doe").fill("Unsaved Stepper Lawyer")
      await page.getByPlaceholder("Doe & Partners LLP").fill("Unsaved Stepper Firm")
      await page.getByLabel("Default jurisdictions", { exact: true }).click()
      await expect(page.getByPlaceholder("Search jurisdictions")).toBeFocused()
      await stepButton(page, "Provider").click()
      await expect(page.locator(".jx-popover")).not.toBeVisible()
      await expect(page.getByRole("heading", { name: "Choose your model provider", exact: true })).toBeFocused()
      await expect(page.getByLabel(/^Model provider/)).toHaveValue("test-models")
      await stepButton(page, "Connection").click()
      await stepButton(page, "Details").click()
      await expect(page.getByPlaceholder("Jane Doe")).toHaveValue("Unsaved Stepper Lawyer")
      await expect(page.getByPlaceholder("Doe & Partners LLP")).toHaveValue("Unsaved Stepper Firm")
      const researchItem = page.getByRole("navigation", { name: "Setup progress" }).locator("li").nth(3)
      await expect(researchItem).not.toHaveClass(/is-complete/)
      await expect(researchItem).not.toContainText("Completed:")
      await expect(page.getByRole("navigation", { name: "Setup progress" }).locator("li").nth(2)).toHaveClass(/is-complete/)
      await stepButton(page, "Research").click()
      await expect(page.getByLabel("Exa API key", { exact: true })).toHaveValue("unsaved-stepper-research")
      await stepButton(page, "Models").focus()
      await page.keyboard.press("Enter")
      await expect(page.getByRole("heading", { name: "Choose your models", exact: true })).toBeFocused()
      await expect(page.getByLabel("Primary model", { exact: true })).toHaveValue("test-models/fast")
      await expect(page.getByLabel("Fast model", { exact: true })).toHaveValue("test-models/primary")
      await page.getByRole("button", { name: "Save and continue", exact: true }).focus()
      await page.keyboard.press("Tab")
      await expect(page.getByRole("button", { name: "Set up later", exact: true })).toBeFocused()
      await page.keyboard.press("Shift+Tab")
      await expect(page.getByRole("button", { name: "Save and continue", exact: true })).toBeFocused()
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      expect(state.completed).toEqual(["models"]); expect(state.skipped).toEqual([])
      expect(state.preferences).toMatchObject({ attorney: "Saved Lawyer", firm: "Saved Firm", searchApiKey: "saved-fixture-exa" })
      expect(state.config).toEqual({ model: "test-models/primary", small_model: "test-models/fast" })
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "stepper-unsaved-roundtrip-requests.json"), JSON.stringify(state.requests, null, 2))
      await screenshot(page, "22-stepper-model-roundtrip")
    } finally { await context.close() }
  })
  await scenario("native environment and public models stay available without automatic sign-in or cloud probes", async () => {
    const state = fixture(); state.publicModels = true; state.nativeEnv = true; state.nativeBedrock = true; state.hostCatalog = true
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      const chooser = page.getByLabel(/^Model provider/)
      await expect(chooser).toHaveValue("")
      await expect(chooser.locator("optgroup").first()).toHaveAttribute("label", "Detected")
      await expect(chooser.locator('optgroup[label="Detected"] option[value="test-models"]')).toContainText("Fixture Models")
      await expect(chooser.locator('optgroup[label="Detected"] option[value="opencode"]')).toContainText("Public models")
      await expect(page.getByText("Reuse a saved OpenCode sign-in (optional)", { exact: true })).toHaveCount(0)
      await page.setViewportSize({ width: 390, height: 844 })
      const providerBounds = await chooser.boundingBox()
      const modalBounds = await page.getByRole("dialog", { name: "Let's get you set up" }).boundingBox()
      expect(providerBounds!.x).toBeGreaterThanOrEqual(modalBounds!.x)
      expect(providerBounds!.x + providerBounds!.width).toBeLessThanOrEqual(modalBounds!.x + modalBounds!.width)
      await screenshot(page, "16-provider-selection-narrow")
      await page.setViewportSize({ width: 1440, height: 1000 })
      await connection(page, "amazon-bedrock")
      await expect(page.getByLabel("AWS profile", { exact: true })).toHaveValue("work")
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeDisabled()
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "google-vertex")
      await expect(page.getByLabel("GCP project id", { exact: true })).toHaveValue("")
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeDisabled()
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "opencode")
      await page.getByText("Connect an account (optional)", { exact: true }).click()
      await expect(page.getByLabel("Connection method").locator("option")).toHaveText(["Console account", "API key"])
      await page.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(page.getByLabel("Primary model", { exact: true }).locator('option[value="opencode/free-fixture"]')).toHaveCount(1)
      await page.getByRole("button", { name: "Back to connection", exact: true }).click()
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "test-models")
      await page.getByText("Change connection (optional)", { exact: true }).click()
      await expect(page.getByLabel("Connection method").locator("option")).toHaveText(["Workspace sign-in", "API key", "Host sign-in"])
      await expect(page.getByRole("heading", { name: "Connection status", exact: true })).toHaveCount(0)
      await expect(page.getByRole("heading", { name: "Saved OpenCode sign-ins", exact: true })).toHaveCount(0)
      await expect(page.getByLabel(/^Saved sign-in/)).toHaveCount(0)
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      expect(state.requests.filter(request => request.path.includes("probe") || request.path.includes("provider-test"))).toHaveLength(0)
      await page.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(page.getByLabel("Primary model", { exact: true }).locator('option[value="test-models/primary"]')).toHaveCount(1)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "native-discovery-readonly-requests.json"), JSON.stringify(state.requests, null, 2))
      await screenshot(page, "10-native-discovery-models")
    } finally { await context.close() }
  })
  await scenario("only the selected provider saved sign-ins appear and provider changes remove stale choices", async () => {
    const state = fixture(); state.reuseMixedStores = true
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      expect(state.requests.filter(request => request.path === "/auth/reuse")).toHaveLength(0)
      await connection(page, "opencode-go")
      const method = page.getByLabel(/^Connection method/)
      await expect(method.locator("option")).toHaveText(["Use saved sign-in", "API key"])
      await expect(page.getByLabel("API key", { exact: true })).toBeVisible()
      await method.selectOption("__saved__")
      const saved = page.getByLabel(/^Saved sign-in/)
      await expect(saved.locator("option")).toHaveText(["Choose a saved sign-in", "Go account · API key (Current OpenCode)"])
      await expect(page.getByText("Choose a saved OpenCode Go sign-in to copy into this installation.", { exact: false })).toBeVisible()
      await expect(page.getByText("Zen account", { exact: false })).toHaveCount(0)
      await expect(page.getByText("ZAI account", { exact: false })).toHaveCount(0)
      await saved.selectOption(JSON.stringify(["current-source", "go-account"]))
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "zai-coding-plan")
      await expect(page.getByLabel(/^Saved sign-in/)).toHaveCount(0)
      await expect(page.getByLabel(/^Connection method/).locator("option")).toContainText(["Use saved sign-in"])
      await page.getByLabel(/^Connection method/).selectOption("__saved__")
      await expect(page.getByLabel(/^Saved sign-in/).locator("option")).toHaveText(["Choose a saved sign-in", "ZAI account · API key (Legacy OpenCode)"])
      await expect(page.getByText("Go account", { exact: false })).toHaveCount(0)
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "test-models")
      await expect(page.getByLabel(/^Connection method/).locator("option")).toHaveText(["Workspace sign-in", "API key", "Host sign-in"])
      await expect(page.getByLabel(/^Saved sign-in/)).toHaveCount(0)
      await expect(page.getByRole("button", { name: "Copy selected sign-in", exact: true })).toHaveCount(0)
      await expect(page.getByText("Use saved sign-in", { exact: true })).toHaveCount(0)
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "opencode-go")
      await page.getByLabel(/^Connection method/).selectOption("__saved__")
      await expect(page.getByLabel(/^Saved sign-in/)).toHaveValue("")
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      const discovery = state.requests.filter(request => request.path === "/auth/reuse")
      expect(discovery.every(request => Boolean(request.query.integrationID))).toBe(true)
      expect(discovery.map(request => request.query.integrationID)).toEqual(expect.arrayContaining(["opencode-go", "zai-coding-plan", "test-login"]))
      const beforeCopy = state.requests.length
      await page.getByLabel(/^Saved sign-in/).selectOption(JSON.stringify(["current-source", "go-account"]))
      await page.getByRole("button", { name: "Copy selected sign-in", exact: true }).click()
      await page.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(page.getByLabel("Primary model", { exact: true })).toHaveValue("opencode-go/go-fixture")
      const refreshed = state.requests.slice(beforeCopy)
      expect(refreshed.filter(request => request.path === "/auth/reuse" && request.method === "POST").map(request => request.body)).toEqual([{ sourceId: "current-source", connectionId: "go-account", expectedIntegrationID: "opencode-go" }])
      for (const endpoint of ["/api/integration", "/api/dochaus/provider-catalog", "/api/provider", "/api/model"]) expect(refreshed.some(request => request.path === endpoint && request.method === "GET")).toBe(true)
      expect(state.requests.filter(request => request.path.includes("probe") || request.path.includes("provider-test"))).toHaveLength(0)
      expect(state.requests.filter(request => request.path === "/engine/config" && request.method === "PATCH")).toHaveLength(0)
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(page.getByRole("heading", { name: "Set up legal research", exact: true })).toBeVisible()
      const writeIndex = state.requests.findIndex(request => request.path === "/engine/config" && request.method === "PATCH")
      expect(state.requests[writeIndex].body).toEqual({ model: "opencode-go/go-fixture", small_model: "opencode-go/go-fixture" })
      expect(state.requests.slice(writeIndex + 1).some(request => request.path === "/engine/config" && request.method === "GET")).toBe(true)
      expect(state.requests.filter(request => request.path === "/auth/reuse" && request.method === "POST")).toHaveLength(1)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "selected-provider-reuse-requests.json"), JSON.stringify(state.requests, null, 2))
      await screenshot(page, "11-selected-go-reuse-model-pair-readback")
    } catch (error) {
      await screenshot(page, "selected-provider-reuse-failure")
      await writeFile(path.join(evidence!, "selected-provider-reuse-failure.json"), JSON.stringify({ requests: state.requests, unexpected: state.unexpected, errors, text: await page.locator("body").innerText() }, null, 2))
      throw error
    } finally { await context.close() }
  })
  await scenario("selected sign-in reuse handles discovery failure and explicit pending copy retry", async () => {
    const state = fixture(); state.reuseAvailable = true; state.reuseDiscoveryUnavailable = true
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await connection(page, "opencode-go")
      await expect(page.getByText("Could not check for a saved sign-in for this provider.", { exact: false })).toBeVisible()
      await expect(page.getByLabel("API key", { exact: true })).toBeVisible()
      await expect(page.getByLabel(/^Connection method/)).toHaveCount(0)
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      state.reuseDiscoveryUnavailable = false
      await page.getByRole("button", { name: "Retry saved sign-ins", exact: true }).click()
      await page.getByLabel(/^Connection method/).selectOption("__saved__")
      await expect(page.getByLabel(/^Saved sign-in/).locator("option")).toHaveText(["Choose a saved sign-in", "Go account · API key (OpenCode)"])
      await expect(page.getByRole("button", { name: "Copy selected sign-in", exact: true })).toBeDisabled()
      await page.getByLabel(/^Saved sign-in/).selectOption(JSON.stringify(["native-source", "go-account"]))
      state.holdReuse = true; state.reuseCopyFailures = 1
      await page.getByRole("button", { name: "Copy selected sign-in", exact: true }).click()
      await expect(page.getByRole("button", { name: "Copying selected sign-in…", exact: true })).toBeDisabled()
      await expect.poll(() => Boolean(state.releaseReuse)).toBe(true)
      await screenshot(page, "08-selected-sign-in-pending")
      state.releaseReuse!(); state.holdReuse = false
      await expect(page.getByRole("alert")).toContainText("selected sign-in could not be copied")
      await expect(page.getByLabel(/^Saved sign-in/)).toHaveValue(JSON.stringify(["native-source", "go-account"]))
      await page.getByRole("button", { name: "Copy selected sign-in", exact: true }).click()
      await page.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(page.getByLabel("Primary model", { exact: true })).toHaveValue("opencode-go/go-fixture")
      expect(state.requests.filter(request => request.path === "/auth/reuse" && request.method === "POST").map(request => request.body)).toEqual(Array(2).fill({ sourceId: "native-source", connectionId: "go-account", expectedIntegrationID: "opencode-go" }))
      expect(state.requests.filter(request => request.path.includes("probe") || request.path.includes("provider-test"))).toHaveLength(0)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await screenshot(page, "09-selected-go-sign-in-discovered")
    } finally { state.releaseReuse?.(); await context.close() }
  })
  await scenario("host connection settings are prefilled; save and failed tests never enable models", async () => {
    const state = fixture(); state.hostCatalog = true; state.hostDetected = true; state.nativeBedrock = true
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await connection(page, "amazon-bedrock")
      await expect(page.getByLabel("AWS profile", { exact: true })).toHaveValue("work")
      await expect(page.getByLabel("AWS region", { exact: true })).toHaveValue("us-east-1")
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeDisabled()
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      await page.getByLabel("AWS profile", { exact: true }).fill("")
      await page.getByLabel("Model to test").selectOption("bedrock-fixture")
      await page.getByRole("button", { name: "Test and enable", exact: true }).click()
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeEnabled()
      expect(state.requests.find(request => request.path === "/engine/config" && request.method === "PATCH")!.body).toMatchObject({ provider: { "amazon-bedrock": { options: { profile: null, region: "us-east-1" } } } })
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "google-vertex")
      await expect(page.getByLabel("GCP project id", { exact: true })).toHaveValue("fixture-project")
      await expect(page.getByLabel("Location", { exact: true })).toHaveValue("us-central1")
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeDisabled()
      expect(state.requests.filter(request => request.path === "/api/dochaus/provider-test")).toHaveLength(1)
      await page.getByLabel("GCP project id", { exact: true }).fill("chosen-fixture-project")
      await page.getByRole("button", { name: "Save connection settings", exact: true }).click()
      await expect(page.getByRole("button", { name: "Test and enable", exact: true })).toBeVisible()
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeDisabled()
      expect(state.requests.filter(request => request.path === "/api/dochaus/provider-test")).toHaveLength(1)
      const saved = state.requests.find(request => request.path === "/engine/config" && request.method === "PATCH" && request.body.provider && "google-vertex" in (request.body.provider as object))!
      expect(saved.body).toMatchObject({ provider: { "google-vertex": { options: { project: "chosen-fixture-project", location: "us-central1" } } } })
      await page.getByLabel("Model to test").selectOption("gemini-fixture")
      state.probeFailures = 1
      await page.getByRole("button", { name: "Test and enable", exact: true }).click()
      await expect(page.getByText("Fixture model access unavailable.", { exact: true })).toBeVisible()
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeDisabled()
      await page.getByRole("button", { name: "Test and enable", exact: true }).click()
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeEnabled()
      expect(state.requests.filter(request => request.path === "/api/dochaus/provider-test").map(request => request.body)).toEqual([{ providerID: "amazon-bedrock", modelID: "bedrock-fixture" }, ...Array(2).fill({ providerID: "google-vertex", modelID: "gemini-fixture" })])
      expect(state.requests.filter(request => request.path === "/host/probe-vertex")).toHaveLength(0)
      expect(state.requests.filter(request => request.path === "/host/gcp-projects" || request.path === "/host/aws-profiles")).toHaveLength(0)
      expect(state.requests.filter(request => request.path === "/auth/reuse")).toHaveLength(0)
      await page.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(section(page, "Primary model").locator("select")).toHaveValue("google-vertex/gemini-fixture")
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "host-explicit-test-requests.json"), JSON.stringify(state.requests, null, 2))
      await screenshot(page, "07-host-provider-discovery")
    } finally { await context.close() }
  })
  await scenario("normal Settings tests Vertex through the native engine and preserves failed model access", async () => {
    const state = fixture(); state.connected = true; state.ready = true; state.hostCatalog = true; state.hostConfigured = true
    state.config = { model: "test-models/primary", small_model: "test-models/fast" }; state.completed = ["models"]; state.skipped = ["research", "details"]
    const { page, context, errors } = await open(state)
    try {
      await page.getByRole("button", { name: "I accept", exact: true }).click()
      await page.getByRole("button", { name: "Settings", exact: true }).click()
      await page.getByRole("tab", { name: "Providers", exact: true }).click()
      const card = page.locator(".settings-needs").filter({ hasText: "Google Vertex AI" })
      await card.getByLabel("GCP project id", { exact: true }).fill("chosen-fixture-project")
      await card.getByLabel("Location", { exact: true }).selectOption("us-central1")
      await card.getByLabel("Model to test").selectOption("gemini-fixture")
      const nativeTests = () => state.requests.filter(request => request.path === "/api/dochaus/provider-test" && request.body.providerID === "google-vertex")
      expect(nativeTests()).toHaveLength(0)
      state.probeFailures = 1
      await card.getByRole("button", { name: "Test and enable", exact: true }).click()
      await expect(card.getByText("Fixture model access unavailable.", { exact: true })).toBeVisible()
      expect(nativeTests().map(request => request.body)).toEqual([{ providerID: "google-vertex", modelID: "gemini-fixture" }])
      expect(state.requests.filter(request => request.path === "/host/probe-vertex")).toHaveLength(0)
      expect(state.requests.find(request => request.path === "/engine/config" && request.method === "PATCH" && request.body.provider && "google-vertex" in (request.body.provider as object))?.body)
        .toMatchObject({ provider: { "google-vertex": { options: { project: "chosen-fixture-project", location: "us-central1" } } } })
      await card.getByRole("button", { name: "Test and enable", exact: true }).click()
      await expect(card).toHaveCount(0)
      expect(nativeTests()).toHaveLength(2)
      expect(state.requests.filter(request => request.path === "/host/probe-vertex")).toHaveLength(0)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "settings-native-vertex-requests.json"), JSON.stringify(state.requests, null, 2))
      await screenshot(page, "26-settings-native-vertex")
    } finally { await context.close() }
  })
  await scenario("native key connection is explicit and advances to the selected provider models", async () => {
    const state = fixture(); const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await connection(page, "test-models")
      await page.getByLabel(/^Connection method/).selectOption("1")
      await page.getByLabel("API key", { exact: true }).fill("synthetic-browser-key")
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      await page.getByRole("button", { name: "Connect", exact: true }).click()
      await expect(page.getByRole("button", { name: "Continue to models", exact: true })).toBeEnabled()
      await expect(stepButton(page, "Models")).toBeEnabled()
      await expect(stepButton(page, "Research")).toBeDisabled()
      await expect(stepButton(page, "Details")).toBeDisabled()
      expect(state.requests.filter(request => request.path.endsWith("/connect/key")).map(request => request.body)).toEqual([{ key: "synthetic-browser-key", answer: {} }])
      expect(state.requests.filter(request => (request.path === "/auth/reuse" && request.method === "POST") || request.path.includes("provider-test"))).toHaveLength(0)
      await page.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(section(page, "Primary model").locator("select")).toHaveValue("")
      await expect(page.getByRole("button", { name: "Save and continue", exact: true })).toBeDisabled()
      await section(page, "Primary model").locator("select").selectOption("test-models/primary")
      await expect(page.getByRole("button", { name: "Save and continue", exact: true })).toBeEnabled()
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(1)
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(page.getByRole("heading", { name: "Set up legal research", exact: true })).toBeVisible()
      await expect(stepButton(page, "Research")).toBeEnabled()
      await expect(stepButton(page, "Details")).toBeEnabled()
      const writeIndex = state.requests.findIndex(request => request.path === "/engine/config" && request.method === "PATCH")
      expect(state.requests[writeIndex].body).toEqual({ model: "test-models/primary", small_model: "test-models/primary" })
      expect(state.requests.slice(writeIndex + 1).some(request => request.path === "/engine/config" && request.method === "GET")).toBe(true)
      expect(state.config).toMatchObject({ model: "test-models/primary", small_model: "test-models/primary" })
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await writeFile(path.join(evidence!, "native-key-model-pair-readback.json"), JSON.stringify(state.requests, null, 2))
      await screenshot(page, "15-native-key-connected")
    } finally { await context.close() }
  })
  await scenario("public and environment providers expose clear native methods and preserve pending OAuth", async () => {
    const state = fixture(); state.publicModels = true; state.nativeEnv = true
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await connection(page, "opencode")
      await page.getByText("Connect an account (optional)", { exact: true }).click()
      await expect(page.getByLabel(/^Connection method/).locator("option")).toHaveText(["Console account", "API key"])
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      await screenshot(page, "17-public-native-sign-in")
      await page.setViewportSize({ width: 390, height: 844 })
      await screenshot(page, "18-public-native-sign-in-narrow")
      await page.setViewportSize({ width: 1440, height: 1000 })
      await page.getByRole("button", { name: "Back to providers", exact: true }).click()
      await connection(page, "test-models")
      await expect(page.getByLabel(/^Connection method/).locator("option")).toHaveText(["Workspace sign-in", "API key", "Host sign-in"])
      const disclosure = page.getByText("Change connection (optional)", { exact: true })
      await disclosure.click()
      await page.getByLabel(/^account/).selectOption("workspace")
      await page.getByLabel("Workspace ID", { exact: true }).fill("fixture-workspace")
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(0)
      await page.getByRole("button", { name: "Connect", exact: true }).click()
      const link = page.getByRole("link", { name: "Open provider sign-in", exact: true })
      await expect(link).toHaveAttribute("href", "https://signin.example/fixture")
      await expect(page.getByText("Finish the fixture sign-in, then enter its code.", { exact: true })).toBeVisible()
      await expect(page.getByLabel(/^Connection method/)).toHaveCount(0)
      await disclosure.click()
      await expect(link).not.toBeVisible()
      await disclosure.click()
      await expect(link).toBeVisible()
      await expect(page.getByText("Finish the fixture sign-in, then enter its code.", { exact: true })).toBeVisible()
      await page.getByText("Environment setup", { exact: true }).click()
      await expect(page.getByText("Set FIXTURE_MODEL_KEY", { exact: false })).toBeVisible()
      await expect(link).toBeVisible()
      await page.getByText("Environment setup", { exact: true }).click()
      await page.setViewportSize({ width: 390, height: 844 })
      await page.getByLabel("Authorization code").scrollIntoViewIfNeeded()
      await expect(link).toHaveAttribute("href", "https://signin.example/fixture")
      await expect(page.getByText("Finish the fixture sign-in, then enter its code.", { exact: true })).toBeVisible()
      expect(state.requests.filter(request => request.path.endsWith("/connect/oauth") && request.method === "POST")).toHaveLength(1)
      expect(state.requests.filter(request => request.method !== "GET")).toHaveLength(1)
      expect(state.requests.filter(request => (request.path === "/auth/reuse" && request.method === "POST") || request.path.includes("probe") || request.path.includes("provider-test"))).toHaveLength(0)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await screenshot(page, "19-environment-native-oauth-pending-narrow")
      await page.setViewportSize({ width: 1440, height: 1000 })
      await screenshot(page, "20-environment-native-oauth-pending")
      await writeFile(path.join(evidence!, "native-sign-in-requests.json"), JSON.stringify(state.requests, null, 2))
    } finally { await context.close() }
  })
  await scenario("consent, native OAuth, atomic models, Research save and failed-save retry", async () => {
    const state = fixture()
    const { page, context, errors } = await open(state)
    try {
      await expect(page.getByRole("heading", { name: "Before you begin" })).toBeVisible()
      expect(state.requests).toHaveLength(0)
      await screenshot(page, "01-launch-consent")
      await accept(page)
      await connection(page, "test-models")
      await page.getByLabel(/^Connection method/).selectOption("0")
      await page.getByLabel(/^account/).selectOption("workspace")
      await page.getByLabel("Workspace ID", { exact: true }).fill("fixture-workspace")
      await page.getByRole("button", { name: "Connect", exact: true }).click()
      await expect(page.getByText("Sign-in pending.", { exact: false })).toBeVisible()
      await page.getByRole("button", { name: "Check status", exact: true }).click()
      const start = state.requests.find(request => request.path.endsWith("/connect/oauth"))!
      expect(start.body).toMatchObject({ methodID: "workspace-oauth", answer: { account: "workspace", tenant: "fixture-workspace" } })
      await page.getByLabel("Authorization code").fill("fixture-code")
      await page.getByRole("button", { name: "Complete sign-in", exact: true }).click()
      await page.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(section(page, "Primary model")).toBeVisible()
      await section(page, "Primary model").locator("select").selectOption("test-models/primary")
      await section(page, "Fast model").locator("select").selectOption("")
      state.policyFailures = 1
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(page.getByRole("alert")).toContainText("That model is excluded by installation policy.")
      state.readbackFailures = 1
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(page.getByRole("alert")).toContainText("not saved")
      expect(state.completed).toEqual([])
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(page.getByRole("heading", { name: "Set up legal research", exact: true })).toBeVisible()
      const modelWrites = state.requests.filter(request => request.path === "/engine/config" && request.method === "PATCH")
      expect(modelWrites.map(request => request.body)).toEqual(Array(3).fill({ model: "test-models/primary", small_model: "test-models/primary" }))
      expect(state.requests.filter(request => request.path === "/updates")).toHaveLength(0)
      expect(state.requests.filter(request => request.path === "/api/dochaus/provider-test")).toHaveLength(0)
      await page.getByLabel("Exa API key", { exact: true }).fill("unsaved-fixture-key")
      await expect(page.getByRole("button", { name: "Skip research setup", exact: true })).toHaveCount(0)
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await currentStage(page, "Details")
      expect(state.preferences.searchApiKey).toBe("unsaved-fixture-key")
      expect(state.requests.filter(request => request.path === "/preferences" && request.method === "PUT")).toHaveLength(1)
      await page.getByPlaceholder("Jane Doe").fill("Entered Lawyer")
      await page.getByPlaceholder("Doe & Partners LLP").fill("Entered Firm")
      state.detailsFailures = 1
      await page.getByRole("button", { name: "Save and finish", exact: true }).click()
      await expect(page.getByRole("alert")).toContainText("Could not save drafting details")
      await expect(page.getByPlaceholder("Jane Doe")).toHaveValue("Entered Lawyer")
      await expect(page.getByPlaceholder("Doe & Partners LLP")).toHaveValue("Entered Firm")
      await screenshot(page, "02-details-save-retry")
      await page.getByRole("button", { name: "Save and finish", exact: true }).click()
      await expect(page.getByRole("dialog", { name: "Let's get you set up" })).toHaveCount(0)
      await expect.poll(() => state.requests.filter(request => request.path === "/updates").length).toBeGreaterThan(0)
      await expect.poll(() => state.requests.filter(request => request.path === "/api/dochaus/provider-test").map(request => request.body)).toEqual([{ providerID: "test-models", modelID: "primary" }])
      expect(state.preferences).toMatchObject({ attorney: "Entered Lawyer", firm: "Entered Firm", searchApiKey: "unsaved-fixture-key" })
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await screenshot(page, "03-setup-completed")
    } finally { await context.close() }
  })
  await scenario("a failed step save clears on a successful retry while a drafting details load error stays", async () => {
    const state = fixture(); state.connected = true; state.ready = true; state.config = { model: "test-models/primary", small_model: "test-models/primary" }; state.completed = ["models"]
    state.jurisdictionsUnavailable = true; state.setupFailures = 1
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await currentStage(page, "Research")
      const loadError = page.getByRole("alert").filter({ hasText: "Could not load drafting details. Retry setup." })
      const saveError = page.getByRole("alert").filter({ hasText: "Research protection is still loading. Retry in a moment." })
      await expect(loadError).toBeVisible()
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(saveError).toBeVisible()
      await currentStage(page, "Research")
      expect(state.completed).toEqual(["models"])
      await page.getByRole("button", { name: "Save and continue", exact: true }).click()
      await currentStage(page, "Details")
      await expect(saveError).toHaveCount(0)
      await expect(loadError).toBeVisible()
      expect(state.completed).toEqual(["models", "research"])
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await screenshot(page, "25-step-save-retry")
    } finally { await context.close() }
  })
  await scenario("later, reload and normal Settings resume without overwriting optional preferences", async () => {
    const state = fixture(); state.connected = true; state.ready = true; state.config = { model: "test-models/primary", small_model: "test-models/primary" }; state.completed = ["models"]; state.skipped = ["research"]
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await expect(page.getByPlaceholder("Jane Doe")).toHaveValue("Saved Lawyer")
      await page.getByRole("button", { name: "Set up later", exact: true }).click()
      await expect(page.getByRole("button", { name: "Resume setup", exact: true })).toBeVisible()
      await page.goto(`${baseURL}/matter/fixture`)
      await expect(page.getByText("Finish setup before opening assistant work.", { exact: false })).toBeVisible()
      await page.reload()
      await expect(page.getByPlaceholder("Jane Doe")).toHaveValue("Saved Lawyer")
      await expect(page.getByRole("button", { name: "Advanced settings", exact: true })).toHaveCount(0)
      await page.getByRole("button", { name: "Set up later", exact: true }).click()
      await page.getByRole("button", { name: "Settings", exact: true }).click()
      await expect(page.getByRole("tab", { name: "Providers", exact: true })).toBeVisible()
      await page.getByRole("button", { name: "Close", exact: true }).click()
      await page.getByRole("button", { name: "Resume setup", exact: true }).click()
      await expect(page.getByPlaceholder("Jane Doe")).toHaveValue("Saved Lawyer")
      expect(state.requests.filter(request => request.path === "/preferences" && request.method === "PUT")).toHaveLength(0)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await screenshot(page, "04-resumed-details")
    } finally { await context.close() }
  })
  await scenario("command attempt resumes across reload without rerunning; environment method is explicit", async () => {
    const state = fixture(); const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await connection(page, "test-models")
      await page.getByLabel(/^Connection method/).selectOption("2")
      await page.getByRole("button", { name: "Run sign-in command", exact: true }).click()
      await expect(page.getByRole("button", { name: "Check status", exact: true })).toBeVisible()
      await page.reload()
      await connection(page, "test-models")
      await expect(page.getByRole("button", { name: "Check status", exact: true })).toBeVisible()
      expect(state.requests.filter(request => request.path.endsWith("/connect/command") && request.method === "POST")).toHaveLength(1)
      await page.getByRole("button", { name: "Check status", exact: true }).click()
      await expect(page.getByText("Continue the fixture host sign-in in its terminal window.", { exact: true })).toBeVisible()
      expect(state.requests.filter(request => request.method === "POST")).toHaveLength(1)
      await page.getByRole("button", { name: "Cancel sign-in", exact: true }).click()
      await page.getByText("Environment setup", { exact: true }).click()
      await expect(page.getByText("Set FIXTURE_MODEL_KEY", { exact: false })).toBeVisible()
      expect(state.requests.filter(request => request.method === "POST")).toHaveLength(1)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await screenshot(page, "05-environment-sign-in")
    } finally { await context.close() }
  })
  await scenario("saved unavailable models remain visible and require an explicit available replacement", async () => {
    const state = fixture(); state.connected = true; state.secondProvider = true
    const { page, context, errors } = await open(state)
    try {
      await accept(page)
      await expect(page.getByLabel(/^Model provider/)).toHaveValue("")
      await expect(page.getByRole("button", { name: "Continue to connection", exact: true })).toBeDisabled()
      await expect(section(page, "Primary model")).toHaveCount(0)
      expect(state.requests.filter(request => request.path === "/auth/reuse")).toHaveLength(0)
      state.config = { model: "test-models/retired", small_model: "test-models/retired" }
      await page.reload()
      await expect(section(page, "Primary model").locator("select")).toHaveValue("test-models/retired")
      await expect(section(page, "Primary model").locator("option:checked")).toContainText("unavailable")
      await expect(page.getByRole("button", { name: "Save and continue", exact: true })).toBeDisabled()
      await section(page, "Primary model").locator("select").selectOption("test-models/primary")
      await expect(page.getByRole("button", { name: "Save and continue", exact: true })).toBeDisabled()
      await section(page, "Fast model").locator("select").selectOption("")
      await expect(page.getByRole("button", { name: "Save and continue", exact: true })).toBeEnabled()
      expect(state.requests.filter(request => request.method === "PATCH")).toHaveLength(0)
      expect(state.unexpected).toEqual([]); expect(errors).toEqual([])
      await screenshot(page, "06-unavailable-model-replacement")
    } finally { await context.close() }
  })
} finally {
  await writeFile(path.join(evidence, "release-setup-results.json"), JSON.stringify(results, null, 2) + "\n")
  await browser.close()
}
