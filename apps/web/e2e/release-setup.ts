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
async function screenshot(page: Awaited<ReturnType<typeof open>>["page"], name: string) { await page.screenshot({ path: path.join(evidence!, `${name}.png`), fullPage: true }) }
async function scenario(name: string, run: () => Promise<void>) {
  if (scenarioPattern && !scenarioPattern.test(name)) return
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}

const stageLabels = ["Provider", "Connection", "Models", "Test", "Research", "Details"]
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
      await stepButton(page, "Details").click(); await stepButton(page, "Research").click()
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
  // Provider/connection/model/test acceptance lives in ai-setup.ts. These
  // scenarios continue to cover the existing Research and Details consumers.

} finally {
  await writeFile(path.join(evidence, "release-setup-results.json"), JSON.stringify(results, null, 2) + "\n")
  await browser.close()
}
