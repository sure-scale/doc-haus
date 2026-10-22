// Synthetic browser acceptance. All backend traffic is intercepted; no paid
// provider requests or real credentials are used.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import type { Locator, Page } from "@playwright/test"
import type { CandidateInput } from "../src/api/opencode/provider-candidate"
import type { Config } from "../src/api/opencode/config"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
const appOrigin = new URL(baseURL).origin
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
type RecordedRequest = { method: string; path: string; body: Record<string, unknown> }
const results: { name: string; status: string; error?: string }[] = []
const allRequests: { name: string; requests: RecordedRequest[] }[] = []
const errors: string[] = []
const unexpected: string[] = []
const scenarioPattern = process.env.DOCHAUS_E2E_SCENARIOS ? new RegExp(process.env.DOCHAUS_E2E_SCENARIOS) : undefined

async function fixture(name: string, width: number, onboarding = false, cloud = false, features: { reuse?: boolean } = {}) {
  const context = await browser.newContext({ viewport: { width, height: 1000 }, reducedMotion: "reduce" })
  await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  page.on("pageerror", error => errors.push(`${name}: ${error.message}`))
  const requests: RecordedRequest[] = []
  allRequests.push({ name, requests })
  const candidates = new Map<string, CandidateInput>()
  const state = {
    config: (cloud
      ? { model: "google-vertex/primary", small_model: "google-vertex/primary", provider: { "google-vertex": { options: { project: "saved-fixture-project", location: "europe-west2" } } } }
      : onboarding ? {} : { model: "fixture-models/primary", small_model: "fixture-models/primary" }) as Config,
    completed: onboarding ? [] as string[] : ["models", "research", "details"],
    ready: !onboarding,
    failFast: false,
    saveFailures: 0,
    progressFailures: 0,
    modelLoadFailures: 0,
    catalogLag: 0,
    oauthComplete: false,
    commandStatus: "pending" as "pending" | "complete",
    preferences: { ...DEFAULT_DRAFTING, attorney: "Synthetic Lawyer", firm: "Synthetic Firm" },
  }
  const progress = () => ({ version: 1, ready: state.ready, completed: state.completed, skipped: [], next: ["models", "research", "details"].find(step => !state.completed.includes(step)) ?? "done" })
  const provider = { id: "fixture-models", integrationID: "fixture-login", name: "Detected Fixture Models", activation: "enabled" }
  const other = { id: "other-fixture", integrationID: "other-login", name: "Other Fixture Models", activation: "auto" }
  const vertex = { id: "google-vertex", integrationID: "google-vertex", name: "Google Vertex AI", activation: "enabled", host: { type: "google-adc", options: { project: "detected-fixture-project", location: "us-central1" }, sources: ["Google ADC file"] } }
  const savedLocalProviders = () => Object.entries(state.config.provider ?? {}).filter(([, entry]) => entry.npm === "@ai-sdk/openai-compatible")
    .map(([id, entry]) => ({ id, name: entry.name ?? id, activation: "enabled" }))
  await context.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || pathname.startsWith("/api/") || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if (!api && url.origin === appOrigin) return route.continue()
    if (!api) { unexpected.push(`${name}: ${request.url()}`); return route.abort() }
    const headers = { "access-control-allow-origin": appOrigin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (request.method() === "OPTIONS") return route.fulfill({ status: 204, headers })
    const method = request.method()
    const body = request.postData() ? request.postDataJSON() as Record<string, unknown> : {}
    requests.push({ method, path: pathname, body })
    const reply = (data: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(data) })
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "POST") {
      const sessionID = `synthetic-candidate-${candidates.size + 1}`
      candidates.set(sessionID, body as CandidateInput)
      return reply({ sessionID })
    }
    if (pathname === "/api/dochaus/provider-candidate/session" && method === "DELETE") return route.fulfill({ status: 204, headers })
    if (pathname === "/api/dochaus/provider-candidate/models" && method === "POST") {
      if (state.modelLoadFailures > 0) { state.modelLoadFailures--; return reply({ error: "Synthetic model catalog is temporarily unavailable." }, 503) }
      const candidate = candidates.get(String(body.sessionID))
      expect(candidate).toBeDefined()
      if (candidate!.local || state.config.provider?.[candidate!.providerID]?.npm === "@ai-sdk/openai-compatible") return reply({ data: [], customModels: true })
      return reply({ data: ["primary", "fast"].map(id => ({ providerID: candidate!.providerID, id, name: id === "primary" ? "Primary Fixture" : "Fast Fixture" })), customModels: false })
    }
    if (pathname === "/api/dochaus/provider-candidate/test" && method === "POST") {
      expect(candidates.has(String(body.sessionID))).toBe(true)
      return state.failFast && body.modelID === "fast" ? reply({ ok: false, kind: "model", error: "Synthetic fast model access is unavailable." }, 400) : reply({ ok: true })
    }
    if (pathname === "/api/dochaus/provider-candidate/auth" && method === "POST") return reply({ data: String(body.action).startsWith("command.")
      ? { attemptID: "synthetic-command", status: body.action === "command.cancel" ? "cancelled" : state.commandStatus, message: "Complete the synthetic host sign-in, then check its status." }
      : { attemptID: "synthetic-oauth", mode: "auto", status: state.oauthComplete ? "complete" : "pending", instructions: "Complete the synthetic browser sign-in, then check its status." },
    })
    if (pathname === "/engine/provider-setup" && method === "POST") {
      if (state.saveFailures > 0) { state.saveFailures--; return reply({ error: "Synthetic save was rejected before activation.", code: "fixture_save_rejected" }, 400) }
      const candidate = candidates.get(String(body.sessionID))
      expect(candidate).toBeDefined()
      state.config = { ...state.config, model: `${candidate!.providerID}/${body.primary}`, small_model: `${candidate!.providerID}/${body.fast}` }
      const existing = state.config.provider?.[candidate!.providerID]
      if (candidate!.local || existing?.npm === "@ai-sdk/openai-compatible") state.config.provider = {
        ...state.config.provider,
        [candidate!.providerID]: {
          ...existing,
          npm: "@ai-sdk/openai-compatible",
          name: candidate!.local?.name ?? existing?.name,
          options: { ...existing?.options, ...candidate!.options },
          models: { ...existing?.models, ...Object.fromEntries([...new Set([String(body.primary), String(body.fast)])].map(id => [id, { name: id }])) },
        },
      }
      state.ready = true
      return reply({ ok: true })
    }
    if (pathname === "/setup") {
      if (method === "PATCH" && state.progressFailures > 0) { state.progressFailures--; return reply({ error: "Synthetic setup progress is temporarily unavailable." }, 503) }
      if (method === "PATCH") state.completed = [...new Set([...state.completed, String(body.step)])]
      return reply(progress())
    }
    if (pathname === "/preferences") {
      if (method === "PUT") state.preferences = { ...state.preferences, ...body }
      return reply(state.preferences)
    }
    if (method !== "GET" && !(pathname === "/api/dochaus/provider-test" && method === "POST")) {
      unexpected.push(`${name}: ${method} ${pathname}`)
      return route.abort()
    }
    if (pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/engine/config") return reply(state.config)
    if (pathname === "/auth/reuse") return reply({ sources: features.reuse ? [
      { id: "source-one", label: "First synthetic store", status: "ready", unavailableCount: 0, connections: [{ id: "account-one", integrationID: "other-login", label: "First fixture account", authType: "key" }] },
      { id: "source-two", label: "Second synthetic store", status: "ready", unavailableCount: 0, connections: [{ id: "account-two", integrationID: "other-login", label: "Second fixture account", authType: "oauth" }] },
    ] : [] })
    if (pathname === "/api/dochaus/provider-catalog") {
      const local = state.catalogLag > 0 ? [] : savedLocalProviders()
      if (state.catalogLag > 0) state.catalogLag--
      return reply({ data: [other, provider, ...(cloud ? [vertex] : []), ...local] })
    }
    if (pathname === "/api/provider") return reply({ data: [provider, ...(cloud ? [vertex] : []), ...savedLocalProviders()] })
    if (pathname === "/api/model") return reply({ data: [
      ...[provider, ...(cloud ? [vertex] : [])].flatMap(item => ["primary", "fast"].map(id => ({ providerID: item.id, id, name: id, enabled: true }))),
      ...savedLocalProviders().flatMap(item => Object.entries(state.config.provider?.[item.id]?.models ?? {}).map(([id, model]) => ({ providerID: item.id, id, name: model.name, enabled: true }))),
    ] })
    if (pathname === "/api/integration") return reply({ data: [
      { id: "fixture-login", name: "Fixture Sign-in", methods: [{ type: "key", label: "New API key" }, { type: "oauth", id: "browser-login", label: "Browser sign-in" }, { type: "command", id: "host-login", label: "Host sign-in", command: ["fixture-auth", "login"] }], connections: [{ type: "credential", id: "saved-fixture", label: "Fixture account", method: "key" }] },
      { id: "other-login", name: "Other sign-in", methods: [{ type: "key", label: "New API key" }], connections: [] },
      ...(cloud ? [{ id: "google-vertex", name: "Google Vertex AI", methods: [], connections: [] }] : []),
    ] })
    if (pathname === "/matters") return reply([])
    if (pathname === "/jurisdictions") return reply([{ code: "england-wales", name: "England and Wales" }])
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: "" })
    unexpected.push(`${name}: ${method} ${pathname}`)
    return route.abort()
  })
  await page.goto(baseURL!)
  return { page, context, state, requests, tests: () => requests.filter(item => item.path === "/api/dochaus/provider-candidate/test"), sessions: () => requests.filter(item => item.path === "/api/dochaus/provider-candidate/session" && item.method === "POST") }
}

async function scenario(name: string, run: () => Promise<void>) {
  if (scenarioPattern && !scenarioPattern.test(name)) return
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); throw error }
}
async function noOverflow(page: Page) {
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true)
  const panel = page.locator(".settings-content:visible, .onboarding-body:visible").first()
  expect(await panel.evaluate(node => node.scrollWidth <= node.clientWidth)).toBe(true)
}
async function hitTarget(control: Locator, minimum = 40) {
  await control.scrollIntoViewIfNeeded()
  const rect = await control.evaluate(node => {
    const box = node.getBoundingClientRect()
    return { width: box.width, height: box.height, uncovered: node.contains(document.elementFromPoint(box.x + box.width / 2, box.y + box.height / 2)) }
  })
  expect(rect.uncovered).toBe(true)
  expect(rect.width).toBeGreaterThanOrEqual(minimum)
  expect(rect.height).toBeGreaterThanOrEqual(minimum)
}

try {
  for (const width of [1440, 390]) await scenario(`Settings prepares, tests and saves exact models at ${width}px`, async () => {
    const f = await fixture(`settings-${width}`, width)
    try {
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      const dialog = f.page.getByRole("dialog", { name: "Settings", exact: true })
      await expect(dialog.getByRole("tab", { name: "AI setup", exact: true })).toHaveAttribute("aria-selected", "true")
      const provider = dialog.getByLabel("Provider", { exact: true })
      await expect(provider).toHaveValue("fixture-models")
      expect(await provider.locator("optgroup").evaluateAll(groups => groups.map(group => ({ label: group.getAttribute("label"), values: Array.from(group.querySelectorAll("option"), option => option.value) })))).toEqual([
        { label: "Detected", values: ["fixture-models"] },
        { label: "Other providers", values: ["other-fixture", "__local_endpoint__"] },
      ])
      await expect(dialog.getByLabel("Primary model", { exact: true })).toHaveValue("fixture-models/primary")
      await expect(dialog.getByLabel("Fast model", { exact: true })).toHaveValue("")
      await provider.scrollIntoViewIfNeeded()
      await f.page.screenshot({ path: path.join(evidence, `settings-provider-${width}.png`) })
      expect(f.sessions()[0]!.body.connection).toEqual({ type: "credential", id: "saved-fixture" })
      const save = dialog.getByRole("button", { name: "Save", exact: true })
      const test = dialog.getByRole("button", { name: "Test connection", exact: true })
      await expect(save).toBeDisabled()
      expect(f.tests()).toHaveLength(0)
      await test.click()
      await expect(save).toBeEnabled()
      expect(f.tests().map(item => item.body.modelID)).toEqual(["primary"])
      await expect(dialog.locator(".ai-setup-result--ready")).toHaveCount(2)
      await expect(dialog.locator(".ai-setup-result--ready").last()).toContainText("Uses the primary model test.")

      await dialog.getByLabel("Fast model", { exact: true }).selectOption("fixture-models/fast")
      await expect(save).toBeDisabled()
      await expect(dialog.locator(".ai-setup-result")).toHaveCount(0)
      f.state.failFast = true
      await test.click()
      await expect(dialog.locator(".ai-setup-result--model")).toContainText("Synthetic fast model access is unavailable.")
      await expect(dialog.locator(".ai-setup-result--ready")).toContainText("fixture-models/primary")
      await expect(save).toBeDisabled()
      expect(f.tests().slice(-2).map(item => item.body.modelID)).toEqual(["primary", "fast"])
      await dialog.locator(".ai-setup-result--model").scrollIntoViewIfNeeded()
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, `settings-model-failure-${width}.png`) })
      f.state.failFast = false
      await test.click()
      await expect(save).toBeEnabled()
      await expect(dialog.locator(".ai-setup-result--ready")).toHaveCount(2)
      expect(f.tests().slice(-2).map(item => item.body.modelID)).toEqual(["primary", "fast"])

      await dialog.getByLabel("Connection method", { exact: true }).selectOption("method:0")
      await expect(save).toBeDisabled()
      await expect(dialog.locator(".ai-setup-result")).toHaveCount(0)
      await dialog.getByLabel("API key", { exact: true }).fill("  synthetic-replacement-key  ")
      expect(f.state.config).toEqual({ model: "fixture-models/primary", small_model: "fixture-models/primary" })
      await dialog.getByRole("button", { name: "Load models", exact: true }).click()
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      expect(f.sessions().at(-1)!.body.connection).toEqual({ type: "key", key: "synthetic-replacement-key", answer: {} })
      await expect(save).toBeDisabled()
      await test.click()
      await expect(save).toBeEnabled()
      const testsBeforeSave = f.tests().length
      const healthBeforeSave = f.requests.filter(item => item.path === "/api/dochaus/provider-test").length
      await hitTarget(provider)
      await hitTarget(dialog.getByLabel("Primary model", { exact: true }))
      await hitTarget(dialog.getByLabel("Fast model", { exact: true }))
      await hitTarget(test)
      await hitTarget(save)
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, `settings-tested-${width}.png`) })
      f.state.saveFailures = 1
      await save.click()
      await expect(dialog.getByRole("alert")).toContainText("Synthetic save was rejected before activation.")
      await expect(save).toBeEnabled()
      expect(f.state.config).toEqual({ model: "fixture-models/primary", small_model: "fixture-models/primary" })
      await save.click()
      await expect(dialog.getByText("AI setup saved.", { exact: true })).toBeVisible()
      expect(f.state.config).toEqual({ model: "fixture-models/primary", small_model: "fixture-models/fast" })
      expect(f.requests.filter(item => item.path === "/engine/provider-setup").at(-1)!.body).toEqual({ sessionID: "synthetic-candidate-2", primary: "primary", fast: "fast" })
      expect(f.tests()).toHaveLength(testsBeforeSave)
      expect(f.requests.filter(item => item.path === "/api/dochaus/provider-test")).toHaveLength(healthBeforeSave)
      await f.page.screenshot({ path: path.join(evidence, `settings-saved-${width}.png`) })
    } finally { await f.context.close() }
  })

  for (const width of [1440, 390]) await scenario(`Onboarding gates and completes all six stages at ${width}px`, async () => {
    const f = await fixture(`onboarding-${width}`, width, true)
    try {
      const dialog = f.page.getByRole("dialog", { name: "Let's get you set up", exact: true })
      const navigation = dialog.getByRole("navigation", { name: "Setup progress", exact: true })
      await expect(dialog.getByRole("heading", { name: "Choose your model provider", exact: true })).toBeVisible()
      await expect(navigation.getByRole("button")).toHaveCount(6)
      for (const stage of ["Connection", "Models", "Test", "Research", "Details"]) await expect(navigation.getByRole("button", { name: new RegExp(`${stage}$`) })).toBeDisabled()
      await expect(dialog.getByRole("button", { name: "Continue to connection", exact: true })).toBeDisabled()
      expect(f.tests()).toHaveLength(0)
      await dialog.getByLabel("Provider", { exact: true }).selectOption("fixture-models")
      await dialog.getByRole("button", { name: "Continue to connection", exact: true }).click()
      await expect(dialog.getByRole("heading", { name: "Connect your provider", exact: true })).toBeVisible()
      await expect(navigation.getByRole("button", { name: /Models$/ })).toBeDisabled()
      await expect(dialog.getByLabel("Connection method", { exact: true })).toHaveValue("credential:saved-fixture")
      await dialog.getByRole("button", { name: "Continue to models", exact: true }).click()
      await expect(dialog.getByRole("heading", { name: "Choose your models", exact: true })).toBeVisible()
      await expect(dialog.getByRole("button", { name: "Continue to test", exact: true })).toBeDisabled()
      await expect(navigation.getByRole("button", { name: /Research$/ })).toBeDisabled()
      await dialog.getByLabel("Primary model", { exact: true }).selectOption("fixture-models/primary")
      await expect(dialog.getByLabel("Fast model", { exact: true })).toHaveValue("")
      await dialog.getByRole("button", { name: "Continue to test", exact: true }).click()
      await expect(dialog.getByRole("heading", { name: "Test your connection", exact: true })).toBeVisible()
      const save = dialog.getByRole("button", { name: "Save and continue", exact: true })
      await expect(save).toBeDisabled()
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, `onboarding-before-test-${width}.png`) })
      await dialog.getByRole("button", { name: "Test connection", exact: true }).click()
      await expect(save).toBeEnabled()
      expect(f.tests().map(item => item.body.modelID)).toEqual(["primary"])
      await expect(navigation.getByRole("button", { name: /Research$/ })).toBeDisabled()
      await hitTarget(save)
      await f.page.screenshot({ path: path.join(evidence, `onboarding-tested-${width}.png`) })
      f.state.progressFailures = 1
      await save.click()
      await expect(dialog.getByRole("alert")).toContainText("AI setup was saved. Setup progress could not be updated.")
      await expect(dialog.getByRole("button", { name: "Save and continue", exact: true })).toBeDisabled()
      await expect(navigation.getByRole("button", { name: /Research$/ })).toBeDisabled()
      await dialog.getByRole("button", { name: "Continue", exact: true }).click()
      await expect(dialog.getByRole("heading", { name: "Set up legal research", exact: true })).toBeVisible()
      expect(f.requests.filter(item => item.path === "/engine/provider-setup")).toHaveLength(1)
      expect(f.state.config).toEqual({ model: "fixture-models/primary", small_model: "fixture-models/primary" })
      expect(f.state.completed).toContain("models")
      await dialog.getByRole("button", { name: "Save and continue", exact: true }).click()
      await expect(dialog.getByRole("heading", { name: "Make doc.haus yours", exact: true })).toBeVisible()
      await dialog.getByLabel("Attorney", { exact: true }).fill("Fixture Counsel")
      await dialog.getByLabel("Firm", { exact: true }).fill("Fixture Practice")
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, `onboarding-details-${width}.png`) })
      await dialog.getByRole("button", { name: /^Save and finish/ }).click()
      await expect(dialog).toHaveCount(0)
      expect(f.state.completed).toEqual(["models", "research", "details"])
      expect(f.state.preferences).toMatchObject({ attorney: "Fixture Counsel", firm: "Fixture Practice" })
    } finally { await f.context.close() }
  })

  await scenario("Saved cloud metadata takes precedence over host detection and edits require a new test", async () => {
    const f = await fixture("cloud-prefill", 1440, false, true)
    try {
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      const dialog = f.page.getByRole("dialog", { name: "Settings", exact: true })
      await expect(dialog.getByLabel("Provider", { exact: true })).toHaveValue("google-vertex")
      await expect(dialog.getByLabel("GCP project id", { exact: true })).toHaveValue("saved-fixture-project")
      await expect(dialog.getByLabel("Location", { exact: true })).toHaveValue("europe-west2")
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      expect(f.sessions()[0]!.body).toEqual({ providerID: "google-vertex", options: { project: "saved-fixture-project", location: "europe-west2" }, connection: { type: "host" } })
      await dialog.getByRole("button", { name: "Test connection", exact: true }).click()
      await expect(dialog.getByRole("button", { name: "Save", exact: true })).toBeEnabled()
      await dialog.getByLabel("GCP project id", { exact: true }).fill(" updated-fixture-project ")
      await expect(dialog.getByRole("button", { name: "Save", exact: true })).toBeDisabled()
      await expect(dialog.locator(".ai-setup-result")).toHaveCount(0)
      await dialog.getByRole("button", { name: "Load models", exact: true }).click()
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      expect(f.sessions().at(-1)!.body.options).toEqual({ project: "updated-fixture-project", location: "europe-west2" })
      expect(f.state.config.provider?.["google-vertex"]?.options?.project).toBe("saved-fixture-project")
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, "settings-cloud-prefill.png") })
    } finally { await f.context.close() }
  })

  await scenario("Pending browser sign-in cannot load models or enable testing before completion", async () => {
    const f = await fixture("pending-oauth", 390)
    try {
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      const dialog = f.page.getByRole("dialog", { name: "Settings", exact: true })
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      await dialog.getByLabel("Connection method", { exact: true }).selectOption("method:1")
      await dialog.getByRole("button", { name: "Sign in", exact: true }).click()
      await expect(dialog.getByText("Sign-in pending.", { exact: true })).toBeVisible()
      await expect(dialog.getByLabel("Primary model", { exact: true })).toHaveCount(0)
      await expect(dialog.getByRole("button", { name: "Test connection", exact: true })).toBeDisabled()
      await expect(dialog.getByRole("button", { name: "Save", exact: true })).toBeDisabled()
      await dialog.getByRole("button", { name: "Check sign-in status", exact: true }).click()
      await expect(dialog.getByText("Sign-in pending.", { exact: true })).toBeVisible()
      expect(f.tests()).toHaveLength(0)
      const sessionsBeforeReload = f.sessions().length
      await f.page.screenshot({ path: path.join(evidence, "settings-oauth-pending-390.png") })
      await f.page.reload()
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      await expect(dialog.getByText("Sign-in pending.", { exact: true })).toBeVisible()
      expect(f.sessions()).toHaveLength(sessionsBeforeReload)
      expect(f.requests.filter(item => item.path === "/api/dochaus/provider-candidate/auth" && item.body.action === "oauth.connect")).toHaveLength(1)
      const pending = await f.page.evaluate(() => Object.entries(sessionStorage).filter(([key]) => key.startsWith("dochaus.setup-sign-in:")).map(([, value]) => JSON.parse(value) as Record<string, unknown>))
      expect(pending).toHaveLength(1)
      expect(Object.keys(pending[0]!).sort()).toEqual(["attempt", "methodID", "options", "providerID", "sessionID"])
      f.state.oauthComplete = true
      f.state.modelLoadFailures = 1
      await dialog.getByRole("button", { name: "Check sign-in status", exact: true }).click()
      await expect(dialog.getByRole("alert")).toContainText("Synthetic model catalog is temporarily unavailable.")
      await expect(dialog.getByLabel("Primary model", { exact: true })).toHaveCount(0)
      const authCalls = f.requests.filter(item => item.path === "/api/dochaus/provider-candidate/auth").length
      await dialog.getByRole("button", { name: "Retry loading models", exact: true }).click()
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      expect(f.sessions()).toHaveLength(sessionsBeforeReload)
      expect(f.requests.filter(item => item.path === "/api/dochaus/provider-candidate/auth")).toHaveLength(authCalls)
      await expect(dialog.getByRole("button", { name: "Test connection", exact: true })).toBeEnabled()
      await expect(dialog.getByRole("button", { name: "Save", exact: true })).toBeDisabled()
      expect(f.sessions().at(-1)!.body.connection).toEqual({ type: "new" })
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, "settings-oauth-complete-390.png") })
    } finally { await f.context.close() }
  })

  await scenario("Local endpoints save without keys and reopen for endpoint and model changes", async () => {
    const f = await fixture("local-endpoint", 390)
    try {
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      const dialog = f.page.getByRole("dialog", { name: "Settings", exact: true })
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      await dialog.getByLabel("Provider", { exact: true }).selectOption("__local_endpoint__")
      await dialog.getByLabel("Connection name", { exact: true }).fill("Fixture local")
      await dialog.getByLabel("Endpoint", { exact: true }).fill("http://127.0.0.1:1234/v1")
      await expect(dialog.getByLabel("API key", { exact: true })).toHaveValue("")
      await dialog.getByRole("button", { name: "Load models", exact: true }).click()
      await expect(dialog.getByLabel("Primary model ID", { exact: true })).toBeVisible()
      expect(f.sessions().at(-1)!.body).toEqual({ providerID: "fixture-local", options: { baseURL: "http://127.0.0.1:1234/v1" }, connection: { type: "host" }, local: { name: "Fixture local", baseURL: "http://127.0.0.1:1234/v1", modelIDs: [] } })
      await dialog.getByLabel("Primary model ID", { exact: true }).fill("served-primary")
      const test = dialog.getByRole("button", { name: "Test connection", exact: true })
      const save = dialog.getByRole("button", { name: "Save", exact: true })
      await expect(save).toBeDisabled()
      await test.click()
      await expect(save).toBeEnabled()
      expect(f.requests.filter(item => item.path === "/api/dochaus/provider-candidate/models" && item.body.modelIDs).at(-1)!.body.modelIDs).toEqual(["served-primary"])
      expect(f.tests().map(item => item.body.modelID)).toEqual(["served-primary"])
      f.state.catalogLag = 3
      await save.click()
      await expect(dialog.getByText("AI setup saved.", { exact: true })).toBeVisible()
      expect(f.state.config).toEqual({ model: "fixture-local/served-primary", small_model: "fixture-local/served-primary", provider: {
        "fixture-local": { npm: "@ai-sdk/openai-compatible", name: "Fixture local", options: { baseURL: "http://127.0.0.1:1234/v1" }, models: { "served-primary": { name: "served-primary" } } },
      } })
      await expect(dialog.getByLabel("Provider", { exact: true })).toHaveValue("fixture-local")
      await expect(dialog.getByLabel("Primary model ID", { exact: true })).toHaveValue("served-primary")
      expect(f.state.catalogLag).toBe(0)
      expect(f.tests().map(item => item.body.modelID)).toEqual(["served-primary"])
      expect(f.requests.filter(item => item.path === "/engine/provider-setup")).toHaveLength(1)
      await dialog.getByRole("button", { name: "Close", exact: true }).click()
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      await expect(dialog.getByLabel("Provider", { exact: true })).toHaveValue("fixture-local")
      await expect(dialog.getByLabel("Endpoint", { exact: true })).toHaveValue("http://127.0.0.1:1234/v1")
      await expect(dialog.getByLabel("Connection method", { exact: true })).toHaveValue("host")
      await expect(dialog.getByLabel("Primary model ID", { exact: true })).toHaveValue("served-primary")
      expect(f.sessions().at(-1)!.body).toEqual({ providerID: "fixture-local", options: { baseURL: "http://127.0.0.1:1234/v1" }, connection: { type: "host" } })
      await test.click()
      await expect(save).toBeEnabled()
      await dialog.getByLabel("Endpoint", { exact: true }).fill("http://127.0.0.1:1235/v1")
      await expect(save).toBeDisabled()
      await expect(dialog.locator(".ai-setup-result")).toHaveCount(0)
      await dialog.getByLabel("Connection method", { exact: true }).selectOption("method:0")
      await dialog.getByLabel("API key", { exact: true }).fill("synthetic-local-key")
      await expect(save).toBeDisabled()
      await dialog.getByRole("button", { name: "Load models", exact: true }).click()
      await expect(dialog.getByLabel("Primary model ID", { exact: true })).toBeVisible()
      expect(f.sessions().at(-1)!.body).toEqual({ providerID: "fixture-local", options: { baseURL: "http://127.0.0.1:1235/v1" }, connection: { type: "key", key: "synthetic-local-key", answer: {} } })
      await dialog.getByLabel("Fast model ID", { exact: true }).fill("served-fast")
      await test.click()
      await expect(save).toBeEnabled()
      expect(f.requests.filter(item => item.path === "/api/dochaus/provider-candidate/models" && item.body.modelIDs).at(-1)!.body.modelIDs).toEqual(["served-primary", "served-fast"])
      expect(f.tests().slice(-2).map(item => item.body.modelID)).toEqual(["served-primary", "served-fast"])
      expect(f.state.config.provider?.["fixture-local"]?.options?.baseURL).toBe("http://127.0.0.1:1234/v1")
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, "settings-local-tested-390.png") })
      await save.click()
      await expect(dialog.getByText("AI setup saved.", { exact: true })).toBeVisible()
      expect(f.state.config).toEqual({ model: "fixture-local/served-primary", small_model: "fixture-local/served-fast", provider: {
        "fixture-local": { npm: "@ai-sdk/openai-compatible", name: "Fixture local", options: { baseURL: "http://127.0.0.1:1235/v1" }, models: { "served-primary": { name: "served-primary" }, "served-fast": { name: "served-fast" } } },
      } })
    } finally { await f.context.close() }
  })

  await scenario("Reusable sign-ins promote their provider and only the selected source becomes a candidate", async () => {
    const f = await fixture("selected-reuse", 1440, false, false, { reuse: true })
    try {
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      const dialog = f.page.getByRole("dialog", { name: "Settings", exact: true })
      const provider = dialog.getByLabel("Provider", { exact: true })
      await expect(provider).toHaveValue("fixture-models")
      expect(await provider.locator('optgroup[label="Detected"] option').evaluateAll(options => options.map(option => (option as HTMLOptionElement).value))).toEqual(["other-fixture", "fixture-models"])
      await provider.selectOption("other-fixture")
      const connection = dialog.getByLabel("Connection method", { exact: true })
      await expect(connection.locator('option[value="reuse:source-one:account-one"]')).toHaveCount(1)
      await expect(connection.locator('option[value="reuse:source-two:account-two"]')).toHaveCount(1)
      await connection.selectOption("reuse:source-two:account-two")
      expect(f.sessions().filter(item => item.body.providerID === "other-fixture")).toHaveLength(0)
      await dialog.getByRole("button", { name: "Load models", exact: true }).click()
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      expect(f.sessions().at(-1)!.body).toEqual({ providerID: "other-fixture", options: {}, connection: { type: "reuse", sourceId: "source-two", connectionId: "account-two" } })
      expect(f.requests.filter(item => item.path === "/auth/reuse" && item.method === "POST")).toHaveLength(0)
      expect(f.requests.filter(item => item.path.startsWith("/api/integration/") && item.method === "POST")).toHaveLength(0)
      await dialog.getByLabel("Primary model", { exact: true }).selectOption("other-fixture/primary")
      await dialog.getByRole("button", { name: "Test connection", exact: true }).click()
      const save = dialog.getByRole("button", { name: "Save", exact: true })
      await expect(save).toBeEnabled()
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, "settings-selected-reuse.png") })
      await save.click()
      await expect(dialog.getByText("AI setup saved.", { exact: true })).toBeVisible()
      expect(f.state.config).toEqual({ model: "other-fixture/primary", small_model: "other-fixture/primary" })
    } finally { await f.context.close() }
  })

  await scenario("Command sign-in survives navigation and restarts only after explicit cancellation", async () => {
    const f = await fixture("command-sign-in", 1440)
    try {
      await f.page.getByRole("button", { name: "Settings", exact: true }).click()
      const dialog = f.page.getByRole("dialog", { name: "Settings", exact: true })
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      await dialog.getByLabel("Connection method", { exact: true }).selectOption("method:2")
      await expect(dialog.getByText("fixture-auth login", { exact: true })).toBeVisible()
      await dialog.getByRole("button", { name: "Run sign-in command", exact: true }).click()
      await expect(dialog.getByText("Sign-in pending.", { exact: true })).toBeVisible()
      const auth = (action: string) => f.requests.filter(item => item.path === "/api/dochaus/provider-candidate/auth" && item.body.action === action)
      expect(auth("command.connect")).toHaveLength(1)
      const sessionsBeforeNavigation = f.sessions().length
      await expect(dialog.getByRole("button", { name: "Test connection", exact: true })).toBeDisabled()
      await dialog.getByRole("tab", { name: "Drafting", exact: true }).click()
      await dialog.getByRole("tab", { name: "AI setup", exact: true }).click()
      await expect(dialog.getByText("Sign-in pending.", { exact: true })).toBeVisible()
      expect(f.sessions()).toHaveLength(sessionsBeforeNavigation)
      expect(auth("command.connect")).toHaveLength(1)
      await dialog.getByRole("button", { name: "Check sign-in status", exact: true }).click()
      await expect(dialog.getByText("Sign-in pending.", { exact: true })).toBeVisible()
      expect(auth("command.status")).toHaveLength(1)
      await dialog.getByRole("button", { name: "Cancel sign-in", exact: true }).click()
      await expect(dialog.getByRole("button", { name: "Run sign-in command", exact: true })).toBeVisible()
      expect(auth("command.cancel")).toHaveLength(1)
      expect(auth("command.connect")).toHaveLength(1)
      await dialog.getByRole("button", { name: "Run sign-in command", exact: true }).click()
      await expect(dialog.getByText("Sign-in pending.", { exact: true })).toBeVisible()
      expect(auth("command.connect")).toHaveLength(2)
      expect(f.sessions()).toHaveLength(sessionsBeforeNavigation + 1)
      f.state.commandStatus = "complete"
      await dialog.getByRole("button", { name: "Check sign-in status", exact: true }).click()
      await expect(dialog.getByLabel("Primary model", { exact: true })).toBeVisible()
      await expect(dialog.getByRole("button", { name: "Test connection", exact: true })).toBeEnabled()
      expect(auth("command.connect")).toHaveLength(2)
      expect(f.tests()).toHaveLength(0)
      await noOverflow(f.page)
      await f.page.screenshot({ path: path.join(evidence, "settings-command-complete.png") })
    } finally { await f.context.close() }
  })
  expect(unexpected).toEqual([])
  expect(errors).toEqual([])
} finally {
  await writeFile(path.join(evidence, "ai-setup-results.json"), JSON.stringify({ results, errors, unexpected, requests: allRequests }, null, 2) + "\n")
  await browser.close()
}
