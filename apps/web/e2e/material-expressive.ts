// Visual and interaction regression against an already-running web preview.
// Every API request is intercepted. Synthetic saved conversations are replayed;
// generation, credentials and unknown requests are refused before any network I/O.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
type Page = import("@playwright/test").Page
type Route = import("@playwright/test").Route
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
const baseline = process.env.DOCHAUS_BASELINE === "1"
const axeScript = process.env.AXE_SCRIPT
const origin = new URL(baseURL).origin
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: Record<string, unknown>[] = []
const screens: Record<string, unknown>[] = []
const allStates: ReturnType<typeof fixture>[] = []
const timestamp = Date.UTC(2026, 9, 21, 10)
const jurisdictions = [{ code: "EW", name: "England and Wales", citationStyle: "oscola" }, { code: "SC", name: "Scotland", citationStyle: "scotland" }]
const matter = { id: "fixture-matter", title: "Alder House · Services agreement", reference: "AH-2026-014", client: "Alder House Studio", jurisdictions: ["EW"], dir: "/fixture/matters/alder-house", created_at: timestamp }
const documents = [
  { id: 1, name: "Services agreement.docx", doc_path: "/fixture/matters/alder-house/Services agreement.docx", created_at: timestamp, role: "source" },
  { id: 2, name: "Client instructions.eml", doc_path: "/fixture/matters/alder-house/Client instructions.eml", created_at: timestamp, role: "source" },
  { id: 3, name: "Review notes.docx", doc_path: "/fixture/matters/alder-house/Review notes.docx", created_at: timestamp, role: "deliverable", label: "notes" },
]
const playbooks = [{ name: "playbook-commercial-services", description: "Preferred positions for services engagements.", content: "# Services agreements\n\n## Payment\nPreferred: clear invoicing and payment terms.", version: "1", last_reviewed: "2026-10-20", builtin: false, updated_at: timestamp, matters: [matter.id] }]
const workflows = [{ name: "fixture-services-review", label: "Services review", description: "Review a services agreement and prepare a clear issue list.", scope: "document", prompt: "Review this agreement.", steps: [{ agent: "legal-review", instructions: "Identify issues grounded in the supplied text." }], created_at: timestamp }]
const session = { id: "ses_fixture", title: "Services agreement review", directory: matter.dir, time: { created: timestamp, updated: timestamp }, metadata: {} }
const transcript = [
  { id: "msg_user", type: "user", text: "Please summarise the commercial points that need our attention before we share this agreement.", time: { created: timestamp } },
  { id: "msg_answer", type: "assistant", agent: "legal-review", finish: "stop", time: { created: timestamp + 1000, completed: timestamp + 2000 }, content: [{ type: "text", text: "## Points to review\n\nThe agreement covers the proposed services, delivery dates and the invoicing process. The following points need your attention:\n\n- **Payment timing:** confirm that the payment period matches the agreed commercial position.\n- **Scope changes:** record how additional work is approved before it begins.\n- **Termination:** review the notice period against the expected project timetable.\n\nThese are review points for your consideration. Confirm the instructions before making changes to the document." }] },
  { id: "msg_idle", type: "idle", outcome: "succeeded", time: { created: timestamp + 2000 } },
]
const setupModels = ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true, capabilities: { tools: true, input: ["text"], output: ["text"] } }))
const column = { id: "payment", question: "When is payment due?" }
function hashQuestion(question: string) { let hash = 5381; for (let i = 0; i < question.length; i++) hash = (hash * 33) ^ question.charCodeAt(i); return (hash >>> 0).toString(36) }
const grid = { columns: [column, { id: "notice", question: "What notice is required to terminate?" }], cells: {
  [`${documents[0]!.name}::payment`]: { answer: "Invoices are payable within 30 days of receipt.", status: "reviewed", questionHash: hashQuestion(column.question), citation: { documentName: documents[0]!.name, docPath: documents[0]!.doc_path, section: "4. Payment", excerpt: "Invoices are payable within 30 days of receipt.", charStart: 0, charEnd: 46 }, comments: [{ id: "comment-fixture", text: "Confirm with the client before approving.", at: timestamp }] },
  [`${documents[1]!.name}::payment`]: { answer: "The client requested monthly invoicing.", status: "filled", questionHash: hashQuestion(column.question), citation: { documentName: documents[1]!.name, docPath: documents[1]!.doc_path, section: "Client email", excerpt: "Please use monthly invoicing.", charStart: 0, charEnd: 28 }, comments: [] },
} }
function fixture(options: { setup?: "done" | "details"; list?: "filled" | "empty" | "error" | "loading" } = {}) {
  return {
    setup: options.setup ?? "done", list: options.list ?? "filled", unexpected: [] as string[], errors: [] as string[], requests: [] as { method: string; path: string }[], writes: [] as { method: string; path: string; body: unknown }[],
    matters: [matter, { ...matter, id: "fixture-second", title: "Northbank · Lease renewal", reference: "NB-2026-008", client: "Northbank Works", created_at: timestamp - 86400000 }],
    created: undefined as typeof matter | undefined,
  }
}
async function open(state: ReturnType<typeof fixture>, width: number, height: number) {
  allStates.push(state)
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", timezoneId: "Europe/London" })
  await context.addInitScript(() => { sessionStorage.setItem("dh.launch-notice.v1", "accepted"); localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1") })
  await context.route("**/*", async (route: Route) => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const method = request.method()
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if (!api && url.origin === origin) return route.continue()
    if (!api) { state.unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (method === "OPTIONS") return route.fulfill({ status: 204, headers })
    state.requests.push({ method, path: pathname })
    const reply = (body: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(body) })
    if (method === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    // AI setup requests: a candidate session per opened Settings, no credentials involved.
    if (method === "POST" && pathname === "/api/dochaus/provider-candidate/session") return reply({ sessionID: `fixture-session-${state.requests.length}` })
    if (method === "POST" && pathname === "/api/dochaus/provider-candidate/models") return reply({ data: setupModels, customModels: false })
    if (method === "DELETE" && pathname === "/api/dochaus/provider-candidate/session") return route.fulfill({ status: 204, headers })
    if (method === "POST" && pathname === "/matters") {
      const body = request.postDataJSON() as typeof matter
      state.writes.push({ method, path: pathname, body })
      state.created = { ...matter, ...body, id: "fixture-created", dir: "/fixture/matters/created" }
      return reply(state.created)
    }
    if (method === "PATCH" && pathname === `/matters/${matter.id}`) {
      const body = request.postDataJSON() as typeof matter
      state.writes.push({ method, path: pathname, body })
      state.matters[0] = { ...state.matters[0]!, ...body }
      return reply(state.matters[0])
    }
    if (method === "PUT" && pathname === `/matters/${matter.id}/grid`) { state.writes.push({ method, path: pathname, body: request.postDataJSON() }); return reply({ ok: true }) }
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply({ version: 1, ready: true, completed: state.setup === "done" ? ["models", "research", "details"] : ["models", "research"], skipped: [], next: state.setup === "done" ? "done" : "details" })
    if (pathname === "/preferences") return reply({ ...DEFAULT_DRAFTING, attorney: "Alex Example", firm: "Example Partners LLP" })
    if (pathname === "/matters") {
      if (state.list === "loading") return // Deliberately paused until this isolated context closes.
      return state.list === "error" ? reply({ error: "The document service is temporarily unavailable." }, 503) : reply(state.list === "empty" ? [] : state.matters)
    }
    if (pathname === `/matters/${matter.id}`) return reply({ ...state.matters[0], documents })
    if (pathname === "/matters/fixture-created" && state.created) return reply({ ...state.created, documents: [] })
    if (/^\/matters\/[^/]+\/drafting-status$/.test(pathname)) return reply(null)
    if (/^\/matters\/[^/]+\/telemetry$/.test(pathname)) return reply({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 1, since_ts: timestamp }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (pathname === `/matters/${matter.id}/grid`) return reply(grid)
    if (/^\/matters\/[^/]+\/answer-drafts\//.test(pathname)) return reply({ answers: {}, revision: 0 })
    if (pathname === `/matters/${matter.id}/documents/text`) return reply({ text: "From: Alex Example <alex@example.test>\nSubject: Services agreement instructions\n\nPlease use monthly invoicing and confirm the delivery timetable before sending the agreement. We would like to keep a clear approval process for any additional services." })
    if (pathname === "/jurisdictions") return reply(jurisdictions)
    if (pathname === "/playbooks") return reply(playbooks)
    if (pathname === "/playbooks/library") return reply({ dir: "/fixture/playbooks", playbooks })
    if (pathname === "/templates") return reply({ dir: "/fixture/templates", templates: [{ name: "Consulting agreement.docx", description: "A starting point for professional services engagements.", placeholders: [{ text: "Client name", kind: "text" }, { text: "Commencement date", kind: "date" }] }, { name: "Mutual confidentiality agreement.docx", description: "Two-way confidentiality terms for early discussions.", placeholders: [] }] })
    if (pathname === "/template-drafts") return reply({ drafts: [] })
    if (pathname === "/workflows") return reply({ dir: "/fixture/workflows", workflows })
    if (pathname === "/skills") return reply({ dir: "/fixture/skills", skills: [{ name: "commercial-terms", description: "Guidance for clear commercial drafting.", content: "# Commercial terms\n\nGround every point in the supplied documents.", builtin: false, enabled: true }, { name: "reference-reading", description: "Read source material before making a recommendation.", content: "# Source reading", builtin: true, enabled: true }] })
    if (pathname === "/agents") return reply({ dir: "/fixture/agents", agents: [{ name: "legal-review", label: "Legal reviewer", description: "A specialist for focused agreement review.", instructions: "Review only the supplied document evidence.", builtin: true, enabled: true, created_at: timestamp }] })
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/engine/config") return reply({ model: "fixture/primary", small_model: "fixture/fast" })
    const provider = { id: "fixture", name: "Fixture Models", integrationID: "fixture", activation: "enabled", package: "@opencode/ai/providers/openai-compatible" }
    if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/fixture" }, data: [provider] })
    if (pathname === "/api/model") return reply({ location: { directory: "/fixture" }, data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
    if (pathname === "/api/integration") return reply({ location: { directory: "/fixture" }, data: [{ id: "fixture", name: "Fixture", methods: [], connections: [{ type: "credential", id: "fixture-account", label: "Fixture", method: "key" }] }] })
    if (pathname === "/api/session/active") return reply({ data: {} })
    if (pathname === "/api/session") return reply({ data: [session], cursor: { previous: null, next: null } })
    if (pathname === `/api/session/${session.id}/message`) return reply({ data: transcript, cursor: { previous: null, next: null } })
    if (/^\/api\/session\/[^/]+\/form$/.test(pathname)) return reply({ data: [] })
    if (pathname === "/api/permission/request") return reply({ location: { directory: matter.dir }, data: [] })
    if (pathname === "/api/agent") return reply({ location: { directory: matter.dir }, data: ["legal-review", "draft", "research", workflows[0]!.name].map(id => ({ id, description: "Synthetic fixture assistant", mode: "primary", hidden: false })) })
    if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: `data: ${JSON.stringify({ id: "evt_fixture", type: "server.connected", data: {} })}\n\n` })
    state.unexpected.push(`${method} ${pathname}`)
    return route.abort()
  })
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  page.on("pageerror", (error: Error) => state.errors.push(error.message))
  return { page, context }
}
async function capture(page: Page, name: string, options: { fullPage?: boolean } = {}) {
  await page.evaluate(async () => { await document.fonts.ready })
  const metrics = await page.evaluate(() => ({ overflowX: document.documentElement.scrollWidth - document.documentElement.clientWidth, dialogs: Array.from(document.querySelectorAll<HTMLElement>('dialog[open], [role="dialog"][aria-modal="true"]')).map(node => { const r = node.getBoundingClientRect(); return { name: node.getAttribute("aria-label") ?? node.getAttribute("aria-labelledby"), left: r.left, right: r.right, top: r.top, bottom: r.bottom, width: innerWidth, height: innerHeight } }) }))
  await page.screenshot({ path: path.join(evidence!, `${name}.png`), fullPage: options.fullPage ?? metrics.dialogs.length === 0 })
  let accessibility: unknown
  if (axeScript) {
    await page.addScriptTag({ path: axeScript })
    accessibility = await page.evaluate(async () => {
      const axe = (window as unknown as { axe: { run: (options: unknown) => Promise<{ violations: { id: string; impact: string; description: string; helpUrl: string; nodes: { target: unknown; failureSummary: string }[] }[] }> } }).axe
      const report = await axe.run({ runOnly: { type: "tag", values: ["wcag2a", "wcag2aa", "wcag21a", "wcag21aa", "wcag22aa"] } })
      return report.violations.map(({ id, impact, description, helpUrl, nodes }) => ({ id, impact, description, helpUrl, nodes: nodes.map(({ target, failureSummary }) => ({ target, failureSummary })) }))
    })
  }
  screens.push({ name, ...metrics, accessibility })
  if (!baseline) {
    expect(metrics.overflowX, `${name}: document horizontal overflow`).toBeLessThanOrEqual(1)
    for (const dialog of metrics.dialogs) { expect(dialog.left, `${name}: dialog left`).toBeGreaterThanOrEqual(-1); expect(dialog.right, `${name}: dialog right`).toBeLessThanOrEqual(dialog.width + 1); expect(dialog.top, `${name}: dialog top`).toBeGreaterThanOrEqual(-1); expect(dialog.bottom, `${name}: dialog bottom`).toBeLessThanOrEqual(dialog.height + 1) }
  }
}
async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); console.log(`FAIL ${name}: ${error instanceof Error ? error.message : String(error)}`) }
}
async function go(page: Page, pathname = "/") { await page.goto(`${baseURL}${pathname}`); await expect(page.locator(".sidebar")).toBeVisible() }
async function closeSettings(page: Page) { await page.getByRole("dialog", { name: "Settings", exact: true }).getByRole("button", { name: "Close", exact: true }).click() }

try {
  for (const [width, height] of [[375, 812], [768, 1024], [1440, 1000]] as const) {
    const size = `${width}x${height}`
    await scenario(`Routes, forms, settled conversation, documents, review and preferences (${size})`, async () => {
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        await go(page)
        await expect(page.locator(".matters-index-list .list-row")).toHaveCount(2)
        await capture(page, `matters-${size}`)
        await page.getByLabel("Filter matters", { exact: true }).fill("Northbank")
        await expect(page.locator(".matters-index-list .list-row")).toHaveCount(1)
        await page.getByLabel("Filter matters", { exact: true }).fill("")
        const createToggle = page.locator(".matter-create-toggle")
        if (await createToggle.isVisible()) await createToggle.click()
        await expect(page.getByRole("textbox", { name: "Matter title", exact: true })).toBeVisible()
        await capture(page, `matter-create-${size}`)
        const edit = page.locator(".matters-index-list .list-row").first().getByRole("button", { name: /^Edit/ })
        await edit.click()
        await expect(page.locator(".matter-edit")).toBeVisible()
        await capture(page, `matter-edit-${size}`)
        if (!baseline) {
          expect(await page.locator(".matter-edit").evaluate((node: HTMLDialogElement) => node.matches(":modal"))).toBe(true)
          for (let i = 0; i < 14; i++) { await page.keyboard.press("Tab"); expect(await page.evaluate(() => document.activeElement === document.body || Boolean(document.activeElement?.closest(".matter-edit")))).toBe(true) }
        }
        await page.keyboard.press("Escape")
        await expect(page.locator(".matter-edit")).toHaveCount(0)
        if (!baseline) await expect(edit).toBeFocused()
        for (const route of ["templates", "playbooks", "workflows", "skills", "agents"]) {
          await go(page, `/${route}`)
          await expect(page.getByRole("heading", { name: route[0]!.toUpperCase() + route.slice(1), exact: true })).toBeVisible()
          await expect(page.locator(".library-collection .list-row").first()).toBeVisible()
          await expect(page.locator(".chat-card")).toBeVisible()
          await capture(page, `${route}-${size}`)
        }
        await go(page, `/matter/${matter.id}?session=${session.id}`)
        await expect(page.getByText("Points to review", { exact: true })).toBeVisible()
        await expect(page.getByRole("textbox", { name: "Message", exact: true })).toBeVisible()
        await capture(page, `matter-conversation-${size}`)
        if (!baseline && width > 700) {
          const bounds = await page.locator(".composer").evaluate(node => { const r = node.getBoundingClientRect(); const bar = document.querySelector(".matter-statusbar")!.getBoundingClientRect(); return { top: r.top, bottom: r.bottom, height: innerHeight, clear: r.bottom <= bar.top + 1 } })
          expect(bounds.top).toBeGreaterThanOrEqual(0); expect(bounds.bottom).toBeLessThanOrEqual(bounds.height); expect(bounds.clear).toBe(true)
        }
        const docs = page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ })
        await docs.click()
        await expect(page.locator("#documents-drawer")).toBeVisible()
        await expect(page.locator("#documents-drawer").getByText("Services agreement.docx", { exact: true })).toBeVisible()
        await capture(page, `documents-drawer-${size}`)
        await page.locator("#documents-drawer").getByRole("button", { name: "Client instructions.eml", exact: true }).click()
        const reader = page.getByRole("dialog", { name: "Client instructions.eml", exact: true })
        await expect(reader.locator(".document-text")).toContainText("Please use monthly invoicing")
        await capture(page, `document-reader-${size}`)
        await reader.getByRole("button", { name: "Close", exact: true }).click()
        await page.locator("#documents-drawer .docs-drawer-close").click()
        await page.getByRole("link", { name: "Review table", exact: true }).click()
        await expect(page.locator(".review-modal")).toBeVisible()
        await expect(page.getByRole("table", { name: "Document review" })).toContainText("Invoices are payable within 30 days")
        await capture(page, `review-${size}`)
        await page.locator(".review-modal").getByRole("button", { name: "Close", exact: true }).click()
        const settings = page.getByRole("button", { name: "Settings", exact: true })
        await settings.click()
        const modal = page.getByRole("dialog", { name: "Settings", exact: true })
        await expect(modal.getByLabel("Primary model", { exact: true })).toBeEnabled()
        await capture(page, `settings-models-${size}`)
        await modal.getByRole(baseline ? "button" : "tab", { name: "Appearance", exact: true }).click()
        await modal.getByRole("button", { name: "Large", exact: true }).click()
        await expect(page.locator("html")).toHaveAttribute("data-text", "large")
        await capture(page, `settings-appearance-large-${size}`)
        await modal.getByRole("button", { name: "Standard", exact: true }).click()
        await closeSettings(page)
        await expect(settings).toBeFocused()
        if (!baseline && width === 1440) {
          const toggle = page.getByRole("button", { name: "Collapse sidebar", exact: true })
          const box = await toggle.boundingBox()
          if (!box) throw new Error("Missing collapse geometry")
          // Exercise both visible rims. The same circle stays on the border.
          for (const fraction of [.08, .92, .08, .92]) {
            const wasExpanded = await page.locator(".sidebar-collapse").getAttribute("aria-expanded")
            const target = await page.locator(".sidebar-collapse").boundingBox()
            if (!target) throw new Error("Missing collapse target")
            await page.mouse.click(target.x + target.width * fraction, target.y + target.height / 2)
            await expect(page.locator(".sidebar-collapse")).toHaveAttribute("aria-expanded", wasExpanded === "true" ? "false" : "true")
          }
        }
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } finally { await context.close() }
    })
    for (const list of ["empty", "error", "loading"] as const) {
      await scenario(`Matters ${list} (${size})`, async () => {
        const state = fixture({ list }); const { page, context } = await open(state, width, height)
        try { await go(page); await expect(page.locator(list === "empty" ? ".empty-state" : list === "error" ? ".load-error" : ".skeleton-list")).toBeVisible(); await capture(page, `matters-${list}-${size}`); expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([]) }
        finally { await context.close() }
      })
    }
    await scenario(`Onboarding details (${size})`, async () => {
      const state = fixture({ setup: "details" }); const { page, context } = await open(state, width, height)
      try { await go(page); await expect(page.getByRole("heading", { name: "Make doc.haus yours", exact: true })).toBeVisible(); await capture(page, `onboarding-${size}`, { fullPage: false }); expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([]) }
      finally { await context.close() }
    })
  }
  await scenario("Narrow, short desktop keeps the composer accessible", async () => {
    const state = fixture(); const { page, context } = await open(state, 1024, 640)
    try { await go(page, `/matter/${matter.id}?session=${session.id}`); await expect(page.getByText("Points to review", { exact: true })).toBeVisible(); await capture(page, "matter-conversation-1024x640"); if (!baseline) { const r = await page.locator(".composer").boundingBox(); expect(r!.y + r!.height).toBeLessThanOrEqual(640) } expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([]) }
    finally { await context.close() }
  })
  if (!baseline) for (const [width, height] of [[375, 812], [768, 1024], [1440, 1000], [720, 500]] as const) {
    await scenario(`Appearance, reduced motion and constrained workspace (${width}x${height})`, async () => {
      const state = fixture(); const { page, context } = await open(state, width, height)
      try {
        await go(page, `/matter/${matter.id}?session=${session.id}`)
        await expect(page.getByText("Points to review", { exact: true })).toBeVisible()
        for (const size of ["Large", "Compact"] as const) {
          await page.getByRole("button", { name: "Settings", exact: true }).click()
          const modal = page.getByRole("dialog", { name: "Settings", exact: true })
          await modal.getByRole("tab", { name: "Appearance", exact: true }).click()
          await modal.getByRole("button", { name: size, exact: true }).click()
          await capture(page, `appearance-${size.toLowerCase()}-${width}x${height}`)
          await closeSettings(page)
          const composer = page.locator(".composer")
          if (height < 600 || width <= 700) await composer.scrollIntoViewIfNeeded()
          const box = await composer.boundingBox()
          expect(box!.y).toBeGreaterThanOrEqual(0)
          expect(box!.y + box!.height).toBeLessThanOrEqual(height + 1)
          await capture(page, `workspace-${size.toLowerCase()}-${width}x${height}`, { fullPage: false })
          await page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ }).click()
          await page.locator("#documents-drawer").getByRole("button", { name: "Client instructions.eml", exact: true }).click()
          const reader = page.getByRole("dialog", { name: "Client instructions.eml", exact: true })
          await expect(reader.locator(".document-text")).toContainText("Please use monthly invoicing")
          await capture(page, `reader-${size.toLowerCase()}-${width}x${height}`)
          await reader.getByRole("button", { name: "Close", exact: true }).click()
          await page.locator("#documents-drawer .docs-drawer-close").click()
        }
        expect(await page.evaluate(() => matchMedia("(prefers-reduced-motion: reduce)").matches)).toBe(true)
        expect(await page.locator(".sidebar").evaluate(node => getComputedStyle(node).transitionDuration.split(",").every(value => parseFloat(value) === 0))).toBe(true)
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } finally { await context.close() }
    })
  }
  if (!baseline) await scenario("Matter edit and create preserve existing API fields", async () => {
    const state = fixture(); const { page, context } = await open(state, 1440, 1000)
    try {
      await go(page)
      await page.locator(".matters-index-list .list-row").first().getByRole("button", { name: /^Edit/ }).click()
      const edit = page.locator(".matter-edit")
      await edit.getByLabel("Title", { exact: true }).fill("Alder House · Updated agreement review")
      await edit.locator(".matter-edit-foot .primary").click()
      await expect(edit).toHaveCount(0)
      expect(state.writes[0]).toMatchObject({ method: "PATCH", path: `/matters/${matter.id}`, body: { title: "Alder House · Updated agreement review", reference: matter.reference, jurisdictions: matter.jurisdictions } })
      expect(Object.hasOwn(state.writes[0]!.body as object, "client")).toBe(false)
      await page.getByRole("textbox", { name: "Matter title", exact: true }).fill("New synthetic engagement")
      await page.getByRole("textbox", { name: "Client you act for (optional)", exact: true }).fill("Example Client")
      await page.getByRole("textbox", { name: "Matter number (optional)", exact: true }).fill("EX-001")
      await page.getByRole("button", { name: "Create matter", exact: true }).click()
      await expect(page).toHaveURL(/\/matter\/fixture-created$/)
      await expect(page.getByRole("heading", { name: "New synthetic engagement" })).toBeVisible()
      expect(state.writes[1]).toEqual({ method: "POST", path: "/matters", body: { title: "New synthetic engagement", reference: "EX-001", client: "Example Client" } })
      expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
    } finally { await context.close() }
  })
} finally {
  await writeFile(path.join(evidence, "material-expressive-results.json"), JSON.stringify({ baseline, results, screens, fixtures: allStates.map(({ errors, unexpected, requests, writes, list, setup }) => ({ errors, unexpected, requests, writes, list, setup })) }, null, 2) + "\n")
  await browser.close()
}
expect(results.filter(result => result.status === "failed")).toEqual([])
if (axeScript) expect(screens.filter(screen => Array.isArray(screen.accessibility) && screen.accessibility.length > 0).map(screen => ({ name: screen.name, violations: screen.accessibility }))).toEqual([])
console.log(`PASS ${results.length} scenarios; ${screens.length} screenshots. ${baseline ? "Baseline capture: new design assertions skipped." : "Responsive and modal bounds checked."} ${axeScript ? "Axe results included in JSON." : "Set AXE_SCRIPT to add axe WCAG checks."}`)
