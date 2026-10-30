// Starter regression against the real web route and an in-memory ingest fixture.
// Every API request is intercepted. Synthetic saved conversations are replayed;
// generation, credentials and unknown requests are refused before any network I/O.
import { mkdir } from "node:fs/promises"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
type Route = import("@playwright/test").Route
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
const origin = new URL(baseURL).origin
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const allStates: ReturnType<typeof fixture>[] = []
const chronology = { schemaVersion: 1, revision: 0, created: false, events: [], coverage: [], proposals: [], sourceDates: {}, suggestion: { status: "unknown", reason: "" }, history: [], framing: "", gaps: [] }
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
    grid: { ...structuredClone(grid), version: 0 },
    created: undefined as typeof matter | undefined,
  }
}
async function open(state: ReturnType<typeof fixture>, width: number, height: number) {
  allStates.push(state)
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", timezoneId: "Europe/London" })
  await context.addInitScript(() => { localStorage.setItem("dh.launch-notice.v1", new Date().toDateString()); localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1") })
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
    if (method === "PUT" && pathname === `/matters/${matter.id}/grid`) {
      const next = request.postDataJSON() as typeof state.grid
      state.writes.push({ method, path: pathname, body: next })
      if (next.version !== state.grid.version) return reply({ grid: state.grid }, 409)
      state.grid = { ...next, version: next.version + 1 }
      return reply(state.grid)
    }
    if (method === "POST" && pathname === `/matters/${matter.id}/grid/sources`) return reply(Object.fromEntries((request.postDataJSON() as { cells: { key: string }[] }).cells.map(cell => [cell.key, "verified"])))
    // The matter page checks the chronology on open and on focus; this matter has none.
    if (pathname === `/matters/${matter.id}/chronology` || (method === "POST" && pathname === `/matters/${matter.id}/chronology/assess`)) return reply(chronology)
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/word/config") return reply({ available: false, origin: null, manifestUrl: null })
    if (/^\/matters\/[^/]+\/documents\/word$/.test(pathname)) return reply({ documents: [] })
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
    if (pathname === `/matters/${matter.id}/grid`) return reply(state.grid)
    if (pathname === `/matters/${matter.id}/documents/versions`) return reply({ name: url.searchParams.get("name"), document: null, versions: [], events: [], working_version: null })
    if (/^\/matters\/[^/]+\/answer-drafts\//.test(pathname)) return reply({ answers: {}, revision: 0 })
    if (pathname === `/matters/${matter.id}/documents/text`) return reply({ text: "From: Alex Example <alex@example.test>\nSubject: Services agreement instructions\n\nPlease use monthly invoicing and confirm the delivery timetable before sending the agreement. We would like to keep a clear approval process for any additional services." })
    if (pathname === "/jurisdictions") return reply(jurisdictions)
    if (pathname === "/clients") return reply([])
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

async function exercise(width: number, height: number) {
  const state = fixture()
  const originalCells = structuredClone(state.grid.cells)
  const { page, context } = await open(state, width, height)
  try {
    await page.goto(`${baseURL}/matter/${matter.id}?session=${session.id}`)
    await expect(page.getByText("Points to review", { exact: true })).toBeVisible()
    await page.getByRole("link", { name: "Review table", exact: true }).click()
    const gridDialog = page.getByRole("dialog", { name: `Review table — ${matter.title}` })
    await expect(gridDialog).toBeVisible()
    // Starters open from the + menu in the header row.
    const addButton = gridDialog.locator(".review-table").getByRole("button", { name: "Add questions", exact: true })
    const openStarters = async () => {
      await addButton.click()
      await page.getByRole("group", { name: "Add questions" }).getByRole("button", { name: "Use a starter", exact: true }).click()
    }
    // The chooser comes first and recommends from source documents only.
    const chooser = page.getByRole("dialog", { name: "Add questions" })
    const starter = page.getByRole("dialog", { name: "Document overview" })
    const chooseOverview = async () => {
      await expect(chooser).toBeVisible()
      await chooser.getByRole("radio", { name: /^Document overview/ }).check()
      await chooser.getByRole("button", { name: "Continue" }).click()
      await expect(starter).toBeVisible()
    }
    await openStarters()
    await expect(chooser).toBeVisible()
    await expect(chooser.getByRole("radio", { name: /^Compare agreements/ })).toBeChecked()
    await expect(chooser).toContainText("Recommended · 2 source documents in this matter")
    await page.screenshot({ path: `${evidence}/starter-chooser-${width}.png` })
    await chooseOverview()
    await expect(starter).toContainText(`3 documents: ${documents.map(document => document.name).join(", ")}`)
    await expect(starter).toContainText("Adds 3 questions · 9 cells to fill. No AI runs until you choose Fill.")
    await page.screenshot({ path: `${evidence}/starter-initial-${width}.png` })
    for (let i = 0; i < 12; i++) {
      await page.keyboard.press("Tab")
      expect(await page.evaluate(() => !!document.activeElement?.closest('dialog[aria-label="Document overview"]'))).toBe(true)
    }
    const bounds = await starter.locator(".picker-panel").boundingBox()
    expect(bounds).not.toBeNull()
    expect(bounds!.x).toBeGreaterThanOrEqual(0)
    expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(width + 1)
    await page.screenshot({ path: `${evidence}/starter-${width}.png` })
    await page.keyboard.press("Escape")
    await expect(starter).toHaveCount(0)
    await expect(addButton).toBeFocused()
    expect(state.writes).toHaveLength(0)

    await openStarters()
    await chooseOverview()
    await starter.getByRole("textbox", { name: "Question 1", exact: true }).fill("What is the main topic?")
    await starter.getByRole("button", { name: "Remove question 2" }).click()
    await expect(starter).toContainText("Adds 2 questions · 6 cells to fill")
    await starter.getByRole("button", { name: "Add questions" }).click()
    await expect(starter).toHaveCount(0)
    expect(state.writes).toHaveLength(1)
    expect(state.grid.cells).toEqual(originalCells)
    expect(state.grid.columns.map(column => column.question)).toEqual([
      "When is payment due?", "What notice is required to terminate?", "What is the main topic?", "What dates or deadlines are stated in this document?",
    ])
    await page.reload()
    await expect(gridDialog).toBeVisible()
    await expect(gridDialog.getByRole("button", { name: "Edit question: What is the main topic?" })).toBeVisible()
    await expect(gridDialog.getByText("Invoices are payable within 30 days of receipt.")).toBeVisible()
    await expect(gridDialog.getByRole("button", { name: "1 comments on Services agreement.docx: When is payment due?" })).toBeVisible()
    await openStarters()
    await chooseOverview()
    await expect(starter).toContainText("Already in the table")
    await starter.getByRole("button", { name: "Cancel" }).click()
    expect(state.writes).toHaveLength(1)

    await gridDialog.getByRole("button", { name: "Fill 10 cells", exact: true }).click()
    const fill = page.getByRole("dialog", { name: "Fill 10 cells?" })
    await expect(fill).toBeVisible()
    await expect(fill).toContainText("10 cells")
    await expect(fill).toContainText("Fill uses AI")
    const firstDoc = fill.locator("section").filter({ has: page.getByRole("heading", { name: documents[0]!.name }) })
    await expect(firstDoc).not.toContainText("When is payment due?")
    await expect(firstDoc).toContainText("What notice is required to terminate?")
    await expect(firstDoc).toContainText("What is the main topic?")
    await expect(fill).toContainText(documents[2]!.name)
    await page.screenshot({ path: `${evidence}/fill-${width}.png` })
    await fill.getByRole("button", { name: "Cancel" }).click()
    await expect(fill).toHaveCount(0)
    expect(state.writes).toHaveLength(1)
    expect(state.requests.filter(request => request.method === "POST" && request.path.startsWith("/api/session"))).toEqual([])
    expect(state.unexpected).toEqual([])
    expect(state.errors).toEqual([])
    console.log(`PASS review starter ${width}x${height}`)
  } finally { await context.close() }
}

try {
  await exercise(1440, 1000)
  await exercise(375, 812)
} finally { await browser.close() }
