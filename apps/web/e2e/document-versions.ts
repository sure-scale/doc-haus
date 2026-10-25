// Document versions (#268) against an already-running web preview. Every API
// request is intercepted: synthetic Word, PDF and email documents with kept
// versions, their version bytes, restore, compare and upload answers are
// fixtures, and generation, credentials and unknown requests are refused. It
// checks the versions list, viewing an older version (and what it switches
// off), the in-page restore confirmation and its request, where focus goes, a
// file changed outside doc.haus, changes made elsewhere, PDF and email
// versions, a version blackline and the return to the list, the compare
// dialog's previous-version default and duplicate uploads, and the upload
// prompt's new-version, duplicate and tracked-changes messages, at 1440×1000
// and 400×860.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../../dochaus/lib/drafting-preferences"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const { zipSync, strToU8 } = (await import(Bun.resolveSync("fflate", path.resolve(import.meta.dir, "../../../dochaus")))) as {
  zipSync: (files: Record<string, Uint8Array>) => Uint8Array; strToU8: (text: string) => Uint8Array
}
type Page = import("@playwright/test").Page
type Route = import("@playwright/test").Route
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
const origin = new URL(baseURL).origin
await mkdir(evidence, { recursive: true })
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: Record<string, unknown>[] = []

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const docx = (paragraphs: string[]) => zipSync(Object.fromEntries(Object.entries({
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/document.xml": `<w:document xmlns:w="${W}"><w:body>${paragraphs.map((text, index) => `<w:p><w:r>${index ? "" : "<w:rPr><w:b/></w:rPr>"}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
}).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)])))

const timestamp = Date.UTC(2026, 9, 24, 9)
const matter = { id: "fixture-matter", title: "Alder House · Services agreement", reference: "AH-2026-014", client: "Alder House Studio", jurisdictions: ["EW"], dir: "/fixture/matters/alder-house", created_at: timestamp }
const NAME = "Services agreement.docx"
const EMAIL = "Client instructions.eml"
const OTHER = "Side letter.docx"
const PDF = "Signed scan.pdf"
// A Word file from before versions were kept: none to list.
const LETTER = "Engagement letter.docx"
const documents = [
  { id: 1, name: NAME, doc_path: `${matter.dir}/${NAME}`, created_at: timestamp, role: "source" },
  { id: 2, name: EMAIL, doc_path: `${matter.dir}/${EMAIL}`, created_at: timestamp - 60_000, role: "source" },
  { id: 3, name: OTHER, doc_path: `${matter.dir}/${OTHER}`, created_at: timestamp - 120_000, role: "source" },
  { id: 4, name: PDF, doc_path: `${matter.dir}/${PDF}`, created_at: timestamp - 180_000, role: "source" },
  { id: 5, name: LETTER, doc_path: `${matter.dir}/${LETTER}`, created_at: timestamp - 240_000, role: "source" },
]
const body = (clause: string) => ["Services agreement", "1. The Supplier shall provide the services described in the schedule.", clause, "3. Either party may end this agreement on thirty days notice."]
const pdfBytes = (text: string) => strToU8(`%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj 3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 144]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj 4 0 obj<</Length 44>>stream\nBT /F1 18 Tf 20 60 Td (${text}) Tj ET\nendstream endobj 5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF`)
const EMAIL_TEXT = ["Subject: Services agreement", "From: Synthetic Client <client@example.invalid>", "", "Please send the agreement with the payment terms we discussed."].join("\n")
const bytes: Record<string, Uint8Array> = {
  "v-2": docx(body("2. Invoices are payable within [REDACTED] days of receipt.")),
  "v-3": docx(body("2. Invoices are payable within sixty days of receipt.")),
  "v-4": docx(body("2. Invoices are payable within forty-five days of receipt.")),
  "o-1": docx(["Side letter", "First draft."]), "o-2": docx(["Side letter", "Second draft."]), "o-3": docx(["Side letter", "Signed form."]),
  "p-1": pdfBytes("Scan, first page set"), "p-2": pdfBytes("Scan, signed set"),
  "e-1": strToU8(`${EMAIL_TEXT}\n`), "e-2": strToU8(`${EMAIL_TEXT}\nWith thanks.\n`),
  [LETTER]: docx(["Engagement letter", "We act for you on the services agreement."]),
}
const version = (number: number, patch: Record<string, unknown>, prefix = "v") => ({
  id: `${prefix}-${number}`, number, sha256: `${prefix}${number}`.repeat(64).slice(0, 64), ext: ".docx", bytes: 12_000, parent: number > 1 ? `${prefix}-${number - 1}` : null,
  derived_from: null, actor: "unknown", run: null, label: null, created_at: timestamp + number * 3_600_000, current: false, removed: null, ...patch,
})
type Version = ReturnType<typeof version>
const initialVersions = (): Record<string, Version[]> => ({
  [NAME]: [
    version(4, { origin: "accept", current: true }),
    version(3, { origin: "revision", actor: "Pat Reviewer", label: "Counterparty markup" }),
    version(2, { origin: "redaction", actor: "agent:legal-review" }),
    version(1, { origin: "upload", removed: "redaction" }),
  ],
  [EMAIL]: [version(2, { origin: "replace", current: true, ext: ".eml" }, "e"), version(1, { origin: "upload", ext: ".eml" }, "e")],
  [OTHER]: [version(3, { origin: "replace", current: true }, "o"), version(2, { origin: "replace" }, "o"), version(1, { origin: "upload" }, "o")],
  [PDF]: [version(2, { origin: "replace", current: true, ext: ".pdf" }, "p"), version(1, { origin: "upload", ext: ".pdf" }, "p")],
  [LETTER]: [],
})
const comparisonFor = (url: URL) => {
  const side = (key: "left" | "right") => {
    const id = url.searchParams.get(`${key}Version`)
    const name = url.searchParams.get(key)!
    const number = id ? Number(id.slice(2)) : 4
    return { name, label: `${name.replace(/\.docx$/, "")} (version ${number})`, trackedChanges: { count: 0, authors: [] }, pendingProposals: 0, ...(id ? { version: { id, number } } : {}) }
  }
  const left = side("left"), right = side("right")
  const hunk = { kind: "modified", location: { scope: "body", paragraph: 3, side: "right", number: "2" }, before: "Invoices are payable within sixty days of receipt.", after: "Invoices are payable within forty-five days of receipt.", untracked: true }
  return { left, right, options: { ignoreFormatting: true, detectMoves: true, headersFooters: true }, attribution: "authors", changes: [{ id: "C1", ...hunk }], untracked: [{ id: "U1", ...hunk }], blacklineName: `Blackline - ${left.label} vs ${right.label}.docx` }
}
const NOTICE = "Saved as version 6, with 2 tracked changes by Pat Reviewer; resolve them in Word before doc.haus edits it."
const proposal = { id: 7, doc_name: NAME, scope: "phrase", find_text: "forty-five", old_text: "forty-five days", new_text: "thirty days", author: "doc.haus", review_comment: "Align with clause 3.", created_at: timestamp }

type Options = {
  // The working file changed outside doc.haus since its current version.
  changed?: boolean
  // Pending proposals on the Word document.
  redlines?: unknown[]
  // How long the versions list takes to answer once a restore has been made.
  slowAfterRestore?: number
  // The first restore fails at the service.
  restoreFails?: boolean
}
type Fixture = {
  opts: Options; docs: Record<string, Version[]>; outside: boolean; restoredOnce: boolean; asked: string[]; texts: number; events: unknown[]
  restores: { path: string; body: unknown }[]; compares: string[]; uploads: { existing: string | null; duplicate: boolean }[]; unexpected: string[]; errors: string[]
}
function fixture(opts: Options = {}): Fixture {
  return {
    opts,
    docs: initialVersions(),
    outside: Boolean(opts.changed),
    restoredOnce: false,
    asked: [] as string[],
    texts: 0,
    events: [] as unknown[],
    restores: [] as { path: string; body: unknown }[],
    compares: [] as string[],
    uploads: [] as { existing: string | null; duplicate: boolean }[],
    unexpected: [] as string[], errors: [] as string[],
  }
}
// Another session restores version 3 of the Word document as its next version.
const restoredElsewhere = (state: Fixture) => restoreAt(state, NAME, "v-3")
// What the restore route does: the working file's own changes are kept first,
// then the version's bytes become the next version.
function restoreAt(state: Fixture, name: string, id: string) {
  const list = state.docs[name]!
  const from = list.find((v) => v.id === id)!
  const prefix = id.split("-")[0]!
  let current = list.find((v) => v.current)!
  let next = Math.max(...list.map((v) => v.number)) + 1
  const kept: Version[] = []
  if (state.outside) {
    kept.push(current = version(next++, { origin: "unknown", parent: current.id, ext: from.ext }, prefix))
    state.outside = false
  }
  const restored = version(next, { origin: "restore", current: true, parent: current.id, sha256: from.sha256, ext: from.ext, label: `Restored from version ${from.number}` }, prefix)
  state.docs[name] = [restored, ...kept.reverse(), ...list.map((v) => ({ ...v, current: false }))]
  state.events.push({ id: state.events.length + 1, kind: "restore", version_id: restored.id, actor: "unknown", detail: { restored_from: from.id, restored_number: from.number }, created_at: timestamp })
  bytes[restored.id] = bytes[from.id]!
  return { from, restored }
}

async function open(state: Fixture, width: number, height: number) {
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", timezoneId: "Europe/London" })
  await context.addInitScript(() => { try { sessionStorage.setItem("dh.launch-notice.v1", "accepted"); localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1") } catch { /* the browser's PDF viewer frame has no storage */ } })
  await context.route("**/*", async (route: Route) => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const method = request.method()
    // The browser's own PDF viewer loads chrome:// resources; blocking them leaves it blank.
    if (/^chrome(-extension)?:/.test(request.url())) return route.continue()
    // The app and its Word runtime come from the preview; every API call goes
    // through the /ingest and /engine prefixes; nothing else may load.
    const api = url.origin === origin && (url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/"))
    if (!api && url.origin === origin) return route.continue()
    if (!api) { state.unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (method === "OPTIONS") return route.fulfill({ status: 204, headers })
    const reply = (value: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(value) })
    const word = (value: Uint8Array, type = `${CT}.document`) => route.fulfill({ status: 200, headers, contentType: type, body: Buffer.from(value) })
    const prefix = `/matters/${matter.id}`
    if (method === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (method === "POST" && pathname === `${prefix}/documents`) {
      const form = request.postDataBuffer()?.toString("latin1") ?? ""
      const existing = /name="existing"\r\n\r\n([^\r]+)/.exec(form)?.[1] ?? null
      const duplicate = /IDENTICAL TO|SIDE LETTER/.test(form)
      state.uploads.push({ existing, duplicate })
      if (form.includes("IDENTICAL TO VERSION 3")) return reply({ error: `${NAME} is already in this matter as version 3 of ${NAME}. Nothing new was stored. Use Restore on version 3 to make it current again.`, code: "duplicate", name: NAME, document_id: "doc-1", version: { id: "v-3", number: 3, current: false } }, 409)
      if (form.includes("SIDE LETTER V2")) return reply({ error: `${OTHER} is already in this matter as version 2 of ${OTHER}. Nothing new was stored. Use Restore on version 2 to make it current again.`, code: "duplicate", name: OTHER, document_id: "doc-3", version: { id: "o-2", number: 2, current: false } }, 409)
      if (form.includes("SIDE LETTER V3")) return reply({ error: `${OTHER} is already in this matter as the current version of ${OTHER}. Nothing new was stored.`, code: "duplicate", name: OTHER, document_id: "doc-3", version: { id: "o-3", number: 3, current: true } }, 409)
      if (!existing) return reply({ error: `${NAME} is already in this matter.`, code: "source_exists", name: NAME }, 409)
      return reply({ name: NAME, docPath: documents[0]!.doc_path, sections: 4, chunks: 4, injection: null, quality: { status: "ok" }, version: { id: "v-6", number: 6, created: true }, tracked_changes: { count: 2, authors: ["Pat Reviewer"] }, notice: NOTICE })
    }
    const restore = /^\/matters\/[^/]+\/documents\/versions\/([^/]+)\/restore$/.exec(pathname)
    if (method === "POST" && restore) {
      state.restores.push({ path: pathname, body: request.postDataJSON() })
      const id = restore[1]!
      const name = Object.keys(state.docs).find((doc) => state.docs[doc]!.some((v) => v.id === id))!
      const list = state.docs[name]!
      if (state.opts.restoreFails && state.restores.length === 1) return reply({ error: "The document service could not save the file. Try again in a moment." }, 503)
      const from = list.find((v) => v.id === id)!
      if (!state.outside && list.find((v) => v.current)!.sha256 === from.sha256) return reply({ error: `${name} already holds version ${from.number}.`, code: "already_current" }, 409)
      const { restored } = restoreAt(state, name, id)
      state.restoredOnce = true
      return reply({ name, docPath: `${matter.dir}/${name}`, version: { id: restored.id, number: restored.number, created: true }, restored_from: { id: from.id, number: from.number } })
    }
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    const name = url.searchParams.get("name") ?? ""
    if (pathname === `${prefix}/documents/versions`) {
      state.asked.push(name)
      const list = state.docs[name]
      if (!list) return reply({ error: `${name} is not a document in this matter.` }, 404)
      if (name === NAME && state.restoredOnce && state.opts.slowAfterRestore) await new Promise((resolve) => setTimeout(resolve, state.opts.slowAfterRestore))
      const held = list.find((v) => v.current)
      return reply({ name, document: held ? { id: `doc-${name}`, name, current_version: held.id } : null, versions: list, events: state.events, working_version: held && !(name === NAME && state.outside) ? held.id : null })
    }
    const content = /^\/matters\/[^/]+\/documents\/versions\/([^/]+)\/content$/.exec(pathname)
    if (content) return bytes[content[1]!] ? word(bytes[content[1]!]!, content[1]!.startsWith("p-") ? "application/pdf" : undefined) : reply({ error: "The kept bytes of this version are missing or damaged.", code: "version_damaged" }, 409)
    if (pathname === `${prefix}/documents/redlined` || pathname === `${prefix}/documents/content`) {
      if (name === PDF) return word(bytes["p-2"]!, "application/pdf")
      if (name === LETTER) return word(bytes[LETTER]!)
      const current = state.docs[name]?.find((v) => v.current)
      return current ? word(bytes[current.id]!) : reply({ error: "Not found" }, 404)
    }
    if (pathname === `${prefix}/documents/text` && name === EMAIL) { state.texts++; return reply({ text: EMAIL_TEXT, injection: null }) }
    if (pathname === `${prefix}/redlines`) return reply(name === NAME ? (state.opts.redlines ?? []) : [])
    if (pathname === `${prefix}/documents/amendments`) return reply(null)
    if (pathname === `${prefix}/compare`) {
      state.compares.push(url.search)
      return url.searchParams.get("format") === "json" ? reply(comparisonFor(url)) : word(bytes["v-4"]!)
    }
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply({ version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" })
    if (pathname === "/preferences") return reply({ ...DEFAULT_DRAFTING, attorney: "Alex Example", firm: "Example Partners LLP" })
    if (pathname === "/matters") return reply([matter])
    if (pathname === prefix) return reply({ ...matter, documents })
    if (pathname === `${prefix}/drafting-status`) return reply(null)
    if (pathname === `${prefix}/telemetry`) return reply({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 0, since_ts: timestamp }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (/^\/matters\/[^/]+\/answer-drafts\//.test(pathname)) return reply({ answers: {}, revision: 0 })
    if (pathname === "/jurisdictions") return reply([{ code: "EW", name: "England and Wales", citationStyle: "oscola" }])
    if (pathname === "/clients") return reply([])
    if (pathname === "/playbooks") return reply([])
    if (pathname === "/workflows") return reply({ dir: "/fixture/workflows", workflows: [] })
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/engine/config") return reply({ model: "fixture/primary", small_model: "fixture/fast" })
    const provider = { id: "fixture", name: "Fixture Models", integrationID: "fixture", activation: "enabled", package: "@opencode/ai/providers/openai-compatible" }
    if (pathname === "/api/dochaus/provider-catalog" || pathname === "/api/provider") return reply({ location: { directory: "/fixture" }, data: [provider] })
    if (pathname === "/api/model") return reply({ location: { directory: "/fixture" }, data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
    if (pathname === "/api/integration") return reply({ location: { directory: "/fixture" }, data: [{ id: "fixture", name: "Fixture", methods: [], connections: [{ type: "credential", id: "fixture-account", label: "Fixture", method: "key" }] }] })
    if (pathname === "/api/session/active") return reply({ data: {} })
    if (pathname === "/api/session") return reply({ data: [], cursor: { previous: null, next: null } })
    if (pathname === "/api/permission/request") return reply({ location: { directory: matter.dir }, data: [] })
    if (pathname === "/api/agent") return reply({ location: { directory: matter.dir }, data: ["legal-review", "draft"].map(id => ({ id, description: "Synthetic fixture assistant", mode: "primary", hidden: false })) })
    if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: `data: ${JSON.stringify({ id: "evt_fixture", type: "server.connected", data: {} })}\n\n` })
    state.unexpected.push(`${method} ${pathname}`)
    return route.abort()
  })
  const page = await context.newPage()
  page.setDefaultTimeout(10000)
  page.on("pageerror", (error: Error) => state.errors.push(error.message))
  page.on("console", (message) => { if (message.type() === "error" && !message.text().startsWith("Failed to load resource")) state.errors.push(`console: ${message.text()}`) })
  return { page, context }
}

let shot = 0
async function capture(page: Page, name: string) {
  await page.evaluate(async () => { await document.fonts.ready })
  const metrics = await page.evaluate(() => ({ overflowX: document.documentElement.scrollWidth - document.documentElement.clientWidth }))
  expect(metrics.overflowX, `${name}: document horizontal overflow`).toBeLessThanOrEqual(1)
  await page.screenshot({ path: path.join(evidence!, `${String(++shot).padStart(2, "0")}-${name}.png`) })
}
// A failing scenario keeps what the page showed and what it asked for.
async function failed(page: Page, state: Fixture, error: unknown) {
  await page.screenshot({ path: path.join(evidence!, `failed-${Date.now()}.png`) }).catch(() => {})
  console.log("unexpected:", state.unexpected, "errors:", state.errors, String(error).slice(0, 300))
}
async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); console.log(`FAIL ${name}: ${error instanceof Error ? error.message : String(error)}`) }
}
// Upload bytes are only read by the stubbed route, so a plain marker says which answer to give.
const upload = (marker: string, name = NAME) => ({ name, mimeType: `${CT}.document`, buffer: Buffer.from(marker) })
async function openDocument(page: Page, name: string) {
  await page.goto(`${baseURL}/matter/${matter.id}`)
  await page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ }).click()
  await page.locator("#documents-drawer").getByRole("button", { name, exact: true }).click()
  return page.getByRole("dialog", { name, exact: true })
}
// The reader's versions list and one of its rows.
function versionsOf(page: Page, reader: ReturnType<Page["getByRole"]>, name: string) {
  const list = reader.getByRole("list", { name: `Versions of ${name}, newest first` })
  return { list, row: (number: number) => list.getByRole("listitem").filter({ has: page.getByText(`Version ${number}`, { exact: true }) }) }
}

try {
  for (const [width, height] of [[1440, 1000], [400, 860]] as const) {
    const size = `${width}`
    await scenario(`versions list, view, restore and blackline (${size})`, async () => {
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        const reader = await openDocument(page, NAME)
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toContainText("forty-five days", { timeout: 30_000 })
        const toggle = reader.getByRole("button", { name: /^Versions/ })
        await expect(toggle).toHaveAttribute("aria-expanded", "false")
        // A closed list is not in the page, so the toggle points at nothing.
        await expect(toggle).not.toHaveAttribute("aria-controls")
        await toggle.click()
        await expect(toggle).toHaveAttribute("aria-controls", "document-versions")
        const { list, row } = versionsOf(page, reader, NAME)
        await expect(list.getByRole("listitem")).toHaveCount(4)
        // On a narrow screen the list sits above the page, and comes first in the tab order too.
        await expect(reader.locator(".viewer-body > *").first()).toHaveClass(width <= 720 ? /versions-panel/ : /viewer-doc/)
        await expect(row(4)).toContainText("Current")
        await expect(row(4)).toContainText("Changes accepted")
        await expect(row(3)).toContainText("Saved from Word · by Pat Reviewer")
        await expect(row(3)).toContainText("Counterparty markup")
        await expect(row(2)).toContainText("Redacted")
        await expect(row(2)).not.toContainText("agent:")
        await expect(row(1)).toContainText("Removed when the document was redacted.")
        await expect(row(1).getByRole("button")).toHaveCount(0)
        await expect(row(4).getByRole("button", { name: "Restore version 4" })).toHaveCount(0)
        await capture(page, `version-list-${size}`)

        // Viewing a version hands focus to the way back, and back to that version's View.
        await row(3).getByRole("button", { name: "View version 3" }).click()
        await expect(reader.locator(".version-banner")).toContainText("Version 3")
        const back = reader.getByRole("button", { name: "Back to current version" })
        await expect(back).toBeFocused()
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toContainText("sixty days")
        await expect(reader.getByRole("button", { name: "Request edits" })).toBeDisabled()
        await expect(reader.getByTitle("Download version 3 exactly as kept")).toBeVisible()
        await capture(page, `version-view-${size}`)
        await back.press("Enter")
        await expect(reader.locator(".version-banner")).toHaveCount(0)
        await expect(row(3).getByRole("button", { name: "View version 3" })).toBeFocused()
        await row(3).getByRole("button", { name: "View version 3" }).press("Enter")
        await expect(back).toBeFocused()

        await row(3).getByRole("button", { name: "Restore version 3" }).click()
        const confirm = row(3).getByRole("group", { name: "Restore version 3" })
        await expect(confirm).toContainText("This saves a copy of version 3 as version 5, the new current version. Nothing is deleted: every version stays in the history.")
        const confirmButton = confirm.getByRole("button", { name: "Restore version 3" })
        await expect(confirmButton).toBeFocused()
        await expect(confirmButton).toHaveAccessibleDescription("This saves a copy of version 3 as version 5, the new current version. Nothing is deleted: every version stays in the history.")
        await capture(page, `restore-confirm-${size}`)
        // Escape cancels the confirmation only, never the viewer.
        await page.keyboard.press("Escape")
        await expect(confirm).toHaveCount(0)
        await expect(reader).toBeVisible()
        await expect(row(3).getByRole("button", { name: "Restore version 3" })).toBeFocused()
        expect(state.restores).toEqual([])
        await row(3).getByRole("button", { name: "Restore version 3" }).click()
        await confirm.getByRole("button", { name: "Restore version 3" }).press("Enter")
        await expect(list.getByRole("listitem")).toHaveCount(5)
        expect(state.restores).toEqual([{ path: `/matters/${matter.id}/documents/versions/v-3/restore`, body: {} }])
        // Focus stays in the list, on its heading, while the pressed button goes.
        await expect(reader.locator("#document-versions-title")).toBeFocused()
        await expect(row(5)).toContainText("Current")
        await expect(row(5)).toContainText("Restored from version 3")
        // Version 3 is what the file now holds: restoring it again would change nothing.
        await expect(row(3).getByRole("button", { name: "Restore version 3" })).toHaveCount(0)
        await expect(row(4).getByRole("button", { name: "Restore version 4" })).toBeEnabled()
        await expect(reader.locator(".version-banner")).toHaveCount(0)
        await expect(page.locator(".toast")).toContainText("Restored version 3 as version 5.")
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toContainText("sixty days")
        await capture(page, `restored-${size}`)

        await row(5).getByRole("button", { name: "Compare with version 4" }).click()
        const blackline = page.getByRole("dialog", { name: /^Blackline - / })
        await expect(blackline.locator(".compare-sides-line, .viewer-title-meta").first()).toBeVisible()
        await expect(blackline.locator(".viewer-title-meta")).toContainText("Services agreement (version 4) → Services agreement (version 5)")
        expect(new URLSearchParams(state.compares.at(-1)!).get("leftVersion")).toBe("v-4")
        expect(new URLSearchParams(state.compares.at(-1)!).get("rightVersion")).toBe("v-5")
        await capture(page, `version-blackline-${size}`)
        await blackline.getByRole("button", { name: "Close", exact: true }).click()
        // Back where the lawyer was: the document with its versions open.
        await expect(reader).toBeVisible()
        await expect(reader.getByRole("button", { name: /^Versions/ })).toHaveAttribute("aria-expanded", "true")
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`an older version open offers no proposals, template or compare (${size})`, async () => {
      const state = fixture({ redlines: [proposal] })
      const { page, context } = await open(state, width, height)
      try {
        const reader = await openDocument(page, NAME)
        await expect(reader.getByRole("button", { name: "Accept", exact: true })).toBeVisible({ timeout: 30_000 })
        const toggle = reader.getByRole("button", { name: /^Versions/ })
        await toggle.click()
        const { row } = versionsOf(page, reader, NAME)
        await row(3).getByRole("button", { name: "View version 3" }).click()
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toContainText("sixty days")
        await toggle.click()
        // The proposals are to the current version; none is offered over an older one.
        await expect(reader.getByRole("button", { name: "Accept", exact: true })).toHaveCount(0)
        await expect(reader.getByRole("button", { name: "Request edits" })).toBeDisabled()
        await reader.locator("summary[aria-label='More actions']").click()
        await expect(reader.getByRole("button", { name: /^Compare with…/ })).toBeDisabled()
        await expect(reader.getByRole("button", { name: /^Create reusable template/ })).toBeDisabled()
        await capture(page, `old-version-actions-${size}`)
        await page.keyboard.press("Escape")
        await reader.getByRole("button", { name: "Back to current version" }).click()
        await expect(reader.getByRole("button", { name: "Accept", exact: true })).toBeVisible()
        await reader.locator("summary[aria-label='More actions']").click()
        await expect(reader.getByRole("button", { name: /^Compare with…/ })).toBeEnabled()
        await expect(reader.getByRole("button", { name: /^Create reusable template/ })).toBeEnabled()
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`a file changed outside doc.haus (${size})`, async () => {
      const state = fixture({ changed: true, slowAfterRestore: 1500 })
      const { page, context } = await open(state, width, height)
      try {
        const reader = await openDocument(page, NAME)
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
        await reader.getByRole("button", { name: /^Versions/ }).click()
        const { list, row } = versionsOf(page, reader, NAME)
        await expect(reader.locator(".versions-panel")).toContainText("The file has changed outside doc.haus since the current version.")
        await expect(row(4).getByRole("button", { name: "Restore version 4" })).toBeVisible()
        await row(3).getByRole("button", { name: "Restore version 3" }).click()
        // doc.haus keeps the outside changes as a version first, so the new number is not foretold.
        const words = "This keeps the file's changes made outside doc.haus as a version, then saves a copy of version 3 as the new current version. Nothing is deleted: every version stays in the history."
        const confirm = row(3).getByRole("group", { name: "Restore version 3" })
        await expect(confirm).toContainText(words)
        await expect(confirm.getByRole("button", { name: "Restore version 3" })).toHaveAccessibleDescription(words)
        await capture(page, `changed-restore-confirm-${size}`)
        await confirm.getByRole("button", { name: "Restore version 3" }).click()
        await expect(page.locator(".toast")).toContainText("Restored version 3 as version 6.")
        // Until the list has caught up, its rows cannot start another restore from stale numbers.
        await expect(list.getByRole("listitem")).toHaveCount(4)
        await expect(row(2).getByRole("button", { name: "Restore version 2" })).toBeDisabled()
        await expect(reader.locator("#document-versions-title")).toBeFocused()
        await expect(list.getByRole("listitem")).toHaveCount(6)
        await expect(row(6)).toContainText("Restored from version 3")
        await expect(row(5)).toContainText("Changed outside doc.haus")
        await expect(row(2).getByRole("button", { name: "Restore version 2" })).toBeEnabled()
        await expect(reader.locator(".versions-panel")).not.toContainText("The file has changed outside doc.haus")
        await capture(page, `changed-restored-${size}`)
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`the list follows versions made elsewhere (${size})`, async () => {
      // Another session's changes are made at the step they belong to, not
      // counted by versions requests: a development build reads the list twice
      // on opening (React StrictMode), a production build once.
      const saved = (state: Fixture) => {
        const list = state.docs[NAME]!
        state.docs[NAME] = [version(5, { origin: "revision", actor: "Sam Partner", current: true }), ...list.map((v) => ({ ...v, current: false }))]
        bytes["v-5"] = bytes["v-4"]!
      }
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        const reader = await openDocument(page, NAME)
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
        // The viewer has read the four versions there were when it opened.
        const toggle = reader.getByRole("button", { name: /^Versions/ })
        await expect(toggle).toHaveText("Versions 4")
        // Another session saves a version after the viewer opened: opening the list shows it.
        saved(state)
        await toggle.click()
        const { list, row } = versionsOf(page, reader, NAME)
        await expect(list.getByRole("listitem")).toHaveCount(5)
        await expect(row(5)).toContainText("Saved from Word · by Sam Partner")
        const restoreThree = row(3).getByRole("button", { name: "Restore version 3" })
        await expect(restoreThree).toBeEnabled()
        // Then another restores version 3, unseen here: the refusal brings the list up to date.
        restoredElsewhere(state)
        await restoreThree.click()
        await row(3).getByRole("group", { name: "Restore version 3" }).getByRole("button", { name: "Restore version 3" }).click()
        await expect(list.getByRole("listitem")).toHaveCount(6)
        await expect(row(3)).toContainText(`${NAME} already holds version 3.`)
        await expect(row(3).getByRole("button", { name: "Restore version 3" })).toHaveCount(0)
        await expect(row(6)).toContainText("Restored from version 3")
        await expect(reader.locator("#document-versions-title")).toBeFocused()
        await capture(page, `already-current-${size}`)
        expect(state.restores).toHaveLength(1)
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`PDF versions (${size})`, async () => {
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        const reader = await openDocument(page, PDF)
        await reader.getByRole("button", { name: /^Versions/ }).click()
        const { list, row } = versionsOf(page, reader, PDF)
        await expect(list.getByRole("listitem")).toHaveCount(2)
        await expect(row(1).getByRole("button")).toHaveText(["View", "Download", "Restore"])
        await row(1).getByRole("button", { name: "View version 1" }).click()
        await expect(reader.locator("iframe.pdf-render")).toHaveAttribute("src", /^blob:/)
        await expect(reader.getByRole("button", { name: "Convert to DOCX" })).toBeDisabled()
        await expect(reader.getByRole("button", { name: "Back to current version" })).toBeFocused()
        // Only the screenshot needs the browser's PDF plugin to have painted the page.
        await page.waitForTimeout(2000)
        await capture(page, `pdf-version-view-${size}`)
        await reader.getByRole("button", { name: "Back to current version" }).press("Enter")
        await expect(reader.locator("iframe.pdf-render")).toHaveAttribute("src", /\/documents\/content\?/)
        await expect(row(1).getByRole("button", { name: "View version 1" })).toBeFocused()
        await expect(reader.getByRole("button", { name: "Convert to DOCX" })).toBeEnabled()
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`email versions beside its text (${size})`, async () => {
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        const reader = await openDocument(page, EMAIL)
        await expect(reader.locator("pre.document-text")).toContainText("Please send the agreement")
        await reader.getByRole("button", { name: /^Versions/ }).click()
        const { list, row } = versionsOf(page, reader, EMAIL)
        await expect(list.getByRole("listitem")).toHaveCount(2)
        // Only the current text can be shown here: an earlier email is downloaded or restored.
        await expect(row(2).getByRole("button")).toHaveText(["Download"])
        await expect(row(1).getByRole("button")).toHaveText(["Download", "Restore"])
        // Opening the list reads it again; its rows wait for that.
        await expect(row(1).getByRole("button", { name: "Restore version 1" })).toBeEnabled()
        await capture(page, `email-versions-${size}`)
        const read = state.texts
        await row(1).getByRole("button", { name: "Restore version 1" }).click()
        await row(1).getByRole("group", { name: "Restore version 1" }).getByRole("button", { name: "Restore version 1" }).click()
        await expect(list.getByRole("listitem")).toHaveCount(3)
        expect(state.restores).toEqual([{ path: `/matters/${matter.id}/documents/versions/e-1/restore`, body: {} }])
        await expect(page.locator(".toast")).toContainText("Restored version 1 as version 3.")
        await expect(row(3)).toContainText("Restored from version 1")
        await expect(reader.locator("#document-versions-title")).toBeFocused()
        // The text is read again for the restored version.
        await expect.poll(() => state.texts).toBeGreaterThan(read)
        await capture(page, `email-restored-${size}`)
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`compare dialog defaults to the previous version (${size})`, async () => {
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        await page.goto(`${baseURL}/matter/${matter.id}`)
        await page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ }).click()
        await page.locator("#documents-drawer").getByRole("button", { name: `Compare ${NAME} with another version` }).click()
        const dialog = page.getByRole("dialog", { name: "Compare versions" })
        const other = dialog.getByRole("combobox", { name: "Other version" })
        await expect(other).toHaveValue("/version/v-3")
        // The previous version is the earlier side; this file the later.
        await expect(dialog.locator(".compare-side").first()).toContainText("Earlier version")
        await expect(dialog.locator("#compare-left-label")).toHaveValue("Services agreement (version 3)")
        await expect(dialog.locator("#compare-right-label")).toHaveValue("Services agreement (version 4)")
        // Every version is offered; one whose copy was removed is listed, not chosen.
        const options = await other.locator("option").evaluateAll((nodes) => nodes.map((node) => ({ text: node.textContent, title: node.getAttribute("title"), disabled: (node as HTMLOptionElement).disabled, group: node.parentElement?.getAttribute("label") ?? null })))
        expect(options.filter((option) => option.group === "Versions of this file").map((option) => [option.text?.split(" · ")[0], option.disabled])).toEqual([["Version 3", false], ["Version 2", false], ["Version 1", true]])
        // The closed picker fits a phone: the year is left to the option's title.
        expect(options.find((option) => option.text?.startsWith("Version 3"))).toMatchObject({ text: "Version 3 · Saved from Word · 24 Oct", title: expect.stringContaining("24 Oct 2026") })
        expect(options.filter((option) => option.group === "Other Word files").map((option) => option.text)).toEqual([OTHER, LETTER])
        await capture(page, `compare-picker-${size}`)
        await dialog.getByRole("button", { name: "Compare", exact: true }).click()
        await expect(page.getByRole("dialog", { name: /^Blackline - / })).toBeVisible()
        const query = new URLSearchParams(state.compares.at(-1)!)
        expect([query.get("left"), query.get("leftVersion"), query.get("right"), query.get("rightVersion")]).toEqual([NAME, "v-3", NAME, null])
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`compare dialog: duplicate uploads and a file without versions (${size})`, async () => {
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        await page.goto(`${baseURL}/matter/${matter.id}`)
        await page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ }).click()
        const drawer = page.locator("#documents-drawer")
        await drawer.getByRole("button", { name: `Compare ${NAME} with another version` }).click()
        const dialog = page.getByRole("dialog", { name: "Compare versions" })
        const other = dialog.getByRole("combobox", { name: "Other version" })
        await expect(other).toHaveValue("/version/v-3")
        // Identical to an earlier version of another file: the picker cannot offer that version, so nothing is picked for it.
        await dialog.locator("input[type=file]").setInputFiles(upload("SIDE LETTER V2", OTHER))
        await expect(dialog.locator(".compare-notice")).toHaveText(`Identical to version 2 of ${OTHER}. Nothing was saved.`)
        await expect(other).toHaveValue("/version/v-3")
        await capture(page, `compare-duplicate-earlier-${size}`)
        // Identical to that file's current version: the file is picked.
        await dialog.locator("input[type=file]").setInputFiles(upload("SIDE LETTER V3", OTHER))
        await expect(dialog.locator(".compare-notice")).toHaveText(`Identical to the current version (version 3) of ${OTHER}. Nothing was saved. It is selected.`)
        await expect(other).toHaveValue(OTHER)
        await dialog.getByRole("button", { name: "Compare", exact: true }).click()
        await expect(page.getByRole("dialog", { name: /^Blackline - / })).toBeVisible()
        const query = new URLSearchParams(state.compares.at(-1)!)
        expect([query.get("left"), query.get("leftVersion"), query.get("right")]).toEqual([OTHER, null, NAME])
        await page.getByRole("dialog", { name: /^Blackline - / }).getByRole("button", { name: "Close", exact: true }).click()

        // A file with no kept versions offers the matter's other Word files, as before versions were kept.
        await drawer.getByRole("button", { name: `Compare ${LETTER} with another version` }).click()
        const plain = page.getByRole("dialog", { name: "Compare versions" })
        const pick = plain.getByRole("combobox", { name: "Other version" })
        await expect(pick.locator("option").first()).toHaveText("Choose a version or file…")
        // Its (empty) versions list has answered, and the dialog has settled on files alone.
        await expect.poll(() => state.asked.includes(LETTER)).toBe(true)
        await page.waitForTimeout(300)
        await expect(pick.locator("optgroup")).toHaveCount(0)
        await expect(plain.getByRole("combobox", { name: "Version of this file" })).toHaveCount(0)
        expect(await pick.locator("option:not([disabled])").allTextContents()).toEqual([NAME, OTHER])
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`upload prompt: new version, duplicate and tracked changes (${size})`, async () => {
      const state = fixture({ restoreFails: true })
      const { page, context } = await open(state, width, height)
      try {
        await page.goto(`${baseURL}/matter/${matter.id}`)
        await page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ }).click()
        const drawer = page.locator("#documents-drawer")
        const input = drawer.locator("input[type=file]")
        await input.setInputFiles(upload("IDENTICAL TO VERSION 3"))
        const rows = drawer.getByRole("list", { name: "Uploads" }).getByRole("listitem")
        await expect(rows.first()).toContainText("Identical to version 3 of Services agreement.docx. Nothing was saved.")
        await expect(rows.first().getByRole("button")).toHaveText(["Restore version 3", "Dismiss"])
        await capture(page, `upload-duplicate-${size}`)
        // An earlier version can be made current again, after the same confirmation.
        await rows.first().getByRole("button", { name: "Restore version 3" }).click()
        const confirm = rows.first().getByRole("group", { name: "Restore version 3" })
        await expect(rows.first()).toContainText("This saves a copy of version 3 as the new current version. Nothing is deleted: every version stays in the history.")
        await expect(confirm.getByRole("button", { name: "Restore version 3" })).toBeFocused()
        await capture(page, `upload-duplicate-restore-confirm-${size}`)
        await page.keyboard.press("Escape")
        await expect(rows.first().getByRole("button", { name: "Restore version 3" })).toBeFocused()
        expect(state.restores).toEqual([])
        // A restore the service could not make says why, and focus stays in the row.
        await rows.first().getByRole("button", { name: "Restore version 3" }).click()
        await confirm.getByRole("button", { name: "Restore version 3" }).press("Enter")
        await expect(rows.first()).toContainText("The document service could not save the file. Try again in a moment.")
        await expect(rows.first().getByRole("button", { name: "Dismiss" })).toBeFocused()
        await rows.first().getByRole("button", { name: "Dismiss" }).click()
        await input.setInputFiles(upload("IDENTICAL TO VERSION 3"))
        await rows.first().getByRole("button", { name: "Restore version 3" }).click()
        await confirm.getByRole("button", { name: "Restore version 3" }).press("Enter")
        await expect(rows).toHaveCount(0)
        await expect(page.locator(".toast").filter({ hasText: "Restored version 3 of Services agreement.docx as version 5." })).toBeVisible()
        // The row goes; focus moves to the documents list, not the page.
        await expect(drawer.getByRole("region", { name: "Documents", exact: true })).toBeFocused()
        expect(state.restores).toEqual([1, 2].map(() => ({ path: `/matters/${matter.id}/documents/versions/v-3/restore`, body: {} })))

        await input.setInputFiles(upload("Their reply"))
        await expect(rows.first()).toContainText("Already in this matter. Save this as a new version of it, or keep both?")
        await expect(rows.first().getByRole("button")).toHaveText(["Save as new version", "Keep both", "Cancel"])
        await capture(page, `upload-save-as-new-version-${size}`)
        await rows.first().getByRole("button", { name: "Save as new version" }).click()
        await expect(rows.first()).toContainText(NOTICE)
        await expect(page.locator(".toast").filter({ hasText: `Services agreement.docx is ready (4 sections). ${NOTICE}` })).toBeVisible()
        await capture(page, `upload-tracked-notice-${size}`)
        expect(state.uploads).toEqual([{ existing: null, duplicate: true }, { existing: null, duplicate: true }, { existing: null, duplicate: false }, { existing: "replace", duplicate: false }])
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })
  }
} finally {
  await writeFile(path.join(evidence, "document-versions-results.json"), JSON.stringify({ results }, null, 2) + "\n")
  await browser.close()
}
expect(results.filter((result) => result.status === "failed")).toEqual([])
console.log(`PASS ${results.length} scenarios; ${shot} screenshots.`)
