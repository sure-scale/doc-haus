// Document versions (#268) against an already-running web preview. Every API
// request is intercepted: a synthetic Word document with four kept versions,
// its version bytes, restore, compare and upload answers are fixtures, and
// generation, credentials and unknown requests are refused. It checks the
// versions list, viewing an older version, the in-page restore confirmation
// and its request, a version blackline and the return to the list, the compare
// dialog's previous-version default, and the upload prompt's new-version,
// duplicate and tracked-changes messages, at 1440×1000 and 400×860.
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
const documents = [
  { id: 1, name: NAME, doc_path: `${matter.dir}/${NAME}`, created_at: timestamp, role: "source" },
  { id: 2, name: "Client instructions.eml", doc_path: `${matter.dir}/Client instructions.eml`, created_at: timestamp - 60_000, role: "source" },
]
const body = (clause: string) => ["Services agreement", "1. The Supplier shall provide the services described in the schedule.", clause, "3. Either party may end this agreement on thirty days notice."]
const bytes: Record<string, Uint8Array> = {
  "v-2": docx(body("2. Invoices are payable within [REDACTED] days of receipt.")),
  "v-3": docx(body("2. Invoices are payable within sixty days of receipt.")),
  "v-4": docx(body("2. Invoices are payable within forty-five days of receipt.")),
}
const version = (number: number, patch: Record<string, unknown>) => ({
  id: `v-${number}`, number, sha256: String(number).repeat(64).slice(0, 64), ext: ".docx", bytes: 12_000, parent: number > 1 ? `v-${number - 1}` : null,
  derived_from: null, actor: "unknown", run: null, label: null, created_at: timestamp + number * 3_600_000, current: false, ...patch,
})
const initialVersions = () => [
  version(4, { origin: "accept", current: true }),
  version(3, { origin: "revision", actor: "Pat Reviewer", label: "Counterparty markup" }),
  version(2, { origin: "redaction", actor: "agent:legal-review" }),
  version(1, { origin: "upload", removed: "redaction" }),
]
const comparisonFor = (url: URL) => {
  const side = (key: "left" | "right") => {
    const id = url.searchParams.get(`${key}Version`)
    const number = id ? Number(id.slice(2)) : 4
    return { name: NAME, label: `Services agreement (version ${number})`, trackedChanges: { count: 0, authors: [] }, pendingProposals: 0, ...(id ? { version: { id, number } } : {}) }
  }
  const left = side("left"), right = side("right")
  const hunk = { kind: "modified", location: { scope: "body", paragraph: 3, side: "right", number: "2" }, before: "Invoices are payable within sixty days of receipt.", after: "Invoices are payable within forty-five days of receipt.", untracked: true }
  return { left, right, options: { ignoreFormatting: true, detectMoves: true, headersFooters: true }, attribution: "authors", changes: [{ id: "C1", ...hunk }], untracked: [{ id: "U1", ...hunk }], blacklineName: `Blackline - ${left.label} vs ${right.label}.docx` }
}
const NOTICE = "Saved as version 6, with 2 tracked changes by Pat Reviewer; resolve them in Word before doc.haus edits it."

function fixture() {
  return {
    versions: initialVersions() as ReturnType<typeof version>[],
    events: [] as unknown[],
    restores: [] as { path: string; body: unknown }[],
    compares: [] as string[],
    uploads: [] as { existing: string | null; duplicate: boolean }[],
    unexpected: [] as string[], errors: [] as string[],
  }
}

async function open(state: ReturnType<typeof fixture>, width: number, height: number) {
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", timezoneId: "Europe/London" })
  await context.addInitScript(() => { sessionStorage.setItem("dh.launch-notice.v1", "accepted"); localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1") })
  await context.route("**/*", async (route: Route) => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    const method = request.method()
    // The app and its Word runtime come from the preview; every API call goes
    // through the /ingest and /engine prefixes; nothing else may load.
    const api = url.origin === origin && (url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/"))
    if (!api && url.origin === origin) return route.continue()
    if (!api) { state.unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    if (method === "OPTIONS") return route.fulfill({ status: 204, headers })
    const reply = (value: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(value) })
    const word = (value: Uint8Array) => route.fulfill({ status: 200, headers, contentType: `${"application/vnd.openxmlformats-officedocument.wordprocessingml"}.document`, body: Buffer.from(value) })
    const prefix = `/matters/${matter.id}`
    if (method === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (method === "POST" && pathname === `${prefix}/documents`) {
      const form = request.postDataBuffer()?.toString("latin1") ?? ""
      const existing = /name="existing"\r\n\r\n([^\r]+)/.exec(form)?.[1] ?? null
      const duplicate = form.includes("IDENTICAL TO VERSION 3")
      state.uploads.push({ existing, duplicate })
      if (duplicate) return reply({ error: `${NAME} is already in this matter as version 3 of ${NAME}. Nothing new was stored.`, code: "duplicate", name: NAME, document_id: "doc-1", version: { id: "v-3", number: 3, current: false } }, 409)
      if (!existing) return reply({ error: `${NAME} is already in this matter.`, code: "source_exists", name: NAME }, 409)
      return reply({ name: NAME, docPath: documents[0]!.doc_path, sections: 4, chunks: 4, injection: null, quality: { status: "ok" }, version: { id: "v-6", number: 6, created: true }, tracked_changes: { count: 2, authors: ["Pat Reviewer"] }, notice: NOTICE })
    }
    const restore = /^\/matters\/[^/]+\/documents\/versions\/([^/]+)\/restore$/.exec(pathname)
    if (method === "POST" && restore) {
      state.restores.push({ path: pathname, body: request.postDataJSON() })
      const from = state.versions.find((v) => v.id === restore[1])!
      const next = Math.max(...state.versions.map((v) => v.number)) + 1
      const current = state.versions.find((v) => v.current)!
      state.versions = [version(next, { origin: "restore", current: true, parent: current.id, label: `Restored from version ${from.number}` }), ...state.versions.map((v) => ({ ...v, current: false }))]
      state.events.push({ id: state.events.length + 1, kind: "restore", version_id: `v-${next}`, actor: "unknown", detail: { restored_from: from.id, restored_number: from.number }, created_at: timestamp })
      bytes[`v-${next}`] = bytes[from.id]!
      return reply({ name: NAME, docPath: documents[0]!.doc_path, version: { id: `v-${next}`, number: next, created: true }, restored_from: { id: from.id, number: from.number } })
    }
    if (method !== "GET") { state.unexpected.push(`${method} ${pathname}`); return route.abort() }
    const current = () => bytes[state.versions.find((v) => v.current)!.id]!
    if (pathname === `${prefix}/documents/versions`) {
      const held = state.versions.find((v) => v.current)!
      return reply({ name: NAME, document: { id: "doc-1", name: NAME, current_version: held.id }, versions: state.versions, events: state.events, working_version: held.id })
    }
    const content = /^\/matters\/[^/]+\/documents\/versions\/([^/]+)\/content$/.exec(pathname)
    if (content) return bytes[content[1]!] ? word(bytes[content[1]!]!) : reply({ error: "The kept bytes of this version are missing or damaged.", code: "version_damaged" }, 409)
    if (pathname === `${prefix}/documents/redlined` || pathname === `${prefix}/documents/content`) return word(current())
    if (pathname === `${prefix}/redlines`) return reply([])
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
async function failed(page: Page, state: ReturnType<typeof fixture>, error: unknown) {
  await page.screenshot({ path: path.join(evidence!, `failed-${Date.now()}.png`) }).catch(() => {})
  console.log("unexpected:", state.unexpected, "errors:", state.errors, String(error).slice(0, 300))
}
async function scenario(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) { results.push({ name, status: "failed", error: error instanceof Error ? error.message : String(error) }); console.log(`FAIL ${name}: ${error instanceof Error ? error.message : String(error)}`) }
}
// Upload bytes are only read by the stubbed route, so a plain marker says which answer to give.
const upload = (marker: string) => ({ name: NAME, mimeType: `${CT}.document`, buffer: Buffer.from(marker) })

try {
  for (const [width, height] of [[1440, 1000], [400, 860]] as const) {
    const size = `${width}`
    await scenario(`versions list, view, restore and blackline (${size})`, async () => {
      const state = fixture()
      const { page, context } = await open(state, width, height)
      try {
        await page.goto(`${baseURL}/matter/${matter.id}`)
        await page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ }).click()
        await page.locator("#documents-drawer").getByRole("button", { name: NAME, exact: true }).click()
        const reader = page.getByRole("dialog", { name: NAME, exact: true })
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toContainText("forty-five days", { timeout: 30_000 })
        const toggle = reader.getByRole("button", { name: /^Versions/ })
        await expect(toggle).toHaveAttribute("aria-expanded", "false")
        await toggle.click()
        const list = reader.getByRole("list", { name: `Versions of ${NAME}, newest first` })
        await expect(list.getByRole("listitem")).toHaveCount(4)
        const row = (number: number) => list.getByRole("listitem").filter({ has: page.getByText(`Version ${number}`, { exact: true }) })
        await expect(row(4)).toContainText("Current")
        await expect(row(4)).toContainText("Changes accepted")
        await expect(row(3)).toContainText("Saved from Word · by Pat Reviewer")
        await expect(row(3)).toContainText("Counterparty markup")
        await expect(row(2)).toContainText("Redacted")
        await expect(row(2)).not.toContainText("agent:")
        await expect(row(1)).toContainText("Removed when the document was redacted.")
        await expect(row(1).getByRole("button")).toHaveCount(0)
        await expect(row(4).getByRole("button", { name: "Restore", exact: true })).toHaveCount(0)
        await capture(page, `version-list-${size}`)

        await row(3).getByRole("button", { name: "View", exact: true }).click()
        await expect(reader.locator(".version-banner")).toContainText("Version 3")
        await expect(reader.locator(".docx-render:not(.docx-skeleton)")).toContainText("sixty days")
        await expect(reader.getByRole("button", { name: "Request edits" })).toBeDisabled()
        await expect(reader.getByTitle("Download version 3 exactly as kept")).toBeVisible()
        await capture(page, `version-view-${size}`)

        await row(3).getByRole("button", { name: "Restore", exact: true }).click()
        const confirm = row(3).getByRole("group", { name: "Restore version 3" })
        await expect(confirm).toContainText("This saves a copy of version 3 as version 5, the new current version. Nothing is deleted: every version stays in the history.")
        await expect(confirm.getByRole("button", { name: "Restore version 3" })).toBeFocused()
        await capture(page, `restore-confirm-${size}`)
        // Escape cancels the confirmation only, never the viewer.
        await page.keyboard.press("Escape")
        await expect(confirm).toHaveCount(0)
        await expect(reader).toBeVisible()
        await expect(row(3).getByRole("button", { name: "Restore", exact: true })).toBeFocused()
        expect(state.restores).toEqual([])
        await row(3).getByRole("button", { name: "Restore", exact: true }).click()
        await row(3).getByRole("button", { name: "Restore version 3" }).click()
        await expect(list.getByRole("listitem")).toHaveCount(5)
        expect(state.restores).toEqual([{ path: `/matters/${matter.id}/documents/versions/v-3/restore`, body: {} }])
        await expect(row(5)).toContainText("Current")
        await expect(row(5)).toContainText("Restored from version 3")
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
        const options = await other.locator("option").evaluateAll((nodes) => nodes.map((node) => ({ text: node.textContent, disabled: (node as HTMLOptionElement).disabled, group: node.parentElement?.getAttribute("label") ?? null })))
        expect(options.filter((option) => option.group === "Versions of this file").map((option) => [option.text?.split(" · ")[0], option.disabled])).toEqual([["Version 3", false], ["Version 2", false], ["Version 1", true]])
        expect(options.filter((option) => option.group === "Other Word files")).toEqual([])
        await capture(page, `compare-picker-${size}`)
        await dialog.getByRole("button", { name: "Compare", exact: true }).click()
        await expect(page.getByRole("dialog", { name: /^Blackline - / })).toBeVisible()
        const query = new URLSearchParams(state.compares.at(-1)!)
        expect([query.get("left"), query.get("leftVersion"), query.get("right"), query.get("rightVersion")]).toEqual([NAME, "v-3", NAME, null])
        expect(state.unexpected).toEqual([]); expect(state.errors).toEqual([])
      } catch (error) { await failed(page, state, error); throw error } finally { await context.close() }
    })

    await scenario(`upload prompt: new version, duplicate and tracked changes (${size})`, async () => {
      const state = fixture()
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
        await rows.first().getByRole("button", { name: "Restore version 3" }).click()
        await confirm.getByRole("button", { name: "Restore version 3" }).click()
        await expect(rows).toHaveCount(0)
        await expect(page.locator(".toast").filter({ hasText: "Restored version 3 of Services agreement.docx as version 5." })).toBeVisible()
        expect(state.restores).toEqual([{ path: `/matters/${matter.id}/documents/versions/v-3/restore`, body: {} }])

        await input.setInputFiles(upload("Their reply"))
        await expect(rows.first()).toContainText("Already in this matter. Save this as a new version of it, or keep both?")
        await expect(rows.first().getByRole("button")).toHaveText(["Save as new version", "Keep both", "Cancel"])
        await capture(page, `upload-save-as-new-version-${size}`)
        await rows.first().getByRole("button", { name: "Save as new version" }).click()
        await expect(rows.first()).toContainText(NOTICE)
        await expect(page.locator(".toast").filter({ hasText: `Services agreement.docx is ready (4 sections). ${NOTICE}` })).toBeVisible()
        await capture(page, `upload-tracked-notice-${size}`)
        expect(state.uploads).toEqual([{ existing: null, duplicate: true }, { existing: null, duplicate: false }, { existing: "replace", duplicate: false }])
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
