// PDF preview acceptance against an already-running web preview (#275).
// A multi-page synthetic PDF is served by an intercepted ingest route; every
// other request is stubbed or refused. The iframe must fill the viewer body at
// desktop and phone widths and the PDF must scroll inside it, not the page.
// Needs full Chromium (the headless shell has no PDF plugin and renders blank):
// set PLAYWRIGHT_CHROMIUM_EXECUTABLE to a full Chromium binary.
import { mkdir } from "node:fs/promises"
import path from "node:path"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
type Route = import("@playwright/test").Route
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
await mkdir(evidence, { recursive: true })
const origin = new URL(baseURL).origin
const timestamp = Date.UTC(2026, 9, 21, 10)
const NAME = "Synthetic proposal.pdf"
const matter = { id: "fixture-matter", title: "Alder House · Services agreement", reference: "AH-2026-014", client: "Alder House Studio", jurisdictions: ["EW"], dir: "/fixture/matters/alder-house", created_at: timestamp }
const documents = [{ id: 1, name: NAME, doc_path: `${matter.dir}/${NAME}`, created_at: timestamp, role: "source" }]

// Minimal multi-page PDF: one Helvetica text block per US Letter page.
function syntheticPdf(pages: number) {
  const objects: string[] = []
  const kids = Array.from({ length: pages }, (_, i) => `${4 + i * 2} 0 R`).join(" ")
  objects.push("<< /Type /Catalog /Pages 2 0 R >>", `<< /Type /Pages /Kids [${kids}] /Count ${pages} >>`, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
  for (let i = 0; i < pages; i++) {
    const text = Array.from({ length: 30 }, (_, line) => `BT /F1 12 Tf 56 ${740 - line * 22} Td (Synthetic page ${i + 1}, line ${line + 1}: the supplier shall deliver the services.) Tj ET`).join("\n")
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${5 + i * 2} 0 R >>`, `<< /Length ${text.length} >>\nstream\n${text}\nendstream`)
  }
  let body = "%PDF-1.4\n"
  const offsets: number[] = []
  objects.forEach((object, i) => { offsets.push(body.length); body += `${i + 1} 0 obj\n${object}\nendobj\n` })
  const xref = body.length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map(o => `${String(o).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Buffer.from(body, "latin1")
}
const pdf = process.env.PDF_FIXTURE ? Buffer.from(await Bun.file(process.env.PDF_FIXTURE).arrayBuffer()) : syntheticPdf(8)

const browser = await chromium.launch({ headless: process.env.PLAYWRIGHT_HEADED !== "1", executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const unexpected: string[] = []
const pageErrors: string[] = []
const measurements: Record<string, unknown> = {}

async function run(label: string, width: number, height: number) {
  const context = await browser.newContext({ viewport: { width, height }, hasTouch: width <= 768, reducedMotion: "reduce", timezoneId: "Europe/London" })
  await context.addInitScript(() => { try { sessionStorage.setItem("dh.launch-notice.v1", "accepted"); localStorage.setItem("dh.sidebar", innerWidth > 960 ? "0" : "1") } catch { /* the browser's PDF viewer frame has no storage */ } })
  await context.route("**/*", async (route: Route) => {
    const request = route.request()
    const url = new URL(request.url())
    const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
    // The browser's own PDF viewer loads chrome:// resources; blocking them leaves it blank.
    if (/^chrome(-extension)?:/.test(request.url())) return route.continue()
    const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType()) || url.pathname.startsWith("/ingest/") || url.pathname.startsWith("/engine/")
    if ((!api || url.pathname.startsWith("/wasm/")) && url.origin === origin) return route.continue()
    if (!api) { unexpected.push(request.url()); return route.abort() }
    const headers = { "access-control-allow-origin": origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
    const method = request.method()
    if (method === "OPTIONS") return route.fulfill({ status: 204, headers })
    const reply = (body: unknown, status = 200) => route.fulfill({ status, headers, contentType: "application/json", body: JSON.stringify(body) })
    if (method === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (method !== "GET") { unexpected.push(`${method} ${pathname}`); return route.abort() }
    if (pathname === `/matters/${matter.id}/documents/content`) return route.fulfill({ status: 200, headers: { ...headers, "content-disposition": "inline" }, contentType: "application/pdf", body: pdf })
    if (pathname === "/auth/reuse") return reply({ sources: [] })
    if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (pathname === "/setup") return reply({ version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" })
    if (pathname === "/preferences") return reply({ attorney: "Alex Example", firm: "Example Partners LLP" })
    if (pathname === "/matters") return reply([matter])
    if (pathname === `/matters/${matter.id}`) return reply({ ...matter, documents })
    if (/^\/matters\/[^/]+\/drafting-status$/.test(pathname)) return reply(null)
    if (/^\/matters\/[^/]+\/telemetry$/.test(pathname)) return reply({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 0, since_ts: timestamp }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (pathname === "/jurisdictions") return reply([{ code: "EW", name: "England and Wales", citationStyle: "oscola" }])
    if (pathname === "/clients" || pathname === "/api/session/active") return reply(pathname === "/clients" ? [] : { data: {} })
    if (pathname === "/engine/config") return reply({ model: "fixture/primary", small_model: "fixture/fast" })
    if (pathname === "/playbooks") return reply([])
    if (pathname === "/workflows") return reply({ dir: "/fixture/workflows", workflows: [] })
    if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    if (pathname === "/api/session") return reply({ data: [], cursor: { previous: null, next: null } })
    if (pathname === "/api/agent") return reply({ location: { directory: matter.dir }, data: [] })
    if (pathname === "/api/permission/request") return reply({ location: { directory: matter.dir }, data: [] })
    if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: `data: ${JSON.stringify({ id: "evt_fixture", type: "server.connected", data: {} })}\n\n` })
    unexpected.push(`${method} ${pathname}`)
    return route.abort()
  })
  const page = await context.newPage()
  page.setDefaultTimeout(15000)
  page.on("pageerror", (error: Error) => pageErrors.push(error.message))
  await page.goto(`${baseURL}/matter/${matter.id}`)
  const docs = page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ })
  await expect(docs).toBeVisible()
  if (await docs.getAttribute("aria-expanded") !== "true") await docs.click()
  await page.getByRole("button", { name: NAME, exact: true }).click()
  const dialog = page.getByRole("dialog", { name: NAME, exact: true })
  const frame = dialog.locator("iframe.pdf-render")
  await expect(frame).toBeVisible()
  await page.waitForTimeout(1500) // let the PDF plugin lay out
  const m = await page.evaluate(() => {
    const box = (selector: string) => { const r = document.querySelector(selector)!.getBoundingClientRect(); return { top: r.top, height: r.height } }
    const scroller = document.querySelector<HTMLElement>(".document-reader .viewer-body")!
    return { iframe: box("iframe.pdf-render"), body: box(".document-reader .viewer-body"), panel: box(".document-reader"), doc: box(".document-reader .viewer-doc"), viewport: innerHeight, bodyScrollable: scroller.scrollHeight > scroller.clientHeight, pageScrollable: document.documentElement.scrollHeight > document.documentElement.clientHeight }
  })
  measurements[label] = m
  await page.screenshot({ path: path.join(evidence!, `pdf-${label}.png`) })
  console.log(label, JSON.stringify(m))
  if (process.env.DOCHAUS_MEASURE_ONLY === "1") return context.close() // report heights without asserting
  // The PDF fills the viewer body (within its padding/status line), not a strip,
  // and the viewer never grows past the screen.
  expect(m.iframe.height).toBeGreaterThanOrEqual(m.viewport * 0.5)
  expect(m.iframe.height).toBeGreaterThanOrEqual(m.body.height * 0.8)
  expect(m.panel.height).toBeLessThanOrEqual(m.viewport + 1)
  expect(m.bodyScrollable).toBe(false) // the PDF scrolls inside its frame, not the viewer
  await context.close()
}

try {
  await run("desktop", 1440, 900)
  await run("narrow", 400, 800)
  expect(pageErrors.slice(0, 5)).toEqual([])
  expect(unexpected.slice(0, 5)).toEqual([])
  console.log(JSON.stringify(measurements, null, 2))
} finally {
  await browser.close()
}
