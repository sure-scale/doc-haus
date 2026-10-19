// Run Details regression against a separately started Vite/preview server.
// Open Details scroll on their own under the run header's title row; a long
// Open items list must keep its summary in view, under its own centre and opaque
// in the header's colour, so the lawyer can fold the list again however far
// they have read. A completed and a paused header cover both summary colours.
// All API traffic is intercepted; this never reaches a live service or account.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"

const { chromium, expect } = await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
await mkdir(evidence, { recursive: true })
const matter = { id: "fixture-matter", title: "Fixture matter", dir: "/fixture-matter", created_at: 1, documents: [] }
const items = [["research_gap", 1], ["placeholder", 17], ["omission", 1], ["left_to_drafter", 2]].flatMap(([kind, count]) =>
  Array.from({ length: Number(count) }, (_, index) => ({ id: `${kind}-${index + 1}`, stage: "gate", kind, status: "disclosed", detail: `The body contains the placeholder [open item: fixture detail ${index + 1}].` })))
const runs = {
  completed: { ...completed, gate: { ...completed.gate, delivery: "review_draft", disclosed: items, items } },
  paused: { ...blockedGate, gate: { ...blockedGate.gate, items } },
}
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: Record<string, unknown>[] = []
try {
  for (const [run, status] of Object.entries(runs)) for (const [width, height] of [[1440, 1000], [1280, 720], [400, 860]]) {
    const context = await browser.newContext({ viewport: { width, height }, reducedMotion: "reduce" })
    await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
    const unexpected: string[] = []
    await context.route("**/*", async (route: {
      request: () => { url: () => string; method: () => string; resourceType: () => string }
      continue: () => Promise<void>; abort: () => Promise<void>; fulfill: (response: unknown) => Promise<void>
    }) => {
      const request = route.request()
      const url = new URL(request.url())
      const pathname = url.pathname.replace(/^\/ingest(?=\/)/, "").replace(/^\/engine(?=\/api\/)/, "")
      const api = ["fetch", "xhr", "eventsource"].includes(request.resourceType())
      if (!api && url.origin === new URL(baseURL).origin) return route.continue()
      if (!api) { unexpected.push(request.url()); return route.abort() }
      const headers = { "access-control-allow-origin": new URL(baseURL).origin, "access-control-allow-credentials": "true", "access-control-allow-headers": "*", "access-control-allow-methods": "*" }
      if (request.method() === "OPTIONS") return route.fulfill({ status: 204, headers })
      const reply = (body: unknown) => route.fulfill({ status: 200, headers, contentType: "application/json", body: JSON.stringify(body) })
      if (request.method() !== "GET") { unexpected.push(`${request.method()} ${pathname}`); return route.abort() }
      const list = { data: [], cursor: { previous: null, next: null } }
      if (pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
      if (pathname === "/setup") return reply({ version: 1, ready: true, completed: ["models", "research", "details"], skipped: [], next: "done" })
      if (pathname === "/matters") return reply([{ ...matter, documents: undefined }])
      if (pathname === `/matters/${matter.id}`) return reply(matter)
      if (pathname === `/matters/${matter.id}/drafting-status`) return reply(status)
      if (pathname === `/matters/${matter.id}/telemetry`) return reply({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 1, since_ts: 1 }, current_session: null, rate_card_as_of: null, unknown_models: [] })
      if (pathname.startsWith(`/matters/${matter.id}/answer-drafts/`)) return reply({ answers: {}, revision: 0 })
      if (pathname === "/playbooks" || pathname === "/jurisdictions") return reply([])
      if (pathname === "/workflows") return reply({ dir: "/fixture-matter/.workflows", workflows: [] })
      if (pathname === "/updates") return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
      if (pathname === "/api/session" || /^\/api\/session\/[^/]+\/message$/.test(pathname)) return reply(list)
      if (pathname === "/api/session/active") return reply({ data: {} })
      if (/^\/api\/session\/[^/]+\/form$/.test(pathname)) return reply({ data: [] })
      if (pathname === "/api/permission/request" || pathname === "/api/agent") return reply({ location: { directory: matter.dir }, data: [] })
      if (pathname === "/api/event") return route.fulfill({ status: 200, headers, contentType: "text/event-stream", body: `data: ${JSON.stringify({ id: "evt_fixture", type: "server.connected", data: {} })}\n\n` })
      unexpected.push(`${request.method()} ${pathname}`)
      return route.abort()
    })
    const page = await context.newPage()
    page.setDefaultTimeout(10000)
    await page.goto(`${baseURL}/matter/${matter.id}?view=chat&session=${status.owner_session_id}`)
    await page.getByRole("button", { name: "Details" }).click()
    const openItems = page.locator(".open-items")
    await openItems.locator("summary").click()
    await page.getByRole("button", { name: `Show all ${items.length}` }).click()
    await page.locator(".run-head-details-body").scrollIntoViewIfNeeded()
    // Scroll the open Details to a reading position, then report what lies under
    // the summary's centre: what a pointer click there lands on.
    const readTo = (position: "past-summary" | "end") => page.evaluate((position: string) => {
      const body = document.querySelector(".run-head-details-body")!
      const summary = document.querySelector(".open-items > summary")!
      body.scrollTop = position === "end" ? body.scrollHeight : body.scrollTop + summary.getBoundingClientRect().top - body.getBoundingClientRect().top + 24
      const view = body.getBoundingClientRect()
      const box = summary.getBoundingClientRect()
      const hit = document.elementFromPoint(box.x + box.width / 2, box.y + box.height / 2)
      const edge = document.elementFromPoint(box.x + box.width / 2, view.top + 1)
      // Hit tests ignore paint: rows passing beneath show through a clear summary.
      const alpha = (colour: string) => Number(/^rgba\(.*,\s*([\d.]+)\)$/.exec(colour)?.[1] ?? (colour.startsWith("rgb(") ? 1 : 0))
      // The header's own colour, or what it sits on where it is transparent.
      let header: Element | null = summary.closest(".run-head")
      while (header && alpha(getComputedStyle(header).backgroundColor) === 0) header = header.parentElement
      const background = getComputedStyle(summary).backgroundColor
      const headerBackground = header ? getComputedStyle(header).backgroundColor : "none"
      return {
        summaryInView: box.top >= view.top - 1 && box.bottom <= view.bottom + 1,
        hitsSummary: Boolean(hit && summary.contains(hit)),
        // No row of the list shows above the summary once it is read past.
        coversTop: Boolean(edge && summary.contains(edge)),
        opaque: alpha(background) === 1,
        matchesHeader: background === headerBackground,
        background: [background, headerBackground],
        landsOn: hit ? `${hit.tagName.toLowerCase()}.${String(hit.className).trim().replace(/\s+/g, ".")}` : "none",
        centre: [Math.round(box.x + box.width / 2), Math.round(box.y + box.height / 2)],
      }
    }, position)
    const past = await readTo("past-summary")
    await page.screenshot({ path: path.join(evidence, `run-open-items-${run}-${width}x${height}-past-summary.png`) })
    // A real pointer click at the summary's centre, only where it is the summary:
    // elsewhere it would land on the header's own controls.
    if (past.hitsSummary) await page.mouse.click(past.centre[0], past.centre[1])
    const folded = await openItems.evaluate((details: HTMLDetailsElement) => !details.open)
    if (folded) await openItems.locator("summary").click()
    const end = await readTo("end")
    await page.screenshot({ path: path.join(evidence, `run-open-items-${run}-${width}x${height}-end.png`) })
    results.push({ run, size: `${width}x${height}`, past, folded, end, unexpected })
    await context.close()
  }
} finally {
  await writeFile(path.join(evidence, "run-open-items-results.json"), JSON.stringify(results, null, 2))
  await browser.close()
}
for (const result of results) {
  console.log(JSON.stringify(result))
  const read = { summaryInView: true, hitsSummary: true, coversTop: true, opaque: true, matchesHeader: true }
  expect(result).toMatchObject({ past: read, folded: true, end: read, unexpected: [] })
}
console.log(`PASS ${results.length} runs and viewports: the Open items summary stays in view, opaque in the header's colour, and folds the list from any reading position`)
