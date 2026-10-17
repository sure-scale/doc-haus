// Matter chat layout regression against a separately started Vite/preview
// server. A blocked run's refinement form makes the run header tall; the
// composer must stay inside the chat card and clear of the status bar, and the
// header's own actions must stay in view while only its requirement list scrolls.
// All API traffic is intercepted; this never reaches a live service or account.
import { mkdir, writeFile } from "node:fs/promises"
import path from "node:path"
import blockedGate from "../../../dochaus/workflows/lean/fixtures/status-blocked-gate.json"

const { chromium, expect } = await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")
const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (!baseURL || !evidence) throw new Error("Set PLAYWRIGHT_BASE_URL and DOCHAUS_EVIDENCE_DIR.")
await mkdir(evidence, { recursive: true })
const matter = { id: "fixture-matter", title: "Fixture matter", dir: "/fixture-matter", created_at: 1, documents: [] }
const parent = (id: string, check: string, reason: string) => ({ parent_id: id, check, target: { document: "d1", role: "body", treatment: "operative" }, finding: { status: "unsatisfied", reason }, scopes: ["target_body"] })
const status = {
  ...blockedGate,
  stages: { ...blockedGate.stages, "checklist-check": { status: "blocked", attempts: 1, started_at: null, finished_at: null } },
  checklist_refinement: { available: true, parents: [
    parent("req-a", "State the agreed payment terms.", "The payment date was not confirmed in the draft or the instructions."),
    parent("req-b", "Address each requested amendment in the draft.", "The earlier inspection did not cover the whole draft."),
    parent("req-c", "Identify the parties by their full legal names.", "One party appears only by a trading name."),
  ] },
}
const browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
const results: Record<string, unknown>[] = []
try {
  for (const [width, height] of [[1512, 806], [1440, 900], [1280, 720], [400, 860]]) {
    const context = await browser.newContext({ viewport: { width, height }, reducedMotion: "reduce" })
    await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
    const unexpected: string[] = []
    const writes: string[] = []
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
      const reply = (body: unknown, code = 200) => route.fulfill({ status: code, headers, contentType: "application/json", body: JSON.stringify(body) })
      if (request.method() !== "GET") {
        writes.push(`${request.method()} ${pathname}`)
        return reply({ error: "Fixture refuses writes." }, 409)
      }
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
    await page.goto(`${baseURL}/matter/${matter.id}?view=chat&session=${blockedGate.owner_session_id}`)
    await expect(page.getByRole("form", { name: "Check a requirement against the full draft" })).toBeVisible()
    const send = page.locator("button.send-btn")
    // A desktop window keeps the header's actions and the composer on screen;
    // a stacked one scrolls the page to each in turn.
    const reachable = (selector: string) => page.evaluate((selector: string) => {
      const node = document.querySelector(selector)!
      const box = node.getBoundingClientRect()
      const hit = document.elementFromPoint(box.x + box.width / 2, box.y + box.height / 2)
      return box.top >= 0 && box.bottom <= innerHeight && box.left >= 0 && box.right <= innerWidth && Boolean(hit && node.contains(hit))
    }, selector)
    const actions: Record<string, boolean> = {}
    for (const [name, selector] of [["continue", ".run-head-actions .run-head-cta"], ["check", ".run-checklist-refinement button[type=submit]"], ["details", ".run-head-toggle"]]) {
      if (width <= 700) await page.locator(selector).scrollIntoViewIfNeeded()
      actions[name] = await reachable(selector)
    }
    await expect(page.getByRole("button", { name: "Continue drafting", exact: true })).toBeVisible()
    await expect(page.getByRole("button", { name: "Check the full draft", exact: true })).toBeVisible()
    await page.getByRole("textbox", { name: "Message" }).fill("Fixture instruction.")
    if (width <= 700) await send.scrollIntoViewIfNeeded()
    const layout = await page.evaluate(() => {
      const rect = (selector: string) => document.querySelector(selector)!.getBoundingClientRect()
      const send = rect("button.send-btn")
      const bar = rect("footer.matter-statusbar")
      const card = rect(".matter-body .chat-card")
      const composer = rect(".composer")
      const hit = document.elementFromPoint(send.x + send.width / 2, send.y + send.height / 2)
      return {
        hitsSend: Boolean(hit?.closest("button.send-btn")),
        hit: hit ? `${hit.tagName.toLowerCase()}.${String(hit.className).trim().replace(/\s+/g, ".")}` : "none",
        composerInCard: composer.top >= card.top - 1 && composer.bottom <= card.bottom + 1,
        clearOfStatusbar: composer.bottom <= bar.top + 1 || composer.top >= bar.bottom - 1,
        statusbarVisible: bar.height > 0 && bar.top >= 0 && bar.bottom <= innerHeight + 1,
        overflowX: document.documentElement.scrollWidth - document.documentElement.clientWidth,
        boxes: Object.fromEntries(Object.entries({ send, bar, card, composer }).map(([name, box]) => [name, [box.top, box.bottom].map(Math.round)])),
      }
    })
    await page.screenshot({ path: path.join(evidence, `matter-composer-${width}x${height}.png`) })
    // Arrowing to the last requirement scrolls it into the list's view.
    const radios = page.getByRole("radio")
    await radios.first().focus()
    for (let index = 1; index < await radios.count(); index++) await page.keyboard.press("ArrowDown")
    const focusInList = await page.evaluate(() => {
      const focused = document.activeElement!.getBoundingClientRect()
      const list = document.querySelector(".run-checklist-choices")!.getBoundingClientRect()
      return document.activeElement === [...document.querySelectorAll(".run-checklist-choices input")].at(-1) && focused.top >= list.top && focused.bottom <= list.bottom
    })
    // A real pointer click: Playwright refuses it while another element intercepts.
    const click = await send.click({ timeout: 3000 }).then(() => "clicked", (error: Error) => error.message.split("\n").find(line => line.includes("intercepts")) ?? error.message.split("\n")[0])
    if (click === "clicked") await expect.poll(() => writes.length).toBeGreaterThan(0)
    results.push({ size: `${width}x${height}`, actions, focusInList, ...layout, click, writes, unexpected })
    await context.close()
  }
} finally {
  await writeFile(path.join(evidence, "matter-composer-results.json"), JSON.stringify(results, null, 2))
  await browser.close()
}
for (const result of results) {
  console.log(JSON.stringify(result))
  expect(result).toMatchObject({ actions: { continue: true, check: true, details: true }, focusInList: true, hitsSend: true, composerInCard: true, clearOfStatusbar: true, overflowX: 0, click: "clicked", unexpected: [] })
  expect(result.writes).toEqual([`POST /matters/${matter.id}/lean/updates`])
  if (!String(result.size).startsWith("400x")) expect(result.statusbarVisible).toBe(true)
}
console.log(`PASS ${results.length} viewports: run actions and Send are in view, hit-testable and clear of the status bar`)
