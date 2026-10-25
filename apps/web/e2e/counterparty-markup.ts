// Counterparty markup triage (#200) against a separately started Vite server.
// The runner starts a real ingest service on DOCHAUS_E2E_INGEST_PORT with a
// temporary data root and settings directory, uploads the synthetic marked-up
// supply agreement (dochaus/test/support/counterparty-markup-fixture.ts) as a
// source and the same bytes as a generated deliverable, and drives the viewer's
// Tracked changes panel: filter the Changes tab to one reviewer, reject the
// shown changes with the Accept all / Reject all menu, accept the other reviewer's replace, keep the
// drafts over a reload, give the attorney name once, and save; then, at phone
// width, the panel sits above the page with Save in reach. Nothing reaches a model: the source's
// embeddings use the local model (DOCHAUS_MODEL_CACHE_DIR, else the per-user
// cache), browser engine traffic is intercepted and unexpected requests fail
// the run. Screenshots and the JSON summary go beneath DOCHAUS_EVIDENCE_DIR.
import { createHash } from "node:crypto"
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises"
import os, { tmpdir } from "node:os"
import path from "node:path"
import type { Browser, Locator } from "@playwright/test"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../../../dochaus/test/support/counterparty-markup-fixture"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")

const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
const ingestPort = Number(process.env.DOCHAUS_E2E_INGEST_PORT)
if (!baseURL || !evidence || !Number.isInteger(ingestPort) || ingestPort < 1024)
  throw new Error("Set PLAYWRIGHT_BASE_URL, DOCHAUS_EVIDENCE_DIR and DOCHAUS_E2E_INGEST_PORT (the port Vite's /ingest proxy targets).")
await mkdir(evidence, { recursive: true })

const NAME = COUNTERPARTY_MARKUP.name
const DELIVERABLE = "supply-agreement-draft.docx"
const [AVERY, MORGAN] = COUNTERPARTY_MARKUP.reviewers as [string, string]
const LAWYER = "Synthetic Reviewing Lawyer"
const FLAG = "Contains text that looks like an instruction — doc.haus did not follow it."

const results: { step: string; status: "passed" | "failed"; error?: string }[] = []
const pageErrors: string[] = []
const consoleErrors: string[] = []
const unexpected: string[] = []
const engineRequests: { method: string; path: string }[] = []
const stubbed: string[] = []
const decisionRequests: { actor: string | undefined; body: unknown; status?: number }[] = []
const hashes: Record<string, string> = {}
let deliverableMarkupStatus: number | undefined
const hash = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")

// The service needs the engine only for its model list (readiness); anything
// else that reaches it is recorded.
const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: (request) => {
  const url = new URL(request.url)
  if (request.method === "GET" && url.pathname === "/api/model")
    return Response.json({ data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
  unexpected.push(`engine fixture ${request.method} ${url.pathname}`)
  return new Response(null, { status: 404 })
} })

const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-counterparty-markup-e2e-"))
const dataRoot = path.join(temporary, "data")
const settings = path.join(temporary, "settings")
await mkdir(path.join(dataRoot, "workspace", ".preferences"), { recursive: true })
await mkdir(settings, { recursive: true })
await writeFile(path.join(settings, "opencode.json"), JSON.stringify({ model: "fixture/primary", small_model: "fixture/fast" }))
// No attorney name yet: the apply bar asks for it once and saves it.
await writeFile(path.join(dataRoot, "workspace", ".preferences", "preferences.json"), JSON.stringify({ firm: "Synthetic Firm" }))
const ingestURL = `http://127.0.0.1:${ingestPort}`
const api = (route: string, init?: RequestInit) => fetch(`${ingestURL}${route}`, { ...init, signal: AbortSignal.timeout(120_000) })
let ingest: ReturnType<typeof Bun.spawn> | undefined

async function portFree(port: number) {
  try { const server = Bun.listen({ hostname: "127.0.0.1", port, socket: { data() {} } }); server.stop(true); return true }
  catch { return false }
}
// Only explicit isolation settings reach the service: no inherited workspace,
// runtime, engine address or credentials from the caller's environment.
async function startIngest() {
  if (!await portFree(ingestPort)) throw new Error(`Port ${ingestPort} is occupied; choose a free DOCHAUS_E2E_INGEST_PORT.`)
  const log = Bun.file(path.join(temporary, "ingest.log"))
  ingest = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, "../../../services/ingest"), stdout: log, stderr: log,
    env: {
      PATH: process.env.PATH ?? "", HOME: process.env.HOME ?? temporary, TMPDIR: process.env.TMPDIR ?? tmpdir(),
      DOCHAUS_DATA_ROOT: dataRoot, WORKSPACE_ROOT: path.join(dataRoot, "workspace"), DOCHAUS_RUNTIME_DIR: path.join(dataRoot, "runtime"), DOCHAUS_DIR: settings, INGEST_HOST: "127.0.0.1", INGEST_PORT: String(ingestPort),
      OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "",
      // The local embedding model's weights are the same everywhere; reuse the download.
      DOCHAUS_MODEL_CACHE_DIR: process.env.DOCHAUS_MODEL_CACHE_DIR || path.join(os.homedir(), ".cache", "doc-haus", "models"),
    },
  })
  const deadline = Date.now() + 30_000
  while (!await fetch(`${ingestURL}/matters`).then(response => response.ok).catch(() => false)) {
    if (ingest.exitCode !== null || Date.now() > deadline) throw new Error(`Isolated ingest did not start; see ${log.name}`)
    await Bun.sleep(100)
  }
}
async function stopIngest() {
  if (!ingest) return
  ingest.kill("SIGTERM")
  await ingest.exited
  ingest = undefined
}

let capture: ((name: string) => Promise<unknown>) | undefined
async function step(name: string, run: () => Promise<void>) {
  try { await run(); results.push({ step: name, status: "passed" }); console.log(`PASS ${name}`) }
  catch (error) {
    results.push({ step: name, status: "failed", error: error instanceof Error ? error.message : String(error) })
    console.log(`FAIL ${name}`)
    await capture?.(`failed-${name.split(" ")[0]}`).catch(() => undefined)
    throw error
  }
}

const file = { path: "", dir: "", matter: "" }
let browser: Browser | undefined
try {
  await startIngest()
  browser = await chromium.launch({ headless: true, executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE })
  const context = await browser!.newContext({ viewport: { width: 1440, height: 1000 }, reducedMotion: "reduce" })
  await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  const origin = new URL(baseURL).origin
  // Engine traffic is answered here; ingest traffic continues through Vite's
  // /ingest proxy to the isolated service.
  await context.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    if (url.origin !== origin) { unexpected.push(`${request.method()} ${request.url()}`); return route.abort() }
    const reply = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    // The release update check runs git in the checkout; it is not part of this flow.
    if (request.method() === "GET" && url.pathname === "/ingest/updates") {
      stubbed.push("GET /ingest/updates")
      return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    }
    if (url.pathname.startsWith("/ingest/")) return route.continue()
    if (!url.pathname.startsWith("/engine/")) return route.continue()
    const pathname = url.pathname.slice("/engine".length)
    engineRequests.push({ method: request.method(), path: pathname })
    if (request.method() === "GET" && pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (request.method() === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (request.method() === "GET" && pathname === "/api/session") return reply({ data: [], cursor: { next: null, previous: null } })
    if (request.method() === "GET" && pathname === "/api/agent") return reply({ location: { directory: file.dir }, data: [] })
    // An idle event stream: left open, as the engine would hold it between events.
    if (request.method() === "GET" && pathname === "/api/event") return
    unexpected.push(`engine ${request.method()} ${pathname}${url.search}`)
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: "not stubbed" }) })
  })
  const page = await context.newPage()
  page.setDefaultTimeout(20_000)
  page.on("pageerror", (error: Error) => pageErrors.push(error.message))
  page.on("console", (message: { type: () => string; text: () => string }) => message.type() === "error" && consoleErrors.push(message.text()))
  page.on("request", request => {
    if (request.method() === "POST" && new URL(request.url()).pathname.endsWith("/documents/markup/decisions"))
      decisionRequests.push({ actor: request.headers()["x-dochaus-actor"], body: request.postDataJSON() })
  })
  const shot = (name: string) => page.screenshot({ path: path.join(evidence, `${name}.png`) })
  capture = shot
  const dialog = page.getByRole("dialog", { name: NAME, exact: true })
  const panel = dialog.locator("#document-markup")
  // Both tabs' panels are rendered; the one not chosen is hidden.
  const shown = panel.locator('[role="tabpanel"]:not([hidden])')
  const changes = shown.locator(".markup-list > li:not(.markup-thread)")
  const threads = shown.locator(".markup-list > li.markup-thread")
  const tab = (name: "Changes" | "Comments") => panel.getByRole("tab", { name: new RegExp(`^${name} \\d+$`) })
  const toggle = dialog.getByRole("button", { name: /^Tracked changes \(\d+\)$/ })
  const render = dialog.locator(".docx-render:not(.docx-skeleton)")
  // The save bar's line: how many of the changes doc.haus can apply are decided, and what is left.
  const progress = panel.locator(".markup-progress p")
  const live = panel.locator(":scope > p.sr-only[role=status]")
  const save = panel.getByRole("button", { name: "Save as new version", exact: true })
  const checked = (card: Locator) => card.getByRole("radio", { checked: true })
  // A change by what changed, as its decision group names it ("Moved paragraph by …").
  const kind = (label: string) => changes.filter({ has: page.getByRole("group", { name: new RegExp(`^Your decision: ${label} by `) }) })
  const change = (selector: string, text: string) => changes.filter({ has: page.locator(selector, { hasText: new RegExp(`^${text}$`) }) })
  const filtersToggle = shown.getByRole("button", { name: /^Filters\b/ })
  const filters = (which: "changes" | "comments") => shown.getByRole("group", { name: `Filter ${which}`, exact: true })
  // The Decision filter counts each decision across the listed changes, the tab's other filters applied.
  const decisionCounts = () => panel.getByLabel("Decision", { exact: true }).locator("option").allTextContents()

  async function openDocument(name: string, navigate = true) {
    if (navigate) await page.goto(`${baseURL!.replace(/\/$/, "")}/matter/${file.matter}`)
    const documents = page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ })
    await expect(documents).toBeVisible()
    if (await documents.getAttribute("aria-expanded") !== "true") await documents.click()
    await page.getByRole("button", { name, exact: true }).click()
    const viewer = page.getByRole("dialog", { name, exact: true })
    await expect(viewer).toBeVisible()
    // The first open loads the Word runtime; later renders are quick.
    await expect(viewer.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
    return viewer
  }

  await step("1 upload the marked-up agreement as a source and the same bytes as a deliverable", async () => {
    const matter = await (await api("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Synthetic Counterparty Markup" }) })).json() as { id: string; dir: string }
    expect(matter.dir.startsWith(dataRoot)).toBe(true)
    Object.assign(file, { matter: matter.id, dir: matter.dir, path: path.join(matter.dir, NAME) })
    const source = new FormData(); source.append("file", new File([counterpartyMarkupDocx()], NAME))
    const uploaded = await api(`/matters/${matter.id}/documents`, { method: "POST", body: source })
    expect(uploaded.status).toBe(200)
    const deliverable = new FormData(); deliverable.append("file", new File([counterpartyMarkupDocx()], DELIVERABLE))
    expect((await api(`/matters/${matter.id}/deliverables`, { method: "POST", body: deliverable })).status).toBe(200)
    hashes.seeded = hash(await Bun.file(file.path).bytes())
    expect(hashes.seeded).toBe(hash(counterpartyMarkupDocx()))
  })

  await step("2 the panel opens by default on Changes, one card per change, and the threads on Comments, in plain words", async () => {
    await openDocument(NAME)
    await expect(toggle).toHaveAccessibleName("Tracked changes (13)")
    await expect(toggle).toHaveAttribute("aria-expanded", "true")
    await expect(panel.getByRole("heading", { name: "Tracked changes", exact: true })).toBeVisible()
    await expect(panel.getByRole("tablist", { name: "Tracked changes and comments", exact: true })).toBeVisible()
    await expect(tab("Changes")).toHaveAccessibleName("Changes 10")
    await expect(tab("Comments")).toHaveAccessibleName("Comments 3")
    await expect(tab("Changes")).toHaveAttribute("aria-selected", "true")
    await expect(changes).toHaveCount(10)
    await expect(threads).toHaveCount(0)
    // Nothing is decided until the lawyer decides it, which the card's edge
    // shows and its decision group says; no internal id is shown.
    for (let index = 0; index < 10; index++) {
      await expect(checked(changes.nth(index))).toHaveCount(0)
      await expect(changes.nth(index).getByRole("group")).toHaveAccessibleDescription("Not yet decided")
    }
    await expect(panel).not.toContainText("rev2-")
    await expect(progress).toHaveText("0 of 10 decided")
    // Every option says how many it lists.
    expect(await decisionCounts()).toEqual(["All decisions (10)", "Not yet decided (10)", "Accept (0)", "Reject (0)", "Leave tracked (0)"])
    // A move is one card with both places, given in its tooltip and to a
    // screen reader; formatting reads as such.
    await expect(kind("Moved paragraph")).toHaveCount(1)
    await expect(kind("Moved paragraph")).toHaveAttribute("title", "Moved from paragraph 3 to paragraph 5")
    await expect(kind("Moved paragraph")).toContainText("Moved from paragraph 3 to paragraph 5")
    await expect(kind("Formatting changed")).toHaveCount(1)
    await expect(kind("Paragraph break deleted")).toHaveCount(1)
    await expect(changes.filter({ hasText: COUNTERPARTY_MARKUP.plantedDeletion.trim() })).toContainText(FLAG)
    // The proposals panel and the versions list are not shown beside it.
    await expect(dialog.locator(".changes-panel")).toHaveCount(1)
    // Clicked, a card shows its change in the document: every part of it
    // outlined, the first scrolled into view.
    const located = render.locator(".markup-located")
    await change(".change-new", "forty-five").locator(".change-new").click()
    await expect(located).toHaveText(["forty-five"])
    await expect(located.first()).toBeInViewport()
    // One insertion Word splits in two runs, and a move's two places.
    await changes.filter({ hasText: "on ten Business Days' notice" }).locator(".change-new").click()
    await expect(located).toHaveText(["on ten Business Days' notice", "at any time"])
    await expect(located.first()).toBeInViewport()
    await kind("Moved paragraph").locator(".markup-text").click()
    await expect(located).toHaveText(["3. Notices. Notices must be in writing.", "3. Notices. Notices must be in writing."])
    await expect(located.first()).toBeInViewport()
    await shot("02-panel-open")
    // Threads: the reply under its comment, the resolved note marked, the quote shown.
    await tab("Comments").click()
    await expect(tab("Comments")).toHaveAttribute("aria-selected", "true")
    await expect(changes).toHaveCount(0)
    await expect(threads).toHaveCount(3)
    const answered = threads.filter({ hasText: COUNTERPARTY_MARKUP.comments[0]!.text })
    const reply = answered.locator(".markup-replies > li")
    await expect(reply).toHaveCount(1)
    await expect(reply).toContainText(MORGAN)
    await expect(reply).toContainText(COUNTERPARTY_MARKUP.comments[1]!.text)
    await expect(answered.locator(".markup-quote")).toHaveText("forty-five")
    await expect(threads.filter({ hasText: COUNTERPARTY_MARKUP.comments[2]!.text })).toContainText("Resolved")
    // Both planted instructions carry the badge, one on each tab, and nothing else does.
    await expect(threads.filter({ hasText: COUNTERPARTY_MARKUP.plantedComment })).toContainText(FLAG)
    await expect(shown.getByText(FLAG, { exact: true })).toHaveCount(1)
    await expect(panel.getByText(FLAG, { exact: true })).toHaveCount(2)
    await shot("02-comments")
    // As in Word with its comments pane open, a comment shows once: while the
    // panel is open the page has no margin cards, only a pin per thread, and a
    // pin opens its thread here.
    const pinned = await answered.getAttribute("data-comment-id")
    await tab("Changes").click()
    await expect(changes).toHaveCount(10)
    await expect(dialog.locator(".comment-margin-note:visible")).toHaveCount(0)
    await expect(dialog.locator("button.doc-comment-pin")).toHaveCount(3)
    await dialog.locator(`button.doc-comment-pin[data-comment-id="${pinned}"]`).click()
    await expect(tab("Comments")).toHaveAttribute("aria-selected", "true")
    await expect(threads.filter({ hasText: COUNTERPARTY_MARKUP.comments[0]!.text })).toBeFocused()
    await tab("Changes").click()
    await expect(changes).toHaveCount(10)
  })

  await step("3 filter to Morgan and reject the shown changes with Reject all shown", async () => {
    await filtersToggle.click()
    const reviewer = filters("changes").getByLabel("Reviewer", { exact: true })
    await expect(reviewer).toBeFocused()
    await reviewer.selectOption(`=${MORGAN}`)
    // Morgan's five changes.
    await expect(changes).toHaveCount(5)
    for (let index = 0; index < 5; index++) await expect(changes.nth(index).locator(".change-author")).toHaveText(MORGAN)
    // Said to a screen reader only; the chips show the filter.
    await expect(shown.locator(".markup-showing.sr-only")).toHaveText("Showing 5 of 10 changes")
    // Closed, the filter keeps filtering and stays in view as a chip.
    await filtersToggle.click()
    await expect(filters("changes")).toBeHidden()
    await expect(filtersToggle).toBeFocused()
    await expect(filtersToggle).toHaveAccessibleName("Filters 1 in use")
    await expect(shown.getByRole("button", { name: `Remove the filter Reviewer: ${MORGAN}`, exact: true })).toBeVisible()
    await expect(changes).toHaveCount(5)
    const decideAll = shown.getByRole("button", { name: "Accept all / Reject all", exact: true })
    await decideAll.click()
    await expect(decideAll).toHaveAttribute("aria-expanded", "true")
    const bulk = shown.getByRole("group", { name: "Accept all / Reject all", exact: true })
    // Filtered, it sets the changes shown; the others keep their decisions.
    await expect(bulk.getByRole("button")).toHaveText(["Accept all shown (5)", "Reject all shown (5)", "Leave all shown tracked (5)"])
    await expect(bulk).toContainText("Your own edits are not changed.")
    await bulk.getByRole("button", { name: "Reject all shown (5)", exact: true }).click()
    await expect(bulk).toHaveCount(0)
    await expect(decideAll).toBeFocused()
    for (let index = 0; index < 5; index++) await expect(checked(changes.nth(index))).toHaveAccessibleName("Reject")
    await expect(live).toHaveText("5 changes rejected.")
    await expect(shown.locator(".markup-undo")).toHaveText("5 changes rejected. Undo")
    await expect(progress).toHaveText("5 of 10 decided")
    await shot("03-morgan-rejected")
    // The Comments tab filters by reviewer too: the two threads Morgan wrote in.
    await tab("Comments").click()
    await filtersToggle.click()
    await filters("comments").getByLabel("Reviewer", { exact: true }).selectOption(`=${MORGAN}`)
    await expect(threads).toHaveCount(2)
    await filters("comments").getByRole("button", { name: "Clear filters", exact: true }).click()
    await expect(threads).toHaveCount(3)
    await expect(filtersToggle).toBeFocused()
    await filtersToggle.click()
    await tab("Changes").click()
    await expect(changes).toHaveCount(5)
  })

  await step("4 accept Avery's replace; Avery's other changes stay undecided", async () => {
    await filtersToggle.click()
    await filters("changes").getByLabel("Reviewer", { exact: true }).selectOption(`=${AVERY}`)
    await expect(changes).toHaveCount(5)
    for (let index = 0; index < 5; index++) await expect(changes.nth(index).locator(".change-author")).toHaveText(AVERY)
    await change(".change-old", "thirty").getByRole("radio", { name: "Accept", exact: true }).click()
    await change(".change-new", "forty-five").getByRole("radio", { name: "Accept", exact: true }).click()
    await expect(checked(change(".change-old", "thirty"))).toHaveAccessibleName("Accept")
    await expect(checked(change(".change-new", "forty-five"))).toHaveAccessibleName("Accept")
    await expect(changes.filter({ hasText: "Not yet decided" })).toHaveCount(3)
    await expect(progress).toHaveText("7 of 10 decided")
    expect(decisionRequests).toEqual([])
    expect(hash(await Bun.file(file.path).bytes())).toBe(hashes.seeded)
  })

  await step("5 the drafts survive a reload, and the bar asks for the attorney name once", async () => {
    await page.reload()
    await openDocument(NAME, false)
    await expect(progress).toHaveText("7 of 10 decided")
    // The filters start afresh, so the Decision filter counts every change.
    await expect.poll(decisionCounts).toEqual(["All decisions (10)", "Not yet decided (3)", "Accept (2)", "Reject (5)", "Leave tracked (0)"])
    await expect(checked(change(".change-old", "thirty"))).toHaveAccessibleName("Accept")
    await expect(checked(change(".change-new", "forty-five"))).toHaveAccessibleName("Accept")
    await expect(panel.getByLabel("Your name, recorded with these decisions", { exact: true })).toBeVisible()
    await expect(panel.locator(".markup-actor")).toHaveCount(0)
    await shot("05-drafts-kept")
  })

  await step("6 save: one new version recorded as the lawyer, the undecided changes still listed", async () => {
    await panel.getByLabel("Your name, recorded with these decisions", { exact: true }).fill(LAWYER)
    await save.click()
    const confirm = panel.getByRole("group", { name: "Save a new version", exact: true })
    await expect(confirm.locator("p").first()).toHaveText(`This saves a new version of ${NAME} with 2 changes accepted and 5 rejected. ` +
      "3 changes you have not decided stay as tracked changes. The current version stays in the history.")
    await expect(confirm.locator(".markup-actor")).toHaveText(`Saved as: ${LAWYER}`)
    await shot("06-confirm")
    const applied = page.waitForResponse(response => new URL(response.url()).pathname.endsWith("/documents/markup/decisions"))
    await confirm.getByRole("button", { name: "Save new version", exact: true }).click()
    const response = await applied
    decisionRequests.at(-1)!.status = response.status()
    expect(response.status()).toBe(200)
    expect(decisionRequests).toHaveLength(1)
    expect(decisionRequests[0]!.actor).toBe(LAWYER)
    const sent = decisionRequests[0]!.body as { base_sha256: string; decisions: { id: string; decision: string }[] }
    expect(sent.base_sha256).toBe(hashes.seeded)
    expect(sent.decisions).toHaveLength(10)
    expect(new Set(sent.decisions.map(row => row.id)).size).toBe(10)
    expect(sent.decisions.map(row => row.decision).sort()).toEqual([...Array(2).fill("accept"), ...Array(3).fill("hold"), ...Array(5).fill("reject")])
    await expect(page.getByText(/^Saved as version 2: 2 accepted, 5 rejected, 3 left tracked\.$/)).toBeVisible()
    // The panel reads the new version: Avery's three undecided changes and every comment remain.
    await expect(toggle).toHaveAccessibleName("Tracked changes (6)", { timeout: 30_000 })
    await expect(tab("Changes")).toHaveAccessibleName("Changes 3")
    await expect(changes).toHaveCount(3)
    for (let index = 0; index < 3; index++) {
      await expect(changes.nth(index).locator(".change-author")).toHaveText(AVERY)
      await expect(checked(changes.nth(index))).toHaveCount(0)
    }
    await expect(kind("Moved paragraph")).toHaveCount(1)
    await expect(changes.filter({ hasText: "on ten Business Days' notice" })).toHaveCount(1)
    await expect(progress).toHaveText("0 of 3 decided")
    // A new version's decisions start afresh, so Save waits for one and says
    // why on hover or focus and to a screen reader.
    await expect(save).toBeDisabled()
    await expect(save).toHaveAccessibleDescription("Accept or reject a change, or approve a reply, to save.")
    await save.hover()
    await expect(page.getByRole("tooltip")).toHaveText("Accept or reject a change, or approve a reply, to save.")
    await page.mouse.move(0, 0)
    await expect(page.getByRole("tooltip")).toHaveCount(0)
    await expect(panel.getByText(FLAG, { exact: true })).toHaveCount(2)
    // Who the decisions are saved as is said when saving.
    await expect(panel.locator(".markup-actor")).toHaveCount(0)
    // The count leads to the first change still to decide.
    await panel.getByRole("button", { name: "0 of 3 decided. Show the next change to decide", exact: true }).click()
    await expect(changes.first().getByRole("radio", { name: "Accept", exact: true })).toBeFocused()
    await expect(tab("Comments")).toHaveAccessibleName("Comments 3")
    // The comment markers after the anchored word ("[0]", or words for a screen
    // reader such as "(comment by …)") are left out of the page's text.
    const pageText = () => render.evaluate(node => {
      const copy = node.cloneNode(true) as HTMLElement
      for (const marker of copy.querySelectorAll("a.comment-marker")) marker.remove()
      return copy.textContent ?? ""
    })
    await expect.poll(pageText).toMatch(/within\s+forty-five\s+days/)
    await expect(render).not.toContainText("thirty")
    await expect(render).not.toContainText("the first")
    hashes.applied = hash(await Bun.file(file.path).bytes())
    expect(hashes.applied).not.toBe(hashes.seeded)
    // The name was saved once, through the drafting preferences.
    expect((await (await api("/preferences")).json() as { attorney?: string }).attorney).toBe(LAWYER)
    await shot("06-applied")
  })

  await step("7 the versions list shows the new version with the actor and replaces the panel", async () => {
    await dialog.getByRole("button", { name: /^Versions/ }).click()
    await expect(panel).toHaveCount(0)
    await expect(toggle).toHaveAttribute("aria-expanded", "false")
    const newest = dialog.locator(".versions-list > li").first()
    await expect(newest).toContainText("Version 2")
    await expect(newest).toContainText("Current")
    // Decided from the panel, not doc.haus's proposals accepted: 5 of these were rejected.
    await expect(newest.locator(".version-meta")).toHaveText(`Tracked changes decided · by ${LAWYER}`)
    const listed = await (await api(`/matters/${file.matter}/documents/versions?name=${encodeURIComponent(NAME)}`)).json() as { versions: { number: number; origin: string; actor: string; sha256: string; current: boolean }[] }
    expect(listed.versions.map(row => [row.number, row.origin, row.actor, row.current])).toEqual([[2, "accept", LAWYER, true], [1, "upload", expect.any(String), false]])
    expect(listed.versions[0]!.sha256).toBe(hashes.applied)
    await shot("07-versions")
    await toggle.click()
    await expect(panel).toBeVisible()
    await expect(dialog.locator("#document-versions")).toHaveCount(0)
  })

  await step("8 a generated deliverable with the same tracked changes offers no panel", async () => {
    await dialog.getByRole("button", { name: "Close", exact: true }).press("Enter")
    // The toggle is decided by the route's answer, so wait for it before checking.
    const read = page.waitForResponse(response => {
      const url = new URL(response.url())
      return url.pathname.endsWith("/documents/markup") && url.searchParams.get("name") === DELIVERABLE
    })
    const other = await openDocument(DELIVERABLE, false)
    deliverableMarkupStatus = (await read).status()
    await expect(other.getByRole("button", { name: /^Versions/ })).toBeVisible()
    await expect(other.getByRole("button", { name: /^Tracked changes/ })).toHaveCount(0)
    await expect(other.locator("#document-markup")).toHaveCount(0)
    await shot("08-deliverable")
  })

  await step("9 at phone width the panel sits above the page, the toggle keeps its count and Save is in reach", async () => {
    await page.setViewportSize({ width: 400, height: 860 })
    // The rail starts collapsed at this width, as on a first visit from a phone.
    await page.evaluate(() => localStorage.setItem("dh.sidebar", "1"))
    const phone = await openDocument(NAME)
    const phonePanel = phone.locator("#document-markup")
    const phoneToggle = phone.locator(".viewer-markup-toggle")
    const phonePage = phone.locator(".docx-render:not(.docx-skeleton)")
    await expect(phonePanel).toBeInViewport()
    await expect(phoneToggle).toHaveAttribute("aria-expanded", "true")
    await expect(phoneToggle).toHaveAccessibleName("Tracked changes (6)")
    await expect(phoneToggle.locator(".viewer-markup-badge")).toBeVisible()
    await expect(phoneToggle.locator(".viewer-markup-badge")).toHaveText("6")
    const [panelBox, pageBox] = [await phonePanel.boundingBox(), await phonePage.boundingBox()]
    expect(panelBox!.y + panelBox!.height).toBeLessThanOrEqual(pageBox!.y)
    // The comments follow the text rather than squeezing it to a few words a line.
    expect(await phonePage.locator(".comment-margin-content").evaluate(node => node.getBoundingClientRect().width)).toBeGreaterThan(250)
    await shot("09-phone-panel")
    const phoneSave = phonePanel.getByRole("button", { name: "Save as new version", exact: true })
    await phoneSave.scrollIntoViewIfNeeded()
    await expect(phoneSave).toBeInViewport()
    // In view and not under the panel's pinned head.
    expect(await phoneSave.evaluate(node => {
      const box = node.getBoundingClientRect()
      const hit = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2)
      return Boolean(hit && node.contains(hit))
    })).toBe(true)
    expect(await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth)).toBeLessThanOrEqual(1)
    await shot("09-phone-save")
  })

  await step("no page errors or unexpected requests", async () => {
    expect(pageErrors).toEqual([])
    expect(unexpected).toEqual([])
  })
}
finally {
  await writeFile(path.join(evidence, "counterparty-markup-results.json"), JSON.stringify({
    results, pageErrors, consoleErrors, unexpected, stubbed: [...new Set(stubbed)],
    engineEndpoints: Object.entries(Object.groupBy(engineRequests, row => `${row.method} ${row.path}`)).map(([endpoint, rows]) => ({ endpoint, count: rows!.length })),
    decisionRequests: decisionRequests.map(({ actor, status, body }) => ({ actor, status, decisions: (body as { decisions?: unknown[] } | null)?.decisions?.length })),
    hashes, deliverableMarkupStatus,
  }, null, 2) + "\n")
  await browser?.close().catch(() => undefined)
  await stopIngest().catch(error => console.error(error))
  engine.stop(true)
  await rm(temporary, { recursive: true, force: true })
}
