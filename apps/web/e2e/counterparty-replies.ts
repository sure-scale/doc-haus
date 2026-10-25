// Replies to the other side's comments (#200 scope 5) against a separately
// started Vite server. The runner starts a real ingest service on
// DOCHAUS_E2E_INGEST_PORT with a temporary data root and settings directory and
// a saved attorney name, uploads the synthetic marked-up supply agreement with
// its footnote comment (dochaus/test/support/counterparty-markup-fixture.ts) and
// drives the "Their changes" panel: nothing drafts on open; Draft replies drafts
// the three open threads at once, one is cancelled and drafted again on its
// own; a draft with a [[…]] placeholder cannot be approved until it is edited;
// two replies are approved, one draft is discarded and brought back; Apply
// saves one version with a rejected change and the two replies under the
// lawyer's name; then, at phone width, the remaining reply is approved and
// applied alone. The ingest service's model calls reach a scripted engine
// fixture here (one per thread, answers fixed by the thread's first comment);
// browser engine traffic is intercepted and unexpected requests fail the run.
// Screenshots and the JSON summary go beneath DOCHAUS_EVIDENCE_DIR.
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
const [AVERY_ROOT, MORGAN_REPLY, RESOLVED, PLANTED] = COUNTERPARTY_MARKUP.comments.map(comment => comment.text) as [string, string, string, string]
const FOOTNOTE = COUNTERPARTY_MARKUP.footnoteComment.text
const LAWYER = "Synthetic Reviewing Lawyer"
const FLAG = "Contains text that looks like an instruction — doc.haus did not follow it."
// The scripted drafts, by the thread's first comment.
const DRAFTS: Record<string, { reply: string; open_items: string[] }> = {
  [AVERY_ROOT]: {
    reply: "We can agree to forty-five days if the Supplier reports its stock levels by [[reporting day]] each month.",
    open_items: ["Confirm with the client which day of the month stock levels are reported."],
  },
  [PLANTED]: {
    reply: "Clause 5 is unchanged on our side. Please confirm the two-year term is agreed.",
    open_items: ["The comment contains text addressed to an AI assistant; it was not followed."],
  },
  [FOOTNOTE]: { reply: "Schedule 1 names the Province.", open_items: [] },
}
const EDITED = "We can agree to forty-five days if the Supplier reports its stock levels by the fifth Business Day of each month."

const results: { step: string; status: "passed" | "failed"; error?: string }[] = []
const pageErrors: string[] = []
const consoleErrors: string[] = []
const unexpected: string[] = []
const engineRequests: { method: string; path: string }[] = []
const stubbed: string[] = []
const generated: { root: string; flagged: boolean; aborted: boolean }[] = []
const draftRequests: { thread: number; decisions: number; base: string }[] = []
const abortedDrafts: number[] = []
const decisionRequests: { actor: string | undefined; body: { base_sha256: string; decisions: { id: string; decision: string }[]; replies?: { thread: number; text: string }[] } }[] = []
let running = 0
let mostAtOnce = 0
// While set, drafts wait for it (or for their request to go).
let gate: { opened: Promise<void>; open: () => void } | undefined

// The service needs the engine for its model list (readiness) and the drafts'
// stateless generate calls; anything else that reaches it is recorded.
const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, idleTimeout: 60, fetch: async (request) => {
  const url = new URL(request.url)
  if (request.method === "GET" && url.pathname === "/api/model")
    return Response.json({ data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
  if (request.method === "POST" && url.pathname === "/api/dochaus/generate") {
    const { prompt } = await request.json() as { prompt: string }
    const thread = JSON.parse(prompt.match(/<thread encoding="json">([\s\S]*?)<\/thread>/)?.[1] ?? "[]") as { text: string }[]
    const row = { root: thread[0]?.text ?? "", flagged: prompt.includes("was flagged as instructions addressed to an AI"), aborted: false }
    generated.push(row)
    const draft = DRAFTS[row.root]
    if (!draft) { unexpected.push(`generate for ${row.root}`); return new Response(null, { status: 500 }) }
    running++
    mostAtOnce = Math.max(mostAtOnce, running)
    try {
      if (gate) await Promise.race([gate.opened, new Promise(resolve => request.signal.addEventListener("abort", resolve, { once: true }))])
      if (request.signal.aborted) { row.aborted = true; return new Response(null, { status: 499 }) }
      return Response.json({ data: { text: JSON.stringify(draft), model: { providerID: "fixture", id: "scripted" }, elapsed_ms: 1, usage: null } })
    } finally {
      running--
    }
  }
  unexpected.push(`engine fixture ${request.method} ${url.pathname}`)
  return new Response(null, { status: 404 })
} })

const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-counterparty-replies-e2e-"))
const dataRoot = path.join(temporary, "data")
const settings = path.join(temporary, "settings")
await mkdir(path.join(dataRoot, "workspace", ".preferences"), { recursive: true })
await mkdir(settings, { recursive: true })
await writeFile(path.join(settings, "opencode.json"), JSON.stringify({ model: "fixture/primary", small_model: "fixture/fast" }))
// The attorney name is saved: replies are signed with it and a thread the
// lawyer answered last awaits nothing.
await writeFile(path.join(dataRoot, "workspace", ".preferences", "preferences.json"), JSON.stringify({ firm: "Synthetic Firm", attorney: LAWYER }))
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

const file = { dir: "", matter: "" }
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
    const pathname = new URL(request.url()).pathname
    if (request.method() !== "POST") return
    if (pathname.endsWith("/documents/markup/decisions")) decisionRequests.push({ actor: request.headers()["x-dochaus-actor"], body: request.postDataJSON() })
    if (pathname.endsWith("/documents/markup/replies/draft")) {
      const body = request.postDataJSON() as { thread: number; decisions?: unknown[]; base_sha256: string }
      draftRequests.push({ thread: body.thread, decisions: body.decisions?.length ?? 0, base: body.base_sha256 })
    }
  })
  page.on("requestfailed", request => {
    if (new URL(request.url()).pathname.endsWith("/documents/markup/replies/draft") && /aborted/i.test(request.failure()?.errorText ?? ""))
      abortedDrafts.push((request.postDataJSON() as { thread: number }).thread)
  })
  const shot = (name: string) => page.screenshot({ path: path.join(evidence, `${name}.png`) })
  // Centred in the panel: scrolled only into view, a card's top sits under the pinned head.
  const centre = (target: Locator) => target.evaluate(node => node.scrollIntoView({ block: "center" }))
  // A tall window for the shot, so the whole card shows between the pinned head and foot.
  const cardShot = async (card: Locator, name: string) => {
    const size = page.viewportSize()!
    await page.setViewportSize({ width: size.width, height: 1800 })
    await centre(card)
    await card.screenshot({ path: path.join(evidence, `${name}.png`) })
    await page.setViewportSize(size)
  }
  capture = shot
  const dialog = page.getByRole("dialog", { name: NAME, exact: true })
  const panel = dialog.locator("#document-markup")
  const changes = panel.locator(".markup-list > li:not(.markup-thread)")
  const threads = panel.locator(".markup-list > li.markup-thread")
  const toggle = dialog.getByRole("button", { name: /^Their changes \(\d+\)$/ })
  const tally = panel.locator(".markup-tally")
  const thread = (text: string) => threads.filter({ has: page.locator(":scope > .markup-comment-text", { hasText: text }) })
  const replyArea = (text: string) => thread(text).getByRole("group", { name: /^Your reply to / })
  const draftAll = panel.getByRole("button", { name: /^Draft replies \(\d+\)$/ })
  const [answered, planted, footnote] = [AVERY_ROOT, PLANTED, FOOTNOTE].map(replyArea) as [Locator, Locator, Locator]

  async function openDocument(navigate = true) {
    if (navigate) await page.goto(`${baseURL!.replace(/\/$/, "")}/matter/${file.matter}`)
    const documents = page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ })
    await expect(documents).toBeVisible()
    if (await documents.getAttribute("aria-expanded") !== "true") await documents.click()
    await page.getByRole("button", { name: NAME, exact: true }).click()
    await expect(dialog).toBeVisible()
    // The first open loads the Word runtime; later renders are quick.
    await expect(dialog.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
    await expect(panel).toBeVisible()
  }

  async function threadReplies() {
    const markup = await (await api(`/matters/${file.matter}/documents/markup?name=${encodeURIComponent(NAME)}`)).json() as { comments: { id: number; author: string; text: string; parent: number | null }[] }
    return markup.comments.filter(comment => comment.author === LAWYER).map(({ parent, text }) => ({ parent, text }))
  }

  await step("1 upload the marked-up agreement with its footnote comment as a source", async () => {
    const matter = await (await api("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Synthetic Counterparty Replies" }) })).json() as { id: string; dir: string }
    expect(matter.dir.startsWith(dataRoot)).toBe(true)
    Object.assign(file, { matter: matter.id, dir: matter.dir })
    const source = new FormData(); source.append("file", new File([counterpartyMarkupDocx({ footnote: true })], NAME))
    expect((await api(`/matters/${matter.id}/documents`, { method: "POST", body: source })).status).toBe(200)
  })

  await step("2 nothing drafts on open: each open thread offers Draft reply and Write a reply", async () => {
    await openDocument()
    await expect(toggle).toHaveAccessibleName("Their changes (14)")
    await expect(threads).toHaveCount(4)
    await expect(draftAll).toHaveAccessibleName("Draft replies (3)")
    for (const area of [answered, planted, footnote]) {
      await expect(area.getByRole("button", { name: "Draft reply", exact: true })).toBeVisible()
      await expect(area.getByRole("button", { name: "Write a reply", exact: true })).toBeVisible()
    }
    // A resolved thread takes no reply here.
    await expect(thread(RESOLVED)).toContainText("Resolved")
    await expect(thread(RESOLVED).getByRole("group", { name: /^Your reply to / })).toHaveCount(0)
    await expect(tally).toHaveText("0 to accept · 0 to reject · 10 held")
    expect(generated).toEqual([])
    await shot("01-no-drafts")
  })

  await step("3 Draft replies drafts the three open threads at once; a cancelled draft never lands", async () => {
    let open = () => {}
    gate = { opened: new Promise<void>(resolve => { open = resolve }), open: () => open() }
    await draftAll.click()
    await expect(panel.getByRole("status").filter({ hasText: "Drafting 3 replies…" })).toBeVisible()
    await expect(panel.getByRole("button", { name: "Cancel drafting", exact: true })).toBeFocused()
    for (const area of [answered, planted, footnote]) await expect(area).toContainText("Drafting a reply…")
    await expect.poll(() => generated.length).toBe(3)
    expect(mostAtOnce).toBe(3)
    // Every quoted string is read by the service; the browser sends ids and the decisions.
    expect(draftRequests.map(row => row.thread).sort()).toEqual([0, 3, 4])
    expect(draftRequests.every(row => row.decisions === 0)).toBe(true)
    await shot("02-drafting")
    await footnote.getByRole("button", { name: "Cancel", exact: true }).click()
    await expect(footnote.getByRole("button", { name: "Draft reply", exact: true })).toBeFocused()
    await expect(panel.getByRole("status").filter({ hasText: "Drafting 2 replies…" })).toBeVisible()
    // The browser ends the request. Vite's dev proxy does not pass that on to
    // the service (the shipped stack calls ingest directly), so whether the
    // model call stopped too is only recorded here.
    await expect.poll(() => abortedDrafts).toEqual([4])
    gate.open()
    gate = undefined
    await expect(answered.getByLabel("Your reply", { exact: true })).toHaveValue(DRAFTS[AVERY_ROOT]!.reply)
    await expect(planted.getByLabel("Your reply", { exact: true })).toHaveValue(DRAFTS[PLANTED]!.reply)
    // The cancelled draft never lands, even once its model call answers.
    await expect.poll(() => generated.filter(row => row.root === FOOTNOTE).length).toBe(1)
    await page.waitForTimeout(500)
    await expect(footnote.getByLabel("Your reply", { exact: true })).toHaveCount(0)
    await expect(panel.getByRole("status").filter({ hasText: "2 draft replies ready — read each one, then approve it." })).toBeVisible()
    await expect(draftAll).toHaveAccessibleName("Draft replies (1)")
  })

  await step("4 a draft lists what to check, and Approve waits until the bracketed part is filled", async () => {
    await expect(answered.locator(".markup-open-items")).toContainText("Check before sending")
    await expect(answered.locator(".markup-open-items")).toContainText(DRAFTS[AVERY_ROOT]!.open_items[0]!)
    const approve = answered.getByRole("button", { name: "Approve", exact: true })
    await expect(approve).toBeDisabled()
    await expect(answered).toContainText("Fill the bracketed parts first.")
    await expect(approve).toHaveAccessibleDescription("Fill the bracketed parts first.")
    await cardShot(thread(AVERY_ROOT), "03-draft-with-open-items")
    await shot("04-placeholder-blocks-approve")
  })

  await step("5 the flagged thread still drafts, with its flag repeated above the draft", async () => {
    expect(generated.find(row => row.root === PLANTED)?.flagged).toBe(true)
    expect(generated.filter(row => row.root !== PLANTED).every(row => !row.flagged)).toBe(true)
    await expect(planted.locator(".markup-flag")).toHaveText(FLAG)
    await expect(planted.locator(".markup-flag + .markup-reply-label")).toHaveCount(1)
    await expect(planted.locator(".markup-open-items")).toContainText(DRAFTS[PLANTED]!.open_items[0]!)
    await cardShot(thread(PLANTED), "05-flagged-thread-draft")
  })

  await step("6 one thread drafted on its own lands in its reply box with focus", async () => {
    await footnote.getByRole("button", { name: "Draft reply", exact: true }).click()
    await expect(footnote.getByLabel("Your reply", { exact: true })).toHaveValue(DRAFTS[FOOTNOTE]!.reply)
    await expect(footnote.getByLabel("Your reply", { exact: true })).toBeFocused()
    await expect(draftAll).toHaveCount(0)
    expect(generated.filter(row => row.root === FOOTNOTE)).toHaveLength(2)
  })

  await step("7 edit the placeholder out and approve two replies", async () => {
    await answered.getByLabel("Your reply", { exact: true }).fill(EDITED)
    await expect(answered.getByText("Fill the bracketed parts first.")).toHaveCount(0)
    await answered.getByRole("button", { name: "Approve", exact: true }).click()
    await expect(answered).toContainText("Reply approved, added when you apply")
    await expect(answered.locator(".markup-reply-final")).toHaveText(EDITED)
    await expect(answered.getByRole("button", { name: "Edit", exact: true })).toBeFocused()
    await planted.getByRole("button", { name: "Approve", exact: true }).click()
    await expect(planted.locator(".markup-reply-final")).toHaveText(DRAFTS[PLANTED]!.reply)
    await expect(tally).toHaveText("0 to accept · 0 to reject · 10 held · 2 replies to add")
    await centre(thread(AVERY_ROOT))
    await shot("06-approved")
    await cardShot(thread(AVERY_ROOT), "06-approved-card")
  })

  await step("8 the replies survive a reload, approved ones still approved", async () => {
    await page.reload()
    await openDocument(false)
    await expect(answered.locator(".markup-reply-final")).toHaveText(EDITED)
    await expect(planted.locator(".markup-reply-final")).toHaveText(DRAFTS[PLANTED]!.reply)
    await expect(footnote.getByLabel("Your reply", { exact: true })).toHaveValue(DRAFTS[FOOTNOTE]!.reply)
    await expect(panel.locator(".markup-hint[role=status]")).toHaveCount(0)
    const stored = await page.evaluate(key => localStorage.getItem(key), `dh.markup-replies.v1/${encodeURIComponent(file.matter)}/${encodeURIComponent(NAME)}`)
    expect(Object.keys((JSON.parse(stored!) as { replies: object }).replies).sort()).toEqual(["0", "3", "4"])
  })

  await step("9 discard a draft, then bring it back with Undo", async () => {
    await footnote.getByRole("button", { name: "Discard", exact: true }).click()
    await expect(footnote.getByRole("status")).toHaveText("Draft discarded. Undo")
    await expect(footnote.getByRole("button", { name: "Draft reply", exact: true })).toBeFocused()
    await expect(draftAll).toHaveAccessibleName("Draft replies (1)")
    await cardShot(thread(FOOTNOTE), "07-discarded-with-undo")
    await footnote.getByRole("button", { name: "Undo", exact: true }).click()
    await expect(footnote.getByLabel("Your reply", { exact: true })).toHaveValue(DRAFTS[FOOTNOTE]!.reply)
    await expect(footnote.getByLabel("Your reply", { exact: true })).toBeFocused()
  })

  await step("10 apply with one Reject: one version with the decision and the two approved replies", async () => {
    const formatting = changes.filter({ hasText: "Formatting changed" })
    await formatting.getByRole("radio", { name: "Reject", exact: true }).click()
    await expect(tally).toHaveText("0 to accept · 1 to reject · 9 held · 2 replies to add")
    await panel.getByRole("button", { name: "Apply decisions…", exact: true }).click()
    const confirm = panel.getByRole("group", { name: "Apply decisions", exact: true })
    await expect(confirm.locator("p").first()).toHaveText(`This saves a new version of ${NAME} with 0 changes accepted, 1 rejected and 2 replies added as ${LAWYER}. ` +
      "9 held changes stay as tracked changes. 1 draft reply is not approved, so it is not added. The current version stays in the history.")
    await confirm.scrollIntoViewIfNeeded()
    await shot("08-apply-confirmation")
    const applied = page.waitForResponse(response => new URL(response.url()).pathname.endsWith("/documents/markup/decisions"))
    await confirm.getByRole("button", { name: "Apply decisions", exact: true }).click()
    const response = await applied
    expect(response.status()).toBe(200)
    expect((await response.json() as { replies: { thread: number }[] }).replies.map(row => row.thread).sort()).toEqual([0, 3])
    expect(decisionRequests).toHaveLength(1)
    const [sent] = decisionRequests
    expect(sent!.actor).toBe(LAWYER)
    expect(sent!.body.decisions).toHaveLength(10)
    expect(sent!.body.decisions.filter(row => row.decision === "reject")).toHaveLength(1)
    // Only the approved replies, as edited; the unapproved draft stays here.
    expect(sent!.body.replies).toEqual([{ thread: 0, text: EDITED }, { thread: 3, text: DRAFTS[PLANTED]!.reply }])
    await expect(page.getByText("Saved as version 2: 0 accepted, 1 rejected, 9 held; 2 replies added.", { exact: true })).toBeVisible()
  })

  await step("11 the threads end with the lawyer's replies; the unsent draft is kept", async () => {
    await expect(toggle).toHaveAccessibleName("Their changes (13)", { timeout: 30_000 })
    const mine = thread(AVERY_ROOT).locator(".markup-replies > li")
    await expect(mine).toHaveCount(2)
    await expect(mine.nth(0)).toContainText(MORGAN_REPLY)
    await expect(mine.nth(1)).toContainText(LAWYER)
    await expect(mine.nth(1)).toContainText(EDITED)
    await expect(thread(PLANTED).locator(".markup-replies > li")).toContainText(DRAFTS[PLANTED]!.reply)
    // Answered last by the lawyer: no draft offered, a reply can still be written.
    for (const area of [answered, planted]) {
      await expect(area.getByRole("button", { name: "Draft reply", exact: true })).toHaveCount(0)
      await expect(area.getByRole("button", { name: "Write a reply", exact: true })).toBeVisible()
    }
    await expect(footnote.getByLabel("Your reply", { exact: true })).toHaveValue(DRAFTS[FOOTNOTE]!.reply)
    await expect(panel.getByText("Carried over")).toHaveCount(0)
    await expect(draftAll).toHaveCount(0)
    expect(await threadReplies()).toEqual([{ parent: 0, text: EDITED }, { parent: 3, text: DRAFTS[PLANTED]!.reply }])
    await centre(thread(AVERY_ROOT))
    await shot("09-after-apply-replies-threaded")
    // A reply written by hand needs text before it can be approved.
    await answered.getByRole("button", { name: "Write a reply", exact: true }).click()
    const typed = answered.getByLabel("Your reply", { exact: true })
    await expect(typed).toBeFocused()
    await expect(answered.getByRole("button", { name: "Approve", exact: true })).toBeDisabled()
    await expect(answered).toContainText("Write the reply first.")
    await typed.fill("Thank you.")
    await expect(answered.getByRole("button", { name: "Approve", exact: true })).toBeEnabled()
    await answered.getByRole("button", { name: "Discard", exact: true }).click()
    await expect(answered.getByRole("button", { name: "Write a reply", exact: true })).toBeFocused()
  })

  await step("12 the versions list shows the new version with the actor and its replies", async () => {
    await dialog.getByRole("button", { name: /^Versions/ }).click()
    const newest = dialog.locator(".versions-list > li").first()
    await expect(newest).toContainText("Version 2")
    await expect(newest.locator(".version-meta")).toHaveText(`Their changes decided · by ${LAWYER}`)
    await expect(newest).toContainText("Tracked changes resolved: 0 accepted, 1 rejected, 9 held; 2 replies added")
    const listed = await (await api(`/matters/${file.matter}/documents/versions?name=${encodeURIComponent(NAME)}`)).json() as { versions: { number: number; origin: string; actor: string }[] }
    expect(listed.versions.map(row => [row.number, row.origin, row.actor])).toEqual([[2, "accept", LAWYER], [1, "upload", expect.any(String)]])
    await shot("10-versions")
    await toggle.click()
    await expect(panel).toBeVisible()
  })

  await step("13 at phone width the last reply is approved and applied alone, with Apply in reach", async () => {
    await page.setViewportSize({ width: 400, height: 860 })
    await page.evaluate(() => localStorage.setItem("dh.sidebar", "1"))
    await openDocument()
    await expect(panel).toBeInViewport()
    await centre(thread(FOOTNOTE))
    await shot("11-phone-draft")
    await footnote.getByRole("button", { name: "Approve", exact: true }).click()
    await expect(tally).toHaveText("0 to accept · 0 to reject · 9 held · 1 reply to add")
    await shot("12-phone-approved")
    const apply = panel.getByRole("button", { name: "Apply decisions…", exact: true })
    await apply.scrollIntoViewIfNeeded()
    await expect(apply).toBeEnabled()
    expect(await apply.evaluate(node => {
      const box = node.getBoundingClientRect()
      const hit = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2)
      return Boolean(hit && node.contains(hit))
    })).toBe(true)
    expect(await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth)).toBeLessThanOrEqual(1)
    await apply.click()
    const confirm = panel.getByRole("group", { name: "Apply decisions", exact: true })
    await expect(confirm.locator("p").first()).toHaveText(`This saves a new version of ${NAME} with 1 reply added as ${LAWYER}. 9 held changes stay as tracked changes. The current version stays in the history.`)
    await shot("13-phone-apply-confirmation")
    await confirm.getByRole("button", { name: "Apply decisions", exact: true }).click()
    await expect(page.getByText("Saved as version 3: 1 reply added.", { exact: true })).toBeVisible()
    await expect(thread(FOOTNOTE).locator(".markup-replies > li")).toContainText(DRAFTS[FOOTNOTE]!.reply, { timeout: 30_000 })
    expect(decisionRequests.at(-1)!.body.replies).toEqual([{ thread: 4, text: DRAFTS[FOOTNOTE]!.reply }])
    expect(decisionRequests.at(-1)!.body.decisions.every(row => row.decision === "hold")).toBe(true)
    const listed = await (await api(`/matters/${file.matter}/documents/versions?name=${encodeURIComponent(NAME)}`)).json() as { versions: { number: number; origin: string; actor: string; label: string | null }[] }
    expect([listed.versions[0]!.number, listed.versions[0]!.origin, listed.versions[0]!.actor, listed.versions[0]!.label]).toEqual([3, "revision", LAWYER, "1 reply added"])
    expect(await threadReplies()).toEqual([{ parent: 0, text: EDITED }, { parent: 3, text: DRAFTS[PLANTED]!.reply }, { parent: 4, text: DRAFTS[FOOTNOTE]!.reply }])
    await centre(thread(FOOTNOTE))
    await shot("14-phone-after-replies-only-apply")
    expect(await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth)).toBeLessThanOrEqual(1)
  })

  await step("14 colour schemes: a dark-scheme request and forced colours", async () => {
    await page.setViewportSize({ width: 1440, height: 1000 })
    await footnote.scrollIntoViewIfNeeded()
    await answered.getByRole("button", { name: "Write a reply", exact: true }).click()
    await answered.getByLabel("Your reply", { exact: true }).fill("Noted, thank you. [[confirm with client]]")
    await planted.getByRole("button", { name: "Write a reply", exact: true }).click()
    await planted.getByLabel("Your reply", { exact: true }).fill("Noted.")
    await planted.getByRole("button", { name: "Approve", exact: true }).click()
    await centre(thread(AVERY_ROOT))
    await page.emulateMedia({ colorScheme: "dark" })
    await shot("15-dark-scheme-request")
    await page.emulateMedia({ colorScheme: "light", forcedColors: "active" })
    await shot("16-forced-colors")
    await page.emulateMedia({ forcedColors: "none" })
  })

  await step("no page errors or unexpected requests", async () => {
    expect(pageErrors).toEqual([])
    expect(unexpected).toEqual([])
  })
}
finally {
  await writeFile(path.join(evidence, "counterparty-replies-results.json"), JSON.stringify({
    results, pageErrors, consoleErrors, unexpected, stubbed: [...new Set(stubbed)],
    engineEndpoints: Object.entries(Object.groupBy(engineRequests, row => `${row.method} ${row.path}`)).map(([endpoint, rows]) => ({ endpoint, count: rows!.length })),
    generated: generated.map(({ root, flagged, aborted }) => ({ thread: root.slice(0, 32), flagged, aborted })), mostAtOnce, draftRequests, abortedDrafts,
    decisionRequests: decisionRequests.map(({ actor, body }) => ({ actor, decisions: body.decisions.length, replies: body.replies?.map(row => row.thread) ?? [] })),
  }, null, 2) + "\n")
  await browser?.close().catch(() => undefined)
  await stopIngest().catch(error => console.error(error))
  engine.stop(true)
  await rm(temporary, { recursive: true, force: true })
}
