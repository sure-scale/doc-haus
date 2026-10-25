// After "Their changes" Apply (#200, #198) against a separately started Vite
// server. The runner starts a real ingest service on DOCHAUS_E2E_INGEST_PORT
// with a temporary data root, a settings directory and a saved attorney name,
// uploads a synthetic services agreement whose clauses Word numbers (clause 5
// cites clause 7, and clause 7 is a tracked deletion), and applies it three
// ways: accepting the deletion shows the notice of the reference it broke and
// an info toast; rejecting it shows nothing; and a check that could not run (the
// response's `consistency` rewritten in the browser, as ingest's own refusal is
// covered by services/ingest/src/document-markup.test.ts) shows one muted line.
// Each state is captured at desktop and 400×860. Nothing reaches a model: the
// sources' embeddings use the local model (DOCHAUS_MODEL_CACHE_DIR, else the
// per-user cache), browser engine traffic is intercepted and unexpected
// requests fail the run. Screenshots and the JSON summary go beneath
// DOCHAUS_EVIDENCE_DIR.
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises"
import os, { tmpdir } from "node:os"
import path from "node:path"
import type { Browser, Locator } from "@playwright/test"
import { numberedMarkupDocx } from "../../../dochaus/test/support/counterparty-markup-fixture"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")

const baseURL = process.env.PLAYWRIGHT_BASE_URL
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
const ingestPort = Number(process.env.DOCHAUS_E2E_INGEST_PORT)
if (!baseURL || !evidence || !Number.isInteger(ingestPort) || ingestPort < 1024)
  throw new Error("Set PLAYWRIGHT_BASE_URL, DOCHAUS_EVIDENCE_DIR and DOCHAUS_E2E_INGEST_PORT (the port Vite's /ingest proxy targets).")
await mkdir(evidence, { recursive: true })

const LAWYER = "Synthetic Reviewing Lawyer"
const BROKEN = "Renumbered agreement.docx"
const CLEAN = "Unchanged agreement.docx"
const UNCHECKED = "Unchecked agreement.docx"
const HEADING = "Check these references and defined terms after your decisions"
const UNCHECKED_LINE = "References and defined terms were not checked for this version."
const DESKTOP = { width: 1440, height: 1000 }
const PHONE = { width: 400, height: 860 }

type Consistency = { status: "checked"; findings: { detail: string }[] } | { status: "not_checked" }
const results: { step: string; status: "passed" | "failed"; error?: string }[] = []
const pageErrors: string[] = []
const consoleErrors: string[] = []
const unexpected: string[] = []
const stubbed: string[] = []
const applied: Record<string, { status: number; consistency?: Consistency; forced?: boolean }> = {}
const shots: string[] = []

// The service needs the engine only for its model list (readiness); anything
// else that reaches it is recorded.
const engine = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: (request) => {
  const url = new URL(request.url)
  if (request.method === "GET" && url.pathname === "/api/model")
    return Response.json({ data: ["primary", "fast"].map(id => ({ providerID: "fixture", id, name: id, enabled: true })) })
  unexpected.push(`engine fixture ${request.method} ${url.pathname}`)
  return new Response(null, { status: 404 })
} })

const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-counterparty-followups-e2e-"))
const dataRoot = path.join(temporary, "data")
const settings = path.join(temporary, "settings")
await mkdir(path.join(dataRoot, "workspace", ".preferences"), { recursive: true })
await mkdir(settings, { recursive: true })
await writeFile(path.join(settings, "opencode.json"), JSON.stringify({ model: "fixture/primary", small_model: "fixture/fast" }))
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
  const context = await browser!.newContext({ viewport: DESKTOP, reducedMotion: "reduce" })
  await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  const origin = new URL(baseURL).origin
  // Engine traffic is answered here; ingest traffic continues through Vite's
  // /ingest proxy to the isolated service.
  await context.route("**/*", async route => {
    const request = route.request()
    const url = new URL(request.url())
    if (url.origin !== origin) { unexpected.push(`${request.method()} ${request.url()}`); return route.abort() }
    const reply = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (request.method() === "GET" && url.pathname === "/ingest/updates") {
      stubbed.push("GET /ingest/updates")
      return reply({ current: { version: 1, name: "Fixture", notes: "" }, available: false })
    }
    if (url.pathname.startsWith("/ingest/")) return route.continue()
    if (!url.pathname.startsWith("/engine/")) return route.continue()
    const pathname = url.pathname.slice("/engine".length)
    if (request.method() === "GET" && pathname === "/api/dochaus/ready") return reply({ ready: true, consentState: "active" })
    if (request.method() === "POST" && pathname === "/api/dochaus/provider-test") return reply({ ok: true })
    if (request.method() === "GET" && pathname === "/api/session") return reply({ data: [], cursor: { next: null, previous: null } })
    if (request.method() === "GET" && pathname === "/api/agent") return reply({ location: { directory: file.dir }, data: [] })
    if (request.method() === "GET" && pathname === "/api/event") return
    unexpected.push(`engine ${request.method()} ${pathname}${url.search}`)
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: "not stubbed" }) })
  })
  const page = await context.newPage()
  page.setDefaultTimeout(20_000)
  page.on("pageerror", (error: Error) => pageErrors.push(error.message))
  page.on("console", (message: { type: () => string; text: () => string }) => message.type() === "error" && consoleErrors.push(message.text()))
  const shot = async (name: string) => { shots.push(`${name}.png`); await page.screenshot({ path: path.join(evidence, `${name}.png`) }) }
  capture = shot

  async function openDocument(name: string) {
    await page.goto(`${baseURL!.replace(/\/$/, "")}/matter/${file.matter}`)
    const documents = page.locator(".sidebar").getByRole("button", { name: /^Documents(?: \d+)?$/ })
    await expect(documents).toBeVisible()
    if (await documents.getAttribute("aria-expanded") !== "true") await documents.click()
    await page.getByRole("button", { name, exact: true }).click()
    const viewer = page.getByRole("dialog", { name, exact: true })
    await expect(viewer).toBeVisible()
    await expect(viewer.locator(".docx-render:not(.docx-skeleton)")).toBeVisible({ timeout: 30_000 })
    return viewer
  }

  // Sets every listed change of `name` to `decision` and applies them, with
  // the response's `consistency` replaced by `forced` when given.
  async function apply(name: string, decision: "Accept" | "Reject", forced?: Consistency) {
    const viewer = await openDocument(name)
    const panel = viewer.locator("#document-markup")
    await expect(panel).toBeVisible()
    await panel.getByRole("group", { name: /^Set all \d+ changes? to$/ }).getByRole("button", { name: decision, exact: true }).click()
    if (forced) await page.route(url => url.pathname.endsWith("/documents/markup/decisions"), async route => {
      const response = await route.fetch()
      await route.fulfill({ response, json: { ...await response.json(), consistency: forced } })
    }, { times: 1 })
    await panel.getByRole("button", { name: "Apply decisions…", exact: true }).click()
    const done = page.waitForResponse(response => new URL(response.url()).pathname.endsWith("/documents/markup/decisions"))
    await panel.getByRole("group", { name: "Apply decisions", exact: true }).getByRole("button", { name: "Apply decisions", exact: true }).click()
    const response = await done
    applied[name] = { status: response.status(), consistency: (await response.json() as { consistency?: Consistency }).consistency, ...(forced ? { forced: true } : {}) }
    expect(response.status()).toBe(200)
    // Every change was decided: the panel goes once the new version is read.
    await expect(panel).toHaveCount(0, { timeout: 30_000 })
    return viewer
  }

  // The same state at phone width: the viewer keeps it while the window narrows.
  async function phone(viewer: Locator, target: string | undefined, name: string) {
    await page.setViewportSize(PHONE)
    if (target) await viewer.locator(target).scrollIntoViewIfNeeded()
    expect(await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth)).toBeLessThanOrEqual(1)
    await shot(name)
    await page.setViewportSize(DESKTOP)
  }

  await step("1 upload three copies of the numbered agreement as sources", async () => {
    const matter = await (await api("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Synthetic Apply Follow-ups" }) })).json() as { id: string; dir: string }
    expect(matter.dir.startsWith(dataRoot)).toBe(true)
    Object.assign(file, { matter: matter.id, dir: matter.dir })
    for (const name of [BROKEN, CLEAN, UNCHECKED]) {
      const body = new FormData(); body.append("file", new File([numberedMarkupDocx()], name))
      expect((await api(`/matters/${matter.id}/documents`, { method: "POST", body })).status).toBe(200)
    }
  })

  await step("2 accepting the deletion of a cited clause shows the reference it broke, with an info toast", async () => {
    const viewer = await apply(BROKEN, "Accept")
    const consistency = applied[BROKEN]!.consistency
    expect(consistency?.status).toBe("checked")
    const notice = viewer.locator(".markup-consistency")
    await expect(notice).toHaveAttribute("role", "status")
    await expect(notice.locator("strong")).toHaveText(HEADING)
    await expect(notice.getByRole("list", { name: HEADING })).toBeVisible()
    await expect(notice.locator("li")).toHaveCount(1)
    await expect(notice.locator("li")).toHaveText(/^Clause 5 refers to clause 7\b.*which was clause 8 before/)
    await expect(page.getByText("1 reference or defined term to check.", { exact: true })).toBeVisible()
    await expect(page.getByText(/^Saved as version 2: \d+ accepted, 0 rejected, 0 held\.$/)).toBeVisible()
    await shot("6b2-notice-desktop")
    await phone(viewer, ".markup-consistency", "6b2-notice-phone")
    // Viewing a kept version ends the notice; it is not brought back.
    await viewer.getByRole("button", { name: /^Versions/ }).click()
    await viewer.getByRole("button", { name: "View version 1", exact: true }).click()
    await expect(viewer.locator(".version-banner")).toBeVisible()
    await expect(notice).toHaveCount(0)
    await viewer.getByRole("button", { name: "Back to current version", exact: true }).click()
    await expect(viewer.locator(".version-banner")).toHaveCount(0)
    await expect(notice).toHaveCount(0)
  })

  await step("3 a clean Apply (the deletion rejected) shows no notice and no toast", async () => {
    const viewer = await apply(CLEAN, "Reject")
    expect(applied[CLEAN]!.consistency).toEqual({ status: "checked", findings: [] })
    await expect(page.getByText(/^Saved as version 2: 0 accepted, \d+ rejected, 0 held\.$/)).toBeVisible()
    await expect(viewer.locator(".markup-consistency, .markup-consistency-unchecked")).toHaveCount(0)
    await expect(page.getByText(/to check\.$/)).toHaveCount(0)
    await shot("6b2-clean-desktop")
    await phone(viewer, undefined, "6b2-clean-phone")
  })

  await step("4 a check that could not run shows one muted line, not a clean result", async () => {
    const viewer = await apply(UNCHECKED, "Accept", { status: "not_checked" })
    const line = viewer.locator(".markup-consistency-unchecked")
    await expect(line).toHaveText(UNCHECKED_LINE)
    await expect(line).toHaveAttribute("role", "status")
    await expect(viewer.locator(".markup-consistency")).toHaveCount(0)
    await expect(page.getByText(/to check\.$/)).toHaveCount(0)
    await shot("6b2-not-checked-desktop")
    await phone(viewer, ".markup-consistency-unchecked", "6b2-not-checked-phone")
  })

  await step("no page errors or unexpected requests", async () => {
    expect(pageErrors).toEqual([])
    expect(unexpected).toEqual([])
  })
}
finally {
  await writeFile(path.join(evidence, "counterparty-followups-results.json"), JSON.stringify({
    results, pageErrors, consoleErrors, unexpected, stubbed: [...new Set(stubbed)], applied, shots,
  }, null, 2) + "\n")
  await browser?.close().catch(() => undefined)
  await stopIngest().catch(error => console.error(error))
  engine.stop(true)
  await rm(temporary, { recursive: true, force: true })
}
