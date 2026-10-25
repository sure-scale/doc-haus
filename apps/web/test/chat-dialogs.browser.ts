// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/chat-dialogs.browser.ts
// Uses the real matter page and chat panel; only the engine's and ingest's HTTP and event stream are mocked.
// Set CHAT_DIALOGS_SCREENSHOTS=<directory> to save each state, captured before it is checked.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Page } from "@playwright/test"
import blockedQuestion from "../../../dochaus/workflows/lean/fixtures/status-blocked-question.json"
import start from "../../../dochaus/workflows/lean/fixtures/start.json"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-chat-dialogs-"))
const shots = process.env.CHAT_DIALOGS_SCREENSHOTS
const sessionID = start.status.owner_session_id
const model = { providerID: "test", id: "model" }
// The run starts working, then stops to ask the lawyer two questions.
const working = start.status
const asking = {
  ...blockedQuestion,
  queued_answer_ids: [],
  questions: [
    blockedQuestion.questions[0],
    { ...blockedQuestion.questions[0], id: "q-storage", question: "Who bears the cost of storage?" },
  ],
}
let drafting: unknown = null

const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter, Route, Routes } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import MatterDetail from ${source("pages/MatterDetail.tsx")}
import { ToastProvider } from ${source("components/Toast.tsx")}
createRoot(document.getElementById("root")).render(createElement(ToastProvider, null,
  createElement(MemoryRouter, { initialEntries: ["/matter/matter?view=chat&session=${sessionID}"] },
    createElement(Routes, null, createElement(Route, { path: "/matter/:id", element: createElement(MatterDetail, { onSessionsChanged: () => {} }) })))))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.path.endsWith(".js"))!.text()
// The app's own stylesheets, in its order, so the captures look like the app.
const main = await Bun.file(path.resolve(import.meta.dir, "../src/main.tsx")).text()
const css = (await Promise.all([...main.matchAll(/^import "\.\/(.+\.css)"$/gm)]
  .map(([, file]) => Bun.file(path.resolve(import.meta.dir, "../src", file!)).text()))).join("\n")
const json = (body: unknown) => Response.json(body)
const session = { id: sessionID, projectID: "project", agent: "qa", model, location: { directory: "/matter" }, cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: 1, updated: 1 } }
const transcript = [{ type: "user", id: "msg_u1", text: "Please prepare our revised supply agreement.", time: { created: 1 } }]
const unknown: string[] = []
const server = Bun.serve({
  port: 0,
  hostname: "127.0.0.1",
  idleTimeout: 0,
  fetch(request) {
    const url = new URL(request.url)
    const route = `${request.method} ${url.pathname}`
    if (route === "GET /harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (route === "GET /") return new Response(`<html><meta charset="utf-8"><style>body{font-family:sans-serif;margin:0}${css}</style><div id="root" style="display:flex;flex-direction:column;height:100vh"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (route === "GET /ingest/matters/matter")
      return json({ id: "matter", title: "Supply agreement", dir: "/matter", created_at: 1, documents: [{ id: 1, name: "Supply terms.docx", doc_path: "/matter/Supply terms.docx", created_at: 1, role: "source" }] })
    if (route === "GET /ingest/jurisdictions" || route === "GET /ingest/playbooks") return json([])
    if (route === "GET /ingest/workflows") return json({ dir: "/workflows", workflows: [] })
    if (route === "GET /api/agent") return json({ data: ["qa", "redline", "research"].map(id => ({ id, description: id, mode: "primary" })) })
    if (route === "GET /ingest/matters/matter/telemetry")
      return json({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 1, since_ts: 1 }, current_session: null, rate_card_as_of: null, unknown_models: [] })
    if (route === "GET /ingest/matters/matter/drafting-status") return json(drafting)
    if (url.pathname.startsWith("/ingest/matters/matter/answer-drafts/")) return json({ answers: {}, revision: 0 })
    if (route === "GET /ingest/engine/config") return json({ model: "test/model", small_model: "test/model" })
    if (route === "GET /api/event") return new Response(new ReadableStream<Uint8Array>(), { headers: { "content-type": "text/event-stream" } })
    if (route === "GET /api/session/active") return json({ data: {} })
    if (route === "GET /api/permission/request") return json({ location: { directory: "/matter" }, data: [] })
    if (route === `GET /api/session/${sessionID}/form`) return json({ data: [] })
    if (route === `GET /api/session/${sessionID}`) return json({ data: session })
    if (route === `GET /api/session/${sessionID}/message`) return json({ data: transcript, cursor: {} })
    // The run's step sessions, read for the run's progress, say nothing here.
    if (route === "GET /api/session" || /^GET \/api\/session\/ses_child\d+\/message$/.test(route)) return json({ data: [], cursor: {} })
    if (route === `POST /api/session/${sessionID}/agent` || route === `POST /api/session/${sessionID}/model`) return new Response(null, { status: 204 })
    unknown.push(route)
    return new Response(null, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const widths = [[1440, 1000], [400, 860]] as const
const failures: string[] = []
// Each check gets its own page and runs even when an earlier one failed.
async function check(name: string, width: number, height: number, run: (page: Page, shot: (label: string) => Promise<void>) => Promise<void>) {
  const context = await browser.newContext({ viewport: { width, height }, reducedMotion: "reduce" })
  await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  const page = await context.newPage()
  page.setDefaultTimeout(5000)
  const shot = async (label: string) => {
    if (shots) await page.screenshot({ path: path.join(shots, `${label}-${width}.png`) })
  }
  try {
    await run(page, shot)
    console.log(`PASS ${name} (${width}px)`)
  } catch (error) {
    failures.push(`${name} (${width}px)`)
    console.log(`FAIL ${name} (${width}px): ${error instanceof Error ? error.message.split("\n").slice(0, 3).join(" | ") : error}`)
  } finally {
    await context.close()
  }
}
const focusWithin = (page: Page, selector: string) => page.evaluate(selector => Boolean(document.activeElement?.closest(selector)), selector)
// The run's status poll runs every five seconds.
const poll = { timeout: 12_000 }
try {
  for (const [width, height] of widths) {
    await check("the assistant picker is a labelled dialog that takes focus, keeps it and gives it back", width, height, async (page, shot) => {
      drafting = null
      await page.goto(server.url.href)
      const trigger = page.getByRole("button", { name: "Auto", exact: true })
      await trigger.click()
      await page.getByText("Choose an assistant", { exact: true }).waitFor()
      await shot("assistant-picker")
      const picker = page.getByRole("dialog", { name: "Choose an assistant" })
      await picker.waitFor()
      assert.equal(await picker.getAttribute("aria-modal"), "true")
      assert.ok(await focusWithin(page, ".picker-panel"), "focus moves into the picker")
      for (let step = 0; step < 6; step++) {
        await page.keyboard.press("Tab")
        assert.ok(await focusWithin(page, ".picker-panel"), "Tab stays inside the picker")
      }
      await page.keyboard.press("Shift+Tab")
      assert.ok(await focusWithin(page, ".picker-panel"), "Shift+Tab stays inside the picker")
      await page.keyboard.press("Escape")
      await picker.waitFor({ state: "detached" })
      assert.ok(await trigger.evaluate(button => button === document.activeElement), "Escape hands focus back to the picker's button")
      await trigger.click()
      await picker.getByRole("button", { name: /^Q&A/ }).click()
      await picker.waitFor({ state: "detached" })
      await page.getByRole("button", { name: "Q&A", exact: true }).waitFor()
      // Choosing an assistant leaves the lawyer in the message box, as before.
      await page.waitForFunction(() => document.activeElement?.getAttribute("aria-label") === "Message")
    })

    await check("a run that stops to ask opens its questions", width, height, async (page, shot) => {
      drafting = working
      await page.goto(server.url.href)
      await page.getByRole("button", { name: "Pause drafting", exact: true }).waitFor()
      // Focus stays in the composer the request was sent from, now empty.
      await page.getByRole("textbox", { name: "Message" }).click()
      drafting = asking
      await page.locator(".question-dock").waitFor(poll)
      await page.waitForTimeout(500)
      await shot("run-asks")
      const questions = page.getByRole("dialog", { name: "Question 1 of 2" })
      await questions.waitFor()
      assert.ok(await focusWithin(page, "dialog.question-dialog"), "focus moves to the question")
      await page.keyboard.press("Escape")
      await questions.waitFor({ state: "hidden" })
      // Closed, the same questions wait in the dock; the next poll does not reopen them.
      await page.waitForTimeout(6000)
      assert.equal(await questions.isVisible(), false, "a closed dialog stays closed for the same questions")
      await shot("run-asks-closed")
      await page.getByRole("button", { name: "Answer 2", exact: true }).click()
      await questions.waitFor()
    })

    await check("questions asked while the lawyer types wait in the dock", width, height, async (page, shot) => {
      drafting = working
      await page.goto(server.url.href)
      const composer = page.getByRole("textbox", { name: "Message" })
      await composer.click()
      await page.keyboard.type("Keep the buyer's liability cap")
      drafting = asking
      await page.locator(".question-dock").waitFor(poll)
      await page.keyboard.type(" as drafted.")
      await shot("run-asks-while-typing")
      assert.equal(await page.getByRole("dialog").count(), 0, "no dialog takes the keyboard")
      assert.equal(await composer.inputValue(), "Keep the buyer's liability cap as drafted.")
      assert.ok(await composer.evaluate(field => field === document.activeElement), "the composer keeps focus")
      await page.getByRole("button", { name: "Answer 2", exact: true }).click()
      await page.getByRole("dialog", { name: "Question 1 of 2" }).waitFor()
    })
  }
  assert.deepEqual(unknown, [])
  assert.deepEqual(failures, [])
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
