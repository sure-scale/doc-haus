// Run from apps/web: bun test/review-starter-handoff.browser.ts
// The real matter page and Workflows page in Chromium, with the engine's and
// ingest's HTTP and event stream mocked: starting a review table from the chat's
// Workflows dialog or the Workflows page lands on the table's Add questions with
// the starter chosen and its example and AI suggestions still on offer, and
// nothing runs before the lawyer's own Start, Suggest or Fill. Synthetic data only.
// Set REVIEW_HANDOFF_SCREENSHOTS=<directory> to save each state before it is checked.
import assert from "node:assert/strict"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Page } from "@playwright/test"
import { CHAT_ASSISTANTS, WORKFLOWS } from "../src/agents"
import type { Grid } from "../src/api/ingest/grid"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-review-handoff-"))
const shots = process.env.REVIEW_HANDOFF_SCREENSHOTS
if (shots) await mkdir(shots, { recursive: true })

const entry = path.join(directory, "harness.tsx")
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import { MemoryRouter, Route, Routes, useLocation, useNavigate } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
import MatterDetail from ${source("pages/MatterDetail.tsx")}
import Workflows from ${source("pages/Workflows.tsx")}
import { ToastProvider } from ${source("components/Toast.tsx")}
// The current address, for checks, and a way to move as a link would.
function Where() {
  const location = useLocation()
  window.where = location.pathname + location.search
  window.go = useNavigate()
  return null
}
createRoot(document.getElementById("root")).render(createElement(ToastProvider, null,
  createElement(MemoryRouter, { initialEntries: [new URLSearchParams(location.search).get("at")] },
    createElement(Where),
    createElement(Routes, null,
      createElement(Route, { path: "/matter/:id", element: createElement(MatterDetail, { onSessionsChanged: () => {} }) }),
      // The library pages sit in the app's content column.
      createElement(Route, { path: "/workflows", element: createElement("main", { className: "content" }, createElement("div", { className: "container" }, createElement(Workflows))) })))))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.path.endsWith(".js"))!.text()
// The app's own stylesheets, in its order, then the ones the pages import.
const main = await Bun.file(path.resolve(import.meta.dir, "../src/main.tsx")).text()
const css = (await Promise.all([...main.matchAll(/^import "\.\/(.+\.css)"$/gm)]
  .map(([, file]) => Bun.file(path.resolve(import.meta.dir, "../src", file!)).text()))).join("\n")
  + (await Promise.all(build.outputs.filter((output) => output.path.endsWith(".css")).map((output) => output.text()))).join("\n")

// Three synthetic matters: two agreements, one agreement, and a table that already has a question.
const doc = (id: number, name: string, role: "source" | "deliverable" = "source") => ({ id, name, doc_path: `/synthetic/${name}`, created_at: 1, role })
const matters = {
  pair: { id: "pair", title: "Supply review", dir: "/synthetic/pair", created_at: 3, documents: [doc(1, "Supply agreement.docx"), doc(2, "Distribution agreement.docx"), doc(3, "Review memo.docx", "deliverable")] },
  solo: { id: "solo", title: "Lease advice", dir: "/synthetic/solo", created_at: 2, documents: [doc(1, "Office lease.docx")] },
  asked: { id: "asked", title: "Services review", dir: "/synthetic/asked", created_at: 1, documents: [doc(1, "Services agreement.docx"), doc(2, "Order form.docx")] },
}
const existing: Grid = { version: 1, columns: [{ id: "existing", question: "Who signed the agreement?" }], cells: {} }
let grids: Record<string, Grid> = {}
let puts = 0
// Engine work the dialogs could start: new conversations and prompts.
let sessions = 0
let prompts: string[] = []
const unknown: string[] = []
const json = (body: unknown) => Response.json(body)
const session = { id: "ses_new", projectID: "project", agent: "legal-review", location: { directory: "/synthetic/solo" }, cost: 0, tokens: { input: 0, output: 0, reasoning: 0, cache: { read: 0, write: 0 } }, time: { created: 1, updated: 1 } }
const server = Bun.serve({
  port: 0,
  hostname: "127.0.0.1",
  idleTimeout: 0,
  async fetch(request) {
    const url = new URL(request.url)
    const route = `${request.method} ${url.pathname}`
    if (route === "GET /harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (route === "GET /") return new Response(`<html><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{margin:0}${css}</style><div id="root" style="display:flex;flex-direction:column;height:100vh"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (route === "GET /ingest/matters") return json(Object.values(matters))
    const matter = url.pathname.match(/^\/ingest\/matters\/(pair|solo|asked)(\/.*)?$/)
    if (matter) {
      const [, id, rest = ""] = matter
      if (route.startsWith("GET") && rest === "") return json(matters[id as keyof typeof matters])
      if (rest === "/grid" && request.method === "GET") return json(grids[id!] ?? { version: 0, columns: [], cells: {} })
      if (rest === "/grid" && request.method === "PUT") {
        puts++
        const incoming = await request.json() as Grid
        grids[id!] = { ...incoming, version: (incoming.version ?? 0) + 1 }
        return json(grids[id!])
      }
      if (rest === "/grid/sources") return json({})
      if (rest === "/telemetry") return json({ current_activity: null, today: { cost_usd: 0, turns: 0 }, matter_total: { cost_usd: 0, turns: 0, sessions: 0, since_ts: 1 }, current_session: null, rate_card_as_of: null, unknown_models: [] })
      if (rest === "/drafting-status") return json(null)
      if (rest.startsWith("/answer-drafts/")) return json({ answers: {}, revision: 0 })
      // No chronology and no open Word documents: the pages carry on without them.
      if (rest === "/chronology" || rest === "/chronology/assess") return new Response(null, { status: 404 })
      if (rest === "/documents/word") return json({ documents: [] })
    }
    if (route === "GET /ingest/word/config") return json({ available: false, origin: null, manifestUrl: null })
    if (route === "GET /ingest/jurisdictions") return json([{ code: "england-wales", name: "England and Wales", citationStyle: "oscola" }])
    if (route === "GET /ingest/playbooks") return json([])
    if (route === "GET /ingest/workflows") return json({ dir: "/synthetic/workflows", workflows: [] })
    if (route === "GET /ingest/preferences") return json({ attorney: "" })
    if (route === "GET /ingest/engine/config") return json({ model: "test/model", small_model: "test/model" })
    if (route === "GET /api/agent") return json({ data: [...CHAT_ASSISTANTS, ...WORKFLOWS].map(({ name }) => ({ id: name, description: name, mode: "primary" })) })
    if (route === "GET /api/event") return new Response(new ReadableStream<Uint8Array>(), { headers: { "content-type": "text/event-stream" } })
    if (route === "GET /api/session/active") return json({ data: {} })
    if (route === "GET /api/permission/request") return json({ data: [] })
    if (route === "GET /api/session") return json({ data: [], cursor: {} })
    if (route === "POST /api/session") { sessions++; return json({ data: session }) }
    if (route === "POST /api/session/ses_new/agent" || route === "POST /api/session/ses_new/model") return new Response(null, { status: 204 })
    if (route === "POST /api/session/ses_new/prompt") {
      const { text } = await request.json() as { text: string }
      prompts.push(text)
      return json({ data: { id: "msg_1", sessionID: "ses_new", time: { created: 1 }, type: "user", payload: { text }, delivery: "queue" } })
    }
    if (route === "POST /api/experimental/session/ses_new/wait") return new Response(null, { status: 204 })
    if (route === "GET /api/session/ses_new") return json({ data: session })
    if (route === "GET /api/session/ses_new/message") return json({ data: [{ type: "idle", id: "msg_i", outcome: "succeeded", time: { created: 1 } }], cursor: {} })
    if (route === "GET /api/session/ses_new/form") return json({ data: [] })
    unknown.push(route)
    return new Response(null, { status: 404 })
  },
})

const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const failures: string[] = []
// Each check gets its own page and fresh fixtures, and runs even when an earlier one failed.
async function check(name: string, [width, height]: readonly [number, number], run: (page: Page, shot: (label: string) => Promise<void>) => Promise<void>) {
  grids = { asked: structuredClone(existing) }
  puts = sessions = 0
  prompts = []
  const context = await browser.newContext({ viewport: { width, height }, reducedMotion: "reduce" })
  await context.addInitScript(() => sessionStorage.setItem("dh.launch-notice.v1", "accepted"))
  const page = await context.newPage()
  page.setDefaultTimeout(6000)
  const shot = async (label: string) => {
    if (shots) await page.screenshot({ path: path.join(shots, `${label}-${width}.png`) })
  }
  try {
    await run(page, shot)
    console.log(`PASS ${name} (${width}px)`)
  } catch (error) {
    failures.push(`${name} (${width}px)`)
    console.log(`FAIL ${name} (${width}px): ${error instanceof Error ? error.message.split("\n").slice(0, 8).join(" | ") : error}`)
  } finally {
    await context.close()
  }
}
const open = (page: Page, at: string) => page.goto(`${server.url.href}?at=${encodeURIComponent(at)}`)
const where = (page: Page) => page.evaluate(() => (window as unknown as { where: string }).where)
const param = async (page: Page, name: string) => new URL(await where(page), "http://x").searchParams.get(name)
const go = (page: Page, to: string) => page.evaluate((to) => (window as unknown as { go: (to: string) => void }).go(to), to)
// Nothing scrolls sideways inside a dialog panel.
const fits = (page: Page) => page.locator(".picker-panel").evaluateAll((panels) => panels.every((panel) => [panel, ...panel.querySelectorAll("*")].every((node) => node.scrollWidth <= node.clientWidth + 1 || getComputedStyle(node).overflowX === "visible")))
const launcher = (page: Page) => page.getByRole("dialog", { name: "Workflows" })
const chooser = (page: Page) => page.getByRole("dialog", { name: "Add questions" })
const table = (page: Page, title: string) => page.getByRole("dialog", { name: `Review table — ${title}` })
const openLauncher = async (page: Page) => {
  await page.getByRole("button", { name: "Workflows", exact: true }).click()
  await expect(launcher(page)).toBeVisible()
}

try {
  for (const size of [[1440, 900], [400, 860]] as const) {
    await check("a chat workflow starts only on Start, from its preview", size, async (page, shot) => {
      await open(page, "/matter/solo?view=chat")
      await openLauncher(page)
      const dialog = launcher(page)
      // One document: chat leads with its workflow preselected, so the review table does not compete as Recommended.
      assert.deepEqual(await dialog.locator("legend").allTextContents(), ["In chat", "Review table"])
      await expect(dialog.getByRole("radio", { name: "Full review" })).toBeChecked()
      await expect(dialog.getByRole("radio", { name: /^Add questions/ })).toHaveAccessibleName("Add questions")
      assert.equal(await dialog.getByText("Recommended", { exact: false }).count(), 0, "no Recommended beside a preselected workflow")
      await expect(dialog.getByRole("region", { name: "About Full review" })).toContainText("Reads every document in the matter")
      await expect(dialog.getByRole("button", { name: "Start" })).toHaveClass(/primary/)
      assert.equal(await dialog.locator("button.primary").count(), 1, "one primary action")
      assert.equal(await fits(page), true, "the dialog fits its width")
      await shot("launcher-workflow")
      // Highlighting entries only changes the preview.
      await dialog.getByRole("radio", { name: /^Add questions/ }).check()
      await expect(dialog.getByRole("region", { name: "About Add questions" })).toContainText("Suggest questions with AI")
      await dialog.getByRole("radio", { name: "Full review" }).check()
      await page.waitForTimeout(300)
      assert.deepEqual([sessions, prompts], [0, []], "nothing runs before Start")
      await dialog.getByRole("button", { name: "Start" }).click()
      await expect(dialog).toBeHidden()
      await expect.poll(() => prompts).toEqual([WORKFLOWS[0]!.prompt])
      assert.equal(sessions, 1, "a matter workflow opens one new conversation")
    })

    await check("a review table starts from the chat's Workflows dialog on Add questions, recommended starter chosen", size, async (page, shot) => {
      await open(page, "/matter/pair?view=chat")
      await openLauncher(page)
      const dialog = launcher(page)
      // Two of the lawyer's own documents: the review table leads, recommended and highlighted.
      assert.deepEqual(await dialog.locator("legend").allTextContents(), ["Review table", "In chat"])
      const review = dialog.getByRole("radio", { name: /^Add questions/ })
      await expect(review).toBeChecked()
      await expect(review).toBeFocused()
      await expect(review).toHaveAccessibleName(/Recommended · 2 source documents in this matter/)
      // The preview names every way the table's Add questions offers, AI included.
      const preview = dialog.getByRole("region", { name: "About Add questions" })
      for (const text of ["Ask the same questions of every document", "Document overview: 3 questions", "Compare agreements: 5 questions, with a fictional example", "Suggest questions with AI", "with Compare agreements selected", "no AI runs until you choose"]) await expect(preview).toContainText(text)
      assert.equal(await dialog.locator("button.primary").count(), 1, "one primary action")
      assert.equal(await fits(page), true, "the dialog fits its width")
      await shot("launcher-starter")
      await dialog.getByRole("button", { name: "Start" }).click()
      // The table's own chooser, not a step past it: the recommended starter chosen,
      // its example and the AI option on offer, Continue the one primary.
      const choose = chooser(page)
      await expect(choose.getByRole("radio", { name: /^Compare agreements/ })).toBeChecked()
      await expect(choose.getByRole("radio", { name: /^Compare agreements/ })).toHaveAccessibleName(/Recommended/)
      await expect(choose.getByRole("button", { name: "See example" })).toBeVisible()
      await expect(choose.getByRole("radio", { name: /^Suggest questions with AI/ })).toBeVisible()
      await expect(choose.locator("button.primary")).toHaveText("Continue")
      await expect.poll(() => param(page, "starter")).toBeNull()
      assert.equal(await param(page, "view"), "review")
      await page.waitForTimeout(300)
      await shot("arrival-chooser")
      // Choosing AI here asks for a description; nothing is suggested until the lawyer asks.
      await choose.getByRole("radio", { name: /^Suggest questions with AI/ }).check()
      await choose.getByRole("button", { name: "Continue" }).click()
      await expect(page.getByRole("dialog", { name: "Suggest questions with AI" }).getByRole("textbox")).toBeFocused()
      await page.waitForTimeout(300)
      assert.deepEqual([sessions, prompts, puts], [0, [], 0], "nothing runs or saves on arrival")
      // Consumed once: closing and reopening the table does not reopen the chooser.
      const describe = page.getByRole("dialog", { name: "Suggest questions with AI" })
      await describe.getByRole("button", { name: "Cancel" }).click()
      await expect(describe).toBeHidden()
      await table(page, "Supply review").getByRole("button", { name: "Close" }).click()
      await expect(table(page, "Supply review")).toBeHidden()
      assert.equal(await param(page, "view"), null)
      await go(page, "/matter/pair?view=review")
      await expect(table(page, "Supply review").locator(".review-toolbar").getByRole("button", { name: "Add questions", exact: true })).toBeEnabled()
      await page.waitForTimeout(400)
      await expect(chooser(page)).toBeHidden()
    })

    await check("the Workflows page lists starters, shows one, and uses it in a matter", size, async (page, shot) => {
      await open(page, "/workflows")
      const section = page.getByRole("region", { name: "Review table" })
      await expect(section.locator(".list-row")).toHaveCount(2)
      await expect(section).toContainText("5 questions")
      await expect(page.locator(".library-count")).toHaveText(String(WORKFLOWS.length + 2))
      await shot("workflows-starters")
      await section.locator(".list-row", { hasText: "Compare agreements" }).getByRole("button", { name: "View" }).click()
      const view = page.getByRole("dialog", { name: "Compare agreements (review table)" })
      for (const text of ["Inputs", "Output", "Questions", "Which law governs the agreement", "this is not a complete legal review"]) await expect(view).toContainText(text)
      assert.equal(await fits(page), true, "the view fits its width")
      await page.waitForTimeout(300)
      await shot("workflows-starter-view")
      // The view ends in its own actions: Close beside the one primary, Use in a matter.
      const actions = view.locator(".compare-actions")
      await expect(actions.locator("button.primary")).toHaveText("Use in a matter")
      await actions.getByRole("button", { name: "Close" }).click()
      await expect(view).toBeHidden()
      await section.locator(".list-row", { hasText: "Document overview" }).getByRole("button", { name: "Use" }).click()
      const picker = page.getByRole("dialog", { name: 'Use "Document overview" in a matter' })
      await picker.getByRole("button", { name: /Supply review/ }).click()
      // The table's Add questions, with the lawyer's choice kept and not second-guessed.
      const choose = chooser(page)
      await expect(choose.getByRole("radio", { name: /^Document overview/ })).toBeChecked()
      assert.equal(await choose.getByText("Recommended", { exact: false }).count(), 0, "the chosen starter is not second-guessed")
      await expect(choose.getByRole("radio", { name: /^Suggest questions with AI/ })).toBeVisible()
      await expect.poll(() => where(page)).toBe("/matter/pair?view=review")
      await page.waitForTimeout(300)
      await shot("workflows-arrival-chooser")
      await choose.getByRole("button", { name: "Continue" }).click()
      await expect(page.getByRole("dialog", { name: "Document overview" }).getByRole("list", { name: "Questions to add" })).toBeVisible()
      assert.deepEqual([sessions, prompts, puts], [0, [], 0], "nothing runs or saves before Add questions")
    })
  }

  const wide = [1440, 900] as const
  await check("Cancel and Escape keep the composer draft and start nothing", wide, async (page) => {
    await open(page, "/matter/solo?view=chat")
    const message = page.getByRole("textbox", { name: "Message" })
    await message.fill("Draft the lease advice note")
    await openLauncher(page)
    await launcher(page).getByRole("button", { name: "Cancel" }).click()
    await expect(launcher(page)).toBeHidden()
    await openLauncher(page)
    await page.keyboard.press("Escape")
    await expect(launcher(page)).toBeHidden()
    await expect(message).toHaveValue("Draft the lease advice note")
    await page.waitForTimeout(300)
    assert.deepEqual([sessions, prompts], [0, []], "closing starts no model work")
  })

  await check("an unknown starter is ignored and removed", wide, async (page) => {
    await open(page, "/matter/pair?view=review&starter=missing-v1")
    await expect(table(page, "Supply review").locator(".review-toolbar").getByRole("button", { name: "Add questions", exact: true })).toBeEnabled()
    await expect.poll(() => param(page, "starter")).toBeNull()
    await page.waitForTimeout(300)
    await expect(chooser(page)).toBeHidden()
  })

  await check("a table with questions opens Add questions with the starter chosen, whose questions are appended", wide, async (page) => {
    await open(page, "/matter/asked?view=review&starter=compare-agreements-v1")
    await expect(chooser(page).getByRole("radio", { name: /^Compare agreements/ })).toBeChecked()
    await chooser(page).getByRole("button", { name: "Continue" }).click()
    const questions = page.getByRole("dialog", { name: "Compare agreements" })
    await expect(questions.getByRole("list", { name: "Questions to add" })).toBeVisible()
    await expect(questions.getByRole("textbox")).toHaveCount(5)
    await expect.poll(() => param(page, "starter")).toBeNull()
    await questions.getByRole("button", { name: "Add questions" }).click()
    await expect(questions).toBeHidden()
    assert.equal(puts, 1)
    assert.deepEqual(grids.asked!.columns.map((column) => column.question)[0], "Who signed the agreement?", "the existing question stays first")
    assert.equal(grids.asked!.columns.length, 6)
  })

  assert.deepEqual([...new Set(unknown)], [], "every request the pages make is mocked")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
if (failures.length) {
  console.log(`${failures.length} failed: ${failures.join("; ")}`)
  process.exit(1)
}
