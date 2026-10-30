// Real ReviewGrid in Chromium with isolated in-memory ingest and a mocked engine:
// the + menu by keyboard, then suggesting questions with AI.
// Run from apps/web: bun test/review-suggest.browser.ts
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Route } from "@playwright/test"
import type { Grid } from "../src/api/ingest/grid"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-review-suggest-"))
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ReviewGrid from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/ReviewGrid.tsx"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-review.css"))}
const docs = ["First.docx", "Second.docx"].map((name, i) => ({ id: i + 1, name, doc_path: name, created_at: 0 }))
createRoot(document.getElementById("root")).render(createElement(ReviewGrid, { matterId: "one", title: "First matter", directory: "/synthetic/one", documents: docs, documentsRevision: 0, onClose: () => {}, onAsk: () => {} }))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://grid.invalid", VITE_OPENCODE_URL: "http://engine.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = await build.outputs.find(output => output.path.endsWith(".css"))?.text() ?? ""
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: req => {
  const pathname = new URL(req.url).pathname
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname === "/harness.css") return new Response(css, { headers: { "content-type": "text/css" } })
  return new Response('<html><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="stylesheet" href="/harness.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>', { headers: { "content-type": "text/html" } })
} })

let grid: Grid = { version: 0, columns: [{ id: "existing", question: "Existing question?" }], cells: {} }
let puts = 0
// Engine: each review-suggest session, its prompt, whether it was interrupted and deleted.
type Session = { id: string; agent?: string; prompt?: string; interrupted: boolean; deleted: boolean }
const sessions: Session[] = []
const replies: string[] = []
let holdWait = false
const heldWaits: Route[] = []
const otherEngineCalls: string[] = []
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 860 } })
  page.setDefaultTimeout(8000)
  await page.route("http://grid.invalid/**", async route => {
    const request = route.request()
    const url = new URL(request.url())
    const reply = (body: unknown, status = 200) => route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) })
    if (url.pathname === "/matters/one/grid") {
      if (request.method() === "GET") return reply(grid)
      if (request.method() === "PUT") {
        puts++
        const incoming = request.postDataJSON() as Grid
        grid = { ...incoming, version: (incoming.version ?? 0) + 1 }
        return reply(grid)
      }
    }
    if (url.pathname === "/matters/one" && request.method() === "GET")
      return reply({ id: "one", title: "First matter", dir: "/synthetic/one", documents: ["First.docx", "Second.docx"].map((name, index) => ({ id: index + 1, name, doc_path: name, created_at: 0 })) })
    if (url.pathname.endsWith("/grid/sources")) return reply({})
    if (url.pathname === "/preferences") return reply({ attorney: "Fixture Lawyer" })
    if (url.pathname === "/engine/config") return reply({ model: "fixture/primary", small_model: "fixture/fast" })
    return reply({ error: `Unexpected ${request.method()} ${url.pathname}` }, 404)
  })
  await page.route("http://engine.invalid/**", route => {
    const request = route.request()
    const url = new URL(request.url())
    const key = `${request.method()} ${url.pathname}`
    const json = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (key === "POST /api/session") {
      const body = request.postDataJSON() as { id: string; agent?: string; model?: unknown; metadata?: unknown }
      assert.deepEqual(body.model, { providerID: "fixture", id: "primary" }, "suggestions run on the primary model")
      assert.deepEqual(body.metadata, { dochaus_role: "review-suggest" })
      sessions.push({ id: body.id, agent: body.agent, interrupted: false, deleted: false })
      return json({ data: { id: body.id, location: { directory: "/synthetic/one" } } })
    }
    const session = sessions.find(item => url.pathname.includes(`/session/${item.id}`))
    if (session) {
      if (key.endsWith("/prompt")) { session.prompt = (request.postDataJSON() as { text: string }).text; return json({ data: { id: "msg_suggest", type: "user", sessionID: session.id, time: { created: 1 }, payload: { text: "suggest" }, delivery: "queue" } }) }
      if (key.endsWith("/wait")) { if (holdWait) { heldWaits.push(route); return } return route.fulfill({ status: 204 }) }
      if (key.endsWith("/interrupt")) { session.interrupted = true; return json({ interrupted: true }) }
      if (key.startsWith("GET") && key.endsWith("/message")) return json({ data: [
        { type: "assistant", id: "msg_reply", agent: "review-suggest", model: { providerID: "fixture", id: "fast" }, finish: "stop", time: { created: 1, completed: 2 }, content: [{ type: "text", text: replies.shift() ?? "" }] },
        { type: "idle", id: "msg_idle", outcome: "succeeded", time: { created: 3 } },
      ], cursor: {} })
      if (key === `DELETE /api/session/${session.id}`) { session.deleted = true; return route.fulfill({ status: 204 }) }
    }
    otherEngineCalls.push(key)
    return route.fulfill({ status: 404, body: `Unexpected ${key}` })
  })

  await page.goto(server.url.href)
  const table = page.getByRole("table", { name: "Document review" })
  const add = table.getByRole("button", { name: "Ways to add questions", exact: true })
  const menu = page.getByRole("group", { name: "Add questions" })
  const focused = () => page.evaluate(() => document.activeElement?.getAttribute("aria-label") ?? document.activeElement?.textContent?.trim())
  await add.waitFor()

  // The + menu by keyboard: arrows open it at its first item and wrap; Escape
  // closes it on the button; a choice opens its dialog, which hands focus back.
  await add.focus()
  await page.keyboard.press("ArrowDown")
  await expect(menu).toBeVisible()
  await expect(menu.getByRole("button", { name: "Write a question" })).toBeFocused()
  await page.keyboard.press("ArrowUp")
  await expect(menu.getByRole("button", { name: "Suggest with AI" })).toBeFocused()
  await page.keyboard.press("ArrowDown")
  await expect(menu.getByRole("button", { name: "Write a question" })).toBeFocused()
  await page.keyboard.press("Escape")
  await expect(menu).toBeHidden()
  await expect(add).toBeFocused()
  await page.keyboard.press("Enter")
  await expect(menu).toBeVisible()
  await expect(add).toHaveAttribute("aria-expanded", "true")
  await page.keyboard.press("ArrowDown")
  await page.keyboard.press("ArrowDown")
  await expect(menu.getByRole("button", { name: "Use a starter" })).toBeFocused()
  await page.keyboard.press("Enter")
  const chooser = page.getByRole("dialog", { name: "Add questions" })
  await expect(chooser).toBeVisible()
  const aiChoice = chooser.getByRole("radio", { name: /^Suggest questions with AI/ })
  await expect(chooser.getByRole("radio")).toHaveCount(3)
  await aiChoice.check()
  const about = chooser.getByRole("region", { name: "About Suggest questions with AI" })
  await expect(about).toContainText("One AI call")
  await expect(about).toContainText("Not the documents' text")
  await page.keyboard.press("Escape")
  await expect(chooser).toHaveCount(0)
  assert.equal(await focused(), "Ways to add questions", "closing the chooser returns focus to the + button")
  await add.click()
  await menu.getByRole("button", { name: "Write a question" }).click()
  const writer = page.getByRole("dialog", { name: "Add question" })
  await expect(writer).toBeVisible()
  await page.keyboard.press("Escape")
  await expect(writer).toHaveCount(0)
  assert.equal(await focused(), "Ways to add questions")
  assert.equal(sessions.length, 0, "opening the menu and chooser calls no model")
  console.log("PASS the + menu opens, moves and closes by keyboard, and each choice opens its dialog and returns focus")

  // Suggest with AI opens at the describe step. Cancel while suggesting stops
  // the call: its session is interrupted and deleted, and the words stay.
  await add.click()
  await menu.getByRole("button", { name: "Suggest with AI" }).click()
  const describe = page.getByRole("dialog", { name: "Suggest questions with AI" })
  const field = describe.getByRole("textbox", { name: "What do you want to find out from these documents?" })
  await expect(field).toBeFocused()
  await expect(describe).toContainText("not the documents' text")
  const suggestButton = describe.getByRole("button", { name: "Suggest questions" })
  await expect(suggestButton).toBeDisabled()
  const wanted = "What each party must deliver and how the arrangement can end"
  await field.fill(wanted)
  holdWait = true
  await suggestButton.click()
  await expect(describe.getByRole("status")).toHaveText("Suggesting questions…")
  await expect(describe.getByRole("button", { name: "Suggest questions" })).toHaveCount(0)
  await expect.poll(() => heldWaits.length).toBe(1)
  await describe.getByRole("button", { name: "Cancel" }).click()
  await expect.poll(() => sessions[0]?.deleted).toBe(true)
  assert.equal(sessions[0]!.interrupted, true, "Cancel interrupts the unsettled call")
  assert.equal(sessions[0]!.agent, "review-suggest")
  await expect(describe).toBeVisible()
  await expect(field).toHaveValue(wanted)
  await expect(field).toBeFocused()
  await expect(describe.getByRole("status")).toHaveCount(0)
  holdWait = false
  await heldWaits.shift()!.abort().catch(() => {})
  console.log("PASS Cancel while suggesting interrupts and deletes the session and keeps the description")

  // A reply outside the contract is one quiet line with Try again.
  replies.push("Here are some questions you might ask.")
  await suggestButton.click()
  const failure = describe.getByRole("status")
  await expect(failure).toContainText("The suggestions could not be read.")
  await expect(describe.getByRole("alert")).toHaveCount(0)
  await expect.poll(() => sessions[1]?.deleted).toBe(true)

  // Try again: the suggestions arrive at the questions step, titled as
  // suggested rather than tagged row by row, without the one already in the table. The AI saw only the description, the names
  // and the table's questions; the session is deleted after.
  replies.push(JSON.stringify({ questions: [
    { text: "Who are the parties to this document?", compare: false },
    { text: "What must each party deliver?", compare: true },
    { text: "existing question?" },
    { text: "How can this document be ended early?" },
  ] }))
  await failure.getByRole("button", { name: "Try again" }).click()
  const suggested = page.getByRole("dialog", { name: "Suggested questions" })
  await expect(suggested).toBeVisible()
  await expect.poll(() => sessions[2]?.deleted).toBe(true)
  assert.equal(sessions[2]!.interrupted, false, "a finished call is not interrupted")
  const firstPrompt = sessions[2]!.prompt!
  for (const text of [JSON.stringify(wanted), JSON.stringify(["First.docx", "Second.docx"]), JSON.stringify(["Existing question?"])]) assert.ok(firstPrompt.includes(text), `the prompt carries ${text}`)
  const list = suggested.getByRole("list", { name: "Questions to add" })
  await expect(list.getByRole("textbox")).toHaveCount(3)
  await expect(list.locator(".review-starter-tag")).toHaveCount(0)
  await expect(suggested.getByRole("textbox", { name: "Question 1" })).toHaveValue("Who are the parties to this document?")
  await expect(suggested).toContainText("Adds 3 questions · 6 cells to fill. No AI runs until you choose Fill.")

  // Edit and remove like starter questions; Suggest more appends below the
  // list and skips what is already listed, starting at the first new one.
  await suggested.getByRole("textbox", { name: "Question 2" }).fill("What must each party deliver, and by when?")
  await suggested.getByRole("button", { name: "Remove question 3" }).click()
  await expect(list.getByRole("textbox")).toHaveCount(2)
  replies.push(JSON.stringify({ questions: [{ text: "Who are the parties to this document?" }, { text: "Which law governs this document?" }] }))
  await suggested.getByRole("button", { name: "Suggest more with AI" }).click()
  await expect(describe).toBeVisible()
  await expect(field).toHaveValue(wanted)
  await describe.getByRole("button", { name: "Suggest questions" }).click()
  await expect(suggested).toBeVisible()
  await expect(list.getByRole("textbox")).toHaveCount(3)
  await expect(suggested.getByRole("textbox", { name: "Question 3" })).toHaveValue("Which law governs this document?")
  await expect(suggested.getByRole("textbox", { name: "Question 3" })).toBeFocused()
  await expect.poll(() => sessions[3]?.deleted).toBe(true)
  assert.ok(sessions[3]!.prompt!.includes(JSON.stringify(["Existing question?", "Who are the parties to this document?", "What must each party deliver, and by when?"])), "Suggest more shows the AI the questions already chosen")

  // Add questions appends through the starter path, keeping the compare flag.
  await suggested.getByRole("button", { name: "Add questions" }).click()
  await expect(suggested).toHaveCount(0)
  await expect.poll(() => puts).toBe(1)
  assert.deepEqual(grid.columns.map(column => [column.question, column.compare]), [
    ["Existing question?", undefined],
    ["Who are the parties to this document?", false],
    ["What must each party deliver, and by when?", undefined],
    ["Which law governs this document?", undefined],
  ])
  assert.equal(await focused(), "Ways to add questions", "adding returns focus to the + button")
  assert.equal(sessions.length, 4)
  assert.ok(sessions.every(session => session.deleted), "every suggest session is deleted")
  assert.deepEqual(otherEngineCalls, [], "nothing else reaches the engine before Fill")

  // Beside a starter's questions an AI suggestion is tagged, until the lawyer
  // edits it into their own question.
  await add.click()
  await menu.getByRole("button", { name: "Use a starter" }).click()
  await chooser.getByRole("radio", { name: /^Document overview/ }).check()
  await chooser.getByRole("button", { name: "Continue" }).click()
  const overview = page.getByRole("dialog", { name: "Document overview" })
  await expect(overview.locator(".review-starter-tag")).toHaveCount(0)
  replies.push(JSON.stringify({ questions: [{ text: "Which law governs each document?" }] }))
  await overview.getByRole("button", { name: "Suggest more with AI" }).click()
  await field.fill(wanted)
  await describe.getByRole("button", { name: "Suggest questions" }).click()
  await expect(overview.locator(".review-starter-tag")).toHaveText(["Suggested by AI"])
  await overview.getByRole("textbox", { name: "Question 4" }).fill("Which law governs this agreement?")
  await expect(overview.locator(".review-starter-tag")).toHaveCount(0)
  await expect.poll(() => sessions[4]?.deleted).toBe(true)
  await overview.getByRole("button", { name: "Cancel" }).click()
  await expect(overview).toHaveCount(0)

  // At 400px the describe step fits the viewport.
  await page.setViewportSize({ width: 400, height: 860 })
  await add.click()
  await menu.getByRole("button", { name: "Suggest with AI" }).click()
  const box = await describe.locator(".picker-panel").boundingBox()
  assert.ok(box && box.x >= 0 && box.x + box.width <= 400, "the describe step fits a 400px viewport")
  await page.keyboard.press("Escape")
  await expect(describe).toHaveCount(0)
  console.log("PASS suggestions arrive de-duplicated, edit and remove, Suggest more appends, Add keeps compare flags, a suggestion beside starter questions is tagged until edited, and every session is deleted")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
