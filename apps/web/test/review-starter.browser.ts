// Real ReviewGrid in Chromium with isolated in-memory ingest and no provider.
// Run from apps/web: bun test/review-starter.browser.ts
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Grid } from "../src/api/ingest/grid"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-review-starter-"))
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ReviewGrid from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/ReviewGrid.tsx"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-review.css"))}
const params = new URLSearchParams(location.search)
const id = params.get("matter") || "one"
const docs = params.has("empty") ? [] : (params.has("long") ? ["long_".repeat(38) + ".docx"] : ["First.docx", "Second.docx"]).map((name, i) => ({ id: i+1, name, doc_path: name, created_at: 0, ...(params.has("deliverable") && i === 1 ? { role: "deliverable" } : {}) }))
createRoot(document.getElementById("root")).render(createElement(ReviewGrid, { matterId: id, title: id === "one" ? "First matter" : "Second matter", directory: "/synthetic/" + id, documents: docs, onClose: () => {}, onAsk: () => {} }))
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
const originalCell = { answer: "Existing answer", status: "reviewed" as const, questionHash: "old", comments: [{ id: "note", text: "Keep this comment", at: 1 }] }
const initial: Grid = { version: 0, columns: [{ id: "existing", question: "Existing question?" }], cells: { "First.docx::existing": originalCell } }
const empty: Grid = { version: 0, columns: [], cells: {} }
const grids: Record<string, Grid> = { one: structuredClone(initial), two: structuredClone(empty), empty: structuredClone(empty), conflict: structuredClone(initial), failure: structuredClone(initial), compare: structuredClone(initial) }
let puts = 0
let engineCalls = 0
const prompts: string[] = []
let sessionCount = 0
let failNext = false
let conflictNext = false
let changedDocuments = false
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage({ viewport: { width: 400, height: 860 } })
  page.setDefaultTimeout(8000)
  await page.route("http://grid.invalid/**", async route => {
    const request = route.request()
    const url = new URL(request.url())
    const reply = (body: unknown, status = 200) => route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) })
    const match = url.pathname.match(/^\/matters\/([^/]+)\/grid$/)
    if (match) {
      const id = match[1]!
      if (request.method() === "GET") return reply(grids[id])
      if (request.method() === "PUT") {
        puts++
        const incoming = request.postDataJSON() as Grid
        if (failNext) { failNext = false; return reply({ error: "Save failed" }, 500) }
        if (conflictNext || incoming.version !== grids[id]!.version) { conflictNext = false; return reply({ error: "Conflict", grid: grids[id] }, 409) }
        grids[id] = { ...incoming, version: (incoming.version ?? 0) + 1 }
        return reply(grids[id])
      }
    }
    const matterMatch = url.pathname.match(/^\/matters\/([^/]+)$/)
    if (matterMatch && request.method() === "GET") {
      const id = matterMatch[1]!
      const names = id === "empty" ? [] : changedDocuments ? ["First.docx", "New.docx"] : ["First.docx", "Second.docx"]
      return reply({ id, title: id === "one" ? "First matter" : "Second matter", dir: `/synthetic/${id}`, documents: names.map((name, index) => ({ id: index + 1, name, doc_path: name, created_at: 0 })) })
    }
    if (url.pathname.endsWith("/grid/sources")) return reply({})
    if (url.pathname === "/preferences") return reply({ attorney: "Fixture Lawyer" })
    if (url.pathname === "/engine/config") return reply({ model: "fixture/model" })
    return reply({ error: `Unexpected ${request.method()} ${url.pathname}` }, 404)
  })
  await page.route("http://engine.invalid/**", route => {
    engineCalls++
    const request = route.request()
    const pathname = new URL(request.url()).pathname
    const reply = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (request.method() === "POST" && pathname === "/api/session") return reply({ data: { id: `session-${++sessionCount}`, location: { directory: "/synthetic/run" } } })
    if (/^\/api\/session\/session-\d+\/(agent|model)$/.test(pathname)) return route.fulfill({ status: 204 })
    if (/^\/api\/session\/session-\d+\/prompt$/.test(pathname)) {
      prompts.push((request.postDataJSON() as { text: string }).text)
      return route.fulfill({ status: 500, body: "Stopped by local fixture" })
    }
    return route.fulfill({ status: 500, body: "Unexpected model call" })
  })
  // Each step names the dialog: the chooser, a starter's example, then its questions.
  const chooser = page.getByRole("dialog", { name: "Choose a starter" })
  const example = page.getByRole("dialog", { name: "Example: Compare agreements" })
  const starter = page.getByRole("dialog", { name: "Document overview" })
  const compare = page.getByRole("dialog", { name: "Compare agreements" })
  const overviewChoice = chooser.getByRole("radio", { name: /^Document overview/ })
  const compareChoice = chooser.getByRole("radio", { name: /^Compare agreements/ })
  const reviewFill = page.getByRole("dialog", { name: "Review fill" })
  // A starter is the primary action of a table without questions, and a quiet one otherwise.
  const starterButton = page.getByRole("button", { name: /^Use a starter$/ })
  const primary = page.locator(".review-toolbar .review-primary")
  const open = async (query: string) => { await page.goto(`${server.url.href}${query}`); await starterButton.waitFor() }
  const showChooser = async () => { await starterButton.click(); await expect(chooser).toBeVisible() }
  const showStarter = async () => {
    await showChooser()
    await overviewChoice.check()
    await chooser.getByRole("button", { name: "Continue" }).click()
    await expect(starter).toBeVisible()
  }

  await open("?matter=one")
  await expect(primary).toHaveCount(1)
  await expect(primary).toHaveText("Fill 1 cell")
  await expect(page.getByRole("button", { name: "Use a starter", exact: true })).toBeVisible()
  const toolbar = page.locator(".review-toolbar")
  assert.equal(await toolbar.evaluate((node) => node.scrollWidth <= node.clientWidth + 1), true, "the toolbar wraps within a 400px viewport")

  // The chooser preselects the starter that fits the matter's source documents.
  await showChooser()
  await expect(compareChoice).toBeChecked()
  await expect(compareChoice).toBeFocused()
  await expect(chooser).toContainText("Recommended · 2 documents in this matter")
  await expect(chooser.getByRole("region", { name: "About Compare agreements" })).toContainText("Which law governs the agreement")
  await expect(chooser.getByRole("button", { name: /^(Continue|Add questions)$/ })).toHaveCount(1)
  const chooserBox = await chooser.boundingBox()
  assert.ok(chooserBox && chooserBox.x >= 0 && chooserBox.x + chooserBox.width <= 400, "chooser must fit a 400px viewport")
  await page.keyboard.press("ArrowUp")
  await expect(overviewChoice).toBeChecked()
  await expect(chooser.getByRole("region", { name: "About Document overview" })).toBeVisible()
  await expect(chooser.getByRole("button", { name: "See example" })).toHaveCount(0)
  await page.keyboard.press("Tab")
  assert.equal(await chooser.evaluate(el => el.contains(document.activeElement)), true, "keyboard focus stays in the chooser")
  await page.keyboard.press("Escape")
  await expect(chooser).toHaveCount(0)

  // The example is read-only: it opens, shows a cited passage and returns to the chooser.
  await showChooser()
  await chooser.getByRole("button", { name: "See example" }).click()
  await expect(example).toBeVisible()
  await expect(example).toContainText("Fictional example")
  await expect(example.getByRole("region", { name: "Example review table" }).getByRole("columnheader")).toHaveCount(6)
  // On a short screen the table scrolls within its own region under its header row.
  const exampleTable = example.getByRole("region", { name: "Example review table" })
  await page.setViewportSize({ width: 400, height: 640 })
  assert.equal(await exampleTable.evaluate((node) => node.scrollHeight > node.clientHeight), true, "the example table scrolls within its own region")
  await exampleTable.evaluate((node) => { node.scrollTop = 120 })
  const headerTop = await exampleTable.getByRole("columnheader").first().evaluate((node) => node.getBoundingClientRect().top - node.closest(".review-example-scroll")!.getBoundingClientRect().top)
  assert.ok(Math.abs(headerTop) <= 1.5, `the example header row stays at the top of its region (${headerTop}px)`)
  await exampleTable.evaluate((node) => { node.scrollTop = 0 })
  await page.setViewportSize({ width: 400, height: 860 })
  // An answer opens its read-only detail; Escape closes only the detail and
  // returns focus to the answer, and a second Escape closes the example.
  const differsCell = example.locator(".review-example-cell button", { hasText: "Differs" }).first()
  await differsCell.click()
  const answerDetail = page.getByRole("dialog", { name: /^Example answer:/ })
  await expect(answerDetail).toBeVisible()
  await expect(answerDetail).toContainText("Fictional example")
  await expect(answerDetail.locator(".review-walk-word")).toContainText("Differs")
  await expect(answerDetail.locator(".citation .excerpt")).toContainText("limited to £250,000 in aggregate")
  await expect(answerDetail.getByRole("button")).toHaveCount(1)
  await page.keyboard.press("Escape")
  await expect(answerDetail).toHaveCount(0)
  await expect(example).toBeVisible()
  await expect(differsCell).toBeFocused()
  await example.locator(".review-example-cell button", { hasText: "Not found" }).first().click()
  await expect(answerDetail.locator(".review-walk-why")).toContainText("The document does not appear to address this question.")
  await expect(answerDetail.getByRole("region", { name: "Source" })).toHaveCount(0)
  await answerDetail.getByRole("button", { name: "Close" }).click()
  await expect(answerDetail).toHaveCount(0)
  await page.keyboard.press("Escape")
  await expect(example).toHaveCount(0)
  assert.equal(puts, 0, "the example must not write")
  await showChooser()
  await chooser.getByRole("button", { name: "See example" }).click()
  await example.getByRole("button", { name: "Back" }).click()
  await expect(chooser).toBeVisible()
  await expect(compareChoice).toBeChecked()
  assert.equal(puts, 0, "the chooser and example must not write")
  assert.equal(engineCalls, 0, "the chooser and example must not call a model")
  await page.keyboard.press("Escape")
  await expect(chooser).toHaveCount(0)
  assert.deepEqual(grids.one, initial)

  await showStarter()
  await expect(starter).toContainText("2 documents: First.docx, Second.docx")
  await expect(starter).toContainText("Adds 3 questions · 6 cells to fill. No AI runs until you choose Fill.")
  const box = await starter.boundingBox()
  assert.ok(box && box.x >= 0 && box.x + box.width <= 400, "starter must fit a 400px viewport")
  await starter.getByRole("textbox", { name: "Question 1" }).focus()
  await page.keyboard.press("Tab")
  assert.equal(await starter.evaluate(el => el.contains(document.activeElement)), true, "keyboard focus stays in the nested preview")
  await page.keyboard.press("Escape")
  await expect(starter).toHaveCount(0)
  assert.equal(puts, 0, "opening or cancelling must not write")
  assert.equal(engineCalls, 0, "opening or cancelling must not call a model")
  assert.deepEqual(grids.one, initial)

  await showStarter()
  await starter.getByRole("textbox", { name: "Question 1" }).fill("  Existing question?  ")
  await expect(starter.getByRole("status")).toContainText("Already in the table")
  await expect(starter).toContainText("Adds 2 questions · 4 cells to fill")
  await starter.getByRole("button", { name: "Remove question 3" }).focus()
  await page.keyboard.press("Enter")
  await expect(starter.getByRole("textbox", { name: "Question 2" })).toBeFocused()
  await starter.getByRole("textbox", { name: "Question 2" }).fill("What is the duration?")
  await expect(starter).toContainText("Adds 1 question · 2 cells to fill")
  await starter.getByRole("button", { name: "Add questions" }).click()
  await expect(starter).toHaveCount(0)
  await expect.poll(() => puts).toBe(1)
  assert.equal(grids.one!.columns.length, 2)
  assert.equal(grids.one!.columns[1]!.question, "What is the duration?")
  assert.deepEqual(grids.one!.cells["First.docx::existing"], originalCell, "append retains reviewed answer and comment")
  assert.equal(engineCalls, 0, "setup must never trigger extraction")
  await page.reload()
  await expect(page.getByRole("button", { name: "Edit question: What is the duration?" })).toBeVisible()
  assert.deepEqual(grids.one!.cells["First.docx::existing"], originalCell)

  await expect(primary).toHaveCount(1)
  await page.getByRole("button", { name: "Fill 3 cells", exact: true }).click()
  await expect(reviewFill).toBeVisible()
  await expect(reviewFill).toContainText("3 cells")
  await expect(reviewFill).toContainText("First.docx")
  await expect(reviewFill).toContainText("Second.docx")
  await expect(reviewFill).toContainText("Existing question?")
  await expect(reviewFill).toContainText("What is the duration?")
  await expect(reviewFill).toContainText("uses AI")
  await reviewFill.getByRole("button", { name: "Cancel" }).click()
  assert.equal(engineCalls, 0, "cancelling Fill must make no model calls")
  assert.equal(puts, 1)

  await showStarter()
  await starter.getByRole("textbox", { name: "Question 1" }).fill("What is the duration?")
  await starter.getByRole("button", { name: "Remove question 3" }).click()
  await starter.getByRole("button", { name: "Remove question 2" }).click()
  await expect(starter.getByRole("button", { name: "Add questions" })).toBeDisabled()
  await starter.getByRole("button", { name: "Remove question 1" }).focus()
  await page.keyboard.press("Enter")
  await expect(starter.getByRole("button", { name: "Cancel" })).toBeFocused()
  await starter.getByRole("button", { name: "Cancel" }).click()
  assert.equal(puts, 1)
  await open("?matter=two")
  await expect(primary).toHaveCount(1)
  await expect(primary).toHaveText("Use a starter")
  // The primary is the only starter button: no quiet duplicate beside it.
  await expect(page.getByRole("button", { name: "Use a starter", exact: true })).toHaveCount(1)
  await showStarter()
  assert.deepEqual(grids.two, empty)
  await starter.getByRole("button", { name: "Cancel" }).click()

  // Without documents nothing can be filled or exported: no primary action.
  await open("?matter=empty&empty=1")
  await expect(primary).toHaveCount(0)
  await expect(page.getByRole("button", { name: /^Export/ })).toHaveCount(0)
  await showStarter()
  await expect(starter).toContainText("Add documents to this matter")
  await expect(starter.getByRole("button", { name: "Add questions" })).toBeDisabled()
  await starter.getByRole("button", { name: "Cancel" }).click()
  assert.equal(puts, 1)

  // A starter whose source documents are mostly deliverables recommends the overview.
  await open("?matter=two&deliverable=1")
  await showChooser()
  await expect(overviewChoice).toBeChecked()
  await expect(chooser).toContainText("Recommended · 1 source document in this matter")
  await chooser.getByRole("button", { name: "Cancel" }).click()
  await expect(chooser).toHaveCount(0)

  // From the example, Use with my matter opens the questions; Add questions only appends.
  await open("?matter=compare")
  await showChooser()
  await chooser.getByRole("button", { name: "See example" }).click()
  await example.getByRole("button", { name: "Use with my matter" }).click()
  await expect(compare).toBeVisible()
  await expect(compare).toContainText("Adds 5 questions · 10 cells to fill")
  await compare.getByRole("button", { name: "Back" }).click()
  await expect(compareChoice).toBeChecked()
  await chooser.getByRole("button", { name: "Continue" }).click()
  await compare.getByRole("button", { name: "Add questions" }).click()
  await expect(compare).toHaveCount(0)
  await expect.poll(() => grids.compare!.columns.length).toBe(6)
  assert.equal(grids.compare!.columns[0]!.question, "Existing question?")
  assert.equal(grids.compare!.columns[5]!.question, "Which law governs the agreement, and how are disputes resolved?")
  assert.deepEqual(grids.compare!.cells, initial.cells, "adding a starter keeps every existing answer and comment")
  assert.equal(engineCalls, 0, "adding starter questions must not call a model")

  await open("?matter=failure")
  await showStarter()
  failNext = true
  await starter.getByRole("button", { name: "Add questions" }).click()
  await expect(starter.getByText("The questions could not be confirmed.", { exact: false })).toBeVisible()
  assert.deepEqual(grids.failure, initial, "failed save does not report success or change stored grid")
  await starter.getByRole("button", { name: "Cancel" }).click()
  await expect(page.getByText("The table could not be saved.", { exact: false })).toBeVisible()

  await open("?matter=conflict")
  await showStarter()
  conflictNext = true
  await starter.getByRole("button", { name: "Add questions" }).click()
  await expect(starter.getByText("The questions could not be confirmed.", { exact: false })).toBeVisible()
  assert.deepEqual(grids.conflict, initial, "conflict preserves stored grid")
  await starter.getByRole("button", { name: "Cancel" }).click()
  await expect(page.getByText("This table conflicts with the saved table.", { exact: false })).toBeVisible()
  assert.equal(engineCalls, 0)

  // A newer saved grid invalidates the disclosed targets before any model call.
  grids.stale = structuredClone(grids.one!)
  await open("?matter=stale")
  await page.getByRole("button", { name: "Fill 3 cells", exact: true }).click()
  grids.stale!.version = (grids.stale!.version ?? 0) + 1
  await reviewFill.getByRole("button", { name: "Start fill" }).click()
  await expect(reviewFill).toHaveCount(0)
  await expect(page.getByText("This table changed elsewhere.", { exact: false })).toBeVisible()
  assert.equal(engineCalls, 0, "a stale grid must not start extraction")

  // Documents can change in another window without changing the grid version.
  // Neither setup nor Fill may proceed against a stale preview.
  grids.changed = structuredClone(grids.one!)
  await open("?matter=changed")
  await showStarter()
  changedDocuments = true
  const beforeMismatch = puts
  await starter.getByRole("button", { name: "Add questions" }).click()
  await expect(starter.getByText("The matter's documents changed.", { exact: false })).toBeVisible()
  assert.equal(puts, beforeMismatch, "changed documents must not write starter questions")
  await starter.getByRole("button", { name: "Cancel" }).click()
  await expect(page.getByText("The matter's documents changed.", { exact: false })).toBeVisible()
  changedDocuments = false
  await open("?matter=changed")
  await page.getByRole("button", { name: "Fill 3 cells", exact: true }).click()
  changedDocuments = true
  await reviewFill.getByRole("button", { name: "Start fill" }).click()
  await expect(reviewFill).toHaveCount(0)
  await expect(page.getByText("The matter's documents changed.", { exact: false })).toBeVisible()
  assert.equal(engineCalls, 0, "changed documents must prevent extraction")
  changedDocuments = false

  // Start the disclosed run on an isolated copy. The local engine records each
  // prompt and refuses it, allowing us to compare actual dispatched targets.
  grids.run = structuredClone(grids.one!)
  await open("?matter=run")
  await page.getByRole("button", { name: "Fill 3 cells", exact: true }).click()
  await expect(reviewFill).toContainText("3 cells")
  await reviewFill.getByRole("button", { name: "Start fill" }).click()
  await expect.poll(() => prompts.length).toBe(3)
  assert.deepEqual([...prompts].sort(), [
    'Document: "First.docx". Question: What is the duration?',
    'Document: "Second.docx". Question: Existing question?',
    'Document: "Second.docx". Question: What is the duration?',
  ].sort(), "the dispatched cells must match the full disclosed run, excluding the reviewed cell")
  const firstMatter = structuredClone(grids.one!)
  await open("?matter=two")
  await showStarter()
  await starter.getByRole("button", { name: "Add questions" }).click()
  await expect(starter).toHaveCount(0)
  assert.equal(grids.two!.columns.length, 3)
  assert.deepEqual(grids.one, firstMatter, "adding to a second matter must not change the first matter")
  assert.deepEqual(grids.two!.cells, {}, "the first matter's answers and comments must not carry over")
  grids.long = structuredClone(initial)
  await open("?matter=long&long=1")
  await page.getByRole("button", { name: "Fill 1 cell", exact: true }).click()
  const fillBody = reviewFill.locator(".review-starter-body")
  await expect(reviewFill.getByRole("heading", { name: `${"long_".repeat(38)}.docx` })).toBeVisible()
  assert.equal(await fillBody.evaluate((node) => node.scrollWidth <= node.clientWidth + 1), true, "long document names wrap without horizontal overflow")
  console.log("PASS starter chooser recommendation and keyboard, read-only example and its answer detail, starter preview and keyboard, safe append and reload, isolation, duplicate handling, Fill disclosure and dispatched targets, changed documents, stale versions, and save failures")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
