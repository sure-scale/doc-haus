// Run from apps/web: bun test/review-followups.browser.ts   (PLAYWRIGHT_MODULE=<playwright module> if needed)
// Mounts the real ReviewGrid with the product's stylesheets, so each #365 §2
// state is measured and photographed where a lawyer sees it: at a wide window
// and at a phone width. Only ingest HTTP is mocked; no extraction runs.
//   2.1 help opened inside a dialog stays inside it, and the progress key and
//       the Fill preview's "Not compared" note keep clear of the content.
//   2.2 deleting a comment offers Undo in its place.
//   2.3 a comment with no name shows no author line, and the first comment
//       offers the name inline.
import assert from "node:assert/strict"
import { mkdirSync } from "node:fs"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { hashQuestion, type Grid } from "../src/api/ingest/grid"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const shots = process.env.REVIEW_SHOTS_DIR ?? path.join(os.tmpdir(), "dochaus-review-help")
mkdirSync(shots, { recursive: true })
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-review-help-harness-"))
const law = "What is the governing law?"
const fee = "What is the fee, and when is it paid?"
const documents = ["Supply terms.docx", "Side letter.docx", "Old note.docx"]
const cite = (name: string, excerpt: string) => ({ documentName: name, docPath: name, section: "", excerpt, charStart: 0, charEnd: 4 })
let grid: Grid = {
  version: 0,
  columns: [{ id: "c1", question: law }, { id: "c2", question: fee, compare: false }],
  cells: {
    "Supply terms.docx::c1": {
      answer: "The law of England and Wales.",
      citation: cite("Supply terms.docx", "This agreement is governed by the law of England and Wales."),
      status: "reviewed",
      questionHash: hashQuestion(law),
      comments: [
        { id: "c1", text: "Check the governing law with the client.", at: Date.UTC(2026, 10, 12, 9), author: "Amina Shah" },
        { id: "c2", text: "Confirmed on the call; no name was recorded.", at: Date.UTC(2026, 10, 13, 10) },
      ],
    },
    "Supply terms.docx::c2": { answer: "£480 an hour, billed monthly.", status: "filled", questionHash: hashQuestion(fee) },
  },
}
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/primitives.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-system.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-documents.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/design-review.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/motion.css"))}
import ${JSON.stringify(path.resolve(import.meta.dir, "../src/styles/modal-frost.css"))}
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ReviewGrid from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/ReviewGrid.tsx"))}
const names = ${JSON.stringify(documents)}
createRoot(document.getElementById("root")).render(createElement(ReviewGrid, {
  matterId: "matter-help", title: "Supplier review", directory: "/matters/help",
  documents: names.map((name, index) => ({ id: index + 1, name, doc_path: name, created_at: 0 })),
  onClose: () => { window.gridCloses = (window.gridCloses ?? 0) + 1 },
}))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://grid.invalid", VITE_OPENCODE_URL: "http://engine.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = (await Promise.all(build.outputs.filter(output => output.kind === "asset" && output.path.endsWith(".css")).map(output => output.text()))).join("\n")
assert.ok(css.includes(".review-toolbar"), "the grid's own stylesheet is part of the harness")
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => {
  const { pathname } = new URL(request.url)
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname === "/harness.css") return new Response(css, { headers: { "content-type": "text/css" } })
  return new Response(`<html><meta charset="utf-8"><link rel="stylesheet" href="/harness.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  const json = (body: unknown) => ({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
  let saves = 0
  await page.route("http://grid.invalid/**", route => {
    const request = route.request()
    const url = new URL(request.url())
    if (url.pathname === "/matters/matter-help/grid" && request.method() === "PUT") {
      grid = { ...(request.postDataJSON() as Grid), version: (grid.version ?? 0) + 1 }
      saves++
      return route.fulfill(json(grid))
    }
    if (url.pathname === "/matters/matter-help/grid") return route.fulfill(json(grid))
    if (url.pathname === "/matters/matter-help") return route.fulfill(json({ documents: documents.map((name, index) => ({ name, id: index + 1 })) }))
    if (url.pathname === "/preferences") return route.fulfill(json({ attorney: "" }))
    if (url.pathname === "/engine/config") return route.fulfill(json({ model: "test/model" }))
    if (url.pathname === "/matters/matter-help/grid/sources") {
      const cells = (request.postDataJSON() as { cells: { key: string }[] }).cells
      return route.fulfill(json(Object.fromEntries(cells.map(cell => [cell.key, "verified"]))))
    }
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: `Unexpected ${url.pathname}` }) })
  })
  await page.route("http://engine.invalid/**", route => route.fulfill({ status: 404, body: "Unexpected" }))

  // Help shown inside a dialog never leaves it, and never lands on the table's
  // column headers; the Fill preview's note stays inside its own dialog too.
  for (const width of [1440, 400]) {
    await page.setViewportSize({ width, height: 900 })
    await page.goto(server.url.href)
    const table = page.getByRole("dialog", { name: "Review table — Supplier review" })
    await table.getByRole("row", { name: /^Side letter\.docx/ }).waitFor()
    await table.getByRole("button", { name: /^Fill \d+ cells$/ }).waitFor({ state: "attached" })
    await table.getByRole("img", { name: /^Review progress/ }).hover()
    const bubble = page.getByRole("tooltip")
    await bubble.waitFor()
    const [tip, dialogBox, head] = await Promise.all([bubble.boundingBox(), table.boundingBox(), page.locator("thead").boundingBox()])
    assert.ok(tip && dialogBox && head)
    const inside = (box: { x: number; y: number; width: number; height: number }) =>
      box.x >= dialogBox.x - 1 && box.x + box.width <= dialogBox.x + dialogBox.width + 1 && box.y >= dialogBox.y - 1 && box.y + box.height <= dialogBox.y + dialogBox.height + 1
    const overlaps = tip.x < head.x + head.width && tip.x + tip.width > head.x && tip.y < head.y + head.height && tip.y + tip.height > head.y
    assert.equal(inside(tip), true, `at ${width}px the progress key stayed inside the dialog`)
    assert.equal(overlaps, false, `at ${width}px the progress key stays clear of the column headers`)
    await page.screenshot({ path: path.join(shots, `progress-key-${width}.png`) })
    await page.mouse.move(0, 0)
    console.log(`PASS at ${width}px the progress key stays inside the dialog, clear of the column headers`)

    // The Fill preview's "Not compared" note, in its own dialog.
    await table.getByRole("button", { name: /^Fill \d+ cells?$/ }).click()
    const fill = page.getByRole("dialog", { name: /^Fill \d+ cells\?$/ })
    await fill.waitFor()
    await fill.getByText("Not compared", { exact: true }).first().hover()
    await bubble.waitFor()
    const [note, fillBox] = await Promise.all([bubble.boundingBox(), fill.boundingBox()])
    assert.ok(note && fillBox)
    assert.equal(inside(note) && note.x >= fillBox.x - 1 && note.x + note.width <= fillBox.x + fillBox.width + 1, true,
      `at ${width}px the Not compared note stayed inside the preview dialog`)
    await page.screenshot({ path: path.join(shots, `fill-not-compared-${width}.png`) })
    await fill.getByRole("button", { name: "Cancel", exact: true }).click()
    await fill.waitFor({ state: "detached" })
    console.log(`PASS at ${width}px the Fill preview's note stays inside the preview dialog`)
  }

  // 2.2 and 2.3 in the cell's answer detail: the first comment offers a name,
  // a comment without one shows its date only, and Delete offers Undo.
  for (const width of [1440, 400]) {
    await page.setViewportSize({ width, height: 900 })
    await page.goto(server.url.href)
    await page.addStyleTag({ content: "*, *::backdrop { backdrop-filter: none !important; -webkit-backdrop-filter: none !important; }" })
    await page.getByRole("button", { name: /^The law of England and Wales\./ }).click()
    const detail = page.getByRole("dialog", { name: /^Answer and comments:/ })
    await detail.waitFor()
    const sign = detail.getByRole("textbox", { name: "Your name, shown on your comments" })
    await sign.waitFor()
    assert.equal(await sign.inputValue(), "")
    const comments = detail.locator(".review-comment")
    assert.equal(await comments.count(), 2)
    assert.equal(await comments.filter({ hasText: "Check the governing law with the client." }).getByText("Amina Shah", { exact: true }).count(), 1)
    assert.equal(await comments.filter({ hasText: "Confirmed on the call" }).locator("strong").count(), 0, "a comment with no name shows no author line")
    assert.equal(await detail.getByText("Unattributed", { exact: true }).count(), 0)
    await page.locator(".review-detail-modal .picker-panel").screenshot({ path: path.join(shots, `comment-name-offer-${width}.png`) })
    await detail.getByRole("button", { name: "Delete comment", exact: true }).first().click()
    await detail.getByText("Comment deleted.", { exact: false }).waitFor()
    assert.equal(await comments.filter({ hasText: "Check the governing law with the client." }).count(), 0)
    await page.locator(".review-detail-modal .picker-panel").screenshot({ path: path.join(shots, `comment-deleted-undo-${width}.png`) })
    await detail.getByRole("button", { name: "Undo", exact: true }).click()
    await comments.filter({ hasText: "Check the governing law with the client." }).waitFor()
    assert.equal(await comments.count(), 2, "Undo restores the comment in its place")
    assert.ok(saves > 0)
    console.log(`PASS at ${width}px the comment name offer, an unnamed comment, and delete with Undo`)
  }
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
