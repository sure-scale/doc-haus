// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/review-grid.browser.ts
// Mounts the real ReviewGrid as MatterDetail mounts it; only ingest HTTP is mocked. No extraction runs.
// Escape closes the innermost layer: a cell's detail first, the review table on the next press.
// Set PLAYWRIGHT_CHROMIUM_PATH=<chromium> when Playwright's own browser is missing.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-review-grid-"))
const grid = {
  columns: [{ id: "c1", question: "What is the governing law?" }],
  cells: {
    "Supply terms.docx::c1": {
      answer: "The law of England and Wales.",
      citation: { documentName: "Supply terms.docx", docPath: "Supply terms.docx", section: "12", excerpt: "This agreement is governed by the law of England and Wales.", charStart: 0, charEnd: 58 },
      status: "filled",
      questionHash: "unused",
    },
  },
}
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ReviewGrid from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/ReviewGrid.tsx"))}
// Closing unmounts the grid, as MatterDetail drops ?view=review.
function Harness() {
  const [open, setOpen] = useState(true)
  return open ? createElement(ReviewGrid, {
    matterId: "matter-grid", title: "Supplier review", directory: "/matters/supplier-review",
    documents: [{ id: 1, name: "Supply terms.docx", doc_path: "Supply terms.docx", created_at: 0 }],
    onClose: () => { window.gridCloses = (window.gridCloses ?? 0) + 1; setOpen(false) },
  }) : null
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://grid.invalid", VITE_OPENCODE_URL: "http://engine.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => new URL(request.url).pathname === "/harness.js"
  ? new Response(script, { headers: { "content-type": "text/javascript" } })
  : new Response(`<html><meta charset="utf-8"><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } }) })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  await page.route("http://grid.invalid/**", route => route.request().method() === "GET"
    ? route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(grid) })
    : route.fulfill({ status: 204 }))
  await page.route("http://engine.invalid/**", route => route.fulfill({ status: 404 }))
  await page.goto(server.url.href)
  const table = page.getByRole("dialog", { name: "Review table — Supplier review" })
  const answer = page.getByRole("button", { name: /^The law of England and Wales\./ })
  const detail = page.locator(".picker-panel").filter({ hasText: "What is the governing law?" })
  const closes = () => page.evaluate(() => (window as unknown as { gridCloses?: number }).gridCloses ?? 0)
  const focused = () => page.evaluate(() => document.activeElement?.tagName)

  // Focus stays on the cell that opened the detail, moves to the dialog when the
  // detail's text is clicked, and falls to the page when a comment is deleted.
  // Repeating the press also covers Chrome refusing to let a later cancel be prevented.
  for (const focus of ["cell", "detail text", "page"] as const) {
    await answer.click()
    await detail.waitFor()
    if (focus === "detail text") await detail.locator(".review-detail-answer").click()
    if (focus === "page") {
      await detail.locator("textarea").fill("Confirm the governing law with the client.")
      await detail.getByRole("button", { name: "Send", exact: true }).click()
      await detail.getByTitle("Delete comment").click()
      await detail.getByText("Confirm the governing law with the client.").waitFor({ state: "detached" })
      assert.equal(await focused(), "BODY")
    }
    await page.keyboard.press("Escape")
    await detail.waitFor({ state: "detached" })
    assert.equal(await table.isVisible(), true, `one Escape (focus on ${focus}) closed the review table, not just the cell detail`)
    assert.equal(await closes(), 0)
    console.log(`PASS Escape with focus on the ${focus} closes the cell detail and keeps the review table open`)
  }

  await page.keyboard.press("Escape")
  await table.waitFor({ state: "detached" })
  assert.equal(await closes(), 1)
  console.log("PASS the next Escape closes the review table once")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
