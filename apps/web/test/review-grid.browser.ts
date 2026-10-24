// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/review-grid.browser.ts
// Mounts the real ReviewGrid as MatterDetail mounts it, and the real DocumentViewer on the Docxodus WASM
// runtime; only ingest HTTP is mocked. No extraction runs.
// Escape closes the innermost layer: the cited document, then a cell's detail, then the review table.
// A cell's citation opens its document at the cited excerpt; a cell whose excerpt is gone reads degraded.
// Set PLAYWRIGHT_CHROMIUM_PATH=<chromium> when Playwright's own browser is missing.
import assert from "node:assert/strict"
import { existsSync, statSync } from "node:fs"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const { zipSync, strToU8 } = (await import(Bun.resolveSync("fflate", path.resolve(import.meta.dir, "../../../dochaus")))) as {
  zipSync: (files: Record<string, Uint8Array>) => Uint8Array; strToU8: (text: string) => Uint8Array
}
const wasmDir = path.resolve(process.env.DOCXODUS_WASM_DIR ?? path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm"))
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-review-grid-"))
const excerpt = "This agreement is governed by the law of England and Wales."
// Synthetic supply terms; the cited sentence is split across runs, as Word stores it.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const supplyTerms = zipSync(Object.fromEntries(Object.entries({
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/document.xml": `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>1. Supply. The Supplier shall supply the Goods.</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">12. This agreement is governed by </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>the law of England and Wales.</w:t></w:r></w:p><w:p><w:r><w:t>13. Notices must be in writing.</w:t></w:r></w:p></w:body></w:document>`,
}).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)])))
const grid = {
  columns: [{ id: "c1", question: "What is the governing law?" }],
  cells: {
    "Supply terms.docx::c1": {
      answer: "The law of England and Wales.",
      citation: { documentName: "Supply terms.docx", docPath: "Supply terms.docx", section: "12", structuralPath: "Supply terms › §12 Governing law", excerpt, charStart: 48, charEnd: 106 },
      status: "filled",
      questionHash: "unused",
    },
    // The side letter was amended after this answer: its excerpt no longer matches.
    "Side letter.docx::c1": {
      answer: "The law of Scotland.",
      citation: { documentName: "Side letter.docx", docPath: "Side letter.docx", section: "4", excerpt: "This letter is governed by the law of Scotland.", charStart: 0, charEnd: 47 },
      status: "filled",
      questionHash: "unused",
    },
    // Left out of the source check's reply, as ingest leaves out a cell it could not check.
    "Old note.docx::c1": {
      answer: "English law.",
      citation: { documentName: "Old note.docx", docPath: "Old note.docx", section: "1", excerpt: "English law applies.", charStart: 0, charEnd: 20 },
      status: "filled",
      questionHash: "unused",
    },
  },
}
const sourceChecks: Record<string, string> = { "Supply terms.docx::c1": "verified", "Side letter.docx::c1": "not_found" }
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
    // Supply terms' extracted text is degraded (#239): its row and the citing cell say so.
    documents: [
      { id: 1, name: "Supply terms.docx", doc_path: "Supply terms.docx", created_at: 0,
        quality: { status: "degraded", signals: [{ code: "replacement_characters", detail: "40 of 900 characters are replacement or control characters" }] } },
      { id: 2, name: "Side letter.docx", doc_path: "Side letter.docx", created_at: 0 },
      { id: 3, name: "Old note.docx", doc_path: "Old note.docx", created_at: 0 },
    ],
    onClose: () => { window.gridCloses = (window.gridCloses ?? 0) + 1; setOpen(false) },
    onAsk: (prompt) => { window.asked = prompt },
  }) : null
}
createRoot(document.getElementById("root")).render(createElement(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://grid.invalid", VITE_OPENCODE_URL: "http://engine.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => {
  const { pathname } = new URL(request.url)
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname.startsWith("/wasm/")) {
    const file = path.join(wasmDir, decodeURIComponent(pathname.slice("/wasm/".length)))
    return file.startsWith(wasmDir + path.sep) && existsSync(file) && statSync(file).isFile() ? new Response(Bun.file(file)) : new Response("Not found", { status: 404 })
  }
  return new Response(`<html><meta charset="utf-8"><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  const checked: string[][] = []
  const pendingChecks: import("@playwright/test").Route[] = []
  let holdChecks = false
  let pendingEdit = false
  let acceptedEdits = 0
  await page.route("http://grid.invalid/**", route => {
    const request = route.request()
    const url = new URL(request.url())
    const json = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (url.pathname === "/matters/matter-grid/grid") return request.method() === "GET" ? json(grid) : route.fulfill({ status: 204 })
    if (url.pathname === "/matters/matter-grid/grid/sources") {
      const cells = (request.postDataJSON() as { cells: { key: string }[] }).cells
      checked.push(cells.map(cell => cell.key))
      if (holdChecks) { pendingChecks.push(route); return }
      return json(Object.fromEntries(cells.map(cell => [cell.key, sourceChecks[cell.key]])))
    }
    if (url.searchParams.get("name") === "Supply terms.docx" && /^\/matters\/matter-grid\/documents\/(redlined|content)$/.test(url.pathname))
      return route.fulfill({ status: 200, contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", body: Buffer.from(supplyTerms) })
    if (url.pathname === "/matters/matter-grid/redlines") return json(pendingEdit ? [{
      id: 1, doc_name: "Supply terms.docx", scope: "phrase", find_text: "Notices must be in writing.",
      old_text: "Notices must be in writing.", new_text: "Notices must be delivered in writing.",
      author: "Reviewer", review_comment: "Clarify delivery.", created_at: 0,
    }] : [])
    if (url.pathname === "/matters/matter-grid/redlines/1/accept") {
      pendingEdit = false
      acceptedEdits++
      return json({})
    }
    if (url.pathname === "/matters/matter-grid/documents/amendments") return json(null)
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: `Unexpected ${url.pathname}` }) })
  })
  await page.route("http://engine.invalid/**", route => route.fulfill({ status: 404 }))
  await page.goto(server.url.href)
  const table = page.getByRole("dialog", { name: "Review table — Supplier review" })
  const answer = page.getByRole("button", { name: /^The law of England and Wales\./ })
  const detail = page.locator(".picker-panel").filter({ hasText: "The law of England and Wales." })
  const detailModal = page.getByRole("dialog", { name: /^Answer and comments:/ })
  const closes = () => page.evaluate(() => (window as unknown as { gridCloses?: number }).gridCloses ?? 0)
  const focusInDetail = () => detailModal.evaluate(element => element.contains(document.activeElement))

  await table.getByRole("rowheader", { name: /Supply terms\.docx/ }).getByText("Text may be incomplete or garbled").waitFor()
  assert.match(await answer.innerText(), /cited text may be incomplete or garbled/)
  console.log("PASS a degraded source and the cells citing it say its text may be incomplete")

  // The native detail owns focus after opening an answer, after clicking its
  // reading surface and after removing a focused comment control. Repeating the
  // press also covers Chrome refusing to prevent a later native cancel.
  for (const focus of ["opening answer", "detail text", "deleted comment"] as const) {
    await answer.click()
    await detail.waitFor()
    assert.equal(await detailModal.evaluate(element => (element as HTMLDialogElement).open), true)
    await detail.getByRole("textbox", { name: "Add a comment", exact: true }).waitFor()
    assert.equal(await detail.getByRole("textbox", { name: "Add a comment", exact: true }).evaluate(element => element === document.activeElement), true, "opening an answer should focus its comment field")
    if (focus === "detail text") await detail.locator(".review-detail-answer").click()
    if (focus === "deleted comment") {
      await detail.locator("textarea").fill("Confirm the governing law with the client.")
      await detail.getByRole("button", { name: "Add comment", exact: true }).click()
      await detail.getByTitle("Delete comment").click()
      await detail.getByText("Confirm the governing law with the client.").waitFor({ state: "detached" })
    }
    assert.equal(await focusInDetail(), true, `focus escaped the active detail after ${focus}`)
    // With an empty composer, Add comment is disabled. The native sequence is
    // Close → textarea → Close, in both directions, without entering the table.
    await detail.locator("textarea").focus()
    await page.keyboard.press("Tab")
    assert.equal(await detail.getByRole("button", { name: "Close", exact: true }).evaluate(element => element === document.activeElement), true, "Tab should wrap to the detail's close button")
    await page.keyboard.press("Shift+Tab")
    assert.equal(await detail.locator("textarea").evaluate(element => element === document.activeElement), true, "Shift+Tab should wrap to the detail's comment field")
    assert.equal(await focusInDetail(), true)
    await page.keyboard.press("Escape")
    await detail.waitFor({ state: "detached" })
    assert.equal(await table.isVisible(), true, `one Escape (focus on ${focus}) closed the review table, not just the cell detail`)
    assert.equal(await closes(), 0)
    assert.equal(await answer.evaluate(element => element === document.activeElement), true, "closing a detail should restore its answer button")
    console.log(`PASS native focus containment, Tab loop, restoration and Escape after ${focus}`)
  }

  // The amended side letter reads degraded in the grid itself, and its detail says why.
  const sideLetterRow = page.getByRole("row", { name: /^Side letter\.docx/ })
  await sideLetterRow.getByText("Source changed", { exact: true }).waitFor()
  assert.equal(await page.getByRole("row", { name: /^Supply terms\.docx/ }).getByText("Source changed").count(), 0, "a verified cell should not read degraded")
  assert.deepEqual(checked.at(-1)?.sort(), ["Old note.docx::c1", "Side letter.docx::c1", "Supply terms.docx::c1"])
  await page.getByRole("button", { name: /^The law of Scotland\./ }).click()
  const sideDetail = page.locator(".picker-panel").filter({ hasText: "The law of Scotland." })
  await sideDetail.getByText("Not found", { exact: true }).waitFor()
  assert.match((await sideDetail.locator(".review-source-check").textContent())!, /no longer contains this excerpt\. It has changed since this answer/)
  await page.keyboard.press("Escape")
  await sideDetail.waitFor({ state: "detached" })
  console.log("PASS a cell whose excerpt is no longer in its document reads degraded in the grid and its detail")

  // A cell the check left out says so instead of waiting forever.
  await page.getByRole("button", { name: /^English law\./ }).click()
  const oldDetail = page.locator(".picker-panel").filter({ hasText: "English law." })
  await oldDetail.getByText("The excerpt could not be checked against the document.", { exact: true }).waitFor()
  await page.keyboard.press("Escape")
  await oldDetail.waitFor({ state: "detached" })
  console.log("PASS a cell the source check could not check reads as unchecked")

  // The citation opens the document over the detail with the cited excerpt
  // highlighted; Escape returns to the same cell's detail, focus on the citation.
  await answer.click()
  await detail.getByText("Verified", { exact: true }).waitFor()
  await detail.getByText("Supply terms › §12 Governing law", { exact: true }).waitFor()
  await detail.locator(".review-source").getByText("cited text may be incomplete or garbled").waitFor()
  const source = detail.getByRole("button", { name: "Open Supply terms.docx at section 12", exact: true })
  await source.click()
  const viewer = page.getByRole("dialog", { name: "Supply terms.docx", exact: true })
  const highlight = viewer.locator(".docx-render mark.citation-passage")
  await highlight.first().waitFor({ timeout: 60_000 })
  const highlighted = (await highlight.allTextContents()).join("").replace(/\s+/g, " ").trim()
  assert.equal(highlighted, excerpt, "the viewer should highlight exactly the cited excerpt")
  assert.equal(await highlight.first().evaluate(element => element === document.activeElement), true, "the highlighted excerpt should take focus")
  await viewer.getByText("The cited excerpt is highlighted.", { exact: false }).waitFor()
  await page.keyboard.press("Escape")
  await viewer.waitFor({ state: "detached" })
  assert.equal(await detail.isVisible(), true, "Escape in the document should return to the cell's detail")
  assert.equal(await source.evaluate(element => element === document.activeElement), true, "closing the document should restore its citation link")
  assert.equal(await table.isVisible(), true)
  assert.equal(await closes(), 0)
  console.log("PASS the citation opens its document at the highlighted excerpt and Escape returns to the same cell")

  // Asking hands chat the cell's question, answer and recorded citation.
  await detail.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
  const asked = await page.evaluate(() => (window as unknown as { asked?: string }).asked ?? "")
  for (const part of ["Question: What is the governing law?", "Answer: The law of England and Wales.", "Cited source: Supply terms.docx § 12 (Supply terms › §12 Governing law)", `Excerpt: "${excerpt}"`, "Source check: Verified."])
    assert.ok(asked.includes(part), `the chat question should carry "${part}"`)
  console.log("PASS asking about a cell hands chat its question, answer and citation")

  // Accept through the real viewer. Keep source-check HTTP replies pending so
  // the old verdict cannot be hidden by a fast replacement response.
  holdChecks = true
  async function editSource() {
    pendingEdit = true
    const nextCheck = page.waitForRequest(request => request.url().endsWith("/grid/sources"))
    await source.click()
    await viewer.getByRole("button", { name: "Accept", exact: true }).click({ timeout: 60_000 })
    await nextCheck
    await viewer.getByRole("button", { name: "Close", exact: true }).click()
    await viewer.waitFor({ state: "detached" })
    assert.equal(await detail.locator(".review-source-badge").count(), 0, "a source edit must immediately invalidate the previous verdict")
    assert.equal(await detail.locator(".review-source-check").textContent(), "Checking the excerpt against the document…")
    await askWithoutVerdict()
  }
  async function askWithoutVerdict() {
    await detail.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
    const prompt = await page.evaluate(() => (window as unknown as { asked?: string }).asked ?? "")
    assert.ok(prompt.includes(`Excerpt: "${excerpt}"`), "the recorded citation is still available")
    assert.doesNotMatch(prompt, /Source check:/, "chat must not receive an obsolete source verdict")
  }
  async function finishCheck(index: number, check: string | undefined) {
    const route = pendingChecks[index]
    assert.ok(route, `source check ${index} should be pending`)
    // Wait for delivery and rendering before checking an obsolete reply's effect.
    const received = page.waitForResponse(response => response.request() === route.request())
    await route.fulfill(check
      ? { status: 200, contentType: "application/json", body: JSON.stringify({ "Supply terms.docx::c1": check }) }
      : { status: 500, contentType: "application/json", body: JSON.stringify({ error: "Source unavailable" }) })
    await (await received).finished()
    await page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))))
  }
  await editSource()
  await finishCheck(0, undefined)
  await detail.getByText("The excerpt could not be checked against the document.", { exact: true }).waitFor()
  assert.equal(await detail.locator(".review-source-badge").count(), 0)
  await askWithoutVerdict()
  console.log("PASS accepting a document edit invalidates Verified while pending and after a failed recheck, including chat")

  await editSource()
  await finishCheck(1, "reanchored")
  await detail.getByText("Re-anchored", { exact: true }).waitFor()
  await detail.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
  assert.match(await page.evaluate(() => (window as unknown as { asked?: string }).asked ?? ""), /Source check: Re-anchored\./)
  console.log("PASS a fresh successful recheck restores its current verdict and chat context")

  // Two newer edits supersede checks still in flight. Neither an obsolete
  // success nor an obsolete failure may overwrite the newest result.
  await editSource()
  await editSource()
  await editSource()
  await finishCheck(4, "not_found")
  await detail.getByText("Not found", { exact: true }).waitFor()
  await finishCheck(2, "verified")
  await finishCheck(3, undefined)
  assert.equal(await detail.locator(".review-source-badge").textContent(), "Not found")
  await detail.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
  assert.match(await page.evaluate(() => (window as unknown as { asked?: string }).asked ?? ""), /Source check: Not found\./)
  assert.equal(acceptedEdits, 5)
  console.log("PASS late success and failure responses cannot replace the latest source verdict")

  await page.keyboard.press("Escape")
  await detail.waitFor({ state: "detached" })
  assert.equal(await table.isVisible(), true)
  await page.keyboard.press("Escape")
  await table.waitFor({ state: "detached" })
  assert.equal(await closes(), 1)
  console.log("PASS the next Escape closes the review table once")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
