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
import { read, utils } from "xlsx"
import { hashQuestion, type Grid } from "../src/api/ingest/grid"

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
const grid: Grid = {
  version: 0,
  columns: [{ id: "c1", question: "What is the governing law?" }],
  cells: {
    "Supply terms.docx::c1": {
      answer: "The law of England and Wales.",
      citation: { documentName: "Supply terms.docx", docPath: "Supply terms.docx", section: "12", structuralPath: "Doc(Supply terms.docx) › §12 Governing law", excerpt, charStart: 48, charEnd: 106 },
      status: "filled",
      questionHash: "unused",
    },
    // The side letter was amended after this answer: its excerpt no longer matches.
    "Side letter.docx::c1": {
      answer: "The law of Scotland.",
      citation: { documentName: "Side letter.docx", docPath: "Side letter.docx", section: "4", excerpt: "This letter is governed by the law of Scotland.", charStart: 0, charEnd: 47 },
      status: "filled",
      questionHash: "unused",
      comments: [{ id: "legacy", text: "Check the earlier version.", at: Date.UTC(2026, 9, 11, 12) }],
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
import { createBrowserRouter, RouterProvider, Link, useLocation, useNavigate } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-router-dom")))}
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
      ...(location.search === "?pool" ? Array.from({ length: 4 }, (_, i) => ({ id: 10 + i, name: "Extra " + i + ".docx", doc_path: "Extra " + i + ".docx", created_at: 0 })) : []),
    ],
    onClose: () => { window.gridCloses = (window.gridCloses ?? 0) + 1; setOpen(false) },
    onAsk: (prompt) => { window.asked = prompt },
  }) : null
}
function NavigationHarness() {
  const location = useLocation()
  const navigate = useNavigate()
  const open = new URLSearchParams(location.search).get("view") === "review"
  return createElement("div", null,
    createElement(Link, { to: "/nav?view=review" }, "Open review"),
    createElement(Link, { to: "/outside" }, "Other page"),
    open && createElement(ReviewGrid, {
      matterId: "matter-grid", title: "Supplier review", directory: "/matters/supplier-review",
      documents: [{ id: 1, name: "Supply terms.docx", doc_path: "Supply terms.docx", created_at: 0 }],
      onClose: () => navigate("/nav"), onAsk: () => navigate("/nav?view=chat"),
    }))
}
const node = document.getElementById("root")
if (location.pathname === "/nav") {
  const router = createBrowserRouter([{ path: "*", element: createElement(NavigationHarness) }])
  createRoot(node).render(createElement(RouterProvider, { router, useTransitions: false }))
} else createRoot(node).render(createElement(Harness))
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
// Edits a column's question in its dialog, as the lawyer does.
async function editQuestion(scope: import("@playwright/test").Page | import("@playwright/test").Locator, from: string, to: string) {
  await scope.getByRole("button", { name: `Edit question: ${from}` }).click()
  const editor = scope.getByRole("dialog", { name: "Edit question" })
  await editor.getByRole("textbox", { name: "Question" }).fill(to)
  await editor.getByRole("button", { name: "Save", exact: true }).click()
  await editor.waitFor({ state: "detached" })
}
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
try {
  const page = await browser.newPage()
  page.setDefaultTimeout(5000)
  const checked: string[][] = []
  const pendingChecks: import("@playwright/test").Route[] = []
  let holdChecks = false
  let pendingEdit = false
  let acceptedEdits = 0
  let storedGrid = structuredClone(grid)
  let attorney = "  Amina Shah  "
  let preferencesFailed = false
  let heldPreferences: import("@playwright/test").Route | undefined
  let holdPreferences = false
  let holdPut = false
  let heldPut: import("@playwright/test").Route | undefined
  let holdGridGet = false
  let heldGridGet: import("@playwright/test").Route | undefined
  let gridGets = 0
  let onExtraction: ((route: import("@playwright/test").Route) => void) | undefined
  const extractionWaits: import("@playwright/test").Route[] = []
  let extractionStarts = 0
  // Review-compare sessions: each minted id, its held wait and its deletion.
  const compareSessions: string[] = []
  const compareDeleted: string[] = []
  const compareWaits: import("@playwright/test").Route[] = []
  let holdCompare = false
  let compareReply = ""
  const putVersions: number[] = []
  const save = (route: import("@playwright/test").Route) => {
    const incoming = route.request().postDataJSON() as Grid
    putVersions.push(incoming.version ?? -1)
    if (incoming.version !== storedGrid.version) return route.fulfill({ status: 409, contentType: "application/json", body: JSON.stringify({ error: "Conflict", grid: storedGrid }) })
    for (const [key, prior] of Object.entries(storedGrid.cells)) {
      const next = incoming.cells[key]
      if (prior.status === "reviewed" && next?.status === "filled" && JSON.stringify({ ...next, status: "reviewed" }) !== JSON.stringify(prior))
        return route.fulfill({ status: 409, contentType: "application/json", body: JSON.stringify({ error: "Unlock separately", grid: storedGrid }) })
    }
    storedGrid = { ...incoming, version: (storedGrid.version ?? 0) + 1 }
    return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(storedGrid) })
  }
  await page.route("http://grid.invalid/**", route => {
    const request = route.request()
    const url = new URL(request.url())
    const json = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (url.pathname === "/matters/matter-grid/grid") {
      if (request.method() === "GET") {
        gridGets++
        if (holdGridGet) { heldGridGet = route; holdGridGet = false; return }
        return json(storedGrid)
      }
      if (holdPut) { heldPut = route; holdPut = false; return }
      return save(route)
    }
    if (url.pathname === "/preferences") {
      if (holdPreferences) { heldPreferences = route; return }
      return preferencesFailed ? route.fulfill({ status: 503 }) : json({ attorney, firm: "Synthetic Review LLP" })
    }
    if (url.pathname === "/matters/matter-grid") return json({ documents: [
      { name: "Supply terms.docx" }, { name: "Side letter.docx" }, { name: "Old note.docx" },
      ...(page.url().endsWith("?pool") ? Array.from({ length: 4 }, (_, i) => ({ name: `Extra ${i}.docx` })) : []),
    ] })
    if (url.pathname === "/engine/config") return json({ model: "test/model" })
    if (url.pathname === "/matters/matter-grid/grid/recheck") return json({ citations: [], passages: "", scope: { passages: 0, read: 0, candidates: 0, reread: 0, terms: [], channels: [] } })
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
  await page.route("http://engine.invalid/**", route => {
    const url = new URL(route.request().url())
    const key = `${route.request().method()} ${url.pathname}`
    const json = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    const compared = compareSessions.find(id => url.pathname.includes(`/session/${id}`))
    if (key === "POST /api/session" && (route.request().postDataJSON() as { agent?: string }).agent === "review-compare") {
      const id = (route.request().postDataJSON() as { id: string }).id
      compareSessions.push(id)
      return json({ data: { id, location: { directory: "/matters/supplier-review" } } })
    }
    if (compared) {
      if (key.endsWith("/prompt")) return json({ data: { id: "msg_compare", type: "user", sessionID: compared, time: { created: 1 }, payload: { text: "compare" }, delivery: "queue" } })
      if (key.endsWith("/wait")) { if (holdCompare) { compareWaits.push(route); return } return route.fulfill({ status: 204 }) }
      if (key.endsWith("/interrupt")) return json({ interrupted: true })
      if (key.startsWith("GET") && key.endsWith("/message")) return json({ data: [
        { type: "assistant", id: "msg_groups", agent: "review-compare", model: { providerID: "test", id: "model" }, finish: "stop", time: { created: 1, completed: 2 }, content: [{ type: "text", text: compareReply }] },
        { type: "idle", id: "msg_idle", outcome: "succeeded", time: { created: 3 } },
      ], cursor: {} })
      if (key === `DELETE /api/session/${compared}`) { compareDeleted.push(compared); return route.fulfill({ status: 204 }) }
    }
    if (key === "POST /api/session") { extractionStarts++; return json({ data: { id: "ses_grid", location: { directory: "/matters/supplier-review" } } }) }
    if (/^POST \/api\/session\/ses_grid\/(agent|model)$/.test(key)) return route.fulfill({ status: 204 })
    if (key === "POST /api/session/ses_grid/prompt") return json({ data: { id: "msg_user", type: "user", sessionID: "ses_grid", time: { created: 1 }, payload: { text: "test" }, delivery: "queue" } })
    if (key === "POST /api/experimental/session/ses_grid/wait") { extractionWaits.push(route); onExtraction?.(route); return }
    if (key === "GET /api/session/ses_grid/message") {
      const assistant = { type: "assistant", id: "msg_answer", agent: "extract", model: { providerID: "test", id: "model" }, finish: "stop", time: { created: 1, completed: 2 }, content: [{ type: "text", text: "New extracted answer" }] }
      const idle = { type: "idle", id: "msg_idle", outcome: "succeeded", time: { created: 3 } }
      return json({ data: url.searchParams.get("order") === "desc" ? [idle, assistant] : [assistant, idle], cursor: {} })
    }
    return route.fulfill({ status: 404, body: `Unexpected ${key}` })
  })
  await page.goto(server.url.href)
  const table = page.getByRole("dialog", { name: "Review table — Supplier review" })
  const answer = page.getByRole("button", { name: /^The law of England and Wales\./ })
  const detail = page.locator(".picker-panel").filter({ hasText: "The law of England and Wales." })
  const detailModal = page.getByRole("dialog", { name: /^Answer and comments:/ })
  const closes = () => page.evaluate(() => (window as unknown as { gridCloses?: number }).gridCloses ?? 0)
  const focusInDetail = () => detailModal.evaluate(element => element.contains(document.activeElement))

  await table.getByRole("rowheader", { name: /Supply terms\.docx/ }).getByText("Text may be incomplete or garbled").waitFor()
  assert.equal(putVersions.length, 0, "opening the grid must not write it")
  // No comments, no comment button: the first comment is added in the answer detail (below).
  assert.equal(await page.getByRole("row", { name: /^Supply terms\.docx/ }).getByRole("button", { name: /comments? on/ }).count(), 0, "a cell without comments shows no comment button")
  assert.equal(await page.getByRole("row", { name: /^Side letter\.docx/ }).getByRole("button", { name: /^1 comment on Side letter\.docx/ }).count(), 1, "a commented cell shows its count")
  // Keep the first save in flight while the lawyer makes a second change.
  holdPut = true
  const firstPut = page.waitForRequest(request => request.url().endsWith("/grid") && request.method() === "PUT")
  const lock = page.getByRole("row", { name: /^Supply terms\.docx/ }).getByTitle("Mark reviewed")
  await lock.click()
  await firstPut
  await page.getByRole("row", { name: /^Supply terms\.docx/ }).getByTitle("Unlock").click()
  await answer.click()
  await detail.getByRole("textbox", { name: "Add a comment" }).fill("Note after unlocking")
  await detail.getByRole("button", { name: "Add comment", exact: true }).click()
  await detail.getByText("Note after unlocking").waitFor()
  page.once("dialog", dialog => dialog.dismiss())
  await detail.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
  assert.equal(await page.evaluate(() => (window as unknown as { asked?: string }).asked), undefined, "chat navigation must not discard a pending save")
  assert.ok(heldPut)
  await save(heldPut)
  await page.waitForFunction(() => document.querySelector('[title="Mark reviewed"]') !== null)
  await page.waitForTimeout(100)
  assert.deepEqual(putVersions, [0, 1, 2], "queued saves must retain the unlock before the following edit")
  assert.equal(storedGrid.cells["Supply terms.docx::c1"]!.status, "filled")
  assert.equal(storedGrid.cells["Supply terms.docx::c1"]!.comments![0]!.text, "Note after unlocking")
  await detail.getByRole("button", { name: "Delete comment", exact: true }).click()
  await page.keyboard.press("Escape")
  await page.waitForTimeout(100)
  console.log("PASS opening is read-only and overlapping edits are saved in order")
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
    // Its one next step, not the comment field, whose keyboard would cover the answer on a phone.
    // The cited excerpt no longer matches the degraded text, so that step is Re-run this cell.
    assert.equal(await detail.locator("button.primary").evaluate(element => element === document.activeElement && element.textContent === "Re-run this cell"), true, "opening an answer should focus its next step")
    if (focus === "detail text") await detail.locator(".review-detail-answer").click()
    if (focus === "deleted comment") {
      await detail.locator("textarea").fill("Confirm the governing law with the client.")
      await detail.getByRole("button", { name: "Add comment", exact: true }).click()
      await detail.getByRole("button", { name: "Delete comment", exact: true }).click()
      await detail.getByText("Confirm the governing law with the client.").waitFor({ state: "detached" })
    }
    assert.equal(await focusInDetail(), true, `focus escaped the active detail after ${focus}`)
    // With an empty composer, Add comment is disabled. The native sequence is
    // Close → … → Re-run this cell → Close, in both directions, without entering the table.
    await detail.getByRole("button", { name: "Re-run this cell", exact: true }).focus()
    await page.keyboard.press("Tab")
    assert.equal(await detail.getByRole("button", { name: "Close", exact: true }).evaluate(element => element === document.activeElement), true, "Tab should wrap to the detail's close button")
    await page.keyboard.press("Shift+Tab")
    assert.equal(await detail.getByRole("button", { name: "Re-run this cell", exact: true }).evaluate(element => element === document.activeElement), true, "Shift+Tab should wrap to the detail's last step")
    assert.equal(await focusInDetail(), true)
    await page.keyboard.press("Escape")
    await detail.waitFor({ state: "detached" })
    assert.equal(await table.isVisible(), true, `one Escape (focus on ${focus}) closed the review table, not just the cell detail`)
    assert.equal(await closes(), 0)
    assert.equal(await answer.evaluate(element => element === document.activeElement), true, "closing a detail should restore its answer button")
    console.log(`PASS native focus containment, Tab loop, restoration and Escape after ${focus}`)
  }

  // Real comment creation snapshots the configured Attorney, including when a
  // later comment uses a new setting. Older comments never borrow that name.
  await answer.click()
  const compose = detail.getByRole("textbox", { name: "Add a comment", exact: true })
  const comments = detail.locator(".review-comment")
  holdPreferences = true
  const adding = page.waitForRequest(request => new URL(request.url()).pathname === "/preferences")
  await compose.fill("Confirm with the client.")
  await detail.getByRole("button", { name: "Add comment", exact: true }).click()
  await adding
  assert.equal(await detail.getByRole("button", { name: "Adding comment…", exact: true }).isDisabled(), true)
  await compose.press("Enter") // A pending submission must not create a duplicate.
  assert.ok(heldPreferences)
  await heldPreferences.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ attorney, firm: "Synthetic Review LLP" }) })
  holdPreferences = false
  await comments.getByText("Amina Shah", { exact: true }).waitFor()
  assert.equal(await comments.count(), 1)
  attorney = "Jordan Chen"
  await compose.fill("<img src=x onerror=window.commentInjected=true>")
  await detail.getByRole("button", { name: "Add comment", exact: true }).click()
  await comments.getByText("Jordan Chen", { exact: true }).waitFor()
  assert.equal(await comments.getByText("Amina Shah", { exact: true }).count(), 1)
  assert.equal(await comments.locator("img").count(), 0, "comment text must render as text")
  attorney = ""
  await compose.fill("No name is configured.")
  await detail.getByRole("button", { name: "Add comment", exact: true }).click()
  await comments.filter({ hasText: "No name is configured." }).getByText("Unattributed", { exact: true }).waitFor()
  preferencesFailed = true
  await compose.fill("Settings are unavailable.")
  await detail.getByRole("button", { name: "Add comment", exact: true }).click()
  await comments.filter({ hasText: "Settings are unavailable." }).getByText("Unattributed", { exact: true }).waitFor()
  await page.waitForTimeout(100)
  assert.deepEqual(storedGrid.cells["Supply terms.docx::c1"]!.comments!.map(comment => comment.author), ["Amina Shah", "Jordan Chen", undefined, undefined])
  assert.equal(storedGrid.cells["Side letter.docx::c1"]!.comments![0]!.author, undefined)
  preferencesFailed = false
  attorney = "Different Current Name"
  await page.reload()
  await answer.click()
  await comments.getByText("Amina Shah", { exact: true }).waitFor()
  assert.equal(await comments.getByText("Jordan Chen", { exact: true }).count(), 1)
  assert.equal(await comments.getByText("Unattributed", { exact: true }).count(), 2)
  assert.equal(await comments.getByText(attorney, { exact: true }).count(), 0)
  await page.keyboard.press("Escape")
  await detail.waitFor({ state: "detached" })
  // Before every cell is reviewed, export is one quiet icon with a menu of formats.
  const exportButton = table.getByRole("button", { name: "Export", exact: true })
  await exportButton.click()
  const formats = table.getByRole("group", { name: "Export as" })
  assert.deepEqual(await formats.getByRole("button").allTextContents(), ["Excel", "CSV", "Word"])
  await page.keyboard.press("ArrowDown")
  assert.equal(await formats.getByRole("button", { name: "Excel", exact: true }).evaluate(element => element === document.activeElement), true, "arrow keys move into the formats")
  await page.keyboard.press("Escape")
  await formats.waitFor({ state: "detached" })
  assert.equal(await exportButton.evaluate(element => element === document.activeElement), true, "Escape returns focus to Export")
  await exportButton.click()
  const downloading = page.waitForEvent("download")
  await formats.getByRole("button", { name: "Excel", exact: true }).click()
  const downloaded = await (await downloading).path()
  assert.ok(downloaded)
  const exported = read(await Bun.file(downloaded).arrayBuffer(), { type: "array" })
  const records = utils.sheet_to_json<string[]>(exported.Sheets.Comments, { header: 1 })
  assert.deepEqual(records.slice(1).map(record => record[3]), ["Amina Shah", "Jordan Chen", "Unattributed", "Unattributed", "Unattributed"])
  assert.equal(records[1]![2], "Supply terms.docx § 12")
  assert.equal(records[2]![5], "<img src=x onerror=window.commentInjected=true>")
  assert.match(exported.Sheets["Tabular review"].B2.c[0].t, /Source: Supply terms\.docx § 12\nComment by Amina Shah/)
  console.log("PASS comments snapshot saved Attorney names, remain unattributed without settings, reload and export with separate sources")

  // The amended side letter reads degraded in the grid itself, and its detail says why.
  const sideLetterRow = page.getByRole("row", { name: /^Side letter\.docx/ })
  await sideLetterRow.getByText("Source changed", { exact: true }).waitFor()
  assert.equal(await page.getByRole("row", { name: /^Supply terms\.docx/ }).getByText("Source changed").count(), 0, "a verified cell should not read degraded")
  // One marker per cell, named by its word; its tooltip gives the reason and the cell's other states.
  const sideMarker = sideLetterRow.getByRole("button", { name: "Source changed", exact: true })
  assert.equal(await sideLetterRow.locator(".review-marker").count(), 1, "one marker per cell")
  await sideMarker.focus()
  assert.match((await page.getByRole("tooltip").textContent())!, /no longer contains the cited excerpt\. .*Also: Question changed\.$/)
  await page.getByRole("row", { name: /^Supply terms\.docx/ }).getByRole("button", { name: "Question changed", exact: true }).waitFor()
  assert.deepEqual(checked.at(-1)?.sort(), ["Old note.docx::c1", "Side letter.docx::c1", "Supply terms.docx::c1"])
  await page.getByRole("button", { name: /^The law of Scotland\./ }).click()
  const sideDetail = page.locator(".picker-panel").filter({ hasText: "The law of Scotland." })
  await sideDetail.locator(".review-comment").getByText("Unattributed", { exact: true }).waitFor()
  await sideDetail.getByText("Source changed", { exact: true }).waitFor()
  assert.match((await sideDetail.locator(".review-source-check").textContent())!, /has changed since this answer and no longer contains this excerpt\. Re-run this cell to answer from the current text\.$/, "an unreviewed cell names its own next step")
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
  // highlighted; Escape returns to the same cell's detail, focus on its link.
  await answer.click()
  await detail.getByText("Verified", { exact: true }).waitFor()
  // The dialog names the document once; its path shows the sections only.
  await detail.getByText("§12 Governing law", { exact: true }).waitFor()
  await detail.locator(".review-source").getByText("cited text may be incomplete or garbled").waitFor()
  const source = detail.getByRole("button", { name: "View in document: Supply terms.docx at section 12", exact: true })
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
  assert.equal(await source.evaluate(element => element === document.activeElement), true, "closing the document should restore focus to View in document")
  assert.equal(await table.isVisible(), true)
  assert.equal(await closes(), 0)
  console.log("PASS the citation opens its document at the highlighted excerpt and Escape returns to the same cell")

  // Asking hands chat the cell's question, answer and recorded citation.
  await detail.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
  const asked = await page.evaluate(() => (window as unknown as { asked?: string }).asked ?? "")
  for (const part of ["Question: What is the governing law?", "Answer: The law of England and Wales.", "Cited source: Supply terms.docx § 12 (Doc(Supply terms.docx) › §12 Governing law)", `Excerpt: "${excerpt}"`, "Source check: Verified."])
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
  await detail.getByText("Source changed", { exact: true }).waitFor()
  await finishCheck(2, "verified")
  await finishCheck(3, undefined)
  assert.equal(await detail.locator(".review-source-badge").textContent(), "Source changed")
  await detail.getByRole("button", { name: "Ask about this answer in chat", exact: true }).click()
  assert.match(await page.evaluate(() => (window as unknown as { asked?: string }).asked ?? ""), /Source check: Source changed\./)
  assert.equal(acceptedEdits, 5)
  console.log("PASS late success and failure responses cannot replace the latest source verdict")

  await page.keyboard.press("Escape")
  await detail.waitFor({ state: "detached" })
  assert.equal(await table.isVisible(), true)

  // A stale save cannot overwrite the other window. Local edits stay visible,
  // a close needs confirmation, and reload is a deliberate discard.
  storedGrid = { ...storedGrid, version: (storedGrid.version ?? 0) + 1 }
  await editQuestion(table, "What is the governing law?", "What is the local question?")
  await table.getByText("Your changes are still shown here and have not been saved.", { exact: false }).waitFor()
  assert.equal(await table.getByRole("button", { name: "Edit question: What is the local question?" }).isDisabled(), true)
  assert.equal(storedGrid.columns[0]!.question, "What is the governing law?")
  page.once("dialog", dialog => dialog.dismiss())
  await table.getByRole("button", { name: "Close", exact: true }).click()
  assert.equal(await table.isVisible(), true)
  holdGridGet = true
  const countBeforeReload = gridGets
  const reloadRequest = page.waitForRequest(request => request.url().endsWith("/grid") && request.method() === "GET")
  page.once("dialog", dialog => dialog.accept())
  await table.getByRole("button", { name: "Reload saved table" }).click()
  await reloadRequest
  await table.getByRole("button", { name: "Reloading…" }).waitFor()
  assert.equal(await table.getByRole("button", { name: "Reloading…" }).isDisabled(), true)
  assert.equal(gridGets, countBeforeReload + 1)
  assert.ok(heldGridGet)
  await heldGridGet.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(storedGrid) })
  await table.getByRole("button", { name: "Edit question: What is the governing law?" }).waitFor()
  console.log("PASS stale changes remain visible until the user deliberately reloads")
  await page.keyboard.press("Escape")
  await table.waitFor({ state: "detached" })
  assert.equal(await closes(), 1)
  console.log("PASS the next Escape closes the review table once")

  // Questions are added, edited and removed in one dialog. A column without
  // answers goes at once; one with answers asks first and says what goes with
  // it, and Escape steps back to the question. Focus lands on the + button.
  storedGrid = structuredClone(grid)
  storedGrid.cells["Old note.docx::c1"]!.status = "reviewed"
  await page.goto(server.url.href)
  const focusedOn = (name: string) => page.waitForFunction(label => document.activeElement?.getAttribute("aria-label") === label || document.activeElement?.textContent === label, name)
  await page.getByRole("table", { name: "Document review" }).getByRole("button", { name: "Ways to add questions", exact: true }).click()
  await page.getByRole("button", { name: "Write a question" }).click()
  const adder = page.getByRole("dialog", { name: "Add question" })
  await page.waitForFunction(() => document.activeElement?.tagName === "TEXTAREA" && document.activeElement.closest("dialog")?.getAttribute("aria-label") === "Add question")
  assert.equal(await adder.getByRole("button", { name: "Add", exact: true }).isDisabled(), true)
  assert.equal(await adder.getByRole("button", { name: "Remove column" }).count(), 0)
  await page.keyboard.type("What is the notice period?")
  await page.keyboard.press("Enter")
  await adder.waitFor({ state: "detached" })
  const added = page.getByRole("button", { name: "Edit question: What is the notice period?" })
  await added.waitFor()
  await focusedOn("Ways to add questions")
  await added.click()
  await page.getByRole("dialog", { name: "Edit question" }).getByRole("button", { name: "Remove column" }).click()
  await added.waitFor({ state: "detached" })
  await focusedOn("Ways to add questions")
  const header = page.getByRole("button", { name: "Edit question: What is the governing law?" })
  await header.click()
  const editor = page.getByRole("dialog", { name: "Edit question" })
  await editor.getByRole("button", { name: "Remove column" }).click()
  const confirm = page.getByRole("dialog", { name: "Remove column" })
  await confirm.getByText("Remove “What is the governing law?” and its 3 answers, 1 of them reviewed, and 1 comment? This cannot be undone.").waitFor()
  await focusedOn("Keep column")
  await page.keyboard.press("Escape")
  await editor.waitFor()
  await focusedOn("Remove column")
  assert.equal(await header.count(), 1)
  await editor.getByRole("button", { name: "Remove column" }).click()
  await confirm.getByRole("button", { name: "Remove column" }).click()
  await header.waitFor({ state: "detached" })
  await focusedOn("Ways to add questions")
  for (let attempt = 0; storedGrid.columns.length && attempt < 100; attempt++) await new Promise(resolve => setTimeout(resolve, 20))
  assert.deepEqual(storedGrid.columns, [])
  assert.deepEqual(storedGrid.cells, {})
  console.log("PASS questions are added, edited and removed in a dialog; removing answers asks first")

  // Real component + mocked engine: release one extraction after changing its
  // live inputs. The unchanged baseline proves the same result is normally applied.
  for (const action of ["none", "review", "question", "delete", "comment"] as const) {
    storedGrid = structuredClone(grid)
    await page.goto(server.url.href)
    const row = page.getByRole("row", { name: /^Supply terms\.docx/ })
    const waiting = new Promise<import("@playwright/test").Route>(resolve => { onExtraction = resolve })
    await row.getByTitle("Re-run this cell").click()
    await page.waitForFunction(() => document.body.textContent?.includes("Filling 1"))
    const heldExtraction = await waiting
    if (action === "review") await row.getByTitle("Mark reviewed").click()
    if (action === "question") await editQuestion(page, "What is the governing law?", "A changed question")
    if (action === "delete") {
      await page.getByRole("button", { name: "Edit question: What is the governing law?" }).click()
      await page.getByRole("dialog", { name: "Edit question" }).getByRole("button", { name: "Remove column" }).click()
      await page.getByRole("dialog", { name: "Remove column" }).getByRole("button", { name: "Remove column" }).click()
    }
    if (action === "comment") {
      await row.getByRole("button", { name: /^The law of England and Wales\./ }).click()
      await page.getByRole("textbox", { name: "Add a comment" }).fill("Preserve this note")
      await page.getByRole("button", { name: "Add comment", exact: true }).click()
      await page.getByText("Preserve this note").waitFor()
      await page.keyboard.press("Escape")
    }
    await heldExtraction.fulfill({ status: 204 })
    await page.getByText("Filling 1", { exact: true }).waitFor({ state: "detached" })
    if (action === "none" || action === "comment") {
      await row.getByRole("button", { name: "New extracted answer" }).waitFor()
      if (action === "comment") await row.getByRole("button", { name: /^1 comment on/ }).waitFor()
    } else if (action === "delete") {
      assert.equal(await page.getByRole("button", { name: /New extracted answer/ }).count(), 0)
      assert.equal(storedGrid.columns.length, 0)
    } else {
      await row.getByRole("button", { name: /^The law of England and Wales\./ }).waitFor()
      assert.equal(await page.getByRole("button", { name: /New extracted answer/ }).count(), 0)
      if (action === "review") assert.equal(await row.getByTitle("Unlock").count(), 1)
    }
    console.log(`PASS extraction result with ${action} during the run`)
  }

  // One primary action at a time, moving on as the table is worked. Review next
  // opens the most urgent cell; the filter and the search narrow the view only.
  storedGrid = structuredClone(grid)
  for (const cell of Object.values(storedGrid.cells)) cell.questionHash = hashQuestion(grid.columns[0]!.question)
  const heldChecks = holdChecks
  holdChecks = false
  await page.goto(server.url.href)
  const toolbar = page.locator(".review-toolbar")
  const primary = toolbar.locator(".review-primary")
  const filter = toolbar.getByRole("button", { name: /^Filter/ })
  const progress = toolbar.getByRole("img", { name: /^Review progress/ })
  await filter.waitFor()
  await toolbar.getByText("0 of 3 reviewed", { exact: true }).waitFor()
  assert.equal(await progress.getAttribute("aria-label"), "Review progress, 3 cells in 3 documents and 1 question: 2 answered but not reviewed, 1 needs attention")
  assert.equal(await primary.count(), 1)
  assert.equal(await primary.textContent(), "Review next")
  assert.equal(await toolbar.getByRole("button", { name: "Review next" }).count(), 1, "the primary action is not repeated")
  assert.equal(await toolbar.getByRole("button", { name: /^Fill/ }).count(), 0, "nothing to fill")
  await primary.click()
  const nextDetail = page.locator(".picker-panel").filter({ hasText: "The law of Scotland." })
  await nextDetail.waitFor()
  await page.keyboard.press("Escape")
  await nextDetail.waitFor({ state: "detached" })
  assert.equal(await primary.evaluate(element => element === document.activeElement), true, "closing the cell returns to Review next")

  await filter.click()
  const menu = toolbar.getByRole("group", { name: "Show only" })
  await menu.getByRole("checkbox", { name: "Source changed, 1 cell", exact: true }).waitFor()
  assert.equal(await filter.getAttribute("aria-expanded"), "true")
  assert.equal(await menu.getByRole("checkbox", { name: /^No source/ }).count(), 0, "a state no cell has is left out")
  await page.keyboard.press("ArrowDown")
  assert.equal(await menu.evaluate(element => element.contains(document.activeElement)), true, "arrow keys move into the states")
  const closesBefore = await closes()
  await page.keyboard.press("Escape")
  await menu.waitFor({ state: "hidden" })
  assert.equal(await filter.evaluate(element => element === document.activeElement), true, "Escape returns focus to Filter")
  assert.equal(await closes(), closesBefore, "Escape closes only the popover")
  await filter.click()
  await menu.getByRole("checkbox", { name: /^Source changed/ }).check()
  await page.getByText("Showing 1 of 3 cells · Source changed", { exact: true }).waitFor()
  assert.equal(await toolbar.getByRole("button", { name: "Filter · 1 selected", exact: true }).count(), 1)
  assert.equal(await page.getByRole("row", { name: /^Supply terms\.docx/ }).count(), 0, "rows without a match are hidden")
  // The search opens from its icon. Escape on an empty field closes it; a search keeps it open.
  const searchToggle = toolbar.getByRole("button", { name: "Search the table", exact: true })
  const search = toolbar.getByRole("searchbox", { name: "Search the table" })
  // The harness has no styles, so the open popover sits in the flow; close it first.
  await page.keyboard.press("Escape")
  await menu.waitFor({ state: "hidden" })
  await searchToggle.click()
  assert.equal(await search.evaluate(element => element === document.activeElement), true, "the search field opens focused")
  await page.keyboard.press("Escape")
  await search.waitFor({ state: "detached" })
  assert.equal(await searchToggle.evaluate(element => element === document.activeElement), true, "Escape closes an empty search")
  assert.equal(await closes(), closesBefore, "Escape closes only the search")
  await searchToggle.click()
  await search.fill("supply")
  await page.getByText("No matching cells.", { exact: true }).waitFor()
  assert.equal(await page.getByText(/^Showing/).count(), 0, "with no match only the No matching cells line shows")
  await filter.focus()
  assert.equal(await search.inputValue(), "supply", "leaving a search keeps it open")
  await page.getByRole("button", { name: "Clear", exact: true }).click()
  await page.getByText("No matching cells.", { exact: true }).waitFor({ state: "detached" })
  assert.equal(await search.count(), 0, "an empty search closes once left")
  assert.equal(await page.locator("tbody tr").count(), 3)
  assert.equal(await filter.evaluate(element => element === document.activeElement), true, "Clear hands focus to Filter")

  // Reviewing in place under a filter keeps the cell where it was, quieter.
  await filter.click()
  await menu.getByRole("checkbox", { name: "Not reviewed, 3 cells", exact: true }).check()
  await page.getByText("Showing 3 of 3 cells · Not reviewed", { exact: true }).waitFor()
  await page.keyboard.press("Escape")
  await menu.waitFor({ state: "hidden" })
  for (const name of ["Supply terms", "Side letter", "Old note"])
    await page.getByRole("row", { name: new RegExp(`^${name}\\.docx`) }).getByTitle("Mark reviewed").click()
  await page.getByText("Showing 0 of 3 cells · Not reviewed", { exact: true }).waitFor()
  assert.equal(await page.locator("td.review-cell.is-dimmed").count(), 3)
  assert.equal(await page.getByRole("row", { name: /^Old note\.docx/ }).getByTitle("Unlock").evaluate(element => element === document.activeElement), true, "the lock keeps focus")
  // Several states widen the view: a cell with either one shows.
  await filter.click()
  await menu.getByRole("checkbox", { name: /^Source changed/ }).check()
  await page.getByText("Showing 1 of 3 cells · Source changed or Not reviewed", { exact: true }).waitFor()
  assert.equal(await toolbar.getByRole("button", { name: "Filter · 2 selected", exact: true }).count(), 1)
  await menu.getByRole("button", { name: "Clear", exact: true }).click()
  await page.getByText(/^Showing/).waitFor({ state: "detached" })
  assert.equal(await filter.evaluate(element => element === document.activeElement), true, "Clear hands focus to Filter")
  await page.keyboard.press("Escape")
  await menu.waitFor({ state: "hidden" })
  // A reviewed cell whose source changed still needs a look.
  await toolbar.getByText("2 of 3 reviewed", { exact: true }).waitFor()
  assert.equal(await progress.getAttribute("aria-label"), "Review progress, 3 cells in 3 documents and 1 question: 2 reviewed, 1 needs attention")
  assert.equal(await primary.count(), 1)
  assert.equal(await primary.getByRole("button", { name: "Export Excel", exact: true }).count(), 1)
  // The changed source still needs a look, so Review next stays, quietly.
  assert.equal(await toolbar.getByRole("button", { name: "Review next" }).count(), 1)
  holdChecks = heldChecks
  console.log("PASS one primary action per state, Review next, the multi-state filter, search and Clear")

  // Review next walks the queue inside one open detail: why each cell needs a
  // look, where it sits in the walk, and one next step. A changed source leads
  // with the lawyer's re-run, which stays on the cell and never marks it. A cell
  // reviewed before its question changed is never re-locked or refreshed from
  // the walk; the end hands back to the toolbar's next step, and the search
  // stays as it was.
  storedGrid = structuredClone(grid)
  for (const cell of Object.values(storedGrid.cells)) cell.questionHash = hashQuestion(grid.columns[0]!.question)
  storedGrid.cells["Old note.docx::c1"] = { ...storedGrid.cells["Old note.docx::c1"]!, status: "reviewed", questionHash: "unused" }
  const walkHeldChecks = holdChecks
  holdChecks = false
  await page.goto(server.url.href)
  await toolbar.getByRole("img", { name: /2 need attention$/ }).waitFor({ state: "attached" })
  assert.equal(await primary.textContent(), "Review next")
  const walkDetail = detailModal.locator(".picker-panel")
  const walkAnswer = (text: string) => walkDetail.locator(".review-detail-answer", { hasText: text }).waitFor()
  const position = walkDetail.locator(".review-walk-position")
  const why = walkDetail.locator(".review-walk-why")
  const walkBar = walkDetail.locator(".review-detail-bar")
  const markNext = walkBar.getByRole("button", { name: "Mark reviewed · Next", exact: true })
  const rerun = walkBar.getByRole("button", { name: "Re-run this cell", exact: true })
  const skip = walkBar.getByRole("button", { name: "Skip", exact: true })
  const previous = walkBar.getByRole("button", { name: "Previous", exact: true })
  const gridPut = () => page.waitForResponse(response => response.url().endsWith("/grid") && response.request().method() === "PUT")
  // Counts the detail's native opens and closes from here on.
  const dialogCalls = () => page.evaluate(() => (window as unknown as { dialogCalls?: string[] }).dialogCalls ?? [])
  await primary.click()
  await walkAnswer("The law of Scotland.")
  assert.equal(await position.textContent(), "1 of 3 to review")
  assert.match(await walkDetail.locator(".viewer-title").innerText(), /Side letter\.docx$/, "the header names the cell's document")
  assert.match(await why.innerText(), /^Source changed\s+The document no longer contains the cited excerpt\./)
  assert.equal(await walkDetail.locator(".review-source-check").count(), 0, "the reason line, not the source, explains a changed source")
  assert.equal(await rerun.evaluate(element => element === document.activeElement && element.classList.contains("primary")), true, "a changed source opens on its re-run")
  assert.equal(await markNext.evaluate(element => element.classList.contains("primary")), false, "Mark reviewed · Next stays as the other step")
  assert.equal(await previous.isDisabled(), true)
  assert.equal(await walkDetail.locator("button.primary").count(), 1, "one primary in review mode")
  assert.equal(await walkDetail.getByRole("button", { name: "Add comment", exact: true }).evaluate(element => element.classList.contains("primary")), false, "Add comment is secondary in review mode")
  await detailModal.evaluate(element => {
    element.setAttribute("data-walk", "same")
    const calls: string[] = (window as unknown as { dialogCalls: string[] }).dialogCalls = []
    for (const name of ["showModal", "close"] as const) {
      const native = HTMLDialogElement.prototype[name]
      HTMLDialogElement.prototype[name] = function (this: HTMLDialogElement) {
        if (this.classList.contains("review-detail-modal")) calls.push(name)
        return native.call(this)
      }
    }
  })
  // The re-run keeps focus while it runs; the old answer cannot be marked meanwhile.
  const heldRerun = new Promise<import("@playwright/test").Route>(resolve => { onExtraction = resolve })
  await page.keyboard.press("Enter")
  const rerunning = await heldRerun
  assert.equal(await walkBar.getByRole("button", { name: "Re-running…", exact: true }).evaluate(element => element === document.activeElement), true, "the re-run keeps focus")
  assert.equal(await markNext.isDisabled(), true, "the answer being replaced cannot be marked reviewed")
  await rerunning.fulfill({ status: 204 })
  onExtraction = undefined
  await walkAnswer("New extracted answer")
  assert.equal(await position.textContent(), "1 of 3 to review", "the re-run stays on its cell")
  await walkBar.locator("button.primary", { hasText: "Mark reviewed · Next" }).waitFor()
  assert.equal(await markNext.evaluate(element => element === document.activeElement), true, "the new answer, still unreviewed, leads with Mark reviewed · Next")
  await page.keyboard.press("Enter")
  await walkAnswer("English law.")
  for (let attempt = 0; storedGrid.cells["Side letter.docx::c1"]!.status !== "reviewed" && attempt < 250; attempt++) await new Promise(resolve => setTimeout(resolve, 20))
  assert.equal(storedGrid.cells["Side letter.docx::c1"]!.status, "reviewed", "Mark reviewed · Next locks the cell")
  assert.equal(await position.textContent(), "2 of 3 to review")
  assert.match(await why.innerText(), /Question changed\s+The question was edited after this cell was reviewed\. Unlock to refresh\.$/)
  assert.equal(await markNext.count(), 0, "a reviewed cell whose question changed offers no Mark reviewed")
  assert.equal(await rerun.count(), 0, "a reviewed cell is never re-run from the walk")
  assert.equal(await skip.count(), 1, "Skip moves on in every state")
  const unlock = walkBar.getByRole("button", { name: "Unlock", exact: true })
  assert.equal(await unlock.evaluate(element => element === document.activeElement && element.classList.contains("primary")), true, "its named remedy, Unlock, leads and keeps focus")
  assert.equal(await skip.evaluate(element => element.classList.contains("primary")), false)
  assert.equal(await focusInDetail(), true, "focus stays in the detail as the walk moves")
  assert.equal(await detailModal.getAttribute("data-walk"), "same", "the detail is not remounted between cells")
  assert.deepEqual(await dialogCalls(), [], "the detail does not close or reopen between cells")

  await previous.click()
  await walkAnswer("New extracted answer")
  assert.equal(await position.textContent(), "1 of 3 to review")
  assert.equal(await markNext.count(), 0, "a cell reviewed in the walk offers no Mark reviewed")
  assert.equal(await walkBar.locator("button.primary").count(), 0, "a reviewed cell with nothing flagged has no primary")
  assert.equal(await walkBar.locator(".review-detail-status").innerText(), "Reviewed")
  assert.equal(await walkBar.getByRole("button", { name: "Unlock", exact: true }).evaluate(element => element.classList.contains("ghost")), true, "its Unlock is quiet")
  assert.equal(await skip.evaluate(element => element === document.activeElement), true, "a move that disables its button hands focus to Skip")
  await skip.click()
  await walkAnswer("English law.")
  await skip.click()
  await walkAnswer("The law of England and Wales.")
  assert.equal(await position.textContent(), "3 of 3 to review")
  assert.equal(await why.count(), 0, "a cell that is only unreviewed needs no reason")
  const tableCloses = await closes()
  await skip.click()
  await detailModal.waitFor({ state: "detached" })
  await toolbar.getByText("End of the review list", { exact: true }).waitFor()
  assert.equal(await primary.textContent(), "Review next")
  assert.equal(await primary.evaluate(element => element === document.activeElement), true, "the end returns focus to the toolbar's next step")
  assert.deepEqual(await dialogCalls(), ["close"], "the walk's end closes the detail once")
  assert.equal(await closes(), tableCloses, "ending the walk keeps the table open")
  assert.equal(storedGrid.cells["Supply terms.docx::c1"]!.status, "filled", "Skip leaves the cell open")
  assert.equal(storedGrid.cells["Old note.docx::c1"]!.status, "reviewed")
  assert.equal(storedGrid.cells["Old note.docx::c1"]!.questionHash, "unused", "the walk never refreshes a reviewed cell")

  // Ending on Mark reviewed leaves nothing to review, so focus lands on Export.
  // The re-run side letter now needs no look, so the walk starts at the old note.
  await primary.click()
  await walkAnswer("English law.")
  assert.equal(await toolbar.getByText("End of the review list").count(), 0, "a new walk clears the last one's word")
  await skip.click()
  await walkAnswer("The law of England and Wales.")
  const lastMark = gridPut()
  await markNext.click()
  await lastMark
  await detailModal.waitFor({ state: "detached" })
  await toolbar.getByText("End of the review list", { exact: true }).waitFor()
  assert.equal(storedGrid.cells["Supply terms.docx::c1"]!.status, "reviewed")
  assert.equal(await primary.getByRole("button", { name: "Export Excel", exact: true }).evaluate(element => element === document.activeElement), true, "with every cell reviewed, the end lands on Export")

  // The walk keeps to the cells the search shows; Escape leaves it as before.
  await searchToggle.click()
  await search.fill("note")
  await page.getByText("Showing 1 of 3 cells · “note”", { exact: true }).waitFor()
  const quietNext = toolbar.getByRole("button", { name: "Review next", exact: true })
  await quietNext.click()
  await walkAnswer("English law.")
  assert.equal(await position.textContent(), "1 of 1 to review", "the walk keeps to the shown cells")
  await page.keyboard.press("Escape")
  await detailModal.waitFor({ state: "detached" })
  assert.equal(await quietNext.evaluate(element => element === document.activeElement), true, "Escape returns focus to where the walk began")
  assert.equal(await toolbar.getByText("End of the review list").count(), 0, "closing early is not the end of the walk")
  assert.equal(await search.inputValue(), "note", "the search stays through the walk")
  await page.getByText("Showing 1 of 3 cells · “note”", { exact: true }).waitFor()
  await search.fill("")
  holdChecks = walkHeldChecks
  console.log("PASS Review next walks the queue in one detail, leads a changed source with its re-run, never re-locks a changed question and hands back to the toolbar")

  // A reviewed Fill compares each filled question's cited answers once, in a
  // hidden review-compare session that is deleted afterwards. The result is
  // stored on the column only; a fresh one marks the differing cell, its header
  // sums it up and its detail says why. The side letter's source is verified
  // here, so Differs is its marker.
  const comparedGrid = () => {
    const next = structuredClone(grid)
    for (const cell of Object.values(next.cells)) cell.questionHash = hashQuestion(grid.columns[0]!.question)
    return next
  }
  const replyFor = (groups: Array<[string[], string]>, summary: string) =>
    JSON.stringify({ groups: groups.map(([documents, reason]) => ({ documents, reason })), summary })
  const threeWay = replyFor([[["Supply terms.docx", "Old note.docx"], "English law governs."], [["Side letter.docx"], "Scots law governs, not English law."]],
    "Two documents choose English law; the side letter chooses Scots law.")
  const heldCompare = async () => {
    for (let attempt = 0; !compareWaits.length && attempt < 250; attempt++) await new Promise(resolve => setTimeout(resolve, 20))
    assert.ok(compareWaits.length, "a compare session should be waiting")
    return compareWaits.shift()!
  }
  const compareHeldChecks = holdChecks
  holdChecks = false
  sourceChecks["Side letter.docx::c1"] = "verified"
  storedGrid = comparedGrid()
  onExtraction = route => void route.fulfill({ status: 204 })
  holdCompare = true
  compareReply = threeWay
  await page.goto(server.url.href + "?pool")
  const head = page.locator("thead th").filter({ hasText: "What is the governing law?" })
  const sideRow = page.getByRole("row", { name: /^Side letter\.docx/ })
  const differsMarker = sideRow.locator(".review-marker", { hasText: "Differs" })
  await head.getByRole("button", { name: "Compare with AI", exact: true }).waitFor()
  assert.equal(compareSessions.length, 0, "opening the table compares nothing")
  assert.equal(await primary.textContent(), "Fill 4 cells")
  await primary.click()
  const fillDialog = page.getByRole("dialog", { name: "Fill 4 cells?" })
  await fillDialog.getByText("After filling, AI compares the cited answers across documents: one AI call for each question with two or more cited answers.", { exact: false }).waitFor()
  const fillHeld: import("@playwright/test").Route[] = []
  onExtraction = route => void fillHeld.push(route)
  await fillDialog.getByRole("button", { name: "Start fill" }).click()
  // While cells fill, the status counts them, answers already in are the one next step, and no compare is offered.
  await toolbar.getByText(/^Filling \d/).waitFor()
  assert.equal(await primary.textContent(), "Review next")
  assert.equal(await toolbar.getByRole("button", { name: "Review next" }).count(), 1, "one Review next while cells fill")
  assert.equal(await head.getByRole("button", { name: /^Compare/ }).count(), 0, "no compare is offered while cells fill")
  onExtraction = route => void route.fulfill({ status: 204 })
  await Promise.all(fillHeld.splice(0).map(route => route.fulfill({ status: 204 })))
  await toolbar.getByText("Comparing 1 question", { exact: true }).waitFor()
  assert.equal(await toolbar.getByText(/^Filling \d/).count(), 0, "comparing is counted apart from filling")
  await head.getByText("Comparing with AI…", { exact: true }).waitFor()
  const firstCompare = await heldCompare()
  assert.equal(compareSessions.length, 1, "one compare for the one filled question")
  assert.equal(await primary.count(), 1)
  await firstCompare.fulfill({ status: 204 })
  const summary = head.getByRole("button", { name: /^AI comparison: 1 of \d+ differs/ })
  await summary.waitFor()
  await toolbar.getByText(/Comparing/).waitFor({ state: "detached" })
  await differsMarker.waitFor()
  assert.equal(await page.getByRole("row", { name: /^Supply terms\.docx/ }).locator(".review-marker", { hasText: "Differs" }).count(), 0, "a consistent answer gets no marker")
  await page.waitForTimeout(100)
  assert.deepEqual(compareDeleted, compareSessions, "the compare session is deleted afterwards")
  const stored = storedGrid.columns[0]!.comparison
  assert.equal(stored?.cells["Side letter.docx"]?.label, "differs")
  assert.equal(stored?.questionHash, hashQuestion(grid.columns[0]!.question))
  assert.equal(JSON.stringify(storedGrid.cells).includes("differs"), false, "the comparison is stored on the column, never in a cell")
  assert.equal(await primary.count(), 1)
  assert.equal(await primary.textContent(), "Review next")

  // The header summary opens the AI summary and the differing cells.
  await summary.click()
  const comparison = head.getByRole("group", { name: "AI comparison" })
  await comparison.getByText("Two documents choose English law; the side letter chooses Scots law.").waitFor()
  // Scrolling the table keeps the popover open; it follows its summary.
  await summary.evaluate((element) => { const view = element.closest(".review-scroll")!; view.scrollTop += 8; view.dispatchEvent(new Event("scroll")) })
  await page.waitForTimeout(50)
  assert.equal(await comparison.isVisible(), true, "scrolling the table keeps the comparison open")
  await page.keyboard.press("ArrowDown")
  assert.equal(await comparison.getByRole("button", { name: /^Side letter\.docx/ }).evaluate(element => element === document.activeElement), true, "arrow keys move into the differing cells")
  const compareCloses = await closes()
  await page.keyboard.press("Escape")
  await comparison.waitFor({ state: "hidden" })
  assert.equal(await summary.evaluate(element => element === document.activeElement), true, "Escape returns focus to the summary")
  assert.equal(await closes(), compareCloses, "Escape closes only the popover")
  await summary.click()
  await comparison.getByRole("button", { name: /^Side letter\.docx/ }).click()
  const compareDetail = detailModal.locator(".picker-panel")
  await compareDetail.getByText("AI comparison: Differs — Scots law governs, not English law.", { exact: true }).waitFor()
  await page.keyboard.press("Escape")
  await detailModal.waitFor({ state: "detached" })
  assert.equal(await summary.evaluate(element => element === document.activeElement), true, "closing the cell returns to the summary")
  await page.getByRole("row", { name: /^Supply terms\.docx/ }).getByRole("button", { name: /^The law of England and Wales\./ }).click()
  await compareDetail.locator(".review-detail-answer").waitFor()
  assert.equal(await compareDetail.getByText(/^AI comparison/).count(), 0, "a consistent cell's detail says nothing about the comparison")
  await page.keyboard.press("Escape")
  await detailModal.waitFor({ state: "detached" })
  console.log("PASS a reviewed Fill compares the question once, marks the differing cell and deletes the session")

  // A changed answer puts the comparison out of date: no Differs marker, and a
  // quiet Compare again. A failed compare stores nothing, stays quiet and offers Try again.
  await page.getByRole("row", { name: /^Old note\.docx/ }).getByTitle("Re-run this cell").click()
  await page.getByRole("row", { name: /^Old note\.docx/ }).getByRole("button", { name: "New extracted answer" }).waitFor()
  await head.getByText("AI comparison out of date.", { exact: false }).waitFor()
  assert.equal(await differsMarker.count(), 0, "an out-of-date comparison marks nothing")
  assert.equal(compareSessions.length, 1, "a single cell's Refresh does not compare")
  compareReply = "not a comparison"
  holdCompare = false
  await head.getByRole("button", { name: "Compare again", exact: true }).click()
  await head.getByText("Could not compare.", { exact: false }).waitFor()
  assert.equal(await page.getByRole("alert").count(), 0, "a failed compare raises no alert")
  assert.equal(storedGrid.columns[0]!.comparison?.comparedAt, stored?.comparedAt, "a failed compare stores nothing")
  compareReply = replyFor([[["Supply terms.docx"], "English law governs."], [["Side letter.docx"], "Scots law governs."]], "The two cited answers choose different laws.")
  await head.getByRole("button", { name: "Try again", exact: true }).click()
  await head.getByRole("button", { name: /^AI comparison: 2 of \d+ differ\b/ }).waitFor()
  await differsMarker.waitFor()
  await page.waitForTimeout(100)
  assert.equal(compareSessions.length, 3)
  assert.deepEqual(compareDeleted, compareSessions, "every compare session is deleted, failed ones included")
  console.log("PASS a changed answer puts the comparison out of date; Compare again runs once and fails quietly, then Try again")

  // An answer that changes while its question is compared drops the result,
  // leaving the column uncompared. Filling and comparing are counted apart.
  storedGrid = comparedGrid()
  holdCompare = true
  compareReply = threeWay
  await page.goto(server.url.href + "?pool")
  await head.getByRole("button", { name: "Compare with AI", exact: true }).waitFor()
  await primary.click()
  await fillDialog.getByRole("button", { name: "Start fill" }).click()
  await toolbar.getByText("Comparing 1 question", { exact: true }).waitFor()
  const midCompare = await heldCompare()
  const supplyRow = page.getByRole("row", { name: /^Supply terms\.docx/ })
  const releaseExtraction = new Promise<import("@playwright/test").Route>(resolve => { onExtraction = resolve })
  await supplyRow.getByTitle("Re-run this cell").click()
  const heldRefresh = await releaseExtraction
  await toolbar.getByText("Filling 1", { exact: true }).waitFor()
  await toolbar.getByText("Comparing 1 question", { exact: true }).waitFor()
  assert.equal(await primary.count(), 1)
  await heldRefresh.fulfill({ status: 204 })
  await supplyRow.getByRole("button", { name: "New extracted answer" }).waitFor()
  await midCompare.fulfill({ status: 204 })
  await toolbar.getByText(/Comparing/).waitFor({ state: "detached" })
  await page.waitForTimeout(150)
  assert.equal(storedGrid.columns[0]!.comparison, undefined, "a result for answers that have since changed is dropped")
  assert.equal(await differsMarker.count(), 0)
  await head.getByRole("button", { name: "Compare with AI", exact: true }).waitFor()
  assert.deepEqual(compareDeleted, compareSessions)
  holdCompare = false
  onExtraction = undefined
  sourceChecks["Side letter.docx::c1"] = "not_found"
  holdChecks = compareHeldChecks
  console.log("PASS a comparison is dropped when an answer changes while it runs")

  // A column marked not to compare (as a starter marks a question whose answers
  // identify each document) is filled but never compared: its header offers no
  // compare even with two cited answers, the Fill preview marks it Not compared,
  // and the Fill opens a compare session only for the other question.
  const partiesQuestion = "Who are the parties?"
  const partiesCell = (doc: string) => ({
    answer: `The parties to ${doc}.`, status: "filled" as const, questionHash: hashQuestion(partiesQuestion),
    citation: { documentName: doc, docPath: doc, section: "1", excerpt: `Parties to ${doc}.`, charStart: 0, charEnd: 10 },
  })
  const flaggedHeldChecks = holdChecks
  holdChecks = false
  storedGrid = comparedGrid()
  storedGrid.columns.push({ id: "c2", question: partiesQuestion, compare: false })
  storedGrid.cells["Supply terms.docx::c2"] = partiesCell("Supply terms.docx")
  storedGrid.cells["Side letter.docx::c2"] = partiesCell("Side letter.docx")
  onExtraction = route => void route.fulfill({ status: 204 })
  const sessionsBefore = compareSessions.length
  await page.goto(server.url.href + "?pool")
  const partiesHead = page.locator("thead th").filter({ hasText: partiesQuestion })
  await head.getByRole("button", { name: "Compare with AI", exact: true }).waitFor()
  assert.equal(await partiesHead.locator(".review-compare-note, .review-compare-trigger").count(), 0, "a column marked not to compare offers no compare")
  assert.equal(await primary.textContent(), "Fill 9 cells")
  await primary.click()
  const partiesDialog = page.getByRole("dialog", { name: "Fill 9 cells?" })
  await partiesDialog.getByText(partiesQuestion).first().waitFor()
  const partiesItems = partiesDialog.getByRole("listitem").filter({ hasText: partiesQuestion })
  assert.deepEqual(await partiesDialog.locator(".fill-preview-note").allTextContents(), Array(await partiesItems.count()).fill("Not compared"), "the Fill preview marks the question it will not compare, and only that one")
  await partiesDialog.getByText("After filling, AI compares the cited answers across documents: one AI call for each question with two or more cited answers.", { exact: false }).waitFor()
  await partiesDialog.getByText("It skips questions marked Not compared.", { exact: false }).waitFor()
  await partiesDialog.getByRole("button", { name: "Start fill" }).click()
  await head.getByRole("button", { name: /^AI comparison: 1 of \d+ differs/ }).waitFor()
  await toolbar.getByText(/Comparing|^Filling \d/).waitFor({ state: "detached" })
  await page.waitForTimeout(100)
  assert.equal(compareSessions.length, sessionsBefore + 1, "only the compared question opens a compare session")
  assert.equal(await partiesHead.locator(".review-compare-note, .review-compare-trigger").count(), 0, "after Fill the column marked not to compare still shows no comparison")
  assert.equal(storedGrid.columns[1]!.comparison, undefined, "nothing is compared or stored for it")
  assert.equal(storedGrid.columns[1]!.compare, false, "Fill keeps the column's flag")
  onExtraction = undefined
  holdChecks = flaggedHeldChecks
  console.log("PASS a column marked not to compare is filled, marked Not compared in the Fill preview and never compared")

  storedGrid = structuredClone(grid)
  extractionWaits.length = 0
  extractionStarts = 0
  onExtraction = undefined
  await page.goto(server.url.href + "?pool")
  await page.getByRole("button", { name: "Fill 7 cells", exact: true }).click()
  await page.getByRole("dialog", { name: "Fill 7 cells?" }).getByRole("button", { name: "Start fill" }).click()
  for (let attempt = 0; extractionWaits.length < 5 && attempt < 250; attempt++) await new Promise(resolve => setTimeout(resolve, 20))
  assert.equal(extractionWaits.length, 5, "five extraction jobs should be in flight")
  assert.equal(extractionStarts, 5, "the extraction pool should hold the remaining work")
  storedGrid = { ...storedGrid, version: (storedGrid.version ?? 0) + 1 }
  await editQuestion(page, "What is the governing law?", "Conflicting question")
  await page.getByText("Your changes are still shown here and have not been saved.", { exact: false }).waitFor()
  page.once("dialog", dialog => dialog.accept())
  await page.getByRole("button", { name: "Reload saved table" }).click()
  await page.getByRole("button", { name: "Edit question: What is the governing law?" }).waitFor()
  await Promise.all(extractionWaits.map(route => route.fulfill({ status: 204 })))
  await page.getByText(/^Filling \d/).waitFor({ state: "detached" })
  assert.equal(extractionStarts, 5, "queued targets from the old table must not start after reload")
  assert.equal(await page.getByRole("button", { name: /New extracted answer/ }).count(), 0)
  console.log("PASS reloading cancels queued extraction targets and ignores in-flight results")

  // A pending settings read is part of a comment submission. Reload must not
  // let that older submission write into the newly loaded table.
  storedGrid = structuredClone(grid)
  await page.goto(server.url.href)
  await page.getByRole("button", { name: /^The law of England and Wales\./ }).click()
  holdPreferences = true
  const waitingPreferences = page.waitForRequest(request => new URL(request.url()).pathname === "/preferences")
  await page.getByRole("textbox", { name: "Add a comment" }).fill("Old pending comment")
  await page.getByRole("button", { name: "Add comment", exact: true }).click()
  await waitingPreferences
  await page.keyboard.press("Escape")
  page.once("dialog", dialog => dialog.dismiss())
  await page.getByRole("dialog", { name: "Review table — Supplier review" }).getByRole("button", { name: "Close", exact: true }).click()
  assert.equal(await page.getByRole("dialog", { name: "Review table — Supplier review" }).isVisible(), true)
  storedGrid = { ...storedGrid, version: (storedGrid.version ?? 0) + 1 }
  await editQuestion(page, "What is the governing law?", "Conflicting question")
  await page.getByText("Your changes are still shown here and have not been saved.", { exact: false }).waitFor()
  page.once("dialog", dialog => dialog.accept())
  await page.getByRole("button", { name: "Reload saved table" }).click()
  await page.getByRole("button", { name: "Edit question: What is the governing law?" }).waitFor()
  assert.ok(heldPreferences)
  const countBeforeRelease = putVersions.length
  await heldPreferences.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ attorney: "Amina Shah" }) })
  holdPreferences = false
  await page.waitForTimeout(100)
  assert.equal(putVersions.length, countBeforeRelease)
  assert.equal(storedGrid.cells["Supply terms.docx::c1"]!.comments, undefined)
  console.log("PASS a pending comment warns before close and cannot write after reload")

  storedGrid = structuredClone(grid)
  await page.goto(new URL("/nav", server.url).href)
  await page.getByRole("link", { name: "Open review" }).click()
  const navTable = page.getByRole("dialog", { name: "Review table — Supplier review" })
  await navTable.getByTitle("Mark reviewed").waitFor()
  holdPut = true
  const navPut = page.waitForRequest(request => request.url().endsWith("/grid") && request.method() === "PUT")
  await navTable.getByTitle("Mark reviewed").click()
  await navPut
  let prompts = 0
  page.once("dialog", dialog => { prompts++; return dialog.dismiss() })
  await page.goBack()
  await navTable.getByTitle("Unlock").waitFor()
  assert.equal(prompts, 1, "Back should ask once before discarding unsaved changes")
  assert.match(page.url(), /view=review/)
  assert.equal(await page.evaluate(() => { const event = new Event("beforeunload", { cancelable: true }); window.dispatchEvent(event); return event.defaultPrevented }), true)
  page.once("dialog", dialog => { prompts++; return dialog.accept() })
  await page.goBack()
  await navTable.waitFor({ state: "detached" })
  assert.equal(prompts, 2, "confirming Back should leave once")
  assert.ok(heldPut)
  await save(heldPut)
  await page.getByRole("link", { name: "Open review" }).click()
  await navTable.getByTitle("Unlock").waitFor()
  holdPut = true
  const secondNavPut = page.waitForRequest(request => request.url().endsWith("/grid") && request.method() === "PUT")
  await navTable.getByTitle("Unlock").click()
  await secondNavPut
  let closePrompts = 0
  page.once("dialog", dialog => { closePrompts++; return dialog.accept() })
  await navTable.getByRole("button", { name: "Close", exact: true }).click()
  await navTable.waitFor({ state: "detached" })
  assert.equal(closePrompts, 1, "closing through the router must not double-prompt")
  assert.ok(heldPut)
  await save(heldPut)
  await page.getByRole("link", { name: "Other page" }).click()
  assert.match(page.url(), /\/outside$/)
  await page.goBack()
  assert.match(page.url(), /\/nav$/)
  await page.goto(new URL("/nav?view=review", server.url).href)
  await navTable.getByTitle("Mark reviewed").waitFor()
  await navTable.getByRole("button", { name: "Close", exact: true }).click()
  await navTable.waitFor({ state: "detached" })
  console.log("PASS Back and router close guard unsaved work without duplicate prompts; unload remains guarded")

  await page.getByRole("link", { name: "Open review" }).click()
  await navTable.getByRole("button", { name: "Edit question: What is the governing law?" }).waitFor()
  storedGrid = { ...storedGrid, version: (storedGrid.version ?? 0) + 1 }
  await editQuestion(navTable, "What is the governing law?", "Conflicting question")
  await navTable.getByText("Your changes are still shown here and have not been saved.", { exact: false }).waitFor()
  holdGridGet = true
  const heldReload = page.waitForRequest(request => request.url().endsWith("/grid") && request.method() === "GET")
  page.once("dialog", dialog => dialog.accept())
  await navTable.getByRole("button", { name: "Reload saved table" }).click()
  await heldReload
  await navTable.getByRole("button", { name: "Reloading…" }).waitFor()
  page.once("dialog", dialog => dialog.accept())
  await navTable.getByRole("button", { name: "Close", exact: true }).click()
  await navTable.waitFor({ state: "detached" })
  assert.ok(heldGridGet)
  await heldGridGet.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ ...storedGrid, columns: [{ id: "c1", question: "Obsolete response" }] }) })
  await page.getByRole("link", { name: "Open review" }).click()
  await navTable.getByRole("button", { name: "Edit question: What is the governing law?" }).waitFor()
  assert.equal(await navTable.getByRole("button", { name: "Edit question: Obsolete response" }).count(), 0)
  console.log("PASS an in-flight reload cannot affect a later opening of the table")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
