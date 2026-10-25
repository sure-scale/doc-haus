// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/document-comments.browser.ts
// Mounts the real DocumentViewer on the Docxodus WASM runtime with the synthetic counterparty markup; only ingest HTTP is mocked.
// While the Tracked changes panel lists the comments, the page shows no cards, only a comment button by each commented
// paragraph, and a button or the commented text opens and focuses its thread in the panel. With the panel closed the
// comments beside the page are Word-style cards: each author's initials and colour match the panel's, a resolved thread
// says so, a screened comment is flagged, a card and its text light up together, and on a narrow page each card sits
// under its paragraph (once) and goes back when the page widens, with citations still found across it.
// Set PLAYWRIGHT_CHROMIUM_PATH=<chromium> when Playwright's own browser is missing; SCREENSHOT_DIR=<dir> saves one screenshot per state.
import assert from "node:assert/strict"
import { existsSync, statSync } from "node:fs"
import { mkdir } from "node:fs/promises"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { strToU8, zipSync } from "fflate"
import { COUNTERPARTY_MARKUP, counterpartyMarkupDocx } from "../../../dochaus/test/support/counterparty-markup-fixture"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const wasmDir = path.resolve(process.env.DOCXODUS_WASM_DIR ?? path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm"))

const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const NAME = COUNTERPARTY_MARKUP.name
const bytes = counterpartyMarkupDocx()
// The lawyer is Morgan here, so Morgan's comments take the brand colour in both places.
const LAWYER = "Morgan Okafor"
// The file's comments as ingest lists them: Word's own initials (Morgan's are
// three letters, so they can only come from here) and a screening finding on
// the comment addressed to an AI assistant.
const markup = {
  name: NAME,
  version: { id: "v1", number: 1, sha256: "synthetic" },
  current: true,
  revisions: [],
  comments: COUNTERPARTY_MARKUP.comments.map((comment) => ({
    id: comment.id, author: comment.author, initials: comment.author === LAWYER ? "MOk" : "AL",
    date: COUNTERPARTY_MARKUP.dates[comment.author], text: comment.text, parent: "parent" in comment ? comment.parent : null,
    resolved: "resolved" in comment ? comment.resolved : null, quote: "", paragraph: null,
    findings: comment.text === COUNTERPARTY_MARKUP.plantedComment ? [{ field: "text", rule: "instruction", detail: "Addressed to an AI assistant." }] : [],
  })),
  authors: COUNTERPARTY_MARKUP.reviewers.map((name) => ({ name, revisions: 0, comments: 2 })),
}

// A small synthetic Word file: body paragraphs (WordprocessingML) and comments, one paragraph each.
type Crafted = { id: number; author: string; text: string }
const DATE = "2026-10-20T09:00:00Z"
const PKG = "http://schemas.openxmlformats.org/package/2006"
const RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const WML = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const escaped = (text: string) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;")
const run = (text: string) => `<w:r><w:t xml:space="preserve">${escaped(text)}</w:t></w:r>`
const reference = (id: number) => `<w:r><w:commentReference w:id="${id}"/></w:r>`
const commented = (id: number, text: string) => `<w:commentRangeStart w:id="${id}"/>${run(text)}<w:commentRangeEnd w:id="${id}"/>${reference(id)}`
const para = (inner: string) => `<w:p>${inner}</w:p>`
const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(inner)}</w:tc>`
function wordFile(body: string[], comments: Crafted[]) {
  const part = (xml: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)
  return zipSync({
    "[Content_Types].xml": part(`<Types xmlns="${PKG}/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${CT}.comments+xml"/></Types>`),
    "_rels/.rels": part(`<Relationships xmlns="${PKG}/relationships"><Relationship Id="rId1" Type="${RELS}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": part(`<Relationships xmlns="${PKG}/relationships"><Relationship Id="rId1" Type="${RELS}/comments" Target="comments.xml"/></Relationships>`),
    "word/document.xml": part(`<w:document ${WML}><w:body>${body.join("")}<w:sectPr/></w:body></w:document>`),
    "word/comments.xml": part(`<w:comments ${WML}>${comments.map(({ id, author, text }) => `<w:comment w:id="${id}" w:author="${escaped(author)}" w:date="${DATE}"><w:p>${run(text)}</w:p></w:comment>`).join("")}</w:comments>`),
  })
}
// What ingest lists for such a file: each comment as screened (hidden characters stripped, findings by field).
type Listed = Crafted & { findings?: { field: string; rule: string; detail: string }[] }
const listing = (comments: Listed[]) => ({
  name: NAME, version: { id: "v1", number: 1, sha256: "synthetic" }, current: true, revisions: [],
  comments: comments.map(({ id, author, text, findings = [] }) => ({ id, author, initials: null, date: DATE, text, parent: null, resolved: null, quote: "", paragraph: null, findings })),
  authors: [...new Set(comments.map((comment) => comment.author))].map((name) => ({ name, revisions: 0, comments: comments.filter((comment) => comment.author === name).length })),
})
const REVIEWERS = ["Avery Lindqvist", "Morgan Okafor", "Sam Rivera", "Jo Park", "Lee Chan"]

const entry = path.join((await import("node:os")).tmpdir(), `dochaus-comments-${process.pid}.tsx`)
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentViewer from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentViewer.tsx"))}
const citation = location.hash.length > 1 ? JSON.parse(decodeURIComponent(location.hash.slice(1))) : undefined
createRoot(document.getElementById("root")).render(createElement(DocumentViewer, { matterId: "matter-comments", name: ${JSON.stringify(NAME)}, citation, onClose: () => {} }))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://viewer.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = ["design-system", "primitives", "design-documents"].map(name => path.resolve(import.meta.dir, `../src/styles/${name}.css`))
const styles = (await Promise.all(css.map(file => Bun.file(file).text()))).join("\n") + "\n" + (await Bun.file(path.resolve(import.meta.dir, "../src/styles.css")).text())
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => {
  const { pathname } = new URL(request.url)
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname === "/app.css") return new Response(styles, { headers: { "content-type": "text/css" } })
  if (pathname.startsWith("/wasm/")) {
    const file = path.join(wasmDir, decodeURIComponent(pathname.slice("/wasm/".length)))
    return file.startsWith(wasmDir + path.sep) && existsSync(file) && statSync(file).isFile() ? new Response(Bun.file(file)) : new Response("Not found", { status: 404 })
  }
  return new Response(`<html><meta charset="utf-8"><link rel="stylesheet" href="/app.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const screenshots = process.env.SCREENSHOT_DIR
if (screenshots) await mkdir(screenshots, { recursive: true })

// Without the file's markup the panel stays closed, so the cards show from the start.
async function open(viewport: { width: number; height: number }, citation?: { section: string; excerpt: string }, withMarkup = true, file: { bytes: Uint8Array; markup: unknown } = { bytes, markup }) {
  const page = await browser.newPage({ viewport })
  pages.push(page)
  page.setDefaultTimeout(60_000)
  const errors: string[] = []
  page.on("pageerror", error => errors.push(error.message))
  await page.route("http://viewer.invalid/**", route => {
    const url = new URL(route.request().url())
    const json = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (/^\/matters\/matter-comments\/documents\/(redlined|content)$/.test(url.pathname))
      return route.fulfill({ status: 200, contentType: `${CT}.document`, body: Buffer.from(file.bytes) })
    if (url.pathname === "/matters/matter-comments/redlines") return json([])
    if (url.pathname === "/matters/matter-comments/documents/amendments") return json(null)
    if (url.pathname === "/matters/matter-comments/documents/markup" && route.request().method() === "GET" && withMarkup) return json(file.markup)
    if (url.pathname === "/preferences") return json({ attorney: LAWYER, firm: "Synthetic Firm" })
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: `Unexpected ${url.pathname}` }) })
  })
  await page.goto(`${server.url.href}${citation ? `#${encodeURIComponent(JSON.stringify(citation))}` : ""}`)
  await page.locator(".docx-render:not(.docx-skeleton) .comment-margin-note .markup-avatar").first().waitFor({ state: "attached" })
  return { page, errors }
}
const note = (page: Awaited<ReturnType<typeof open>>["page"], id: number) => page.locator(`.docx-render .comment-margin-note[data-comment-id="${id}"]`)
// Each byline's author, avatar classes and initials, in the panel or beside the page.
const avatars = (page: Awaited<ReturnType<typeof open>>["page"], scope: string) => page.locator(scope).evaluateAll((nodes) => nodes.map((node) => {
  const avatar = node.querySelector(".markup-avatar")!
  const author = node.matches(".comment-margin-note") ? node.getAttribute("data-author") : node.querySelector(".change-author")?.textContent
  return `${author}: ${avatar.className.replace(" is-small", "")} ${avatar.textContent}`
}))
const linked = (page: Awaited<ReturnType<typeof open>>["page"]) => page.locator(".docx-render .is-linked").evaluateAll(nodes => nodes.map(node => `${node.classList.contains("comment-highlight") ? "text" : node.classList.contains("doc-comment-pin") ? "pin" : "card"}#${node.getAttribute("data-comment-id")}`).sort())
const closePanel = async (page: Awaited<ReturnType<typeof open>>["page"]) => {
  await page.getByRole("button", { name: "Close tracked changes" }).click()
  await page.locator('.docx-render:not([data-comments="pane"]) .comment-margin-note').first().waitFor({ state: "visible" })
}
// Each comment button, scrolled into view: its thread, the button at its centre, and whether it covers the document's own text.
const reach = (page: Awaited<ReturnType<typeof open>>["page"]) => page.locator(".docx-render .doc-comment-pin").evaluateAll(pins => pins.map(pin => {
  pin.scrollIntoView({ block: "center" })
  const box = pin.getBoundingClientRect()
  const walker = document.createTreeWalker(pin.closest(".comment-margin-content") ?? pin.closest(".docx-render")!, NodeFilter.SHOW_TEXT)
  const range = document.createRange()
  let covers = false
  for (let node = walker.nextNode(); node && !covers; node = walker.nextNode()) {
    if (!node.textContent?.trim() || node.parentElement?.closest("a.comment-marker, .doc-comment-pin")) continue
    range.selectNodeContents(node)
    covers = [...range.getClientRects()].some(text => text.left < box.right - 0.5 && box.left < text.right - 0.5 && text.top < box.bottom - 0.5 && box.top < text.bottom - 0.5)
  }
  const under = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2)?.closest<HTMLElement>(".doc-comment-pin")
  return { id: (pin as HTMLElement).dataset.commentId, hit: under?.dataset.commentId ?? "none", covers }
}))
// Each crafted file's case runs on its own and reports, so one failure does not hide the next.
const pages: import("@playwright/test").Page[] = []
const failures: string[] = []
async function check(name: string, body: () => Promise<void>) {
  const from = pages.length
  try {
    await body()
    console.log(`PASS ${name}`)
  } catch (error) {
    failures.push(name)
    console.log(`FAIL ${name}\n${error instanceof Error ? error.message : String(error)}`)
  } finally {
    for (const page of pages.splice(from)) await page.close()
  }
}
const HIDDEN = /[\u200B-\u200F\u202A-\u202E\u2060-\u2064\u2066-\u2069\uFEFF]/u

try {
  const { page, errors } = await open({ width: 1440, height: 900 })
  // Once the panel has read the lawyer's name, both places colour Morgan as the lawyer.
  await page.locator(`#document-markup .markup-byline:has(.change-author:text-is("${LAWYER}")) .markup-avatar.tone-self`).first().waitFor({ state: "attached" })
  await note(page, 2).locator(".markup-avatar.tone-self").waitFor({ state: "attached" })
  const panel = new Set(await avatars(page, "#document-markup .markup-thread .markup-byline"))
  const margin = await avatars(page, ".docx-render .comment-margin-note")
  assert.equal(margin.length, 4, "a card for each root comment and the reply")
  for (const byline of margin) assert.ok(panel.has(byline), `the panel shows ${byline} too (panel: ${[...panel].join(", ")})`)
  assert.ok(margin.includes(`${LAWYER}: markup-avatar tone-self MOK`), "the lawyer's own comments are in the brand colour, with Word's initials")
  assert.equal(await note(page, 1).locator(".markup-avatar").getAttribute("class").then(value => value?.includes("is-small")), true, "a reply's initials are small")
  assert.equal(await note(page, 0).locator(":scope > .comment-margin-note-header > .markup-avatar").getAttribute("class").then(value => value?.includes("is-small")), false)
  assert.match(await note(page, 0).locator(":scope > .comment-margin-note-header time").textContent() ?? "", /^20 Oct 2026, \d\d:\d\d$/)

  // The panel opens by default and lists the comments, so the page shows none of its cards and keeps its own
  // width; each thread has a button in the right margin by its paragraph.
  const render = page.locator(".docx-render:not(.docx-skeleton)")
  await page.locator('.docx-render[data-comments="pane"]').waitFor({ state: "attached" })
  assert.equal(await page.locator(".docx-render .comment-margin-note:visible").count(), 0, "no cards beside the page while the panel is open")
  assert.equal(await render.evaluate(root => getComputedStyle(root).maxWidth), "760px")
  const pins = page.locator(".docx-render button.doc-comment-pin")
  assert.deepEqual(await pins.evaluateAll(nodes => nodes.map(node => node.getAttribute("aria-label"))), [
    "Comment by Avery Lindqvist", `Resolved comment by ${LAWYER}`, "Comment by Avery Lindqvist",
  ])
  assert.equal(await pins.evaluateAll(nodes => nodes.every(node => {
    const pin = node.getBoundingClientRect(), block = node.parentElement!.getBoundingClientRect(), page = node.closest(".docx-render")!.getBoundingClientRect()
    return pin.width >= 22 && pin.left >= block.right && pin.right <= page.right
  })), true, "each button sits in the page's right margin")
  if (screenshots) await page.screenshot({ path: path.join(screenshots, "comments-pane.png") })
  // Hovering a thread in the panel lights its text and button.
  await page.locator('#document-markup li.markup-thread[data-comment-id="3"]').hover()
  assert.deepEqual(await linked(page), ["card#3", "pin#3", "text#3"])
  await page.mouse.move(2, 2)
  assert.deepEqual(await linked(page), [])
  // A click on text no comment covers opens no thread.
  await page.locator(".docx-render .comment-margin-content p").filter({ hasText: /^9\. Governing law/ }).click({ position: { x: 4, y: 4 } })
  await page.waitForTimeout(500)
  assert.equal(await page.evaluate(() => document.activeElement?.closest("#document-markup .markup-thread")?.getAttribute("data-comment-id") ?? null), null, "no thread takes focus")
  // A button opens its thread in the Comments tab and focuses the card; so does the commented text.
  await pins.nth(2).click()
  await page.waitForFunction(() => document.activeElement?.matches('#document-markup li.markup-thread[data-comment-id="3"]'))
  assert.equal(await page.locator("#document-markup [role=tab][aria-selected=true]").textContent().then(text => text?.startsWith("Comments")), true)
  assert.deepEqual(await linked(page), ["card#3", "pin#3", "text#3"], "the focused card's text stays lit")
  if (screenshots) await page.screenshot({ path: path.join(screenshots, "comments-pane-revealed.png") })
  await page.locator('.docx-render .comment-highlight[data-comment-id="0"]').click()
  await page.waitForFunction(() => document.activeElement?.matches('#document-markup li.markup-thread[data-comment-id="0"]'))
  await page.locator('#document-markup li.markup-thread[data-comment-id="0"]').evaluate(card => (card as HTMLElement).blur())
  await pins.first().click()
  await page.waitForFunction(() => document.activeElement?.matches('#document-markup li.markup-thread[data-comment-id="0"]'))

  // Closed, the panel gives the comments back to the page as cards.
  await page.getByRole("button", { name: "Close tracked changes" }).click()
  await page.locator('.docx-render:not([data-comments="pane"]) .comment-margin-note').first().waitFor({ state: "visible" })
  assert.equal(await pins.count(), 0)
  assert.equal(await page.locator(".docx-render .comment-margin-note:visible").count(), 4)
  await page.mouse.move(2, 2)
  if (screenshots) await page.screenshot({ path: path.join(screenshots, "comments-closed.png") })

  // Resolved in words, in the card and its name; the screened comment flagged.
  assert.equal(await note(page, 2).locator(".comment-resolved-badge").isVisible(), true)
  assert.equal(await note(page, 2).locator(".comment-resolved-badge").textContent(), "Resolved")
  assert.equal(await note(page, 2).getAttribute("aria-label"), `Comment by ${LAWYER}, resolved`)
  const flag = note(page, 3).locator(".markup-flag")
  assert.equal(await flag.isVisible(), true)
  assert.match(await flag.textContent() ?? "", /looks like an instruction — doc\.haus did not follow it/)
  assert.equal(await note(page, 0).locator(".markup-flag").count(), 0)
  // No "[0]" in the text: each marker is words for a screen reader only.
  assert.deepEqual(await page.locator(".docx-render a.comment-marker").allTextContents(), [
    "(comment by Avery Lindqvist)", `(reply by ${LAWYER})`, `(resolved comment by ${LAWYER})`, "(comment by Avery Lindqvist)",
  ])
  assert.equal(await page.locator(".docx-render a.comment-marker").evaluateAll(nodes => nodes.every(node => node.getBoundingClientRect().width <= 1)), true)
  assert.equal(await page.locator(".docx-render .comment-margin-backref").evaluateAll(nodes => nodes.every(node => (node as HTMLElement).offsetParent === null)), true, "no ↩ links")

  // A card and its text light up together, by hover or focus; a reply lights its thread.
  await note(page, 0).locator(":scope > .comment-margin-note-body").hover()
  assert.deepEqual(await linked(page), ["card#0", "text#0"])
  if (screenshots) await page.screenshot({ path: path.join(screenshots, "comments-linked.png") })
  await note(page, 1).hover()
  assert.deepEqual(await linked(page), ["card#0", "text#0"])
  await page.locator('.docx-render .comment-highlight[data-comment-id="2"]').hover()
  assert.deepEqual(await linked(page), ["card#2", "text#2"])
  await page.mouse.move(2, 2)
  assert.deepEqual(await linked(page), [])
  await note(page, 3).focus()
  assert.deepEqual(await linked(page), ["card#3", "text#3"])
  await note(page, 3).blur()
  assert.deepEqual(await linked(page), [])

  // A narrow page: each card under its paragraph, its text there once; back in the margin when it widens.
  await page.setViewportSize({ width: 400, height: 860 })
  await page.locator(".docx-render .doc-comment-stack").first().waitFor({ state: "attached" })
  const stacked = await page.locator(".docx-render").evaluate(root => ({
    column: root.querySelector(".comment-margin-column")!.children.length,
    stacks: [...root.querySelectorAll(".doc-comment-stack")].map(stack => [...stack.children].map(card => {
      const id = card.getAttribute("data-comment-id")
      return `${id}:${Boolean(stack.previousElementSibling?.querySelector(`.comment-highlight[data-comment-id="${id}"]`))}`
    }).join(",")),
  }))
  assert.deepEqual(stacked, { column: 0, stacks: ["0:true", "2:true", "3:true"] })
  const text = await page.locator(".docx-render").textContent() ?? ""
  for (const { text: comment } of COUNTERPARTY_MARKUP.comments) assert.equal(text.split(comment).length - 1, 1, `once: ${comment}`)
  assert.ok(await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth) <= 1, "nothing wider than the phone")
  if (screenshots) await page.screenshot({ path: path.join(screenshots, "comments-narrow.png") })
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.waitForFunction(() => !document.querySelector(".docx-render .doc-comment-stack"))
  assert.deepEqual(await page.locator(".docx-render .comment-margin-column > .comment-margin-note").evaluateAll(nodes => nodes.map(node => node.getAttribute("data-comment-id"))), ["0", "2", "3"])
  assert.deepEqual(errors, [])
  console.log("PASS the open panel takes the comments, opened from their buttons and text; closed, the cards match its colours and initials, resolved and flagged in words, linked to their text, stacked and restored")
  await page.close()

  // A citation across a card stacked between two paragraphs and a marker still finds the document's text.
  const cited = await open({ width: 400, height: 860 }, { section: "1", excerpt: "forty-five days of each Order. 2. Price. The Price" }, false)
  await cited.page.locator(".docx-render .doc-comment-stack").first().waitFor({ state: "attached" })
  await cited.page.locator("mark.citation-passage").first().waitFor()
  assert.equal((await cited.page.locator("mark.citation-passage").allTextContents()).join(" ").replace(/\u00a0/g, " ").replace(/\s+/g, " "), "forty-five days of each Order. 2. Price. The Price")
  assert.equal(await cited.page.locator(".doc-comment-stack mark.citation-passage").count(), 0)
  console.log("PASS a citation across a stacked card selects only the document's text")
  await cited.page.close()

  await check("a card shows its comment without the hidden characters its flag says were removed", async () => {
    // A right-to-left override shows "0001" as "1000"; ingest strips it, and a zero-width space in the name, and says so.
    const author = "Avery Lindqvist", text = "Price is USD 0001 per unit."
    const file = {
      bytes: wordFile([para(run("1. Price. ") + commented(0, "The unit price is fixed."))], [{ id: 0, author: "Avery\u200B Lindqvist", text: "Price is USD \u202E0001\u202C per unit." }]),
      markup: listing([{ id: 0, author, text, findings: [
        { field: "author", rule: "zero-width", detail: "zero-width characters invisible to a human reader (1 removed)" },
        { field: "text", rule: "bidi-control", detail: "Unicode bidirectional controls that can visually reorder or mask text (2 removed)" },
      ] }]),
    }
    const { page, errors } = await open({ width: 1440, height: 900 }, undefined, true, file)
    const pin = page.locator('.docx-render[data-comments="pane"] .doc-comment-pin')
    await pin.waitFor({ state: "attached" })
    assert.equal(await pin.getAttribute("aria-label"), `Comment by ${author}`)
    await closePanel(page)
    const card = note(page, 0)
    assert.equal(await card.locator(".comment-margin-note-body").textContent(), text)
    assert.equal(await card.locator(".comment-margin-author").textContent(), author)
    assert.equal(HIDDEN.test(await card.textContent() ?? ""), false, "no hidden character anywhere in the card")
    assert.equal(await card.getAttribute("aria-label"), `Comment by ${author}`)
    assert.match(await card.locator(".markup-flag").textContent() ?? "", /Hidden characters were removed from this text/)
    assert.deepEqual(await page.locator(".docx-render a.comment-marker").allTextContents(), [`(comment by ${author})`])
    assert.deepEqual(errors, [])
  })

  await check("a bookmark named like a comment's marker moves neither its button nor its card", async () => {
    const comments = [{ id: 1, author: REVIEWERS[0]!, text: "Agreed, no change." }]
    const file = {
      bytes: wordFile([
        para('<w:bookmarkStart w:id="90" w:name="comment-ref-1"/>' + run("1. Delivery. The goods ship within ten days.") + '<w:bookmarkEnd w:id="90"/>'),
        para(run("2. Price. ") + commented(1, "The price is fixed for the term.")),
      ], comments),
      markup: listing(comments),
    }
    const { page, errors } = await open({ width: 1440, height: 900 }, undefined, true, file)
    const pin = page.locator('.docx-render[data-comments="pane"] .doc-comment-pin')
    await pin.waitFor({ state: "attached" })
    assert.match(await pin.evaluate(node => node.parentElement!.textContent ?? ""), /^2\. Price\./, "the button is by the commented clause")
    await closePanel(page)
    await page.setViewportSize({ width: 400, height: 860 })
    const stack = page.locator(".docx-render .doc-comment-stack")
    await stack.waitFor({ state: "attached" })
    assert.equal(await stack.evaluate(node => Boolean(node.previousElementSibling?.querySelector('.comment-highlight[data-comment-id="1"]'))), true, "the card is under the commented clause")
    assert.deepEqual(errors, [])
  })

  // Clicks each thread's button at its centre, as a lawyer would, and checks the panel focuses that thread.
  const clickEach = async (page: Awaited<ReturnType<typeof open>>["page"], ids: number[]) => {
    for (const id of ids) {
      const pin = page.locator(`.docx-render .doc-comment-pin[data-comment-id="${id}"]`)
      await pin.scrollIntoViewIfNeeded()
      const box = (await pin.boundingBox())!
      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2)
      await page.waitForFunction(id => document.activeElement?.matches(`#document-markup li.markup-thread[data-comment-id="${id}"]`), id, { timeout: 5_000 })
        .catch(async () => assert.fail(`the button for thread ${id} opened ${await page.evaluate(() => document.activeElement?.getAttribute("data-comment-id") ?? document.activeElement?.tagName)}`))
    }
  }

  await check("several buttons on one short clause can each be clicked, over no text, on a desktop or a phone", async () => {
    const comments = [0, 1, 2, 3, 4].map(id => ({ id, author: REVIEWERS[id]!, text: `Comment ${id}.` }))
    const file = {
      bytes: wordFile([
        para(run("1. ") + commented(0, "Short") + commented(1, " clause") + commented(2, " here.")),
        para(run("2. ") + commented(3, "Next clause.")),
        para(run("3. ") + commented(4, "Third clause.")),
      ], comments),
      markup: listing(comments),
    }
    for (const width of [1440, 400]) {
      const { page, errors } = await open({ width, height: 900 }, undefined, true, file)
      await page.locator('.docx-render[data-comments="pane"] .doc-comment-pin').first().waitFor({ state: "attached" })
      assert.deepEqual(await reach(page), comments.map(({ id }) => ({ id: String(id), hit: String(id), covers: false })), `${width}px`)
      await clickEach(page, comments.map(({ id }) => id))
      assert.ok(await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth) <= 1, `${width}px: nothing wider than the window`)
      assert.deepEqual(errors, [])
    }
  })

  await check("a table cell's buttons sit inside the cell, clickable and over none of its text", async () => {
    const comments = [{ id: 0, author: REVIEWERS[0]!, text: "Left." }, { id: 1, author: REVIEWERS[1]!, text: "Right." }, { id: 2, author: REVIEWERS[2]!, text: "Right again." }]
    const table = '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>' +
      cell(commented(0, "The Supplier delivers the goods to the Buyer's warehouse within ten business days of each order.")) +
      cell(commented(1, "Price") + run(" and ") + commented(2, "terms")) + "</w:tr></w:tbl>"
    const file = { bytes: wordFile([para(run("1. Schedule.")), table, para(run("2. After the schedule."))], comments), markup: listing(comments) }
    for (const width of [1440, 400]) {
      const { page, errors } = await open({ width, height: 900 }, undefined, true, file)
      await page.locator('.docx-render[data-comments="pane"] .doc-comment-pin').first().waitFor({ state: "attached" })
      assert.deepEqual(await reach(page), comments.map(({ id }) => ({ id: String(id), hit: String(id), covers: false })), `${width}px`)
      assert.deepEqual(await page.locator(".docx-render .doc-comment-pin").evaluateAll(pins => pins.map(pin => {
        const box = pin.getBoundingClientRect(), cell = pin.parentElement!.getBoundingClientRect()
        return `${pin.parentElement!.tagName}:${box.left >= cell.left && box.right <= cell.right && box.top >= cell.top && box.bottom <= cell.bottom}`
      })), ["TD:true", "TD:true", "TD:true"], `${width}px: each button inside its cell`)
      await clickEach(page, comments.map(({ id }) => id))
      assert.deepEqual(errors, [])
    }
  })

  assert.deepEqual(failures, [], "every crafted file's case passes")
} finally {
  await browser.close()
  server.stop(true)
}
