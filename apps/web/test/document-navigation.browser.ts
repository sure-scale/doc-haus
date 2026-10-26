// Run from apps/web: bun test/document-navigation.browser.ts
// Real DocumentViewer + Docxodus + app CSS; only ingest HTTP is mocked.
// SCREENSHOT_DIR keeps desktop/narrow evidence outside the source checkout.
import assert from "node:assert/strict"
import { existsSync, statSync } from "node:fs"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { strToU8, zipSync } from "fflate"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const wasmDir = path.resolve(process.env.DOCXODUS_WASM_DIR ?? path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm"))
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-document-navigation-"))
const screenshots = process.env.SCREENSHOT_DIR
if (screenshots) await mkdir(screenshots, { recursive: true })
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const run = (text: string, bold = false) => `<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`
const paragraph = (text: string) => `<w:p>${run(text)}</w:p>`
const heading = (text: string, level: number) => `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/><w:numPr><w:ilvl w:val="${level - 1}"/><w:numId w:val="1"/></w:numPr></w:pPr>${run(text)}</w:p>`
const citation = "The payment terms remain subject to written agreement."
function document(older = false, hardBreak = false) {
  const body = older ? paragraph("Earlier text with no headings.") : [
    heading("Scope", 1),
    hardBreak
      ? `<w:p>${run("The payment", true)}<w:r><w:br/></w:r>${run("terms apply to each invoice.")}</w:p>`
      : `<w:p>${run("The pay", true)}${run("ment terms apply to each invoice.")}</w:p>`,
    ...Array.from({ length: 12 }, (_, index) => paragraph(`Supporting paragraph ${index + 1}. The parties shall record their instructions in writing.`)),
    heading("Charges", 2), paragraph(citation),
    paragraph("The fee [insert amount] is payable on receipt."),
    ...Array.from({ length: 12 }, (_, index) => paragraph(`Further paragraph ${index + 1}. The services include preparation and review of the agreed documents.`)),
    heading("Notices", 1),
    `<w:p><w:ins w:id="1" w:author="Synthetic Counsel" w:date="2026-10-25T10:00:00Z">${run("PAYMENT TERMS may be revised in writing.")}</w:ins></w:p>`,
    paragraph("A literal [charge] is shown here."),
  ].join("")
  const parts = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${CT}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${CT}.numbering+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdN" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`,
    "word/styles.xml": `<w:styles xmlns:w="${W}">${[1, 2].map(level => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:pPr><w:outlineLvl w:val="${level - 1}"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>`).join("")}</w:styles>`,
    "word/numbering.xml": `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${[0, 1].map(level => `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="${level ? "%1.%2" : "%1."}"/><w:suff w:val="space"/></w:lvl>`).join("")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    "word/document.xml": `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  }
  return zipSync(Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8"?>${xml}`)])))
}
const bytes = document(), earlier = document(true), brokenLine = document(false, true)
const kept = (number: number, current: boolean) => ({ id: `v-${number}`, number, sha256: String(number).repeat(64), ext: ".docx", bytes: bytes.length, parent: null, derived_from: null, origin: "upload", actor: "unknown", run: null, label: null, created_at: 0, current, removed: null })
const proposal = { id: 1, doc_name: "Terms.docx", scope: "phrase", find_text: "payment terms", old_text: "payment terms", new_text: "agreed payment terms", author: "Synthetic Counsel", review_comment: "Confirm payment terms before signing.", created_at: 0 }
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement as h } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentViewer from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentViewer.tsx"))}
const text = location.hash === "#text"
createRoot(document.getElementById("root")).render(h(DocumentViewer, { matterId: "navigation", name: text ? "Instructions.eml" : "Terms.docx", citation: text ? undefined : { section: "Charges", excerpt: ${JSON.stringify(citation)} }, onClose: () => { document.body.dataset.closed = "true" } }))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://viewer.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const styles = (await Promise.all([
  "styles/design-system.css", "styles/primitives.css", "styles/design-documents.css", "styles.css", "components/DocumentNavigation.css", "components/ChangeTimeline.css",
].map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch(request) {
  const { pathname } = new URL(request.url)
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname === "/app.css") return new Response(styles, { headers: { "content-type": "text/css" } })
  if (pathname.startsWith("/wasm/")) {
    const file = path.join(wasmDir, decodeURIComponent(pathname.slice("/wasm/".length)))
    return file.startsWith(wasmDir + path.sep) && existsSync(file) && statSync(file).isFile() ? new Response(Bun.file(file)) : new Response("Not found", { status: 404 })
  }
  return new Response('<html><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="stylesheet" href="/app.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>', { headers: { "content-type": "text/html" } })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const writes: string[] = []
async function open(width = 1440, text = false, fallback = false, hardBreak = false) {
  const page = await browser.newPage({ viewport: { width, height: 900 } })
  page.setDefaultTimeout(30_000)
  if (fallback) await page.addInitScript(() => { Object.defineProperty(window, "Highlight", { value: undefined }) })
  await page.route("http://viewer.invalid/**", route => {
    const request = route.request(), url = new URL(request.url())
    if (request.method() !== "GET") { writes.push(request.method() + " " + url.pathname); return route.fulfill({ status: 405 }) }
    if (/^\/matters\/navigation\/documents\/(redlined|content)$/.test(url.pathname)) return route.fulfill({ status: 200, body: Buffer.from(hardBreak ? brokenLine : bytes), contentType: `${CT}.document` })
    const json = (body: unknown) => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify(body) })
    if (url.pathname === "/matters/navigation/redlines") return json([proposal])
    if (url.pathname === "/matters/navigation/documents/amendments") return json(null)
    if (url.pathname === "/matters/navigation/documents/versions") return json({ name: "Terms.docx", versions: text ? [] : [kept(2, true), kept(1, false)], events: [], working_version: text ? null : "v-2", document: null })
    if (url.pathname === "/matters/navigation/documents/versions/v-1/content") return route.fulfill({ status: 200, body: Buffer.from(earlier), contentType: `${CT}.document` })
    if (url.pathname === "/matters/navigation/documents/text") return json({ text: "Subject: Payment terms\nPayment terms remain open.\n", injection: null })
    if (url.pathname === "/preferences") return json({ attorney: "Synthetic Counsel" })
    return route.fulfill({ status: 404, contentType: "application/json", body: "{}" })
  })
  await page.goto(server.url.href + (text ? "#text" : ""))
  await page.locator(".docx-render:not(.docx-skeleton)").waitFor()
  if (!text) await page.locator("mark.citation-passage").first().waitFor()
  return page
}
try {
  const page = await open()
  const documentRoot = page.locator(".docx-render")
  await page.getByRole("button", { name: "Outline", exact: true }).click()
  const outline = page.getByRole("navigation", { name: "Document outline" })
  const labels = await outline.locator("li button").allTextContents()
  console.log("Numbered outline:", JSON.stringify(labels))
  assert.deepEqual(labels.map(label => label.replace(/\s+/g, " ").trim()), ["1. Scope", "1.1 Charges", "2. Notices"])
  await outline.getByRole("button", { name: "2. Notices", exact: true }).click()
  await expect(documentRoot.locator("h1").last()).toBeFocused()
  assert.equal(await documentRoot.locator("h1").last().evaluate(node => { const rect = node.getBoundingClientRect(), viewport = node.closest(".viewer-doc")!.getBoundingClientRect(); return rect.top >= viewport.top && rect.top < viewport.bottom }), true, "outline scrolls its rendered heading into view")
  // Retain actual node identities: Find cannot remount citation, run or change
  // markup, even when its query spans their boundaries.
  await documentRoot.evaluate(root => { (window as unknown as { preserved: Node[] }).preserved = [...root.querySelectorAll("mark.citation-passage, ins.rev-ins, .doc-marker, h1")] })
  await page.keyboard.press("ControlOrMeta+f")
  const input = page.getByRole("searchbox", { name: "Find in document" })
  await expect(input).toBeFocused()
  await input.fill("payment terms")
  const count = page.locator(".document-find-count")
  await expect(count).toHaveText("1 of 3")
  await input.press("Shift+Enter")
  await expect(count).toHaveText("3 of 3")
  await input.press("Enter")
  await expect(count).toHaveText("1 of 3")
  await page.getByRole("button", { name: "Next match", exact: true }).click()
  await expect(count).toHaveText("2 of 3")
  await page.getByRole("button", { name: "Previous match", exact: true }).click()
  await expect(count).toHaveText("1 of 3")
  const highlighted = await page.evaluate(() => [...CSS.highlights.entries()].filter(([name]) => name.endsWith("-find")).flatMap(([, highlight]) => [...highlight].map(range => range.toString())))
  assert.deepEqual(highlighted.map(text => text.toLowerCase()), ["payment terms", "payment terms", "payment terms"], "split runs and tracked insertion are searched; sidebar/proposal copies stay outside scope")
  assert.equal(await documentRoot.evaluate(root => (window as unknown as { preserved: Node[] }).preserved.every(node => root.contains(node))), true, "Find preserves existing citation, marker, heading and tracked-change nodes")
  assert.equal(await page.locator(".changes-panel .change-card").count(), 1)
  assert.equal((await page.locator("mark.citation-passage").allTextContents()).join(""), citation)
  // Hover/markup decorators may change classes while the reader scrolls.
  // Repainting those live ranges must not jump back to the active match.
  const scrolled = await documentRoot.evaluate(root => {
    const viewport = root.closest(".viewer-doc")!
    viewport.scrollTop += 180
    root.classList.add("synthetic-decoration-change")
    return viewport.scrollTop
  })
  await page.waitForTimeout(100)
  assert.equal(await documentRoot.evaluate(root => root.closest(".viewer-doc")!.scrollTop), scrolled, "decoration updates preserve the reader's scroll position")
  await input.fill("[charge]")
  await expect(count).toHaveText("1 of 1")
  await input.fill("unfindable text")
  await expect(count).toHaveText("No matches")
  await expect(page.getByRole("button", { name: "Next match", exact: true })).toBeDisabled()
  await input.fill("payment terms")
  await expect(count).toHaveText("1 of 3")
  if (screenshots) await page.screenshot({ path: path.join(screenshots, "viewer-navigation-desktop.png") })
  await input.press("Escape")
  await expect(input).toHaveCount(0)
  await expect(page.getByRole("button", { name: "Find", exact: true })).toBeFocused()
  assert.equal(await page.locator("body").getAttribute("data-closed"), null, "Escape closes Find without closing the viewer")
  assert.equal(await page.evaluate(() => [...CSS.highlights.keys()].filter(name => name.startsWith("document-navigation-")).length), 0, "closing Find releases its highlights")
  await page.keyboard.press("Escape")
  await expect(outline).toHaveCount(0)
  await expect(page.getByRole("button", { name: "Outline", exact: true })).toBeFocused()
  assert.equal(await page.locator("body").getAttribute("data-closed"), null)
  // Search and outline follow the displayed kept version, not the hidden
  // working file. This also exercises replacement of the converted document.
  await page.getByRole("button", { name: "Outline", exact: true }).click()
  await page.getByRole("button", { name: /^Versions/ }).click()
  await page.locator('[data-version-view="v-1"]').click()
  await expect(outline).toContainText("No headings in this document.")
  await page.keyboard.press("ControlOrMeta+f")
  await input.fill("payment terms")
  await expect(count).toHaveText("No matches")
  await page.close()
  console.log("PASS numbered outline, scoped split-run Find, wrap/keyboard, preserved markup, Escape and version refresh")

  // Change only the inline split-run paragraph to a real Word hard break.
  // Its rendered break separates words, while inline runs above join them.
  const lineBreak = await open(1440, false, false, true)
  await expect(lineBreak.locator(".docx-render br")).toHaveCount(1)
  await lineBreak.keyboard.press("ControlOrMeta+f")
  const breakInput = lineBreak.getByRole("searchbox")
  await breakInput.fill("payment terms")
  await expect(lineBreak.locator(".document-find-count")).toHaveText("1 of 3")
  assert.equal(await lineBreak.evaluate(() => [...CSS.highlights.entries()].filter(([name]) => name.endsWith("-find")).some(([, highlight]) => [...highlight].some(range => (range as Range).cloneContents().querySelector("br")))), true, "the match range crosses the actual rendered break")
  await breakInput.fill("paymentterms")
  await expect(lineBreak.locator(".document-find-count")).toHaveText("No matches")
  await lineBreak.close()
  console.log("PASS Word hard break separates words without changing inline split-run matching")

  const narrow = await open(400)
  await narrow.getByRole("button", { name: "Outline", exact: true }).click()
  await narrow.getByRole("button", { name: "Find", exact: true }).click()
  await narrow.getByRole("searchbox").fill("payment terms")
  await expect(narrow.locator(".document-find-count")).toHaveText("1 of 3")
  assert.equal(await narrow.locator(".viewer-panel").evaluate(panel => panel.scrollWidth <= panel.clientWidth), true, "navigation fits the narrow viewer")
  if (screenshots) await narrow.screenshot({ path: path.join(screenshots, "viewer-navigation-400px.png") })
  await narrow.getByRole("navigation", { name: "Document outline" }).getByRole("button", { name: "2. Notices", exact: true }).click()
  await expect(narrow.getByRole("navigation", { name: "Document outline" })).toHaveCount(0)
  await expect(narrow.locator(".docx-render h1").last()).toBeFocused()
  assert.equal(await narrow.locator(".docx-render h1").last().evaluate(node => { const rect = node.getBoundingClientRect(), viewport = node.closest(".viewer-body")!.getBoundingClientRect(); return rect.top >= viewport.top && rect.top < viewport.bottom }), true)
  await narrow.close()
  console.log("PASS narrow layout and focused heading jump")

  const text = await open(400, true)
  await text.keyboard.press("ControlOrMeta+f")
  await text.getByRole("searchbox").fill("payment terms")
  await expect(text.locator(".document-find-count")).toHaveText("1 of 2")
  assert.equal(await text.getByRole("button", { name: "Outline", exact: true }).count(), 0)
  await text.close()
  const fallback = await open(1440, false, true)
  await fallback.keyboard.press("ControlOrMeta+f")
  await fallback.getByRole("searchbox").fill("payment terms")
  await expect(fallback.locator(".document-find-count")).toHaveText("1 of 3")
  await expect(fallback.getByRole("searchbox")).toBeFocused()
  assert.equal((await fallback.evaluate(() => getSelection()?.toString()))?.toLowerCase(), "payment terms")
  await fallback.getByRole("button", { name: "Close Find" }).click()
  assert.equal(await fallback.evaluate(() => getSelection()?.toString()), "")
  await fallback.close()
  assert.deepEqual(writes, [], "navigation never writes document content")
  console.log("PASS extracted text and browser selection fallback")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
