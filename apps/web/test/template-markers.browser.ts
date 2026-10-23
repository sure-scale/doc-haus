// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/template-markers.browser.ts
// Mounts the real TemplateViewer on the Docxodus WASM runtime with a synthetic template; nothing else is mocked.
// The side panel lists the same markers the document shows (a [REDACTED] mark is a redaction, not a placeholder),
// an entry and its chips highlight together from either side, and a click walks through a marker's occurrences.
// Set PLAYWRIGHT_CHROMIUM_PATH=<chromium> when Playwright's own browser is missing; SCREENSHOT_DIR=<dir> saves one screenshot per state.
import assert from "node:assert/strict"
import { existsSync, statSync } from "node:fs"
import { mkdir } from "node:fs/promises"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const { zipSync, strToU8 } = (await import(Bun.resolveSync("fflate", path.resolve(import.meta.dir, "../../../dochaus")))) as {
  zipSync: (files: Record<string, Uint8Array>) => Uint8Array; strToU8: (text: string) => Uint8Array
}
const wasmDir = path.resolve(process.env.DOCXODUS_WASM_DIR ?? path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm"))

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const filler = (n: number) => Array.from({ length: n }, (_, i) => [`Clause ${i + 1}. The parties agree the general terms in this clause apply throughout.`])
// Paragraphs as lists of runs. The party name occurs twice, far apart; the fee is split across two runs.
const paragraphs: string[][] = [
  ["This agreement is made between [insert party one name] and the Supplier."],
  ["The witness ", "[REDACTED]", " and the referee [REDACTED] attended."],
  ["The fee of [insert fee", " amount] is due on receipt."],
  ...filler(30),
  ["[optional: Non-solicitation] Neither party shall solicit staff."],
  ["Signed for [insert party one name] by its director [REDACTED]."],
]
const run = (text: string, bold = false) => `<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`
const body = paragraphs.map(runs => `<w:p>${runs.map((text, i) => run(text, i % 2 === 1)).join("")}</w:p>`).join("")
const files: Record<string, string> = {
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  "word/document.xml": `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
}
const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)])))

const entry = path.join((await import("node:os")).tmpdir(), `dochaus-template-markers-${process.pid}.tsx`)
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import TemplateViewer from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/TemplateViewer.tsx"))}
const loadBytes = () => fetch("/template.docx").then(r => r.arrayBuffer()).then(b => new Uint8Array(b))
createRoot(document.getElementById("root")).render(createElement(TemplateViewer, { name: "Synthetic template.docx", loadBytes, onClose: () => {} }))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://viewer.invalid" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = ["design-system", "primitives", "design-documents", "design-library"].map(name => path.resolve(import.meta.dir, `../src/styles/${name}.css`))
const styles = (await Promise.all(css.map(file => Bun.file(file).text()))).join("\n") + "\n" + (await Bun.file(path.resolve(import.meta.dir, "../src/styles.css")).text())
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", fetch: request => {
  const { pathname } = new URL(request.url)
  if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
  if (pathname === "/app.css") return new Response(styles, { headers: { "content-type": "text/css" } })
  if (pathname === "/template.docx") return new Response(bytes, { headers: { "content-type": `${CT}.document` } })
  if (pathname.startsWith("/wasm/")) {
    const file = path.join(wasmDir, decodeURIComponent(pathname.slice("/wasm/".length)))
    return file.startsWith(wasmDir + path.sep) && existsSync(file) && statSync(file).isFile() ? new Response(Bun.file(file)) : new Response("Not found", { status: 404 })
  }
  return new Response(`<html><meta charset="utf-8"><link rel="stylesheet" href="/app.css"><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
} })
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const screenshots = process.env.SCREENSHOT_DIR
if (screenshots) await mkdir(screenshots, { recursive: true })
const shot = async (page: import("@playwright/test").Page, name: string) => { if (screenshots) await page.screenshot({ path: path.join(screenshots, name) }) }

try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } })
  page.setDefaultTimeout(60_000)
  await page.goto(server.url.href)
  const panel = page.locator(".changes-panel")
  await panel.locator(".marker-entry").first().waitFor()
  const heads = await panel.locator(".changes-head").allTextContents()
  const chips = (kind: string) => page.locator(`.docx-render .doc-marker-${kind}:is([data-marker-part="only"], [data-marker-part="first"])`)
  // Counts agree with the chips in the document: the redaction marks are not placeholders.
  assert.deepEqual(heads, [`${await chips("placeholder").count() - 1} placeholders`, "1 optional clause", `${await chips("redaction").count()} redactions`])
  assert.deepEqual(heads, ["3 placeholders", "1 optional clause", "3 redactions"])
  const entries = await panel.locator(".marker-entry-text").allTextContents()
  assert.deepEqual(entries, ["[insert party one name]", "[insert fee amount]", "Non-solicitation", "[REDACTED]"], "one entry per distinct marker; the split fee is one")
  assert.equal(await panel.locator(".marker-entry", { hasText: "[insert party one name]" }).locator(".marker-entry-count").textContent(), "×2")
  // Same chip style as the document, and never colour alone.
  const look = (locator: import("@playwright/test").Locator) => locator.evaluate(node => { const s = getComputedStyle(node); return { background: s.backgroundColor, color: s.color, border: s.borderTopStyle } })
  const placeholderEntry = panel.locator(".marker-entry.doc-marker-placeholder").first(), redactionEntry = panel.locator(".marker-entry.doc-marker-redaction")
  assert.deepEqual(await look(placeholderEntry), await look(chips("placeholder").first()))
  assert.deepEqual(await look(redactionEntry), await look(chips("redaction").first()))
  assert.equal((await look(placeholderEntry)).border, "dashed"); assert.equal((await look(redactionEntry)).border, "solid")
  assert.match(await redactionEntry.getAttribute("aria-label") ?? "", /^\[REDACTED\], redaction, 3 occurrences\. Show in document$/)
  await shot(page, "template-panel.png")
  console.log("PASS the panel lists the document's markers with matching counts and chip styles")

  // Hover a panel entry: every piece of its chips highlights, including both runs of the split fee.
  // Highlights follow React state, so wait for each to settle.
  const settled = (selector: string, count: number, message: string) => page.waitForFunction(([selector, count]) => document.querySelectorAll(selector).length === count, [selector, count] as const, { timeout: 5000 }).catch(() => assert.fail(message))
  const fee = panel.locator(".marker-entry", { hasText: "[insert fee amount]" })
  await fee.hover()
  await settled(".docx-render .doc-marker.is-linked", 2, "hovering the fee entry highlights its chip")
  assert.deepEqual(await page.locator(".docx-render .doc-marker.is-linked").allTextContents().then(t => t.map(s => s.replace(/\u00a0/g, " "))), ["[insert fee", " amount]"])
  await shot(page, "template-hover-entry.png")
  await page.mouse.move(5, 5)
  await settled(".docx-render .is-linked", 0, "leaving the entry clears the highlight")
  // Keyboard focus behaves like hover.
  await redactionEntry.focus()
  await settled(".docx-render .doc-marker-redaction.is-linked", 3, "focusing the entry highlights every occurrence")
  await redactionEntry.blur()
  await settled(".docx-render .is-linked", 0, "blur clears the highlight")
  // Hover a chip in the document: its panel entry highlights; plain text clears it.
  await chips("redaction").first().hover()
  await settled(".marker-entry.is-linked", 1, "hovering a chip highlights its entry")
  assert.equal(await panel.locator(".marker-entry.is-linked").textContent(), "[REDACTED]×3")
  await settled(".docx-render .doc-marker-redaction.is-linked", 3, "all occurrences highlight with it")
  await shot(page, "template-hover-chip.png")
  await page.locator(".docx-render p", { hasText: "Clause 1." }).first().hover()
  await settled(".marker-entry.is-linked", 0, "plain text clears the entry highlight")
  console.log("PASS hover and focus highlight an entry and its chips from either side")

  // Click walks through occurrences: first, then second (far down), then back to the first.
  const party = panel.locator(".marker-entry", { hasText: "[insert party one name]" })
  const doc = page.locator(".viewer-doc")
  const flashed = () => page.locator(".docx-render .doc-marker.is-flash").evaluateAll(nodes => nodes.map(node => node.closest("p")?.textContent?.slice(0, 20)))
  await party.click()
  assert.deepEqual(await flashed(), ["This agreement is ma"])
  assert.equal(await party.locator(".marker-entry-count").textContent(), "1/2")
  await party.click()
  assert.deepEqual(await flashed(), ["Signed for [insert p"])
  assert.equal(await party.locator(".marker-entry-count").textContent(), "2/2")
  await page.waitForFunction(() => { const chip = document.querySelector(".docx-render .doc-marker.is-flash")!.getBoundingClientRect(), box = document.querySelector(".viewer-doc")!.getBoundingClientRect(); return chip.top >= box.top && chip.bottom <= box.bottom })
  assert.ok(await doc.evaluate(node => node.scrollTop) > 0, "the document scrolled to the second occurrence")
  await shot(page, "template-click-second.png")
  await party.click()
  assert.deepEqual(await flashed(), ["This agreement is ma"])
  await page.waitForFunction(() => document.querySelectorAll(".docx-render .is-flash").length === 0, undefined, { timeout: 5000 })
  console.log("PASS clicking an entry scrolls to each occurrence in turn and briefly highlights it")
  await page.close()

  // Reduced motion: the jump is instant.
  const still = await browser.newPage({ viewport: { width: 1280, height: 800 }, reducedMotion: "reduce" })
  await still.goto(server.url.href)
  await still.locator(".changes-panel .marker-entry").first().waitFor()
  const stillParty = still.locator(".changes-panel .marker-entry", { hasText: "[insert party one name]" })
  await stillParty.click(); await stillParty.click()
  assert.ok(await still.locator(".viewer-doc").evaluate(node => node.scrollTop) > 0, "scrolled at once without smooth motion")
  console.log("PASS reduced motion scrolls without animation")
  await still.close()

  // Forced colours drop backgrounds and shadows: border styles and an outline carry the meaning.
  const forced = await browser.newPage({ viewport: { width: 1280, height: 800 }, forcedColors: "active" })
  await forced.goto(server.url.href)
  await forced.locator(".changes-panel .marker-entry").first().waitFor()
  await forced.locator(".changes-panel .marker-entry", { hasText: "[insert fee amount]" }).hover()
  await forced.waitForFunction(() => document.querySelectorAll(".docx-render .doc-marker.is-linked").length === 2, undefined, { timeout: 5000 })
  const ring = await forced.locator(".docx-render .doc-marker.is-linked").first().evaluate(node => getComputedStyle(node).outlineStyle)
  assert.equal(ring, "solid", "the linked chip keeps a visible outline")
  assert.equal(await forced.locator(".marker-entry.doc-marker-redaction").evaluate(node => getComputedStyle(node).borderTopStyle), "solid")
  assert.equal(await forced.locator(".marker-entry.doc-marker-placeholder").first().evaluate(node => getComputedStyle(node).borderTopStyle), "dashed")
  await shot(forced, "template-forced-colors.png")
  console.log("PASS forced colours keep the border styles and the linked outline")
  await forced.close()
} finally {
  await browser.close()
  server.stop(true)
}
