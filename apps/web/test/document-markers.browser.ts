// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/document-markers.browser.ts
// Mounts the real DocumentViewer on the Docxodus WASM runtime; only ingest HTTP is mocked.
// Redaction marks and placeholders get two chip classes on exactly their spans, the text is unchanged, and a citation on adjacent or overlapping text still selects the right characters.
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
// Paragraphs as lists of runs; the fee placeholder is split across two runs.
const paragraphs: string[][] = [
  ["Prepared for [insert Client name] by counsel."],
  ["The witness ", "[REDACTED]", " shall pay the invoice."],
  ["Hearing set for [open item: confirm date] unless moved."],
  ["The fee of [insert fee", " amount] is due on receipt."],
  ["Nothing in this paragraph needs attention."],
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

const entry = path.join((await import("node:os")).tmpdir(), `dochaus-markers-${process.pid}.tsx`)
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentViewer from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentViewer.tsx"))}
const citation = location.hash.length > 1 ? JSON.parse(decodeURIComponent(location.hash.slice(1))) : undefined
createRoot(document.getElementById("root")).render(createElement(DocumentViewer, { matterId: "matter-markers", name: "markers.docx", citation, onClose: () => {} }))
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

async function open(citation?: { section: string; excerpt: string }) {
  const page = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  page.setDefaultTimeout(60_000)
  await page.route("http://viewer.invalid/**", route => {
    const url = new URL(route.request().url())
    if (/^\/matters\/matter-markers\/documents\/(redlined|content)$/.test(url.pathname))
      return route.fulfill({ status: 200, contentType: `${CT}.document`, body: Buffer.from(bytes) })
    if (url.pathname === "/matters/matter-markers/redlines") return route.fulfill({ status: 200, contentType: "application/json", body: "[]" })
    if (url.pathname === "/matters/matter-markers/documents/amendments") return route.fulfill({ status: 200, contentType: "application/json", body: "null" })
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: `Unexpected ${url.pathname}` }) })
  })
  await page.goto(`${server.url.href}${citation ? `#${encodeURIComponent(JSON.stringify(citation))}` : ""}`)
  await page.locator(".docx-render:not(.docx-skeleton) .doc-marker").first().waitFor()
  return page
}
const spans = (page: Awaited<ReturnType<typeof open>>, kind: string) => page.locator(`.docx-render .doc-marker-${kind}`).allTextContents().then(texts => texts.map(text => text.replace(/\u00a0/g, " ")))
try {
  const page = await open()
  assert.deepEqual(await spans(page, "redaction"), ["[REDACTED]"])
  assert.deepEqual(await spans(page, "placeholder"), ["[insert Client name]", "[open item: confirm date]", "[insert fee", " amount]"])
  const text = (await page.locator(".docx-render").textContent())!.replace(/ /g, " ")
  for (const runs of paragraphs) assert.ok(text.includes(runs.join("")), `text unchanged: ${runs.join("")}`)
  assert.equal(await page.locator(".docx-render .doc-marker").evaluateAll(nodes => nodes.filter(node => node.closest("p")?.textContent?.includes("Nothing in this paragraph")).length), 0, "plain text is not decorated")
  const fee = page.locator(".doc-marker[data-marker-id]").filter({ hasText: "[insert fee" })
  assert.equal(await page.locator('.doc-marker-placeholder[data-marker-part="first"], .doc-marker-placeholder[data-marker-part="only"]').count(), 3, "split-run marker is one chip with the others")
  assert.equal(await fee.evaluate(node => node.getAttribute("aria-label")), "[insert fee amount]")
  assert.equal(await page.locator('.doc-marker[aria-hidden="true"]').count(), 1, "only the second piece of the split marker is hidden from the accessibility tree")
  assert.equal(await page.locator('.doc-marker[data-marker-part="last"]').evaluate(node => getComputedStyle(node).borderLeftWidth), "0px", "no inner border")
  const style = (kind: string) => page.locator(`.doc-marker-${kind}`).first().evaluate(node => { const s = getComputedStyle(node); return { background: s.backgroundColor, color: s.color, border: s.borderTopStyle, title: node.getAttribute("title"), role: node.getAttribute("aria-roledescription") } })
  const redaction = await style("redaction"), placeholder = await style("placeholder")
  assert.notEqual(redaction.background, placeholder.background)
  assert.notEqual(redaction.border, "none"); assert.notEqual(placeholder.border, "none")
  assert.notEqual(redaction.border, placeholder.border, "not colour alone: border styles differ")
  assert.equal(redaction.title, "Redacted text"); assert.equal(placeholder.role, "Placeholder: to complete")
  assert.match(await page.locator(".viewer-title-markers").textContent() ?? "", /^3 placeholders · 1 redaction$/)
  await page.locator(".docx-render").evaluate(node => node.scrollIntoView())
  if (screenshots) await page.screenshot({ path: path.join(screenshots, "markers-chips.png") })
  console.log("PASS chips on exactly the marker spans, text unchanged", JSON.stringify({ redaction, placeholder }))
  await page.close()

  // A citation on the text next to a marker selects exactly that text.
  const adjacent = await open({ section: "2", excerpt: "shall pay the invoice." })
  await adjacent.locator("mark.citation-passage").first().waitFor()
  assert.equal((await adjacent.locator("mark.citation-passage").allTextContents()).join("").replace(/\u00a0/g, " "), "shall pay the invoice.")
  if (screenshots) await adjacent.screenshot({ path: path.join(screenshots, "markers-citation-adjacent.png") })
  console.log("PASS citation next to a redaction selects its own text")
  await adjacent.close()

  // A citation crossing a marker boundary selects the same characters as without chips.
  const crossing = await open({ section: "2", excerpt: "[REDACTED] shall pay" })
  await crossing.locator("mark.citation-passage").first().waitFor()
  assert.equal((await crossing.locator("mark.citation-passage").allTextContents()).join("").replace(/\u00a0/g, " "), "[REDACTED] shall pay")
  assert.deepEqual(await spans(crossing, "redaction"), ["[REDACTED]"])
  if (screenshots) await crossing.screenshot({ path: path.join(screenshots, "markers-citation-over-marker.png") })
  console.log("PASS citation across a marker selects the exact characters")
  await crossing.close()
} finally {
  await browser.close()
  server.stop(true)
}
