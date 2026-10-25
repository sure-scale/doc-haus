// Run from apps/web: bun test/docx-render-css.browser.ts
// Values from a Word file are written into the rendered page's stylesheet.
// Mounts the real DocumentViewer, BlacklineViewer and TemplateViewer on the
// Docxodus WASM runtime with a synthetic file in which each case below tries to
// leave its value and add rules: one outlining that case's probe on the app
// page, outside the document, and (where noted) one loading a remote image
// named after the case. None may apply in any viewer; ordinary font names
// (Latin and Japanese) must still reach the stylesheet and a file's own style
// must still apply. Only ingest HTTP is mocked.
// DOCHAUS_EVIDENCE_DIR optionally keeps a screenshot per viewer.
import assert from "node:assert/strict"
import { existsSync, statSync } from "node:fs"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { strToU8, zipSync } from "fflate"

const { chromium } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const wasmDir = path.resolve(process.env.DOCXODUS_WASM_DIR ?? path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm"))
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-docx-render-css-"))
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
const evidence = process.env.DOCHAUS_EVIDENCE_DIR
if (evidence) await mkdir(evidence, { recursive: true })

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const remote = (name: string) => `http://remote.invalid/${name}.png`
const xml = (text: string) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;")
const probe = (name: string, colour: string) => `#${name}-probe { outline: 6px solid ${colour} }`
// A font name is written inside quotes (font-family: 'Name', serif).
const font = `X'; } p { background-image: url('${remote("font")}') } ${probe("font", "rgb(220, 0, 0)")} span.zz { font-family: 'Y`
// A colour is written bare (color: #RRGGBB), so a semicolon or brace is enough.
const colour = `000000; } ${probe("colour", "rgb(0, 0, 220)")} span.zz { color: #000000`
// A style id is written as selector text (p.docx-<id> {).
const style = `X { } ${probe("style", "rgb(0, 140, 0)")} p::before { content: url('${remote("style")}') } p.Y`
const probes = { font: "font name", colour: "colour", style: "style id" }
const sentence = "The Supplier shall deliver the goods within fourteen days."
const styled = "Clause styled by the file's own paragraph style."
const ordinary = "The Customer shall pay each invoice within thirty days."
const docx = zipSync(Object.fromEntries(Object.entries({
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${CT}.styles+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
  "word/styles.xml": `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${xml(font)}" w:hAnsi="${xml(font)}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="${xml(style)}"><w:name w:val="Clause"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`,
  "word/document.xml": `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:color w:val="${xml(colour)}"/></w:rPr><w:t>${sentence}</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t>${ordinary}</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="MS 明朝" w:hAnsi="MS 明朝"/></w:rPr><w:t xml:space="preserve"> Payment is due on receipt.</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="${xml(style)}"/></w:pPr><w:r><w:t>${styled}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
}).map(([name, part]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${part}`)])))

const params = { left: "Supplier agreement.docx", right: "Counterparty response.docx", leftLabel: "Sent", rightLabel: "Returned", ignoreFormatting: true, detectMoves: true, headersFooters: true }
const comparison = {
  left: { name: params.left, label: "Sent", trackedChanges: { count: 0, authors: [] }, pendingProposals: 0 },
  right: { name: params.right, label: "Returned", trackedChanges: { count: 0, authors: [] }, pendingProposals: 0 },
  options: { ignoreFormatting: true, detectMoves: true, headersFooters: true }, attribution: "authors",
  changes: [], untracked: [], blacklineName: "Blackline - Sent vs Returned.docx",
}
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement as h } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentViewer from ${source("components/DocumentViewer.tsx")}
import BlacklineViewer from ${source("components/BlacklineViewer.tsx")}
import TemplateViewer from ${source("components/TemplateViewer.tsx")}
const loadBytes = () => fetch("/template.docx").then(r => r.arrayBuffer()).then(b => new Uint8Array(b))
const viewers = {
  document: () => h(DocumentViewer, { matterId: "material", name: "Counterparty response.docx", onClose: () => {} }),
  blackline: () => h(BlacklineViewer, { matterId: "material", params: ${JSON.stringify(params)}, onClose: () => {} }),
  template: () => h(TemplateViewer, { name: "Counterparty template.docx", loadBytes, onClose: () => {} }),
}
createRoot(document.getElementById("root")).render(viewers[location.hash.slice(1)]())
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "/ingest", VITE_OPENCODE_URL: "/" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = (await Promise.all([
  "styles.css", "styles/primitives.css", "styles/design-system.css", "styles/design-documents.css",
].map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const server = Bun.serve({
  port: 0, hostname: "127.0.0.1", fetch(request) {
    const url = new URL(request.url), { pathname } = url
    if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (pathname.startsWith("/wasm/")) {
      const file = path.join(wasmDir, decodeURIComponent(pathname.slice("/wasm/".length)))
      return file.startsWith(wasmDir + path.sep) && existsSync(file) && statSync(file).isFile() ? new Response(Bun.file(file)) : new Response("Not found", { status: 404 })
    }
    if (pathname === "/") return new Response(`<!doctype html><html lang="en"><meta charset="utf-8"><style>${css}body{margin:0;padding:24px}</style>${Object.entries(probes).map(([name, label], i) => `<div id="${name}-probe" style="position:fixed;left:24px;bottom:${24 + i * 56}px;z-index:2147483647;background:#fff;padding:8px 12px;font:14px system-ui">App page, outside the document (${label})</div>`).join("")}<div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (request.method !== "GET") return Response.json({ error: "Read-only harness" }, { status: 405 })
    const word = () => new Response(docx, { headers: { "content-type": `${CT}.document` } })
    if (pathname === "/template.docx" || /^\/ingest\/matters\/material\/documents\/(redlined|content)$/.test(pathname)) return word()
    if (pathname === "/ingest/matters/material/compare") return url.searchParams.get("format") === "json" ? Response.json(comparison) : word()
    if (pathname === "/ingest/matters/material/redlines") return Response.json([])
    if (pathname === "/ingest/matters/material/documents/amendments") return Response.json(null)
    return Response.json({ error: "Not in this harness" }, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const failures: string[] = []
try {
  for (const viewer of ["document", "blackline", "template"]) {
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 } })
    page.setDefaultTimeout(60_000)
    const remote: string[] = []
    const errors: string[] = []
    page.on("pageerror", error => errors.push(error.message))
    await page.route("**/*", route => {
      if (new URL(route.request().url()).origin === server.url.origin) return route.continue()
      remote.push(route.request().url())
      return route.abort()
    })
    await page.goto(`${server.url.href}#${viewer}`)
    const rendered = page.locator(".docx-render:not(.docx-skeleton)").filter({ hasText: sentence })
    await rendered.first().waitFor()
    await page.waitForLoadState("networkidle")
    const outlines = await page.evaluate(names => names.map(name => [name, getComputedStyle(document.getElementById(`${name}-probe`)!).outlineStyle]), Object.keys(probes))
    const weight = await rendered.first().getByText(styled).evaluate(node => getComputedStyle(node).fontWeight)
    const sheet = await page.evaluate(() => [...document.querySelectorAll(".docx-render style")].map(style => style.textContent).join("\n"))
    if (evidence) await page.screenshot({ path: path.join(evidence, `${viewer}-viewer.png`), fullPage: true })
    const problem = [
      ...errors.map(error => `page error: ${error}`),
      ...outlines.filter(([, outline]) => outline !== "none").map(([name]) => `a rule from the file's ${probes[name as keyof typeof probes]} restyled the page outside the document`),
      ...remote.map(url => `remote request: ${url}`),
      ...["Times New Roman", "MS 明朝"].filter(name => !sheet.includes(name)).map(name => `ordinary font name "${name}" missing from the stylesheet`),
      ...(Number(weight) >= 700 ? [] : [`the file's own paragraph style no longer applies (font-weight ${weight})`]),
    ]
    if (problem.length) failures.push(`${viewer}: ${problem.join("; ")}`)
    console.log(`${problem.length ? "FAIL" : "PASS"} ${viewer} viewer${problem.length ? `: ${problem.join("; ")}` : ""}`)
    await page.close()
  }
  assert.deepEqual(failures, [], "No value in a Word file adds CSS rules to the page in any viewer")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
