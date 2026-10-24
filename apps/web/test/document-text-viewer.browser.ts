// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/document-text-viewer.browser.ts
// Mounts the real DocumentViewer in Chromium with only ingest HTTP mocked. An email or spreadsheet opens as the
// text ingest extracted, read-only, with a cited excerpt highlighted and focused; it never takes the Word path.
// A Word file is the control: it still renders through Docxodus from the redlined bytes. An older kept version
// of it open in the viewer carries no citation highlight or status: the citation is to the current text (#268).
// Set PLAYWRIGHT_CHROMIUM_PATH=<chromium> when Playwright's own browser is missing, DOCXODUS_WASM_DIR=<wasm dir> to try another runtime build.
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
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-document-text-viewer-"))

// Synthetic extracted text in the shapes ingest produces: email headers above the body, one CSV block per sheet.
const texts: Record<string, string> = {
  "Client instructions.eml": [
    "Subject: Refurbishment invoices",
    "From: Synthetic Client <client@example.invalid>",
    "To: Synthetic Counsel <counsel@example.invalid>",
    "Date: Mon, 12 Oct 2026 09:00:00 +0000",
    "",
    "Please chase the three unpaid invoices.",
    "The works were signed off on",
    "2 March 2026 by the site manager.",
  ].join("\n"),
  "Cap table.xlsx": ["## Holders", "Holder,Class,Shares", "Synthetic Founder Ltd,Ordinary,6000", "Synthetic Fund LP,Preferred,2500"].join("\n"),
}
type Case = { name: string; excerpt?: string; shown: string[]; marked?: string }
const cases: Case[] = [
  // The excerpt's line break differs from the text's; the citation match is whitespace-tolerant.
  { name: "Client instructions.eml", excerpt: "The works were signed off on 2 March 2026 by the site manager.", shown: ["Subject: Refurbishment invoices", "Please chase the three unpaid invoices."], marked: "The works were signed off on\n2 March 2026 by the site manager." },
  { name: "Cap table.xlsx", excerpt: "Synthetic Fund LP,Preferred,2500", shown: ["## Holders", "Synthetic Founder Ltd,Ordinary,6000"], marked: "Synthetic Fund LP,Preferred,2500" },
  { name: "Client instructions.eml", shown: ["Please chase the three unpaid invoices."] },
]

// Control: a minimal synthetic Word package still takes the Docxodus path.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const sentence = "The Supplier shall deliver the goods within fourteen days."
const wordFile = (...paragraphs: string[]) => zipSync(Object.fromEntries(Object.entries({
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/document.xml": `<w:document xmlns:w="${W}"><w:body>${paragraphs.map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join("")}</w:body></w:document>`,
}).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)])))
const docx = wordFile(sentence)
// An older kept version that also holds the cited sentence.
const earlier = wordFile("Earlier version of the terms.", sentence)
const kept = (number: number, current: boolean) => ({ id: `v-${number}`, number, sha256: String(number).repeat(64), ext: ".docx", bytes: 1000, parent: number > 1 ? `v-${number - 1}` : null,
  derived_from: null, origin: number > 1 ? "replace" : "upload", actor: "unknown", run: null, label: null, created_at: Date.UTC(2026, 9, 24, 9 + number), current, removed: null })

const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentViewer from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentViewer.tsx"))}
const params = new URLSearchParams(location.search)
const excerpt = params.get("excerpt")
createRoot(document.getElementById("root")).render(createElement(DocumentViewer, {
  matterId: "matter-viewer", name: params.get("name"), citation: excerpt === null ? undefined : { section: "cited passage", excerpt }, onClose: () => {},
}))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "http://viewer.invalid" }) } })
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
const failures: string[] = []
const normalize = (value: string) => value.replace(/\s+/g, " ").trim()

async function open(name: string, excerpt?: string) {
  const page = await browser.newPage()
  page.setDefaultTimeout(60_000)
  const calls: string[] = []
  await page.route("http://viewer.invalid/**", route => {
    const request = route.request()
    const url = new URL(request.url())
    calls.push(`${request.method()} ${url.pathname}`)
    if (request.method() !== "GET") return route.fulfill({ status: 405 })
    const requested = url.searchParams.get("name")
    if (url.pathname === "/matters/matter-viewer/documents/text" && requested && requested in texts)
      return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ text: texts[requested], injection: null }) })
    if (requested === "terms.docx" && /^\/matters\/matter-viewer\/documents\/(redlined|content)$/.test(url.pathname))
      return route.fulfill({ status: 200, contentType: `${CT}.document`, body: Buffer.from(docx) })
    if (url.pathname === "/matters/matter-viewer/redlines") return route.fulfill({ status: 200, contentType: "application/json", body: "[]" })
    if (url.pathname === "/matters/matter-viewer/documents/amendments") return route.fulfill({ status: 200, contentType: "application/json", body: "null" })
    if (url.pathname === "/matters/matter-viewer/documents/versions") {
      const versions = requested === "terms.docx" ? [kept(2, true), kept(1, false)] : []
      return route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ name: requested, document: null, versions, events: [], working_version: versions.length ? "v-2" : null }) })
    }
    if (url.pathname === "/matters/matter-viewer/documents/versions/v-1/content") return route.fulfill({ status: 200, contentType: `${CT}.document`, body: Buffer.from(earlier) })
    return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: `Unexpected ${url.pathname}` }) })
  })
  const query = new URLSearchParams({ name, ...(excerpt === undefined ? {} : { excerpt }) })
  await page.goto(`${server.url.href}?${query}`)
  return { page, calls }
}

try {
  for (const { name, excerpt, shown, marked } of cases) {
    const label = `${name}${excerpt ? " (cited)" : ""}`
    const { page, calls } = await open(name, excerpt)
    const text = page.locator("pre.document-text")
    const error = page.locator(".viewer-doc > p.muted")
    await text.or(error).first().waitFor()
    // Let a wrong path settle into its error before judging.
    await page.waitForTimeout(500)
    const body = (await text.count()) ? (await text.textContent())! : ""
    const mark = page.locator("pre.document-text mark.citation-passage")
    const status = (await page.locator('.viewer-doc > p[role="status"]').allTextContents()).join(" ")
    const problem = [
      ...(await error.allTextContents()).map(message => `error shown: ${message}`),
      ...shown.filter(line => !body.includes(line)).map(line => `missing text "${line}"`),
      ...calls.filter(call => call !== "GET /matters/matter-viewer/documents/text" && call !== "GET /matters/matter-viewer/documents/versions").map(call => `unexpected request: ${call}`),
      ...((await page.getByRole("button", { name: /Request edits|Convert to DOCX|Download/ }).count()) ? ["Word or PDF actions offered"] : []),
    ]
    if (marked) {
      const highlighted = (await mark.allTextContents()).join("")
      if (highlighted !== marked) problem.push(`highlighted ${JSON.stringify(highlighted)}, expected ${JSON.stringify(marked)}`)
      if (normalize(highlighted) !== normalize(excerpt!)) problem.push("highlight differs from the excerpt")
      // Count first: evaluate on a missing mark would wait out the page timeout.
      if (!(await mark.count()) || !(await mark.first().evaluate(node => node === document.activeElement))) problem.push("highlight not focused")
      if (!status.includes("The cited excerpt is highlighted.")) problem.push(`status ${JSON.stringify(status)}`)
    } else if (await mark.count()) problem.push("highlight without a citation")
    if (problem.length) failures.push(`${label}: ${problem.join("; ")}`)
    console.log(`${problem.length ? "FAIL" : "PASS"} ${label}: ${problem.length ? problem.join("; ") : `text view read-only (${calls.join(", ")})${marked ? `, highlighted ${JSON.stringify(normalize(marked))}` : ""}`}`)
    await page.close()
  }

  const { page, calls } = await open("terms.docx")
  const rendered = page.locator(".docx-render:not(.docx-skeleton):not(.document-text)")
  const error = page.locator(".viewer-doc > p.muted")
  await rendered.or(error).first().waitFor()
  const problem = [
    ...(await error.allTextContents()).map(message => `error shown: ${message}`),
    ...((await rendered.count()) && normalize((await rendered.textContent())!).includes(sentence) ? [] : [`missing text "${sentence}"`]),
    ...(calls.includes("GET /matters/matter-viewer/documents/redlined") ? [] : ["redlined bytes never requested"]),
    ...(calls.some(call => call.endsWith("/documents/text")) ? ["Word file read as extracted text"] : []),
    ...((await page.getByRole("button", { name: "Request edits" }).count()) ? [] : ["Request edits missing"]),
  ]
  if (problem.length) failures.push(`terms.docx: ${problem.join("; ")}`)
  console.log(`${problem.length ? "FAIL" : "PASS"} terms.docx (control): ${problem.length ? problem.join("; ") : `Word path unchanged (${calls.join(", ")})`}`)
  await page.close()

  {
    const { page } = await open("terms.docx", sentence)
    const mark = page.locator(".docx-render mark.citation-passage")
    const status = page.locator('.viewer-doc > p[role="status"]')
    const problem: string[] = []
    await mark.first().waitFor()
    await page.getByRole("button", { name: /^Versions/ }).click()
    await page.locator(".version-row").filter({ hasText: "Version 1" }).getByRole("button", { name: /^View/ }).click()
    await page.locator(".docx-render:not(.docx-skeleton)").filter({ hasText: "Earlier version of the terms." }).waitFor()
    // Let the citation pass run over the older text before judging.
    await page.waitForTimeout(500)
    if (await mark.count()) problem.push("cited excerpt highlighted in an older version")
    const said = (await status.allTextContents()).join(" ")
    if (said.includes("Source section")) problem.push(`citation status shown over an older version: ${JSON.stringify(said)}`)
    await page.getByRole("button", { name: "Back to current version" }).click()
    await page.locator(".docx-render:not(.docx-skeleton)").filter({ hasNotText: "Earlier version of the terms." }).waitFor()
    if (!(await mark.count())) problem.push("cited excerpt not highlighted again in the current version")
    if (problem.length) failures.push(`terms.docx (older version): ${problem.join("; ")}`)
    console.log(`${problem.length ? "FAIL" : "PASS"} terms.docx (cited, older version): ${problem.length ? problem.join("; ") : "no citation over the older version; back on the current one it is highlighted"}`)
    await page.close()
  }
  assert.deepEqual(failures, [], "Emails and spreadsheets open as highlighted read-only text; Word files keep the Word viewer")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
