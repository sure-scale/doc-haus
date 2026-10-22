// Run from apps/web: bun test/comparison.browser.ts
// Real comparison UI and Docxodus WASM; synthetic comparison JSON/Word bytes
// replace the read-only ingest responses. This does not test the comparison algorithm.
// MATERIAL_EVIDENCE_DIR optionally writes responsive screenshots outside source.
import assert from "node:assert/strict"
import { existsSync, statSync } from "node:fs"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const { zipSync, strToU8 } = (await import(Bun.resolveSync("fflate", path.resolve(import.meta.dir, "../../../dochaus")))) as {
  zipSync: (files: Record<string, Uint8Array>) => Uint8Array; strToU8: (text: string) => Uint8Array
}
const wasmDir = path.resolve(process.env.DOCXODUS_WASM_DIR ?? path.join(path.dirname(fileURLToPath(import.meta.resolve("docxodus"))), "wasm"))
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-comparison-"))
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
const evidence = process.env.MATERIAL_EVIDENCE_DIR
if (evidence) await mkdir(evidence, { recursive: true })

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const blackline = zipSync(Object.fromEntries(Object.entries({
  "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/document.xml": `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Supplier agreement</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">The Supplier shall deliver the goods within </w:t></w:r><w:del w:id="1" w:author="Not tracked (Returned)" w:date="2026-10-20T09:00:00Z"><w:r><w:delText>fourteen</w:delText></w:r></w:del><w:ins w:id="2" w:author="Not tracked (Returned)" w:date="2026-10-20T09:00:00Z"><w:r><w:t>seven</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> days.</w:t></w:r></w:p><w:p><w:r><w:t>Either party may end this agreement on thirty days notice.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
}).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)])))
const params = { left: "Supplier agreement.docx", right: "Counterparty response.docx", leftLabel: "Sent", rightLabel: "Returned", ignoreFormatting: true, detectMoves: true, headersFooters: true }
const hunk = { kind: "modified", location: { scope: "body", paragraph: 2, side: "right", heading: "Delivery" }, before: "The Supplier shall deliver the goods within fourteen days.", after: "The Supplier shall deliver the goods within seven days.", untracked: true }
const fixture = {
  left: { name: params.left, label: "Sent", trackedChanges: { count: 0, authors: [] }, pendingProposals: 0 },
  right: { name: params.right, label: "Returned", trackedChanges: { count: 0, authors: [] }, pendingProposals: 0 },
  options: { ignoreFormatting: true, detectMoves: true, headersFooters: true }, attribution: "authors",
  changes: [{ id: "C1", ...hunk }], untracked: [{ id: "U1", ...hunk }], blacklineName: "Blackline - Sent vs Returned.docx",
}
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement as h, Fragment, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import CompareDialog from ${source("components/CompareDialog.tsx")}
import BlacklineViewer from ${source("components/BlacklineViewer.tsx")}
import ModalHeader from ${source("components/ModalHeader.tsx")}

const params = ${JSON.stringify(params)}
const documents = [
  { name: params.left, created_at: 10 },
  { name: params.right, created_at: 30 },
  { name: "Previous draft.docx", created_at: 20 },
  { name: "Source evidence.pdf", created_at: 40 },
]
function Documents({ onClose }) {
  const [viewer, setViewer] = useState(null)
  return h("div", { className: "viewer-overlay" },
    h("section", { className: "picker-panel", "aria-label": "Documents" },
      h(ModalHeader, { title: "Documents", onClose }),
      h("div", { className: "picker-body" },
        h("button", { type: "button", onClick: () => setViewer(params) }, "Open synthetic blackline"),
        h("button", { type: "button", onClick: () => setViewer({ ...params, right: "Unavailable.docx" }) }, "Open unavailable blackline")),
      viewer && h(BlacklineViewer, { matterId: "material", params: viewer, onClose: () => setViewer(null) })))
}
function Harness() {
  const [compare, setCompare] = useState(false)
  const [documentsOpen, setDocumentsOpen] = useState(false)
  const [result, setResult] = useState(null)
  return h(Fragment, null,
    h("button", { type: "button", onClick: () => setCompare(true) }, "Compare versions"),
    h("button", { type: "button", onClick: () => setDocumentsOpen(true) }, "Open documents"),
    h("output", { "data-testid": "compare-params" }, result ? JSON.stringify(result) : ""),
    compare && h(CompareDialog, { matterId: "material", name: params.left, documents,
      onUploaded: () => { throw new Error("Unexpected upload") },
      onCompare: value => { setResult(value); setCompare(false) }, onClose: () => setCompare(false) }),
    documentsOpen && h(Documents, { onClose: () => setDocumentsOpen(false) }))
}
createRoot(document.getElementById("root")).render(h(Harness))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "/ingest", VITE_OPENCODE_URL: "/" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = (await Promise.all([
  "styles.css", "narration.css", "narration-live.css", "styles/primitives.css", "styles/upload.css",
  "styles/shell.css", "styles/chat-log.css", "styles/drafting.css", "styles/drafting-items.css",
  "styles/run-summary.css", "styles/motion.css", "styles/design-system.css", "styles/design-shell.css",
  "styles/design-documents.css", "styles/design-conversation.css", "styles/design-library.css", "styles/design-review.css",
].map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")
const calls: string[] = []
const server = Bun.serve({
  port: 0, hostname: "127.0.0.1", fetch(request) {
    const url = new URL(request.url), { pathname } = url
    if (pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (pathname.startsWith("/wasm/")) {
      const file = path.join(wasmDir, decodeURIComponent(pathname.slice("/wasm/".length)))
      return file.startsWith(wasmDir + path.sep) && existsSync(file) && statSync(file).isFile() ? new Response(Bun.file(file)) : new Response("Not found", { status: 404 })
    }
    if (pathname === "/") return new Response(`<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><style>${css}body{margin:0;padding:24px}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    calls.push(`${request.method} ${pathname}?${url.searchParams}`)
    if (request.method === "GET" && pathname === "/ingest/matters/material/compare") {
      if (url.searchParams.get("right") === "Unavailable.docx") return Response.json({ error: "missing_document", message: "The later version is no longer available in this matter." }, { status: 404 })
      if (url.searchParams.get("format") === "json") return Response.json(fixture)
      return new Response(blackline, { headers: { "content-type": `${CT}.document` } })
    }
    return Response.json({ error: "Unexpected request" }, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const errors: string[] = []
try {
  const page = await browser.newPage({ reducedMotion: "reduce" })
  page.setDefaultTimeout(20_000)
  page.on("pageerror", error => errors.push(error.message))
  await page.route("**/*", route => {
    if (new URL(route.request().url()).origin === server.url.origin) return route.continue()
    errors.push(`Unexpected external request: ${route.request().url()}`)
    return route.abort()
  })
  await page.goto(server.url.href)
  const withinViewport = async (element: import("@playwright/test").Locator, width: number, height: number) => {
    const bounds = await element.boundingBox()
    assert.ok(bounds && bounds.x >= -1 && bounds.y >= -1 && bounds.x + bounds.width <= width + 1 && bounds.y + bounds.height <= height + 1,
      `modal stays inside ${width}x${height}: ${JSON.stringify(bounds)}`)
    assert.equal(await element.evaluate(node => node.scrollWidth > node.clientWidth + 1), false, "modal has no horizontal overflow")
  }
  for (const width of [375, 768, 1440]) {
    const height = width === 375 ? 812 : 900
    await page.setViewportSize({ width, height })
    await page.getByRole("button", { name: "Compare versions", exact: true }).click()
    const picker = page.getByRole("dialog", { name: "Compare versions", exact: true })
    await expect(picker).toBeVisible()
    await withinViewport(picker, width, height)
    await expect(picker.getByRole("button", { name: "Compare", exact: true })).toBeDisabled()
    const choices = await picker.getByRole("combobox", { name: "Other version" }).locator("option").allTextContents()
    assert.deepEqual(choices, ["Choose a file…", "Counterparty response.docx", "Previous draft.docx"], "only other Word files appear, newest first")
    await picker.getByRole("combobox", { name: "Other version" }).selectOption(params.right)
    await picker.locator("#compare-left-label").fill("  Sent version  ")
    await picker.locator("#compare-right-label").fill(" Returned version ")
    await picker.getByRole("button", { name: "Swap", exact: true }).click()
    await expect(picker.locator("#compare-left-label")).toHaveValue(" Returned version ")
    await expect(picker.locator("#compare-right-label")).toHaveValue("  Sent version  ")
    for (const name of ["Ignore formatting-only changes", "Detect moved text", "Include headers and footers"]) await expect(picker.getByRole("checkbox", { name })).toBeChecked()
    await picker.getByRole("checkbox", { name: "Ignore formatting-only changes" }).uncheck()
    await picker.getByRole("checkbox", { name: "Include headers and footers" }).uncheck()
    const compare = picker.getByRole("button", { name: "Compare", exact: true })
    await compare.focus()
    await page.keyboard.press("Tab")
    await expect(picker.getByRole("button", { name: "Close", exact: true })).toBeFocused()
    await page.keyboard.press("Shift+Tab")
    await expect(compare).toBeFocused()
    if (evidence) await page.screenshot({ path: path.join(evidence, `compare-picker-${width}.png`) })
    await page.keyboard.press("Enter")
    await expect(picker).toHaveCount(0)
    assert.deepEqual(JSON.parse((await page.getByTestId("compare-params").textContent())!), {
      left: params.right, right: params.left, leftLabel: "Returned version", rightLabel: "Sent version",
      ignoreFormatting: false, detectMoves: true, headersFooters: false,
    }, "selection, swapped labels, trimmed labels and comparison options preserve the expected payload")
    await expect(page.getByRole("button", { name: "Compare versions", exact: true })).toBeFocused()

    await page.getByRole("button", { name: "Open documents", exact: true }).click()
    const documents = page.getByRole("dialog", { name: "Documents", exact: true })
    await documents.getByRole("button", { name: "Open synthetic blackline", exact: true }).click()
    const viewer = page.getByRole("dialog", { name: fixture.blacklineName, exact: true })
    await expect(viewer).toBeVisible()
    const rendered = viewer.locator(".docx-render:not(.docx-skeleton)")
    await expect(rendered).toContainText("The Supplier shall deliver the goods within")
    await expect(rendered.locator("ins.rev-ins")).toContainText("seven")
    await expect(rendered.locator("del.rev-del")).toContainText("fourteen")
    const readingWidth = await rendered.evaluate(element => {
      const style = getComputedStyle(element)
      return element.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight)
    })
    assert.ok(readingWidth >= Math.min(320, width - 80), `comparison leaves readable document text width at ${width}px: ${readingWidth}px`)
    await expect(viewer.getByRole("complementary", { name: "Changes", exact: true })).toContainText("Edits made without tracked changes")
    await expect(viewer.getByRole("button", { name: "Download blackline", exact: true })).toBeEnabled()
    await withinViewport(viewer, width, height)
    if (evidence) await page.screenshot({ path: path.join(evidence, `blackline-fixture-${width}.png`) })
    await expect(viewer.getByRole("button", { name: "Close", exact: true })).toBeFocused()
    const lastChange = viewer.getByRole("button", { name: /Delivery/ }).last()
    await lastChange.focus()
    await page.keyboard.press("Tab")
    await expect(viewer.getByRole("button", { name: "Download blackline", exact: true })).toBeFocused()
    await page.keyboard.press("Shift+Tab")
    await expect(lastChange).toBeFocused()
    await documents.getByRole("button", { name: "Open unavailable blackline", exact: true }).evaluate(element => (element as HTMLButtonElement).focus())
    await expect(viewer.getByRole("button", { name: "Download blackline", exact: true })).toBeFocused()
    await page.keyboard.press("Escape")
    await expect(viewer).toHaveCount(0)
    await expect(documents.getByRole("button", { name: "Open synthetic blackline", exact: true })).toBeFocused()

    await documents.getByRole("button", { name: "Open unavailable blackline", exact: true }).click()
    const refused = page.getByRole("dialog", { name: "Blackline - Supplier agreement vs Unavailable.docx", exact: true })
    await expect(refused).toContainText("The later version is no longer available in this matter.")
    await expect(refused.getByRole("button", { name: "Download blackline", exact: true })).toBeDisabled()
    await withinViewport(refused, width, height)
    await page.keyboard.press("Tab")
    await expect(refused.getByRole("button", { name: "Close", exact: true })).toBeFocused()
    if (evidence) await page.screenshot({ path: path.join(evidence, `blackline-refusal-${width}.png`) })
    await page.keyboard.press("Escape")
    await expect(refused).toHaveCount(0)
    await expect(documents.getByRole("button", { name: "Open unavailable blackline", exact: true })).toBeFocused()
    await page.keyboard.press("Escape")
    await expect(documents).toHaveCount(0)
    await expect(page.getByRole("button", { name: "Open documents", exact: true })).toBeFocused()
    console.log(`PASS ${width}px comparison picker payload/keyboard/bounds, synthetic Word blackline render/trap/restore, and API refusal state.`)
  }
  assert.ok(calls.length === 12 && calls.every(call => call.startsWith("GET /ingest/matters/material/compare?")), `only twelve expected read-only comparison responses: ${calls.join("\n")}`)
  assert.deepEqual(errors, [], "no browser errors or external requests")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
