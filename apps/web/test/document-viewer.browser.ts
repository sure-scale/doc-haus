// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/document-viewer.browser.ts
// Mounts the real DocumentViewer on the Docxodus WASM runtime served as the web build serves /wasm/; only ingest HTTP is mocked.
// A Word package may omit its styles and settings parts (#216); each variant must render its text read-only.
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
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-document-viewer-"))

// Synthetic two-paragraph agreement; the optional parts are added only when kept.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const sentences = ["The Supplier shall deliver the goods within fourteen days.", "Either party may end this agreement on thirty days notice."]
const optional = {
  styles: `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
  settings: `<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/></w:settings>`,
}
type Part = keyof typeof optional
function docx(keep: Part[]) {
  const body = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>The Supplier</w:t></w:r><w:r><w:t xml:space="preserve"> shall deliver the goods within fourteen days.</w:t></w:r></w:p><w:p><w:r><w:t>${sentences[1]}</w:t></w:r></w:p>`
  const files: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/>${keep.map(part => `<Override PartName="/word/${part}.xml" ContentType="${CT}.${part}+xml"/>`).join("")}</Types>`,
    "_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${keep.map((part, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/${part}" Target="${part}.xml"/>`).join("")}</Relationships>`,
    "word/document.xml": `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    ...Object.fromEntries(keep.map(part => [`word/${part}.xml`, optional[part]])),
  }
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, xml]) => [name, strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`)])))
}
// The complete package is the control: it shows the harness renders a normal file.
const variants: Record<string, Part[]> = { "complete.docx": ["styles", "settings"], "no-styles.docx": ["settings"], "no-settings.docx": ["styles"], "no-styles-or-settings.docx": [] }
const files = new Map(Object.entries(variants).map(([name, keep]) => [name, docx(keep)]))

const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentViewer from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentViewer.tsx"))}
createRoot(document.getElementById("root")).render(createElement(DocumentViewer, { matterId: "matter-viewer", name: decodeURIComponent(location.hash.slice(1)), onClose: () => {} }))
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
try {
  console.log(`Docxodus runtime: ${wasmDir}`)
  for (const name of files.keys()) {
    const page = await browser.newPage()
    page.setDefaultTimeout(60_000)
    const calls: string[] = []
    const writes: string[] = []
    // The viewer only reads: the redlined bytes (clean when nothing is pending), the change list and the stored file.
    await page.route("http://viewer.invalid/**", route => {
      const request = route.request()
      const url = new URL(request.url())
      calls.push(`${request.method()} ${url.pathname}`)
      if (request.method() !== "GET") {
        writes.push(`${request.method()} ${url.pathname}`)
        return route.fulfill({ status: 405 })
      }
      const bytes = url.searchParams.get("name") === name ? files.get(name) : undefined
      if (bytes && /^\/matters\/matter-viewer\/documents\/(redlined|content)$/.test(url.pathname))
        return route.fulfill({ status: 200, contentType: `${CT}.document`, body: Buffer.from(bytes) })
      if (url.pathname === "/matters/matter-viewer/redlines") return route.fulfill({ status: 200, contentType: "application/json", body: "[]" })
      return route.fulfill({ status: 404, contentType: "application/json", body: JSON.stringify({ error: `Unexpected ${url.pathname}` }) })
    })
    await page.goto(`${server.url.href}#${encodeURIComponent(name)}`)
    const rendered = page.locator(".docx-render:not(.docx-skeleton)")
    const error = page.locator(".viewer-doc > p.muted")
    await rendered.or(error).first().waitFor()
    const shown = (await rendered.count()) ? (await rendered.textContent())!.replace(/\s+/g, " ") : ""
    const problem = [
      ...(await error.allTextContents()),
      ...sentences.filter(sentence => !shown.includes(sentence)).map(sentence => `missing text "${sentence}"`),
      ...writes.map(write => `write sent: ${write}`),
      ...(calls.includes("GET /matters/matter-viewer/documents/redlined") ? [] : ["redlined bytes never requested"]),
    ]
    if (problem.length) failures.push(`${name}: ${problem.join("; ")}`)
    console.log(`${problem.length ? "FAIL" : "PASS"} ${name}: ${problem.length ? problem.join("; ") : `rendered "${sentences[0]}" read-only (${calls.join(", ")})`}`)
    await page.close()
  }
  assert.deepEqual(failures, [], "Every Word package, with or without its styles and settings parts, renders its text")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
