// Run from apps/web: bun test/chronology-intake.browser.ts
// Actual upload UI/XHR; only ingest responses are fixtures. This does not test
// backend DOCX-origin recognition or extraction.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Document } from "../src/api/ingest/documents"

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test") as typeof import("@playwright/test")
const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-chronology-intake-"))
const entry = path.join(temporary, "intake.tsx")
await Bun.write(entry, `
import { createElement, useEffect, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentUpload from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentUpload.tsx"))}
function App() {
  const [documents, setDocuments] = useState([])
  useEffect(() => { fetch("/ingest/matters/m/documents").then(r => r.json()).then(setDocuments) }, [])
  return createElement(DocumentUpload, {matterId: "m", documents, onView: () => {},
    onUploaded: () => fetch("/ingest/matters/m/documents").then(r => r.json()).then(setDocuments)})
}
createRoot(document.getElementById("root")).render(createElement(App))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "/ingest", VITE_OPENCODE_URL: "/" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const javascript = await build.outputs.find(output => output.path.endsWith(".js"))!.text()
const requests: { name: string; bytes: string; existing: FormDataEntryValue | null; origin: FormDataEntryValue | null }[] = []
const documents: Document[] = []
const unknown: string[] = []
const perFile = new Map<string, number>()
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", async fetch(request) {
  const route = `${request.method} ${new URL(request.url).pathname}`
  if (route === "GET /") return new Response('<html><meta charset="utf-8"><style>body{font-family:sans-serif;max-width:440px}.doc-row-meta{display:flex;flex-wrap:wrap;gap:8px}.upload-actions{display:flex;gap:8px}</style><div id="root"></div><script src="/intake.js" type="module"></script></html>', { headers: { "content-type": "text/html" } })
  if (route === "GET /intake.js") return new Response(javascript, { headers: { "content-type": "text/javascript" } })
  if (route === "GET /ingest/matters/m/documents") return Response.json(documents)
  if (route === "POST /ingest/matters/m/documents") {
    const form = await request.formData()
    const file = form.get("file") as File
    const sent = { name: file.name, bytes: await file.text(), existing: form.get("existing"), origin: form.get("chronology_origin") }
    requests.push(sent)
    const count = (perFile.get(file.name) ?? 0) + 1
    perFile.set(file.name, count)
    const conflict = (code: string) => Response.json({ code, name: file.name, error: "Fixture conflict.", original_name: "Chronology.docx" }, { status: 409 })
    if (file.name === "Collision-first.docx" && count === 1) return conflict("source_exists")
    if (!sent.origin) return conflict("derived_chronology")
    if (!sent.existing) return conflict("source_exists")
    const name = sent.existing === "keep-both" ? file.name.replace(".docx", " (2).docx") : file.name
    documents.push({ id: documents.length + 1, name, doc_path: `/m/${name}`, created_at: Date.now(), role: "source", source_provenance: "supplied-chronology" })
    return Response.json({ name, docPath: `/m/${name}`, sections: 1, chunks: 1, injection: null, source_provenance: "supplied-chronology" })
  }
  if (route === "GET /favicon.ico") return new Response(null, { status: 204 })
  unknown.push(route)
  return new Response("Unexpected fixture request", { status: 404 })
} })
let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined
try {
  browser = await chromium.launch({ headless: true, channel: process.env.PLAYWRIGHT_CHANNEL })
  const page = await browser.newPage({ viewport: { width: 440, height: 800 } })
  const errors: string[] = []
  page.on("pageerror", error => errors.push(error.message))
  await page.goto(server.url.href)
  const chooseFile = (name: string) => page.locator('input[type="file"]').setInputFiles({ name, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from("edited synthetic chronology") })

  await chooseFile("Warning-first.docx")
  await page.getByRole("button", { name: "Add as supplied chronology", exact: true }).waitFor()
  assert.equal(requests.length, 1)
  assert.equal(requests[0]!.origin, null)
  assert.match(await page.locator(".upload-status").innerText(), /not independent corroboration/)
  await page.getByRole("button", { name: "Add as supplied chronology", exact: true }).click()
  await page.getByRole("button", { name: "Keep both", exact: true }).waitFor()
  assert.equal(
    await page.locator('.docs-drawer-body[aria-label="Documents"]').evaluate(element => element === document.activeElement),
    true,
    "focus stays in Documents when acknowledgement changes the row",
  )
  assert.equal(requests[1]!.origin, "supplied-chronology")
  assert.equal(requests[1]!.existing, null)
  await page.getByRole("button", { name: "Keep both", exact: true }).click()
  await page.getByRole("button", { name: "Warning-first (2).docx", exact: true }).waitFor()
  assert.equal(
    await page.locator('.docs-drawer-body[aria-label="Documents"]').evaluate(element => element === document.activeElement),
    true,
    "focus stays in Documents after keeping both",
  )
  assert.deepEqual(requests[2], { name: "Warning-first.docx", bytes: "edited synthetic chronology", existing: "keep-both", origin: "supplied-chronology" })
  assert.match(await page.locator(".docs-timeline").innerText(), /Supplied chronology/)

  await chooseFile("Collision-first.docx")
  await page.getByRole("button", { name: "Save as new version", exact: true }).click()
  await page.getByRole("button", { name: "Add as supplied chronology", exact: true }).waitFor()
  assert.equal(requests[4]!.existing, "replace")
  assert.equal(requests[4]!.origin, null)
  await page.getByRole("button", { name: "Add as supplied chronology", exact: true }).click()
  await page.getByRole("button", { name: "Collision-first.docx", exact: true }).waitFor()
  assert.deepEqual(requests[5], { name: "Collision-first.docx", bytes: "edited synthetic chronology", existing: "replace", origin: "supplied-chronology" })

  await chooseFile("Cancel.docx")
  await page.getByRole("button", { name: "Add as supplied chronology", exact: true }).waitFor()
  const count = requests.length
  await page.locator(".upload-row.is-chronology").getByRole("button", { name: "Cancel", exact: true }).click()
  assert.equal(await page.locator(".upload-row.is-chronology").count(), 0)
  assert.equal(requests.length, count)

  await page.reload()
  await page.getByRole("button", { name: "Warning-first (2).docx", exact: true }).waitFor()
  assert.match(await page.locator(".docs-timeline").innerText(), /Supplied chronology/)
  assert.equal(documents.length, 2)
  assert.equal(unknown.length, 0, unknown.join("\n"))
  assert.deepEqual(errors, [])
  console.log("Chronology intake browser PASS: warning/choice in both orders, same-file retries, cancel without retry, persisted renamed source classification. Ingest responses mocked.")
} finally {
  await browser?.close()
  server.stop(true)
  await rm(temporary, { recursive: true, force: true })
}
