// Run from apps/web: bun test/chronology-intake.browser.ts
// Actual upload UI/XHR; only ingest responses are fixtures. This does not test
// backend DOCX-origin recognition or extraction.
import assert from "node:assert/strict"
import { mkdir, mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Document } from "../src/api/ingest/documents"

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test") as typeof import("@playwright/test")
const temporary = await mkdtemp(path.join(os.tmpdir(), "dochaus-chronology-intake-"))
const captures = process.env.INTAKE_SCREENSHOTS
if (captures) await mkdir(captures, { recursive: true })
const entry = path.join(temporary, "intake.tsx")
const styles = [...(await readFile(path.resolve(import.meta.dir, "../src/main.tsx"), "utf8")).matchAll(/^import "([^"]+\.css)"/gm)]
  .map(([, source]) => `import ${JSON.stringify(source!.startsWith(".") ? path.resolve(import.meta.dir, "../src", source!) : Bun.resolveSync(source!, path.resolve(import.meta.dir, "..")))}`)
await Bun.write(entry, `
import { createElement, useEffect, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentUpload from ${JSON.stringify(path.resolve(import.meta.dir, "../src/components/DocumentUpload.tsx"))}
${styles.join("\n")}
function App() {
  const [documents, setDocuments] = useState([])
  useEffect(() => { fetch("/ingest/matters/m/documents").then(r => r.json()).then(setDocuments) }, [])
  return createElement("aside", {className: "docs-drawer", style: {height: "100dvh"}}, createElement(DocumentUpload, {matterId: "m", documents, onView: () => {},
    onUploaded: () => fetch("/ingest/matters/m/documents").then(r => r.json()).then(setDocuments)}))
}
createRoot(document.getElementById("root")).render(createElement(App))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "/ingest", VITE_OPENCODE_URL: "/" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const assets = new Map(build.outputs.map(output => [`/${path.basename(output.path)}`, output]))
const requests: { name: string; bytes: string; existing: FormDataEntryValue | null; origin: FormDataEntryValue | null }[] = []
const documents: Document[] = []
const unknown: string[] = []
const perFile = new Map<string, number>()
let releaseSlow: (() => void) | undefined
const server = Bun.serve({ port: 0, hostname: "127.0.0.1", async fetch(request) {
  const route = `${request.method} ${new URL(request.url).pathname}`
  if (route === "GET /") return new Response('<html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/intake.css"><div id="root"></div><script src="/intake.js" type="module"></script></html>', { headers: { "content-type": "text/html" } })
  const asset = assets.get(new URL(request.url).pathname)
  if (asset) return new Response(asset)
  if (route === "GET /ingest/matters/m/documents") return Response.json(documents)
  if (route === "POST /ingest/matters/m/documents") {
    const form = await request.formData()
    const file = form.get("file") as File
    const sent = { name: file.name, bytes: await file.text(), existing: form.get("existing"), origin: form.get("chronology_origin") }
    requests.push(sent)
    const count = (perFile.get(file.name) ?? 0) + 1
    perFile.set(file.name, count)
    const conflict = (code: string) => Response.json({ code, name: file.name, error: "Fixture conflict.", original_name: "Chronology.docx" }, { status: 409 })
    if (file.name === "Generated-name.docx") return conflict("deliverable_exists")
    if (["Collision-first.docx", "Failure.docx"].includes(file.name) && count === 1) return conflict("source_exists")
    if (file.name !== "Ordinary.pdf") {
      if (!sent.origin) return conflict("derived_chronology")
      if (file.name === "Recognized-generated.docx") return conflict("deliverable_exists")
      if (file.name === "Slow.docx") await new Promise<void>(resolve => { releaseSlow = resolve })
      else if (!sent.existing) return conflict("source_exists")
      if (file.name === "Failure.docx" && count === 3) return Response.json({ error: "Temporary reader failure." }, { status: 503 })
    }
    const name = sent.existing === "keep-both" ? file.name.replace(".docx", " (2).docx") : file.name
    const provenance = sent.origin ? { source_provenance: "supplied-chronology" as const } : {}
    documents.push({ id: documents.length + 1, name, doc_path: `/m/${name}`, created_at: Date.now(), role: "source", ...provenance })
    return Response.json({ name, docPath: `/m/${name}`, sections: 1, chunks: 1, injection: null, ...provenance })
  }
  if (route === "GET /favicon.ico") return new Response(null, { status: 204 })
  unknown.push(route)
  return new Response("Unexpected fixture request", { status: 404 })
} })
let browser: Awaited<ReturnType<typeof chromium.launch>> | undefined
try {
  browser = await chromium.launch({ headless: true, channel: process.env.PLAYWRIGHT_CHANNEL })
  const page = await browser.newPage({ viewport: { width: 375, height: 900 }, isMobile: true, hasTouch: true })
  const errors: string[] = []
  page.on("pageerror", error => errors.push(error.message))
  await page.goto(server.url.href)
  const chooseFile = (name: string) => page.locator('input[type="file"]').setInputFiles({ name, mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: Buffer.from("edited synthetic chronology") })
  const capture = async (name: string) => {
    if (!captures) return
    await page.evaluate(async () => { await document.fonts.ready })
    await page.screenshot({ path: path.join(captures, name), animations: "disabled" })
  }

  await chooseFile("Warning-first.docx")
  await page.getByRole("button", { name: "Add as supplied chronology", exact: true }).waitFor()
  assert.equal(requests.length, 1)
  assert.equal(requests[0]!.origin, null)
  const review = page.getByRole("dialog", { name: "Add this chronology?", exact: true })
  await review.waitFor()
  assert.match(await review.innerText(), /not independent corroboration/)
  assert.match(await review.innerText(), /will not replace the chronology workbench/)
  assert.match(await review.innerText(), /Warning-first.docx/)
  await page.getByRole("heading", { name: "Add this chronology?", exact: true }).waitFor()
  assert.equal(await review.evaluate(element => element.contains(document.activeElement)), true)
  await review.evaluate(element => { (element as HTMLElement).dataset.reviewProbe = "same-dialog" })
  await page.keyboard.press("Tab")
  assert.equal(await review.getByRole("button", { name: "Skip this file", exact: true }).evaluate(element => element === document.activeElement), true)
  await page.keyboard.press("Shift+Tab")
  assert.equal(await review.getByRole("button", { name: "Add as supplied chronology", exact: true }).evaluate(element => element === document.activeElement), true, "focus wraps inside the review")
  const geometry = await review.evaluate(element => ({ width: element.getBoundingClientRect().width, scroll: element.scrollWidth, client: element.clientWidth, buttons: [...element.querySelectorAll("button")].map(button => button.getBoundingClientRect().height), page: document.documentElement.scrollWidth }))
  assert.ok(geometry.width <= 375 && geometry.scroll <= geometry.client && geometry.page <= 375, JSON.stringify(geometry))
  assert.ok(geometry.buttons.every(height => height >= 44), JSON.stringify(geometry))
  await capture("intake-warning-375.png")
  await page.getByRole("button", { name: "Add as supplied chronology", exact: true }).click()
  await page.getByRole("button", { name: "Keep both", exact: true }).waitFor()
  assert.equal(
    await review.evaluate(element => element.contains(document.activeElement)),
    true,
    "the same review retains focus when acknowledgement reveals a collision",
  )
  assert.equal(requests[1]!.origin, "supplied-chronology")
  assert.equal(requests[1]!.existing, null)
  assert.equal(await review.getAttribute("data-review-probe"), "same-dialog", "acknowledgement and collision share one dialog instance")
  await capture("intake-collision-375.png")
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
  await review.getByRole("button", { name: "Skip this file", exact: true }).click()
  assert.equal(await page.locator(".upload-row.is-chronology").count(), 0)
  assert.equal(requests.length, count)

  // A second exception queues behind the active file; an ordinary upload in
  // the same batch completes without waiting for either decision.
  const input = (name: string) => ({ name, mimeType: "application/octet-stream", buffer: Buffer.from("synthetic batch source") })
  await page.locator('input[type="file"]').setInputFiles([input("Batch-one.docx"), input("Ordinary.pdf"), input("Batch-two.docx")])
  await page.getByRole("button", { name: "Ordinary.pdf", exact: true }).waitFor()
  await page.waitForFunction(() => document.querySelector(".upload-review-dialog")?.textContent?.includes("1 more file"))
  assert.match(await review.innerText(), /Batch-one.docx/)
  assert.equal(await page.locator("dialog[open]").count(), 1)
  const batchRequests = requests.filter(request => request.name.startsWith("Batch-"))
  assert.deepEqual(batchRequests.map(request => request.origin), [null, null], "no acknowledgement applies to another file")
  const beforeSkip = requests.length
  await page.keyboard.press("Escape")
  assert.match(await review.innerText(), /Batch-two.docx/)
  assert.equal(await page.locator("dialog[open]").count(), 1)
  await review.getByRole("button", { name: "Skip this file", exact: true }).click()
  assert.equal(await review.count(), 0)
  assert.equal(requests.length, beforeSkip)
  assert.equal(await page.locator('.docs-drawer-body[aria-label="Documents"]').evaluate(element => element === document.activeElement), true)

  // Escape cannot claim that an already-sent request was cancelled.
  await chooseFile("Slow.docx")
  await review.getByRole("button", { name: "Add as supplied chronology", exact: true }).click()
  await page.waitForFunction(() => document.querySelector(".upload-review-dialog button")?.hasAttribute("disabled"))
  assert.ok(releaseSlow)
  await page.keyboard.press("Escape")
  assert.equal(await review.count(), 1)
  assert.equal(await review.getByRole("button", { name: "Skip this file", exact: true }).isDisabled(), true)
  releaseSlow!()
  await page.getByRole("button", { name: "Slow.docx", exact: true }).waitFor()
  assert.equal(await review.count(), 0)

  // A temporary failure retains the file, acknowledgement and earlier choice.
  await chooseFile("Failure.docx")
  await page.getByRole("button", { name: "Keep both", exact: true }).click()
  await review.getByRole("button", { name: "Add as supplied chronology", exact: true }).click()
  const failed = page.locator(".upload-row.is-failed").filter({ hasText: "Failure.docx" })
  await failed.getByRole("button", { name: "Retry", exact: true }).waitFor()
  assert.equal(await review.count(), 0)
  await failed.getByRole("button", { name: "Retry", exact: true }).click()
  await page.getByRole("button", { name: "Failure (2).docx", exact: true }).waitFor()
  assert.deepEqual(requests.filter(request => request.name === "Failure.docx").map(request => [request.existing, request.origin]), [[null, null], ["keep-both", null], ["keep-both", "supplied-chronology"], ["keep-both", "supplied-chronology"]])

  await chooseFile("Generated-name.docx")
  const generated = page.locator(".upload-row.is-exists").filter({ hasText: "Generated-name.docx" })
  await generated.getByRole("button", { name: "Keep both", exact: true }).waitFor()
  assert.equal(await generated.getByRole("button", { name: "Save as new version", exact: true }).count(), 0)
  await generated.getByRole("button", { name: "Cancel", exact: true }).click()

  await chooseFile("Recognized-generated.docx")
  await review.getByRole("button", { name: "Add as supplied chronology", exact: true }).click()
  await review.getByRole("button", { name: "Keep both", exact: true }).waitFor()
  assert.equal(await review.getByRole("button", { name: "Save as new version", exact: true }).count(), 0)
  await review.getByRole("button", { name: "Skip this file", exact: true }).click()

  // Appearance scaling and long names must not push a phone review offscreen.
  await page.evaluate(() => document.documentElement.style.setProperty("--ui-scale", "1.25"))
  await chooseFile(`${"Long supplied chronology name ".repeat(8)}.docx`)
  await review.waitFor()
  const scaled = await review.evaluate(element => ({ left: element.getBoundingClientRect().left, right: element.getBoundingClientRect().right, scroll: element.scrollWidth, client: element.clientWidth, page: document.documentElement.scrollWidth }))
  assert.ok(scaled.left >= 0 && scaled.right <= 375 && scaled.scroll <= scaled.client && scaled.page <= 375, JSON.stringify(scaled))
  await page.keyboard.press("Escape")
  await page.evaluate(() => document.documentElement.style.removeProperty("--ui-scale"))

  await page.reload()
  await page.getByRole("button", { name: "Warning-first (2).docx", exact: true }).waitFor()
  assert.match(await page.locator(".docs-timeline").innerText(), /Supplied chronology/)
  await capture("intake-accepted-375.png")
  assert.equal(documents.length, 5)
  assert.equal(unknown.length, 0, unknown.join("\n"))
  if (captures) {
    const desktop = await browser.newPage({ viewport: { width: 1440, height: 1000 } })
    desktop.on("pageerror", error => errors.push(error.message))
    await desktop.goto(server.url.href)
    await desktop.locator('input[type="file"]').setInputFiles(input("Equipment chronology.docx"))
    await desktop.getByRole("dialog", { name: "Add this chronology?", exact: true }).waitFor()
    await desktop.evaluate(async () => { await document.fonts.ready })
    await desktop.screenshot({ path: path.join(captures, "intake-warning-1440.png"), animations: "disabled" })
    await desktop.close()
  }
  assert.deepEqual(errors, [])
  console.log("Chronology intake browser PASS: coordinated modal and both conflict orders; batch exception queue with ordinary progress; per-file Skip/Escape; focus loop/return; 375px touch targets/overflow; in-flight and failed retries; generated deliverable choices; persisted renamed provenance. Actual product CSS; ingest responses mocked.")
} finally {
  await browser?.close()
  server.stop(true)
  await rm(temporary, { recursive: true, force: true })
}
