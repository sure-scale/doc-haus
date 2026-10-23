// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/email-attachments.browser.ts
// The real documents list with only the ingest matter and delete routes mocked.
// No model or engine calls. PLAYWRIGHT_CHROMIUM_PATH can select an installed Chromium.
// Set EMAIL_ATTACHMENTS_SCREENSHOTS=<directory> to save the tree views.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Page } from "@playwright/test"
import type { Document } from "../src/api/ingest/documents"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const shots = process.env.EMAIL_ATTACHMENTS_SCREENSHOTS
// The drawer scrolls inside the page, so a tall viewport shows the whole list.
async function shot(page: Page, name: string) {
  if (!shots) return
  await page.setViewportSize({ width: 375, height: 1500 })
  // Let entry animations (the confirm's fade-in) finish.
  await page.waitForTimeout(500)
  await page.screenshot({ path: path.join(shots, `${name}.png`) })
  await page.setViewportSize({ width: 375, height: 812 })
}
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-email-attachments-"))
const source = (file: string) => JSON.stringify(path.resolve(import.meta.dir, "../src", file))
const entry = path.join(directory, "harness.tsx")
await Bun.write(entry, `
import { createElement as h, useState } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import { createRoot } from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import DocumentUpload from ${source("components/DocumentUpload.tsx")}
import { ToastProvider } from ${source("components/Toast.tsx")}

function Harness({ initial }) {
  const [documents, setDocuments] = useState(initial)
  const reload = () => fetch("/ingest/matters/m").then(r => r.json()).then(m => setDocuments(m.documents))
  return h(ToastProvider, null, h(DocumentUpload, { matterId: "m", documents, onUploaded: reload, onView: () => {} }))
}
fetch("/ingest/matters/m").then(r => r.json()).then(m => createRoot(document.getElementById("root")).render(h(Harness, { initial: m.documents })))
`)
const build = await Bun.build({
  entrypoints: [entry], target: "browser",
  define: { "import.meta.env": JSON.stringify({ VITE_INGEST_URL: "/ingest", VITE_OPENCODE_URL: "/" }) },
})
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find(output => output.kind === "entry-point")!.text()
const css = (await Promise.all([
  "styles.css", "styles/primitives.css", "styles/upload.css", "styles/design-system.css", "styles/design-documents.css",
].map(file => Bun.file(path.resolve(import.meta.dir, "../src", file)).text()))).join("\n")

// Synthetic matter: an email with a document, an archive holding a further
// document (a grandchild) and an unsupported file; a second email with a
// document and an unreadable one; a third with a long list of attachments; and
// a document uploaded on its own between them.
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
let documents: Document[] = [
  {
    id: 1, name: "Thread.eml", doc_path: "/m/Thread.eml", created_at: 4, role: "source",
    attachments: [
      { name: "Schedule.docx", type: DOCX, size: 7000, status: "indexed", document: "Schedule.docx" },
      { name: "Exhibits.zip", type: "application/zip", size: 900, status: "indexed", document: "Exhibits.zip" },
      { name: "bundle.zip", type: "application/zip", size: 120, status: "unsupported" },
    ],
    embedded_images: 1,
    removal: { attachments: ["Schedule.docx", "Exhibits.zip"], kept: [{ name: "Shared.pdf", reason: "linked" }] },
  },
  {
    id: 2, name: "Schedule.docx", doc_path: "/m/Schedule.docx", created_at: 3, role: "source", attachment_of: { name: "Thread.eml", removed: false },
    quality: { status: "degraded", signals: [{ code: "replacement_characters", detail: "40 of 900 characters are replacement or control characters" }] },
  },
  {
    id: 7, name: "Exhibits.zip", doc_path: "/m/Exhibits.zip", created_at: 4, role: "source", attachment_of: { name: "Thread.eml", removed: false },
    // A zip's own outcomes, including one this build has no words for.
    attachments: [
      { name: "Exhibit A.pdf", type: "application/pdf", size: 900, status: "indexed", document: "Exhibit A.pdf" },
      { name: "Locked.docx", type: DOCX, size: 900, status: "encrypted" as unknown as "failed" },
      { name: "Later.bin", type: "", size: 1, status: "from_a_newer_ingest" as unknown as "failed" },
    ],
    removal: { attachments: ["Exhibit A.pdf"], kept: [] },
  },
  { id: 8, name: "Exhibit A.pdf", doc_path: "/m/Exhibit A.pdf", created_at: 4, role: "source", attachment_of: { name: "Exhibits.zip", removed: false } },
  {
    id: 5, name: "Scan.pdf", doc_path: "/m/Scan.pdf", created_at: 1, role: "source", attachment_of: { name: "Other.eml", removed: false },
    quality: { status: "no_readable_text", signals: [{ code: "empty", detail: "No text could be extracted" }] },
  },
  { id: 9, name: "Lease.docx", doc_path: "/m/Lease.docx", created_at: 3, role: "source" },
  {
    id: 3, name: "Other.eml", doc_path: "/m/Other.eml", created_at: 2, role: "source",
    attachments: [
      { name: "Note.pdf", type: "application/pdf", size: 900, status: "indexed", document: "Note.pdf" },
      { name: "Scan.pdf", type: "application/pdf", size: 4000, status: "no_readable_text", document: "Scan.pdf" },
    ],
    removal: { attachments: ["Note.pdf"], kept: [] },
  },
  { id: 4, name: "Note.pdf", doc_path: "/m/Note.pdf", created_at: 1, role: "source", attachment_of: { name: "Other.eml", removed: false } },
  {
    id: 10, name: "Photos.eml", doc_path: "/m/Photos.eml", created_at: 0, role: "source",
    attachments: Array.from({ length: 7 }, (_, index) => ({ name: `IMG_${index + 1}.heic`, type: "image/heic", size: 1, status: "unsupported" as const })),
  },
]
const requests: string[] = []
const server = Bun.serve({
  port: 0, hostname: "127.0.0.1",
  fetch(request) {
    const url = new URL(request.url)
    if (url.pathname === "/harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (url.pathname === "/") return new Response(`<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><style>${css}
      body { margin: 0; padding: 16px; }
    </style><div id="root" class="docs-drawer"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (request.method === "GET" && url.pathname === "/ingest/matters/m") return Response.json({ id: "m", documents })
    // An upload with no readable text: kept and listed, nothing indexed.
    if (request.method === "POST" && url.pathname === "/ingest/matters/m/documents") {
      const quality = { status: "no_readable_text" as const, signals: [{ code: "empty", detail: "No text could be extracted" }] }
      documents = [...documents, { id: 6, name: "Blank.docx", doc_path: "/m/Blank.docx", created_at: 9, role: "source", quality }]
      return Response.json({ name: "Blank.docx", docPath: "/m/Blank.docx", sections: 0, chunks: 0, injection: null, quality })
    }
    requests.push(`${request.method} ${url.pathname}${url.search}`)
    if (request.method === "DELETE" && url.pathname === "/ingest/matters/m/documents") {
      const name = url.searchParams.get("name")!
      const keep = url.searchParams.get("attachments") === "keep"
      const removal = documents.find(d => d.name === name)?.removal
      const removed = keep ? [] : removal?.attachments ?? []
      const gone = [name, ...removed]
      documents = documents
        .filter(d => !gone.includes(d.name))
        .map(d => d.attachments ? { ...d, attachments: d.attachments.map(a => a.document && gone.includes(a.document) ? { ...a, status: "removed" as const } : a) } : d)
        .map(d => d.attachment_of?.name === name ? { ...d, attachment_of: { name, removed: true } } : d)
      return Response.json({ ok: true, name, role: "source", attachments: { removed, kept: keep ? [] : removal?.kept ?? [] } })
    }
    return Response.json({ error: "Unexpected request" }, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const errors: string[] = []
try {
  const page = await browser.newPage({ viewport: { width: 375, height: 812 }, reducedMotion: "reduce", deviceScaleFactor: shots ? 2 : 1 })
  page.setDefaultTimeout(5000)
  page.on("pageerror", error => errors.push(error.message))
  await page.goto(server.url.href)

  // Each attachment is one row under its email, never again at the top level;
  // an archive's file sits one level further in. Roots keep newest-first order.
  const thread = page.getByRole("list", { name: "Attachments of Thread.eml" })
  await expect(thread.locator(":scope > li > .doc-row")).toHaveCount(3)
  await expect(page.getByRole("button", { name: "Schedule.docx", exact: true })).toHaveCount(1)
  await expect(thread.getByRole("button", { name: "Schedule.docx", exact: true })).toBeVisible()
  await expect(thread.locator(":scope > li > .doc-row.is-entry")).toContainText(["bundle.zip"])
  await expect(thread.locator(":scope > li > .doc-row.is-entry .doc-child-status")).toHaveText(["Not supported"])
  await expect(thread).toContainText("1 embedded image in the message body")
  await expect(page.getByText(/Attachment of/)).toHaveCount(0)
  const archive = page.getByRole("list", { name: "Files in Exhibits.zip" })
  await expect(archive.getByRole("button", { name: "Exhibit A.pdf", exact: true })).toBeVisible()
  await expect(archive.locator(".doc-row.is-entry .doc-child-status")).toHaveText(["Password-protected", "Not added"])
  // A zip opens in the text viewer as its file list, so its name is a link.
  await expect(page.getByRole("button", { name: "Exhibits.zip", exact: true })).toHaveCount(1)
  const left = async (name: string) => (await page.locator(".doc-row", { has: page.getByText(name, { exact: true }) }).first().locator(".doc-type").boundingBox())!.x
  const [rootX, childX, grandchildX] = [await left("Thread.eml"), await left("Exhibits.zip"), await left("Exhibit A.pdf")]
  assert.ok(rootX < childX && childX < grandchildX, `each level indents further: ${rootX} < ${childX} < ${grandchildX}`)
  const names = await page.locator(".docs-timeline-files > li > .doc-row .doc-row-main").evaluateAll(rows => rows.map(row => row.textContent))
  assert.deepEqual(names, ["Thread.eml", "Lease.docx", "Other.eml", "Photos.eml"], "only parents and lone uploads are top-level, newest first")
  await expect(page.locator(".matter-files-count")).toHaveText(String(documents.length))
  console.log("PASS Attachments are listed once, under their email; an archive's file is indented twice; roots keep their order.")

  // A long attachment list starts closed behind its count and opens on demand.
  const toggle = page.getByRole("button", { name: "7 attachments" })
  await expect(toggle).toHaveAttribute("aria-expanded", "false")
  await expect(page.getByRole("list", { name: "Attachments of Photos.eml" })).toHaveCount(0)
  await toggle.click()
  await expect(page.getByRole("button", { name: "Hide 7 attachments" })).toHaveAttribute("aria-expanded", "true")
  await expect(page.getByRole("list", { name: "Attachments of Photos.eml" }).locator(".doc-row.is-entry")).toHaveCount(7)
  await shot(page, "1-tree-expanded")
  await page.getByRole("button", { name: "Hide 7 attachments" }).click()
  await shot(page, "2-tree-long-list-closed")
  console.log("PASS A long attachment list is collapsed behind its count and expands.")

  // Extraction quality (#239): an attachment with no readable text is its own
  // row under its email, marked not indexed; degraded and unreadable documents
  // say so under their names.
  await expect(page.getByRole("list", { name: "Attachments of Other.eml" }).getByRole("button", { name: "Scan.pdf", exact: true })).toBeVisible()
  await expect(page.locator(".doc-row", { has: page.getByRole("button", { name: "Schedule.docx", exact: true }) }).locator(".document-quality"))
    .toHaveText("Text may be incomplete or garbled")
  const scanNote = page.locator(".doc-row", { has: page.getByRole("button", { name: "Scan.pdf", exact: true }) }).locator(".document-quality")
  await expect(scanNote).toHaveText("Not indexed: no readable text")
  await expect(scanNote).toHaveAttribute("title", "No text could be extracted")
  const noteBounds = await scanNote.boundingBox()
  assert.ok(noteBounds && noteBounds.x + noteBounds.width <= 375, `quality note fits a 375px drawer: ${JSON.stringify(noteBounds)}`)
  await page.locator('input[type="file"]').setInputFiles({ name: "Blank.docx", mimeType: "application/octet-stream", buffer: Buffer.from("synthetic") })
  await expect(page.locator(".upload-row.is-not_indexed .upload-status")).toHaveText("Added, not indexed: no readable text")
  await expect(page.getByText("Blank.docx was added but not indexed: no readable text (no text could be extracted).", { exact: false })).toBeVisible()
  await expect(page.getByText("is ready")).toHaveCount(0)
  await expect(page.locator(".doc-row", { has: page.getByRole("button", { name: "Blank.docx", exact: true }) }).locator(".document-quality"))
    .toHaveText("Not indexed: no readable text")
  console.log("PASS Unreadable and degraded documents say so in the list, the attachment list, the upload row and the toast.")

  // Removing an email offers both choices, naming what goes and what stays.
  await page.getByRole("button", { name: "Remove Thread.eml" }).click()
  const confirm = page.getByRole("group", { name: "Remove Thread.eml" })
  await expect(confirm).toContainText("Thread.eml came with Schedule.docx, Exhibits.zip.")
  await expect(confirm).toContainText("Shared.pdf is kept: another email or zip also has it.")
  await expect(confirm.getByRole("button")).toHaveText(["Remove email and 2 attachments", "Remove email only", "Cancel"])
  await expect(confirm.getByRole("button", { name: "Remove email and 2 attachments" })).toBeFocused()
  await shot(page, "3-remove-email-confirm")
  const bounds = await confirm.boundingBox()
  assert.ok(bounds && bounds.x >= 0 && bounds.x + bounds.width <= 375, `confirmation fits a 375px drawer: ${JSON.stringify(bounds)}`)
  await confirm.getByRole("button", { name: "Cancel" }).click()
  await expect(confirm).toHaveCount(0)
  assert.deepEqual(requests, [], "cancel removes nothing")

  // A zip's confirmation names its files, not attachments.
  await page.getByRole("button", { name: "Remove Exhibits.zip" }).click()
  const zipConfirm = page.getByRole("group", { name: "Remove Exhibits.zip" })
  await expect(zipConfirm).toContainText("Exhibits.zip contains Exhibit A.pdf.")
  await expect(zipConfirm.getByRole("button")).toHaveText(["Remove zip and 1 file", "Remove zip only", "Cancel"])
  await shot(page, "4-remove-zip-confirm")
  await zipConfirm.getByRole("button", { name: "Cancel" }).click()

  await page.getByRole("button", { name: "Remove Thread.eml" }).click()
  await confirm.getByRole("button", { name: "Remove email only" }).click()
  await expect(page.getByText("From removed email Thread.eml")).toHaveCount(2)
  await expect(page.locator(".docs-timeline-files > li > .doc-row", { has: page.getByRole("button", { name: "Schedule.docx", exact: true }) })).toBeVisible()
  // The archive's file stays under the archive, which is still here.
  await expect(page.getByRole("list", { name: "Files in Exhibits.zip" }).getByRole("button", { name: "Exhibit A.pdf", exact: true })).toBeVisible()
  console.log("PASS Remove email only keeps the attachments at the top level, shown as from a removed email.")

  // A child is removed on its own with the single confirm, not the email choice.
  await page.getByRole("button", { name: "Remove Exhibit A.pdf" }).click()
  await page.locator(".doc-row", { has: page.getByRole("button", { name: "Exhibit A.pdf", exact: true }) }).getByRole("button", { name: "Confirm" }).click()
  await expect(page.getByRole("button", { name: "Exhibit A.pdf", exact: true })).toHaveCount(0)
  // Its entry stays in the zip's list, as removed.
  await expect(page.getByRole("list", { name: "Files in Exhibits.zip" }).locator(".doc-row.is-entry", { hasText: "Exhibit A.pdf" })).toContainText("Removed from the matter")
  console.log("PASS An attachment row removes only itself.")

  await page.getByRole("button", { name: "Remove Other.eml" }).click()
  await page.getByRole("group", { name: "Remove Other.eml" }).getByRole("button", { name: "Remove email and 1 attachment" }).click()
  await expect(page.getByRole("button", { name: "Note.pdf", exact: true })).toHaveCount(0)
  await expect(page.getByText("Removed Other.eml and 1 attachment.")).toBeVisible()
  assert.deepEqual(requests, [
    "DELETE /ingest/matters/m/documents?name=Thread.eml&attachments=keep",
    "DELETE /ingest/matters/m/documents?name=Exhibit%20A.pdf",
    "DELETE /ingest/matters/m/documents?name=Other.eml",
  ])
  assert.deepEqual(errors, [])
  console.log("PASS Remove email and attachments is the default and removes them in the same request.")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
