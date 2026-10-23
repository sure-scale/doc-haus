// Run from apps/web: PLAYWRIGHT_MODULE=<installed playwright module> bun test/email-attachments.browser.ts
// The real documents list with only the ingest matter and delete routes mocked.
// No model or engine calls. PLAYWRIGHT_CHROMIUM_PATH can select an installed Chromium.
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Document } from "../src/api/ingest/documents"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
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

// Synthetic matter: two emails, each with one extracted attachment.
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
let documents: Document[] = [
  {
    id: 1, name: "Thread.eml", doc_path: "/m/Thread.eml", created_at: 4, role: "source",
    attachments: [
      { name: "Schedule.docx", type: DOCX, size: 7000, status: "indexed", document: "Schedule.docx" },
      { name: "bundle.zip", type: "application/zip", size: 120, status: "unsupported" },
    ],
    embedded_images: 1,
    removal: { attachments: ["Schedule.docx"], kept: [{ name: "Shared.pdf", reason: "linked" }] },
  },
  { id: 2, name: "Schedule.docx", doc_path: "/m/Schedule.docx", created_at: 3, role: "source", attachment_of: { name: "Thread.eml", removed: false } },
  {
    id: 3, name: "Other.eml", doc_path: "/m/Other.eml", created_at: 2, role: "source",
    attachments: [{ name: "Note.pdf", type: "application/pdf", size: 900, status: "indexed", document: "Note.pdf" }],
    removal: { attachments: ["Note.pdf"], kept: [] },
  },
  { id: 4, name: "Note.pdf", doc_path: "/m/Note.pdf", created_at: 1, role: "source", attachment_of: { name: "Other.eml", removed: false } },
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
    requests.push(`${request.method} ${url.pathname}${url.search}`)
    if (request.method === "DELETE" && url.pathname === "/ingest/matters/m/documents") {
      const name = url.searchParams.get("name")!
      const keep = url.searchParams.get("attachments") === "keep"
      const removal = documents.find(d => d.name === name)?.removal
      const removed = keep ? [] : removal?.attachments ?? []
      documents = documents
        .filter(d => d.name !== name && !removed.includes(d.name))
        .map(d => d.attachment_of?.name === name ? { ...d, attachment_of: { name, removed: true } } : d)
      return Response.json({ ok: true, name, role: "source", attachments: { removed, kept: keep ? [] : removal?.kept ?? [] } })
    }
    return Response.json({ error: "Unexpected request" }, { status: 404 })
  },
})
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
const errors: string[] = []
try {
  const page = await browser.newPage({ viewport: { width: 375, height: 812 }, reducedMotion: "reduce" })
  page.setDefaultTimeout(5000)
  page.on("pageerror", error => errors.push(error.message))
  await page.goto(server.url.href)

  // The email lists every attachment with its status; the child names its email.
  const attachments = page.getByRole("list", { name: "Attachments" }).first()
  await expect(attachments).toContainText("bundle.zip")
  await expect(attachments).toContainText("Not supported")
  await expect(attachments.getByRole("button", { name: "Added" })).toBeVisible()
  await expect(attachments).toContainText("1 embedded image in the message body")
  await expect(page.getByText("Attachment of Thread.eml")).toBeVisible()
  console.log("PASS An email lists its attachments with their status; an extracted attachment names its email.")

  // Removing an email offers both choices, naming what goes and what stays.
  await page.getByRole("button", { name: "Remove Thread.eml" }).click()
  const confirm = page.getByRole("group", { name: "Remove Thread.eml" })
  await expect(confirm).toContainText("Thread.eml came with Schedule.docx.")
  await expect(confirm).toContainText("Shared.pdf is kept: another email also has it.")
  await expect(confirm.getByRole("button")).toHaveText(["Remove email and 1 attachment", "Remove email only", "Cancel"])
  await expect(confirm.getByRole("button", { name: "Remove email and 1 attachment" })).toBeFocused()
  const bounds = await confirm.boundingBox()
  assert.ok(bounds && bounds.x >= 0 && bounds.x + bounds.width <= 375, `confirmation fits a 375px drawer: ${JSON.stringify(bounds)}`)
  await confirm.getByRole("button", { name: "Cancel" }).click()
  await expect(confirm).toHaveCount(0)
  assert.deepEqual(requests, [], "cancel removes nothing")

  await page.getByRole("button", { name: "Remove Thread.eml" }).click()
  await confirm.getByRole("button", { name: "Remove email only" }).click()
  await expect(page.getByText("From removed email Thread.eml")).toBeVisible()
  await expect(page.getByRole("button", { name: "Schedule.docx", exact: true })).toBeVisible()
  console.log("PASS Remove email only keeps the attachment and shows it as from a removed email.")

  await page.getByRole("button", { name: "Remove Other.eml" }).click()
  await page.getByRole("group", { name: "Remove Other.eml" }).getByRole("button", { name: "Remove email and 1 attachment" }).click()
  await expect(page.getByRole("button", { name: "Note.pdf", exact: true })).toHaveCount(0)
  await expect(page.getByText("Removed Other.eml and 1 attachment.")).toBeVisible()
  assert.deepEqual(requests, [
    "DELETE /ingest/matters/m/documents?name=Thread.eml&attachments=keep",
    "DELETE /ingest/matters/m/documents?name=Other.eml",
  ])
  assert.deepEqual(errors, [])
  console.log("PASS Remove email and attachments is the default and removes them in the same request.")
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
