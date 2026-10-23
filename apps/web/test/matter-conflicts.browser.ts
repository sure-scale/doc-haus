// Run from apps/web: bun test/matter-conflicts.browser.ts (optionally PLAYWRIGHT_MODULE / PLAYWRIGHT_CHROMIUM_PATH).
// The conflicts check on the new-matter form, the edit-matter dialog and the
// matter header's dialog, through the real Matters page, MatterHeader and
// ConflictsDialog; only ingest's matter HTTP is mocked, by a small in-memory
// store that validates like the endpoint. Set CONFLICTS_SCREENSHOTS=<directory>
// to save views.
import assert from "node:assert/strict"
import { mkdir, mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Page } from "@playwright/test"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-conflicts-"))
const shots = process.env.CONFLICTS_SCREENSHOTS
const source = path.resolve(import.meta.dir, "../src")
const file = (name: string) => JSON.stringify(path.join(source, name))
const dependency = (name: string) => JSON.stringify(fileURLToPath(import.meta.resolve(name)))

type Conflicts = { status: string; by?: string; owner?: string; due?: string; reason?: string; recorded_at: number }
type Row = { id: string; title: string; conflicts?: Conflicts; dir: string; created_at: number }
let matters: Row[] = []
let attorney = ""
let failConflicts = false
const writes: { route: string; body: Record<string, unknown> }[] = []

const entry = path.join(directory, "harness.tsx")
// The matter route stands in for MatterDetail with the same conflicts wiring:
// the chip opens the dialog, and a saved record replaces the matter's.
await Bun.write(entry, `
import { createElement, useEffect, useState } from ${dependency("react")}
import { createRoot } from ${dependency("react-dom/client")}
import { MemoryRouter, Route, Routes, useParams } from ${dependency("react-router-dom")}
import Matters from ${file("pages/Matters.tsx")}
import MatterHeader from ${file("components/matter/MatterHeader.tsx")}
import ConflictsDialog from ${file("components/matter/ConflictsDialog.tsx")}
import { ToastProvider } from ${file("components/Toast.tsx")}
import { getMatter } from ${file("api/ingest/matters.ts")}
const noop = () => {}
const done = async () => {}
function MatterPage() {
  const { id } = useParams()
  const [matter, setMatter] = useState(null)
  const [recording, setRecording] = useState(false)
  useEffect(() => { getMatter(id).then(setMatter) }, [id])
  if (!matter) return null
  return createElement("div", null,
    createElement(MatterHeader, {
      matter, jurisdictions: [],
      editingRef: false, setEditingRef: noop, draftRef: "", setDraftRef: noop, commitRef: done,
      editingClient: false, setEditingClient: noop, draftClient: "", setDraftClient: noop, commitClient: done,
      renaming: false, setRenaming: noop, draftTitle: "", setDraftTitle: noop, commitRename: done,
      patch: done, onEditConflicts: () => setRecording(true),
    }),
    recording && createElement(ConflictsDialog, {
      matterId: matter.id, current: matter.conflicts,
      onSaved: (saved) => setMatter((m) => m && { ...m, conflicts: saved.conflicts }),
      onClose: () => setRecording(false),
    }))
}
const start = new URLSearchParams(location.search).get("path") ?? "/"
createRoot(document.getElementById("root")).render(createElement(ToastProvider, null,
  createElement(MemoryRouter, { initialEntries: [start] },
    createElement("main", { className: "content" }, createElement("div", { className: "container" },
      createElement(Routes, null,
        createElement(Route, { path: "/", element: createElement(Matters) }),
        createElement(Route, { path: "/matter/:id", element: createElement(MatterPage) })))))))
`)
const build = await Bun.build({ entrypoints: [entry], target: "browser", define: { "import.meta.env": JSON.stringify({ VITE_OPENCODE_URL: "/engine", VITE_INGEST_URL: "/ingest" }) } })
assert.equal(build.success, true, build.logs.join("\n"))
const script = await build.outputs.find((output) => output.kind === "entry-point")!.text()
const imports = [...(await readFile(path.join(source, "main.tsx"), "utf8")).matchAll(/import "(\.\/[^"\n]+\.css)"/g)]
const css = (await Promise.all(imports.map(([, name]) => readFile(path.resolve(source, name!), "utf8")))).join("\n")

const server = Bun.serve({
  port: 0,
  hostname: "127.0.0.1",
  async fetch(request) {
    const url = new URL(request.url)
    const route = `${request.method} ${url.pathname}`
    if (route === "GET /harness.js") return new Response(script, { headers: { "content-type": "text/javascript" } })
    if (route === "GET /") return new Response(`<!doctype html><html><meta charset="utf-8"><style>${css}</style><div id="root"></div><script type="module" src="/harness.js"></script></html>`, { headers: { "content-type": "text/html" } })
    if (route === "GET /ingest/matters") return Response.json(matters)
    if (route === "GET /ingest/jurisdictions") return Response.json([])
    if (route === "GET /ingest/preferences") return Response.json({ attorney })
    if (route === "POST /ingest/matters") {
      const body = (await request.json()) as Record<string, string>
      writes.push({ route, body })
      const row: Row = { id: `m${matters.length + 1}`, title: body.title!, dir: "/tmp/matter", created_at: matters.length + 1 }
      matters.push(row)
      return Response.json(row)
    }
    const conflicts = url.pathname.match(/^\/ingest\/matters\/(\w+)\/conflicts$/)
    if (conflicts && request.method === "POST") {
      const body = (await request.json()) as Record<string, string>
      writes.push({ route: `POST conflicts ${conflicts[1]}`, body })
      if (failConflicts) return Response.json({ error: "The matter file is locked" }, { status: 500 })
      if (body.status === "pending" && !body.owner) return Response.json({ error: 'conflicts "pending" requires "owner" — who is responsible for completing the check' }, { status: 400 })
      const row = matters.find((m) => m.id === conflicts[1])!
      row.conflicts = { status: body.status!, by: body.by, owner: body.owner, due: body.due, reason: body.reason, recorded_at: Date.UTC(2026, 9, 23, 12) }
      return Response.json(row)
    }
    const one = url.pathname.match(/^\/ingest\/matters\/(\w+)$/)
    const row = one && matters.find((m) => m.id === one[1])
    if (row && request.method === "GET") return Response.json({ ...row, documents: [] })
    if (row && request.method === "PATCH") {
      const body = (await request.json()) as Record<string, string>
      writes.push({ route: `PATCH ${row.id}`, body })
      row.title = body.title!
      return Response.json(row)
    }
    return Response.json({ error: "not mocked" }, { status: 404 })
  },
})

const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
if (shots) await mkdir(shots, { recursive: true })
const passed: string[] = []

async function scenario(name: string, setup: { seed?: Row[]; attorney?: string; failConflicts?: boolean }, run: (page: Page) => Promise<void>, start = "/") {
  matters = (setup.seed ?? []).map((row) => ({ ...row }))
  attorney = setup.attorney ?? ""
  failConflicts = setup.failConflicts ?? false
  writes.length = 0
  const page = await browser.newPage({ viewport: { width: 1280, height: 1000 }, reducedMotion: "reduce" })
  page.setDefaultTimeout(5000)
  const errors: string[] = []
  page.on("pageerror", (error) => errors.push(error.message))
  try {
    await page.goto(`${server.url.href}?${new URLSearchParams({ path: start })}`)
    await run(page)
    assert.deepEqual(errors, [])
    passed.push(name)
    console.log(`PASS ${name}`)
  } finally {
    await page.close()
  }
}

const shot = (page: Page, name: string, selector?: string) =>
  shots ? (selector ? page.locator(selector).first().screenshot({ path: path.join(shots, `${name}.png`) }) : page.screenshot({ path: path.join(shots, `${name}.png`) })) : Promise.resolve(undefined)
const routes = () => writes.map((w) => w.route)
const createForm = (page: Page) => page.locator("#new-matter-fields")
const cleared: Row = { id: "a", title: "Lease review", conflicts: { status: "cleared", by: "A. Lawyer", recorded_at: Date.UTC(2026, 9, 1, 12) }, dir: "/tmp/a", created_at: 1 }
const unrecorded: Row = { id: "b", title: "Board minutes", dir: "/tmp/b", created_at: 2 }

try {
  await scenario("a new matter leaves the check unrecorded by default and creates in one request", {}, async (page) => {
    const group = createForm(page).getByRole("radiogroup", { name: "Conflicts check" })
    await expect(group.getByRole("radio", { name: "Not recorded" })).toBeChecked()
    await shot(page, "create-default", ".matter-create-panel")
    await page.getByRole("textbox", { name: "Matter title" }).fill("Option deed")
    await page.getByRole("button", { name: "Create matter" }).click()
    await expect(page.getByRole("button", { name: /Conflicts: not recorded/ })).toBeVisible()
    assert.deepEqual(routes(), ["POST /ingest/matters"])
  })

  await scenario("a pending check with no owner says so inline, then is recorded with the new matter", {}, async (page) => {
    await page.getByRole("textbox", { name: "Matter title" }).fill("Share purchase")
    await createForm(page).getByRole("radio", { name: "Pending" }).check()
    await page.getByRole("button", { name: "Create matter" }).click()
    const owner = createForm(page).getByRole("textbox", { name: "Owner" })
    await expect(owner).toBeFocused()
    await expect(owner).toHaveAttribute("aria-invalid", "true")
    await expect(createForm(page).getByText("Name who will complete the check.")).toBeVisible()
    assert.deepEqual(routes(), [], "a blank owner must not create the matter")
    await shot(page, "create-pending-owner-missing", ".matter-create-panel")
    await owner.fill("B. Partner")
    await expect(createForm(page).getByText("Name who will complete the check.")).toHaveCount(0)
    await page.getByRole("button", { name: "Create matter" }).click()
    await expect(page.getByRole("button", { name: /Conflicts: pending — owner B\. Partner/ })).toBeVisible()
    assert.deepEqual(routes(), ["POST /ingest/matters", "POST conflicts m1"])
    assert.deepEqual(writes[1]!.body, { status: "pending", owner: "B. Partner" })
    await shot(page, "create-pending-header-chip", ".matter-page-head")
  })

  await scenario("choosing a state fills a blank name with the firm's attorney", { attorney: "A. Lawyer" }, async (page) => {
    await page.getByRole("textbox", { name: "Matter title" }).fill("Supply note")
    await createForm(page).getByRole("radio", { name: "Cleared" }).check()
    await expect(createForm(page).getByRole("textbox", { name: "Checked by" })).toHaveValue("A. Lawyer")
    await shot(page, "create-cleared", ".matter-create-panel")
    await page.getByRole("button", { name: "Create matter" }).click()
    await expect(page.getByRole("button", { name: /Conflicts: cleared by A\. Lawyer/ })).toBeVisible()
  })

  await scenario("if recording fails, the matter is still created and opened, and the failure is said", { failConflicts: true }, async (page) => {
    await page.getByRole("textbox", { name: "Matter title" }).fill("Urgent filing")
    await createForm(page).getByRole("radio", { name: "Bypassed" }).check()
    await createForm(page).getByRole("textbox", { name: "Bypassed by" }).fill("C. Counsel")
    await createForm(page).getByRole("textbox", { name: "Reason" }).fill("Court deadline today")
    await page.getByRole("button", { name: "Create matter" }).click()
    await expect(page.getByText(/Created matter "Urgent filing", but its conflicts check was not recorded \(The matter file is locked\)/)).toBeVisible()
    await expect(page.getByRole("button", { name: /Conflicts: not recorded/ })).toBeVisible()
    assert.deepEqual(routes(), ["POST /ingest/matters", "POST conflicts m1"])
    await shot(page, "create-record-failed")
  })

  await scenario("the edit dialog shows the recorded check and records only a change", { seed: [cleared, unrecorded] }, async (page) => {
    await page.getByRole("button", { name: "Edit Lease review" }).click()
    const dialog = page.getByRole("dialog", { name: "Edit matter" })
    const group = dialog.getByRole("radiogroup", { name: "Conflicts check" })
    await expect(group.getByRole("radio", { name: "Cleared" })).toBeChecked()
    // A recorded check can be replaced but not un-recorded.
    await expect(group.getByRole("radio", { name: "Not recorded" })).toHaveCount(0)
    await expect(dialog.getByRole("textbox", { name: "Checked by" })).toHaveValue("A. Lawyer")
    await shot(page, "edit-recorded", "dialog.matter-edit")
    await dialog.getByRole("button", { name: "Save changes" }).click()
    await expect(dialog).toHaveCount(0)
    assert.deepEqual(routes(), ["PATCH a"], "an unchanged check is not re-recorded")

    writes.length = 0
    await page.getByRole("button", { name: "Edit Lease review" }).click()
    await group.getByRole("radio", { name: "Bypassed" }).check()
    await dialog.getByRole("textbox", { name: "Reason" }).fill("")
    await dialog.getByRole("button", { name: "Save changes" }).click()
    await expect(dialog.getByRole("textbox", { name: "Reason" })).toBeFocused()
    await expect(dialog.getByText("Give the reason for the bypass.")).toBeVisible()
    assert.deepEqual(routes(), [])
    await shot(page, "edit-bypass-reason-missing", "dialog.matter-edit")
    await dialog.getByRole("textbox", { name: "Reason" }).fill("Client instruction")
    await dialog.getByRole("button", { name: "Save changes" }).click()
    await expect(dialog).toHaveCount(0)
    assert.deepEqual(routes(), ["PATCH a", "POST conflicts a"])
    assert.deepEqual(writes[1]!.body, { status: "bypassed", by: "A. Lawyer", reason: "Client instruction" })

    // An unrecorded matter opens on "Not recorded", which saves no check.
    writes.length = 0
    await page.getByRole("button", { name: "Edit Board minutes" }).click()
    await expect(group.getByRole("radio", { name: "Not recorded" })).toBeChecked()
    await dialog.getByRole("button", { name: "Save changes" }).click()
    await expect(dialog).toHaveCount(0)
    assert.deepEqual(routes(), ["PATCH b"])
  })

  await scenario("the header dialog shows a blank owner as a field message, and the chip follows the save", { seed: [unrecorded] }, async (page) => {
    await page.getByRole("button", { name: /Conflicts: not recorded/ }).click()
    const dialog = page.getByRole("dialog", { name: "Conflicts check" })
    await dialog.getByRole("radio", { name: "Pending" }).check()
    await dialog.getByRole("textbox", { name: "Owner" }).fill("")
    await dialog.getByRole("button", { name: "Record" }).click()
    await expect(dialog.getByText("Name who will complete the check.")).toBeVisible()
    await expect(dialog.getByRole("alert")).toHaveCount(0)
    assert.deepEqual(routes(), [], "the raw endpoint error is never reached")
    await shot(page, "header-dialog-owner-missing", ".compare-dialog")
    await dialog.getByRole("textbox", { name: "Owner" }).fill("B. Partner")
    await dialog.getByRole("button", { name: "Record" }).click()
    await expect(dialog).toHaveCount(0)
    await expect(page.getByRole("button", { name: /Conflicts: pending — owner B\. Partner/ })).toBeVisible()
  }, "/matter/b")

  console.log(`PASS ${passed.length} matter conflicts cases with no browser errors`)
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
