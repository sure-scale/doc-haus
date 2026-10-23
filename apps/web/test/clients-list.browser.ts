// Run from apps/web: bun test/clients-list.browser.ts (optionally PLAYWRIGHT_MODULE / PLAYWRIGHT_CHROMIUM_PATH).
// The Clients list and the type-to-suggest matter Client field through the
// real sidebar, Matters and Clients pages and matter header; only ingest's
// matter HTTP is mocked, by a small in-memory store. Set
// CLIENTS_SCREENSHOTS=<directory> to save views.
import assert from "node:assert/strict"
import { mkdir, mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Page } from "@playwright/test"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-clients-"))
const shots = process.env.CLIENTS_SCREENSHOTS
const source = path.resolve(import.meta.dir, "../src")
const file = (name: string) => JSON.stringify(path.join(source, name))
const dependency = (name: string) => JSON.stringify(fileURLToPath(import.meta.resolve(name)))

type Row = { id: string; title: string; client?: string; reference?: string; dir: string; created_at: number }
let matters: Row[] = []
const writes: { method: string; id?: string; body: Record<string, unknown> }[] = []

const entry = path.join(directory, "harness.tsx")
// The matter route stands in for MatterDetail with the same header wiring: an
// edit flag and draft per field, and a client commit that sends only a change.
await Bun.write(entry, `
import { createElement, useEffect, useState } from ${dependency("react")}
import { createRoot } from ${dependency("react-dom/client")}
import { MemoryRouter, Route, Routes, useParams } from ${dependency("react-router-dom")}
import Sidebar from ${file("components/Sidebar.tsx")}
import Matters from ${file("pages/Matters.tsx")}
import Clients from ${file("pages/Clients.tsx")}
import MatterHeader from ${file("components/matter/MatterHeader.tsx")}
import { ToastProvider } from ${file("components/Toast.tsx")}
import { getMatter, renameMatter } from ${file("api/ingest/matters.ts")}
const noop = () => {}
const done = async () => {}
function MatterPage() {
  const { id } = useParams()
  const [matter, setMatter] = useState(null)
  const [editingClient, setEditingClient] = useState(false)
  const [draftClient, setDraftClient] = useState("")
  useEffect(() => { getMatter(id).then(setMatter) }, [id])
  if (!matter) return null
  async function commitClient() {
    setEditingClient(false)
    const next = draftClient.trim()
    if (next === (matter.client ?? "")) return
    const updated = await renameMatter(matter.id, matter.title, matter.reference, matter.jurisdictions, matter.playbook, next)
    setMatter({ ...matter, client: updated.client })
  }
  return createElement(MatterHeader, {
    matter, jurisdictions: [],
    editingRef: false, setEditingRef: noop, draftRef: "", setDraftRef: noop, commitRef: done,
    editingClient, setEditingClient, draftClient, setDraftClient, commitClient,
    renaming: false, setRenaming: noop, draftTitle: "", setDraftTitle: noop, commitRename: done,
    patch: done, onEditConflicts: noop,
  })
}
const start = new URLSearchParams(location.search).get("path") ?? "/"
createRoot(document.getElementById("root")).render(createElement(ToastProvider, null,
  createElement(MemoryRouter, { initialEntries: [start] },
    createElement("div", { className: "shell" },
      createElement(Sidebar, { onOpenSettings: noop, sessionsVersion: 0 }),
      createElement("main", { className: "content" }, createElement("div", { className: "container" },
        createElement(Routes, null,
          createElement(Route, { path: "/", element: createElement(Matters) }),
          createElement(Route, { path: "/clients", element: createElement(Clients) }),
          createElement(Route, { path: "/matter/:id", element: createElement(MatterPage) }))))))))
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
    if (route === "POST /ingest/matters") {
      const body = (await request.json()) as Record<string, string | undefined>
      writes.push({ method: "POST", body })
      const row: Row = { id: `m${matters.length + 1}`, title: body.title!, client: body.client, reference: body.reference, dir: "/tmp/matter", created_at: matters.length + 1 }
      matters.push(row)
      return Response.json(row)
    }
    const one = url.pathname.match(/^\/ingest\/matters\/(\w+)$/)
    const row = one && matters.find((m) => m.id === one[1])
    if (one && !row) return Response.json({ error: "not found" }, { status: 404 })
    if (row && request.method === "GET") return Response.json({ ...row, documents: [] })
    if (row && request.method === "PATCH") {
      const body = (await request.json()) as Record<string, string | undefined>
      writes.push({ method: "PATCH", id: row.id, body })
      Object.assign(row, { title: body.title, reference: body.reference })
      if (body.client !== undefined) row.client = body.client.trim() || undefined
      return Response.json(row)
    }
    // The sidebar's conversation list, and anything else, is not part of this check.
    return Response.json({ error: "not mocked" }, { status: 404 })
  },
})

const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH } : {}) })
if (shots) await mkdir(shots, { recursive: true })
const passed: string[] = []

async function scenario(name: string, seed: Row[], run: (page: Page) => Promise<void>, start = "/") {
  matters = seed.map((row) => ({ ...row }))
  writes.length = 0
  const page = await browser.newPage({ viewport: { width: 1280, height: 860 }, reducedMotion: "reduce" })
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

const shot = (page: Page, name: string) => (shots ? page.screenshot({ path: path.join(shots, `${name}.png`) }) : Promise.resolve(undefined))
const seed: Row[] = [
  { id: "a", title: "Lease review", client: "Northwind Traders", dir: "/tmp/a", created_at: 1 },
  { id: "b", title: "Supply note", client: "northwind  traders", reference: "2026-0042", dir: "/tmp/b", created_at: 2 },
  { id: "c", title: "Mutual NDA", client: "Contoso", dir: "/tmp/c", created_at: 3 },
  { id: "d", title: "Board minutes", dir: "/tmp/d", created_at: 4 },
]

try {
  await scenario("with no client on any matter, Clients shows the empty state", [{ ...seed[3]! }], async (page) => {
    await page.getByRole("navigation", { name: "Main navigation" }).getByRole("link", { name: "Clients" }).click()
    await expect(page.getByRole("heading", { name: "No clients yet" })).toBeVisible()
    await expect(page.getByText("Add a client on a matter.")).toBeVisible()
    await shot(page, "clients-empty")
  })

  await scenario("the Clients entry sits under Matters and lists names, counts and matter links", seed, async (page) => {
    const nav = page.getByRole("navigation", { name: "Main navigation" }).getByRole("link")
    assert.deepEqual(await nav.allTextContents(), ["Matters", "Clients"])
    await nav.filter({ hasText: "Clients" }).click()
    const list = page.getByRole("list", { name: "Clients", exact: true })
    const rows = list.locator(":scope > li")
    await expect(rows).toHaveCount(2)
    // Case and spacing variants are one client, under its most common (here first) spelling.
    await expect(rows.nth(0)).toContainText("Contoso")
    await expect(rows.nth(0)).toContainText("1 matter")
    await expect(rows.nth(1)).toContainText("Northwind Traders")
    await expect(rows.nth(1)).toContainText("2 matters")
    const northwind = page.getByRole("list", { name: "Matters for Northwind Traders" })
    assert.deepEqual(await northwind.getByRole("link").allTextContents(), ["Supply note", "Lease review"])
    await shot(page, "clients-list")
    // Clicking a client narrows the page to its matters.
    await list.getByRole("link", { name: "Northwind Traders" }).click()
    await expect(page.getByRole("heading", { name: "Northwind Traders" })).toBeVisible()
    await expect(page.getByRole("list", { name: "Matters for Northwind Traders" }).locator(":scope > li")).toHaveCount(2)
    await shot(page, "clients-filtered")
    await page.getByRole("link", { name: "Supply note" }).click()
    await expect(page.getByRole("button", { name: "Edit client: northwind  traders" })).toBeVisible()
  })

  await scenario("a new matter's Client suggests an existing client, picked by keyboard and saved", seed, async (page) => {
    await page.getByRole("textbox", { name: "Matter title" }).fill("Option deed")
    const field = page.getByRole("combobox", { name: "Client you act for (optional)" })
    await field.pressSequentially("nor")
    const listbox = page.getByRole("listbox", { name: "Existing clients" })
    await expect(listbox.getByRole("option")).toHaveText(["Northwind Traders"])
    await expect(field).toHaveAttribute("aria-expanded", "true")
    // Nothing is preselected, so Enter on free text would keep that text.
    assert.equal(await field.getAttribute("aria-activedescendant"), null)
    await page.keyboard.press("ArrowDown")
    const option = listbox.getByRole("option", { name: "Northwind Traders" })
    await expect(option).toHaveAttribute("aria-selected", "true")
    assert.equal(await field.getAttribute("aria-activedescendant"), await option.getAttribute("id"))
    await shot(page, "client-suggest-new-matter")
    await page.keyboard.press("Enter")
    await expect(field).toHaveValue("Northwind Traders")
    await expect(field).toHaveAttribute("aria-expanded", "false")
    assert.equal(writes.length, 0, "picking a suggestion must not submit the form")
    await page.keyboard.press("Enter")
    await expect.poll(() => writes.length).toBe(1)
    assert.equal(writes[0]!.body.client, "Northwind Traders")
  })

  await scenario("the edit dialog's Client suggests, closes on Escape without closing the dialog, and saves a pick", seed, async (page) => {
    await page.getByRole("button", { name: "Edit Board minutes" }).click()
    const dialog = page.getByRole("dialog", { name: "Edit matter" })
    const field = dialog.getByRole("combobox", { name: "Client" })
    await field.pressSequentially("tos")
    const listbox = page.getByRole("listbox", { name: "Existing clients" })
    await expect(listbox.getByRole("option")).toHaveText(["Contoso"])
    await page.keyboard.press("Escape")
    await expect(listbox).toHaveCount(0)
    await expect(dialog).toBeVisible()
    // ArrowDown reopens a closed list without choosing anything.
    await page.keyboard.press("ArrowDown")
    await expect(listbox.getByRole("option")).toHaveText(["Contoso"])
    assert.equal(await field.getAttribute("aria-activedescendant"), null)
    // A cleared field lists every client.
    await field.fill("")
    await expect(listbox.getByRole("option")).toHaveText(["Contoso", "Northwind Traders"])
    await page.keyboard.press("ArrowDown")
    await page.keyboard.press("ArrowDown")
    await page.keyboard.press("Enter")
    await expect(field).toHaveValue("Northwind Traders")
    await expect(dialog).toBeVisible()
    await page.keyboard.press("Enter")
    await expect(dialog).toHaveCount(0)
    assert.deepEqual(writes.map((w) => [w.method, w.id, w.body.client]), [["PATCH", "d", "Northwind Traders"]])
  })

  await scenario("in the matter header, a click picks a suggestion and a new free-text client saves and joins the Clients list", seed, async (page) => {
    await page.getByRole("button", { name: "Add the client you act for" }).click()
    const field = page.getByRole("combobox", { name: "Client you act for" })
    await field.pressSequentially("con")
    await shot(page, "client-suggest-header")
    await page.getByRole("listbox", { name: "Existing clients" }).getByRole("option", { name: "Contoso" }).click()
    await expect(field).toHaveValue("Contoso")
    await expect(field).toBeFocused()
    // Typing past a suggestion keeps the typed text: Enter saves it, not the suggestion.
    await field.pressSequentially(" Group")
    await expect(page.getByRole("listbox", { name: "Existing clients" })).toHaveCount(0)
    await field.fill("Fabrikam")
    await page.keyboard.press("Enter")
    await expect(page.getByRole("button", { name: "Edit client: Fabrikam" })).toBeVisible()
    assert.deepEqual(writes.map((w) => [w.method, w.id, w.body.client]), [["PATCH", "d", "Fabrikam"]])
    await page.getByRole("link", { name: "All matters", exact: true }).click()
    await page.getByRole("navigation", { name: "Main navigation" }).getByRole("link", { name: "Clients" }).click()
    const rows = page.getByRole("list", { name: "Clients", exact: true }).locator(":scope > li")
    await expect(rows).toHaveCount(3)
    await expect(rows.nth(1)).toContainText("Fabrikam")
    await expect(page.getByRole("list", { name: "Matters for Fabrikam" }).getByRole("link")).toHaveText(["Board minutes"])
  }, "/matter/d")

  console.log(`PASS ${passed.length} clients list cases with no browser errors`)
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
