// Run from apps/web: bun test/clients-list.browser.ts (optionally PLAYWRIGHT_MODULE / PLAYWRIGHT_CHROMIUM_PATH).
// The Clients list and the type-to-suggest matter Client field through the
// real sidebar, Matters and Clients pages and matter header; only ingest's
// matter and client HTTP is mocked, by a small in-memory store. Set
// CLIENTS_SCREENSHOTS=<directory> to save views.
import assert from "node:assert/strict"
import { mkdir, mkdtemp, readFile, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import type { Page } from "@playwright/test"
import { clientKey } from "../src/lib/clients"

const { chromium, expect } = (await import(process.env.PLAYWRIGHT_MODULE ?? "@playwright/test")) as typeof import("@playwright/test")
const directory = await mkdtemp(path.join(os.tmpdir(), "dochaus-clients-"))
const shots = process.env.CLIENTS_SCREENSHOTS
const source = path.resolve(import.meta.dir, "../src")
const file = (name: string) => JSON.stringify(path.join(source, name))
const dependency = (name: string) => JSON.stringify(fileURLToPath(import.meta.resolve(name)))

type Row = { id: string; title: string; client?: string; reference?: string; dir: string; created_at: number }
type Client = { id: number; name: string; reference?: string; created_at: number }
let matters: Row[] = []
let clients: Client[] = []
const writes: { method: string; id?: string; body: Record<string, unknown> }[] = []
const clientWrites: { method: string; body?: Record<string, unknown>; id?: string }[] = []

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
    if (route === "GET /ingest/clients") return Response.json(clients)
    // Like ingest: a name that is already a client returns it with 200.
    if (route === "POST /ingest/clients") {
      const body = (await request.json()) as { name: string; reference?: string }
      clientWrites.push({ method: "POST", body })
      const existing = clients.find((c) => clientKey(c.name) === clientKey(body.name))
      if (existing) return Response.json(existing)
      const client: Client = { id: clients.length + 1, name: body.name.trim(), ...(body.reference ? { reference: body.reference } : {}), created_at: 1 }
      clients.push(client)
      return Response.json(client, { status: 201 })
    }
    const clientRoute = url.pathname.match(/^\/ingest\/clients\/(\d+)$/)
    // Like ingest: an edit renames every matter of the old client, and a name
    // that is already another client merges into it.
    if ((clientRoute && request.method === "PATCH") || route === "POST /ingest/clients/rename") {
      const body = (await request.json()) as { from?: string; name: string; reference: string }
      clientWrites.push({ method: request.method, id: clientRoute?.[1], body })
      const current = clientRoute ? clients.find((c) => String(c.id) === clientRoute[1]) : clients.find((c) => clientKey(c.name) === clientKey(body.from!))
      const oldKey = clientKey(current?.name ?? body.from!)
      const target = clientKey(body.name) === oldKey ? undefined : clients.find((c) => clientKey(c.name) === clientKey(body.name))
      let client: Client
      if (target) {
        clients = clients.filter((c) => c !== current)
        client = target
      } else if (current) {
        client = Object.assign(current, { name: body.name, reference: body.reference || undefined })
      } else {
        client = { id: clients.length + 10, name: body.name, ...(body.reference ? { reference: body.reference } : {}), created_at: 1 }
        clients.push(client)
      }
      const moved = matters.filter((m) => m.client && clientKey(m.client) === oldKey)
      for (const m of moved) m.client = client.name
      return Response.json({ client, matters: moved.length, merged: Boolean(target) })
    }
    if (clientRoute && request.method === "DELETE") {
      clientWrites.push({ method: "DELETE", id: clientRoute[1] })
      clients = clients.filter((c) => String(c.id) !== clientRoute[1])
      return Response.json({ ok: true })
    }
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

async function scenario(name: string, seed: Row[], run: (page: Page) => Promise<void>, start = "/", storedSeed: Client[] = []) {
  matters = seed.map((row) => ({ ...row }))
  clients = storedSeed.map((row) => ({ ...row }))
  writes.length = 0
  clientWrites.length = 0
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
    await expect(page.getByText("Add the people and companies you act for here, or name a client on a matter.")).toBeVisible()
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
    await expect(listbox.getByRole("option")).toHaveText(["Northwind Traders", "Add “nor” as a client"])
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
    await expect(listbox.getByRole("option")).toHaveText(["Contoso", "Add “tos” as a client"])
    await page.keyboard.press("Escape")
    await expect(listbox).toHaveCount(0)
    await expect(dialog).toBeVisible()
    // ArrowDown reopens a closed list without choosing anything.
    await page.keyboard.press("ArrowDown")
    await expect(listbox.getByRole("option")).toHaveText(["Contoso", "Add “tos” as a client"])
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
    // Typing past a suggestion keeps the typed text: Enter saves it, not the
    // suggestion, and not the offer to add it as a client.
    await field.pressSequentially(" Group")
    await expect(page.getByRole("listbox", { name: "Existing clients" }).getByRole("option")).toHaveText(["Add “Contoso Group” as a client"])
    await field.fill("Fabrikam")
    await page.keyboard.press("Enter")
    await expect(page.getByRole("button", { name: "Edit client: Fabrikam" })).toBeVisible()
    assert.deepEqual(writes.map((w) => [w.method, w.id, w.body.client]), [["PATCH", "d", "Fabrikam"]])
    assert.deepEqual(clientWrites, [])
    await page.getByRole("link", { name: "All matters", exact: true }).click()
    await page.getByRole("navigation", { name: "Main navigation" }).getByRole("link", { name: "Clients" }).click()
    const rows = page.getByRole("list", { name: "Clients", exact: true }).locator(":scope > li")
    await expect(rows).toHaveCount(3)
    await expect(rows.nth(1)).toContainText("Fabrikam")
    await expect(page.getByRole("list", { name: "Matters for Fabrikam" }).getByRole("link")).toHaveText(["Board minutes"])
  }, "/matter/d")

  await scenario("a new client typed in the Client field is added by keyboard, selected, and suggested from then on", seed, async (page) => {
    await page.getByRole("textbox", { name: "Matter title" }).fill("Option deed")
    const field = page.getByRole("combobox", { name: "Client you act for (optional)" })
    await field.pressSequentially("Fabrikam  Inc")
    const listbox = page.getByRole("listbox", { name: "Existing clients" })
    await expect(listbox.getByRole("option")).toHaveText(["Add “Fabrikam Inc” as a client"])
    // Still nothing preselected: Enter here would keep the typed text.
    assert.equal(await field.getAttribute("aria-activedescendant"), null)
    await page.keyboard.press("ArrowDown")
    const add = listbox.getByRole("option", { name: "Add “Fabrikam Inc” as a client" })
    await expect(add).toHaveAttribute("aria-selected", "true")
    await shot(page, "client-add-from-field")
    await page.keyboard.press("Enter")
    await expect(field).toHaveValue("Fabrikam Inc")
    await expect(field).toHaveAttribute("aria-expanded", "false")
    assert.deepEqual(clientWrites, [{ method: "POST", body: { name: "Fabrikam Inc" } }])
    assert.equal(writes.length, 0, "adding a client must not submit the matter form")
    // The added client is now a suggestion, with no offer to add it again.
    await field.fill("fab")
    await expect(listbox.getByRole("option")).toHaveText(["Fabrikam Inc", "Add “fab” as a client"])
    await field.fill("fabrikam inc.")
    await expect(listbox).toHaveCount(0)
    await field.fill("Fabrikam Inc")
    await page.keyboard.press("Enter")
    await expect.poll(() => writes.length).toBe(1)
    assert.equal(writes[0]!.body.client, "Fabrikam Inc")
  })

  await scenario("adding a client someone else already added selects the existing client", seed, async (page) => {
    const field = page.getByRole("combobox", { name: "Client you act for (optional)" })
    await expect(page.getByRole("heading", { name: "Your matters" })).toBeVisible()
    // Added after this page loaded its clients, under another spelling.
    clients.push({ id: 9, name: "Tailspin Toys Ltd", created_at: 1 })
    await field.pressSequentially("tailspin toys ltd.")
    await page.getByRole("listbox", { name: "Existing clients" }).getByRole("option", { name: "Add “tailspin toys ltd.” as a client" }).click()
    await expect(field).toHaveValue("Tailspin Toys Ltd")
    await expect(field).toBeFocused()
    assert.equal(clients.length, 1)
  })

  await scenario("the Clients page adds a client from the form on the right, refuses a blank name and folds a duplicate", seed, async (page) => {
    const panel = page.getByRole("region", { name: "Add a new client" })
    const list = page.getByRole("list", { name: "Clients", exact: true })
    await expect(list.locator(":scope > li")).toHaveCount(2)
    const [form, collection] = await Promise.all([panel.boundingBox(), page.getByRole("region", { name: "Your clients" }).boundingBox()])
    assert.ok(form && collection && form.x > collection.x + collection.width - 1, "the add form sits to the right of the list")
    await panel.getByRole("button", { name: "Add client" }).click()
    await expect(panel.getByRole("alert")).toHaveText("Name the client to add it.")
    await expect(panel.getByRole("textbox", { name: "Client name" })).toBeFocused()
    await panel.getByRole("textbox", { name: "Client name" }).fill("Wide World Importers")
    await panel.getByRole("textbox", { name: "Client number (optional)" }).fill("C-0007")
    await page.keyboard.press("Enter")
    const rows = list.locator(":scope > li")
    await expect(rows).toHaveCount(3)
    await expect(rows.nth(2)).toContainText("Wide World Importers")
    await expect(rows.nth(2)).toContainText("C-0007")
    await expect(rows.nth(2)).toContainText("No matters yet")
    await expect(panel.getByRole("textbox", { name: "Client name" })).toHaveValue("")
    await shot(page, "clients-page")
    // The same client again is no second row and no error.
    await panel.getByRole("textbox", { name: "Client name" }).fill("wide world importers.")
    await panel.getByRole("button", { name: "Add client" }).click()
    await expect(page.getByText("\"Wide World Importers\" is already a client.")).toBeVisible()
    await expect(rows).toHaveCount(3)
    // An added client with no matters can be deleted; clients with matters cannot.
    await expect(rows.nth(0).getByRole("button")).toHaveText(["Edit"])
    await rows.nth(2).getByRole("button", { name: "Delete Wide World Importers" }).click()
    await rows.nth(2).getByRole("button", { name: "Confirm delete Wide World Importers" }).click()
    await expect(rows).toHaveCount(2)
    assert.deepEqual(clientWrites.map((w) => w.method), ["POST", "POST", "DELETE"])
  }, "/clients")

  await scenario("an added client joins the matter client of the same name and a client with no matters says so", seed, async (page) => {
    const rows = page.getByRole("list", { name: "Clients", exact: true }).locator(":scope > li")
    await expect(rows).toHaveCount(3)
    await expect(rows.nth(1)).toContainText("Fabrikam")
    await expect(rows.nth(1)).toContainText("No matters yet")
    // Stored "NORTHWIND TRADERS" is the name shown for the matters' client.
    await expect(rows.nth(2)).toContainText("NORTHWIND TRADERS")
    await expect(rows.nth(2)).toContainText("2 matters")
    await rows.nth(1).getByRole("link", { name: "Fabrikam" }).click()
    await expect(page.getByText("No matters for Fabrikam yet.")).toBeVisible()
  }, "/clients", [{ id: 1, name: "Fabrikam", created_at: 1 }, { id: 2, name: "NORTHWIND TRADERS", created_at: 1 }])

  await scenario("Matters and Clients share one layout for side-by-side review", seed, async (page) => {
    await expect(page.getByRole("heading", { name: "Your matters" })).toBeVisible()
    await shot(page, "matters-page")
    await page.getByRole("navigation", { name: "Main navigation" }).getByRole("link", { name: "Clients" }).click()
    await expect(page.getByRole("list", { name: "Clients", exact: true }).locator(":scope > li")).toHaveCount(3)
    await shot(page, "clients-page-with-added")
    // Narrow, the add form folds behind a New client toggle, as on Matters.
    await page.setViewportSize({ width: 420, height: 860 })
    await page.getByRole("button", { name: "New client" }).click()
    await expect(page.getByRole("textbox", { name: "Client name" })).toBeVisible()
    await shot(page, "clients-page-narrow")
  }, "/", [{ id: 1, name: "Fabrikam", reference: "C-0101", created_at: 1 }])

  await scenario("every client row has Edit; editing a matter-only client renames it on its matters and stores its number", seed, async (page) => {
    const list = page.getByRole("list", { name: "Clients", exact: true })
    const rows = list.locator(":scope > li")
    await expect(rows.nth(0).getByRole("button")).toHaveText(["Edit"])
    await expect(rows.nth(1).getByRole("button")).toHaveText(["Edit"])
    await rows.nth(1).getByRole("button", { name: "Edit Northwind Traders" }).click()
    const dialog = page.getByRole("dialog", { name: "Edit client" })
    const name = dialog.getByRole("textbox", { name: "Client name" })
    await expect(name).toHaveValue("Northwind Traders")
    await expect(name).toBeFocused()
    await expect(dialog.getByText("A new name also changes the Client on its 2 matters.")).toBeVisible()
    // A blank name is refused in place, as on the matter dialog.
    await name.fill(" ")
    await dialog.getByRole("button", { name: "Save changes" }).click()
    await expect(dialog.getByRole("alert")).toHaveText("Name the client to save it.")
    await expect(name).toBeFocused()
    await name.fill("Northwind Traders Ltd")
    await dialog.getByRole("textbox", { name: "Client number" }).fill("C-0042")
    await shot(page, "clients-edit-dialog")
    await page.keyboard.press("Enter")
    await expect(dialog).toHaveCount(0)
    await expect(page.getByText("Updated client \"Northwind Traders Ltd\".")).toBeVisible()
    assert.deepEqual(clientWrites, [{ method: "POST", id: undefined, body: { from: "Northwind Traders", name: "Northwind Traders Ltd", reference: "C-0042" } }])
    await expect(rows.nth(1)).toContainText("Northwind Traders Ltd")
    await expect(rows.nth(1)).toContainText("C-0042")
    await expect(rows.nth(1)).toContainText("2 matters")
    await expect(page.getByRole("list", { name: "Matters for Northwind Traders Ltd" }).getByRole("link")).toHaveText(["Supply note", "Lease review"])
    // The Matters page and the matter header show the new name.
    await page.getByRole("navigation", { name: "Main navigation" }).getByRole("link", { name: "Matters" }).click()
    await expect(page.locator(".matter-client")).toHaveText(["Contoso", "Northwind Traders Ltd", "Northwind Traders Ltd"])
    await page.getByRole("link", { name: /^Supply note/ }).click()
    await expect(page.getByRole("button", { name: "Edit client: Northwind Traders Ltd" })).toBeVisible()
  }, "/clients")

  await scenario("renaming a client to another client's name says it will merge, then moves its matters there", seed, async (page) => {
    const rows = page.getByRole("list", { name: "Clients", exact: true }).locator(":scope > li")
    await expect(rows).toHaveCount(3)
    await rows.nth(0).getByRole("button", { name: "Edit Contoso" }).click()
    const dialog = page.getByRole("dialog", { name: "Edit client" })
    await dialog.getByRole("textbox", { name: "Client name" }).fill("northwind traders.")
    await expect(dialog.getByRole("status")).toHaveText("Northwind Traders already exists; Contoso’s 1 matter will move to it.")
    await shot(page, "clients-edit-merge")
    await dialog.getByRole("button", { name: "Save changes" }).click()
    await expect(dialog).toHaveCount(0)
    await expect(page.getByText("Merged \"Contoso\" into \"Northwind Traders\".")).toBeVisible()
    assert.deepEqual(clientWrites, [{ method: "PATCH", id: "1", body: { name: "Northwind Traders", reference: "C-1" } }])
    await expect(rows).toHaveCount(2)
    await expect(rows.nth(1)).toContainText("Northwind Traders")
    await expect(rows.nth(1)).toContainText("3 matters")
    // Cancel and Escape leave a client unchanged.
    await rows.nth(1).getByRole("button", { name: "Edit Northwind Traders" }).click()
    await dialog.getByRole("textbox", { name: "Client name" }).fill("Someone else")
    await dialog.getByRole("button", { name: "Cancel" }).click()
    await rows.nth(0).getByRole("button", { name: "Edit Fabrikam" }).click()
    await page.keyboard.press("Escape")
    await expect(dialog).toHaveCount(0)
    assert.equal(clientWrites.length, 1)
  }, "/clients", [{ id: 1, name: "Contoso", reference: "C-1", created_at: 1 }, { id: 2, name: "Fabrikam", created_at: 1 }, { id: 3, name: "Northwind Traders", created_at: 1 }])

  await scenario("the Matters and Clients edit dialogs match side by side", seed, async (page) => {
    await page.getByRole("button", { name: "Edit Supply note" }).click()
    const matterDialog = page.getByRole("dialog", { name: "Edit matter" })
    await expect(matterDialog).toBeVisible()
    const matterBox = await matterDialog.boundingBox()
    await expect(matterDialog.getByRole("textbox", { name: "Title" })).toBeFocused()
    await shot(page, "matters-edit-dialog")
    await page.keyboard.press("Escape")
    await page.getByRole("navigation", { name: "Main navigation" }).getByRole("link", { name: "Clients" }).click()
    await page.getByRole("button", { name: "Edit Northwind Traders" }).click()
    const clientDialog = page.getByRole("dialog", { name: "Edit client" })
    await expect(clientDialog).toBeVisible()
    await expect(clientDialog.getByRole("textbox", { name: "Client name" })).toBeFocused()
    const clientBox = await clientDialog.boundingBox()
    assert.equal(clientBox?.width, matterBox?.width, "both dialogs share the 480px width")
    await shot(page, "clients-edit-dialog-side-by-side")
    // Narrow, the dialog keeps its gutter as on Matters.
    await page.setViewportSize({ width: 420, height: 860 })
    const narrow = await clientDialog.boundingBox()
    assert.ok(narrow && narrow.x >= 15 && narrow.x + narrow.width <= 405, "the dialog fits a narrow screen")
    await shot(page, "clients-edit-dialog-narrow")
    await page.keyboard.press("Escape")
    await expect(page.getByRole("button", { name: "Edit Northwind Traders" })).toBeVisible()
    await shot(page, "clients-page-narrow-rows")
  }, "/", [{ id: 1, name: "Fabrikam", reference: "C-0101", created_at: 1 }])

  console.log(`PASS ${passed.length} clients list cases with no browser errors`)
} finally {
  await browser.close()
  server.stop(true)
  await rm(directory, { recursive: true, force: true })
}
