import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement, StrictMode } from "react"
import { MemoryRouter } from "react-router-dom"
import type { BatchSheet } from "../src/api/ingest/template-batch"

// Synthetic template fields and spreadsheet; no real client data.
const fields: BatchSheet["fields"] = [
  { key: "[Client Name]", label: "Client Name", kind: "fill" },
  { key: "[insert notice date]", label: "Notice date", kind: "fill" },
  { key: "[optional: Non-solicit]", label: "Include: Non-solicit", kind: "optional" },
]
const sheet = (mapping: BatchSheet["mapping"]): BatchSheet => ({
  template: { name: "Mutual NDA.docx", sha256: "abc" },
  fields,
  unnamedBlanks: 0,
  columns: ["Client Name", "Notice date", "Due"],
  rows: [
    { row: 2, values: ["Northwind Ltd", "1 March 2026", "x"] },
    { row: 3, values: ["Fabrikam LLP", "", "x"] },
    { row: 4, values: ["Contoso plc", "2 March 2026", "x"] },
  ],
  dropped: { blank: 0, header: 0 },
  mapping,
  remembered: false,
  identifyingKey: "[Client Name]",
  source: "clients.csv",
})

async function render(read: BatchSheet, props: { matter?: string; onCreated?: () => void } = {}) {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "HTMLDialogElement", "Node", "Event", "getComputedStyle", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true,
    value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const originalFetch = globalThis.fetch
  const created: unknown[] = []
  const requests: string[] = []
  globalThis.fetch = (async (input, init) => {
    const url = String(input)
    requests.push(url)
    if (url.endsWith("/matters")) return Response.json([{ id: "m1", title: "Supplier NDAs", dir: "/m1", created_at: 2 }])
    if (url.endsWith("/templates/batch/read")) return Response.json(read)
    if (url.endsWith("/templates/batch/preview")) return new Response(new Uint8Array([1]), { headers: { "x-dochaus-batch-missing": encodeURIComponent("[]") } })
    if (url.endsWith("/matters/m1/template-batches")) {
      created.push(JSON.parse(String(init?.body)))
      return Response.json({ id: "b1", folder: "Mutual NDA – 2026-10-31", documents: [{ name: "a.docx", row: 2, missing: [] }, { name: "b.docx", row: 3, missing: ["Notice date"] }], skipped: [4], failed: [] })
    }
    return new Response("not found", { status: 404 })
  }) as typeof fetch
  const { createRoot } = await import("react-dom/client")
  const { default: TemplateBatchDialog } = await import("../src/components/TemplateBatchDialog")
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  await act(async () => {
    // StrictMode as in main.tsx: its double mount once dropped every result.
    root.render(createElement(StrictMode, null, createElement(MemoryRouter, null, createElement(TemplateBatchDialog, { template: { name: "Mutual NDA.docx", description: "", placeholders: [] }, onClose: () => {}, ...props }))))
  })
  const drop = async () => {
    const input = document.body.querySelector<HTMLInputElement>("input[type=file]")!
    Object.defineProperty(input, "files", { configurable: true, value: [new File(["x"], "clients.csv")] })
    await act(async () => { input.dispatchEvent(new browser.Event("change", { bubbles: true }) as unknown as Event) })
  }
  const button = (label: string | RegExp) => Array.from(document.body.querySelectorAll("button")).find(item => typeof label === "string" ? item.textContent === label : label.test(item.textContent ?? ""))
  const cleanup = async () => {
    await act(async () => root.unmount())
    container.remove()
    globalThis.fetch = originalFetch
    await browser.happyDOM.close()
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
  return { drop, button, created, requests, cleanup, text: () => document.body.textContent ?? "" }
}

test("every fill field matched skips the Match step; Create counts and sends only rows not skipped", async () => {
  const view = await render(sheet({ "[Client Name]": "Client Name", "[insert notice date]": "Notice date", "[optional: Non-solicit]": null }))
  try {
    await view.drop()
    expect(view.text()).not.toContain("Match")
    expect(view.text()).toContain("1 of 3")
    expect(view.text()).toContain("Northwind Ltd")
    // Fabrikam leaves the notice date blank: counted, not created blank.
    expect(view.text()).toContain("2 ready")
    expect(view.text()).toContain("1 need input")
    expect(view.button("Create 3 documents")).toBeDefined()

    await act(async () => { document.body.querySelector<HTMLButtonElement>("[aria-label='Next row']")!.click() })
    await act(async () => { document.body.querySelector<HTMLButtonElement>("[aria-label='Next row']")!.click() })
    expect(view.text()).toContain("Contoso plc")
    await act(async () => { document.body.querySelector<HTMLButtonElement>("[role=switch]")!.click() })
    expect(view.text()).toContain("1 skipped")
    expect(view.text()).toContain("1 ready")
    await act(async () => { view.button("Create 2 documents")!.click() })

    expect(view.created).toHaveLength(1)
    expect(view.created[0]).toMatchObject({ name: "Mutual NDA.docx", sha256: "abc", source: "clients.csv", skipped: [4] })
    expect(view.text()).toContain("2 documents created in Mutual NDA – 2026-10-31")
    expect(view.button("Open matter")).toBeDefined()
  } finally {
    await view.cleanup()
  }
})

test("an unmatched fill field asks only about that field, then previews", async () => {
  const view = await render(sheet({ "[Client Name]": "Client Name", "[insert notice date]": null, "[optional: Non-solicit]": null }))
  try {
    await view.drop()
    expect(view.text()).toContain("Match 1 field")
    const labels = Array.from(document.body.querySelectorAll(".template-batch-match label")).map(label => label.textContent)
    expect(labels).toEqual(["Notice date"])
    expect(view.text()).toContain("1 field matched")
    // Left unmatched, every row needs the notice date.
    await act(async () => { view.button("Preview")!.click() })
    expect(view.text()).toContain("0 ready")
    expect(view.text()).toContain("3 need input")
    expect(view.button("Change matching")).toBeDefined()
  } finally {
    await view.cleanup()
  }
})

test("opened inside a matter, it asks for no matter and creates into that one", async () => {
  let refreshed = 0
  const view = await render(sheet({ "[Client Name]": "Client Name", "[insert notice date]": "Notice date", "[optional: Non-solicit]": null }), { matter: "m1", onCreated: () => { refreshed++ } })
  try {
    expect(document.body.querySelector("select")).toBeNull()
    expect(view.requests.some(url => url.endsWith("/matters"))).toBe(false)
    await view.drop()
    await act(async () => { view.button("Create 3 documents")!.click() })
    expect(view.created).toHaveLength(1)
    expect(refreshed).toBe(1)
    expect(view.button("View documents")).toBeDefined()
    expect(view.button("Open matter")).toBeUndefined()
  } finally {
    await view.cleanup()
  }
})
