import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement, StrictMode } from "react"
import { MemoryRouter } from "react-router-dom"
import type { Template } from "../src/api/ingest/library"

// Synthetic library entries; no real client data.
const letter: Template = { name: "Consent request letter.docx", description: "", placeholders: [{ text: "[Supplier Name]", kind: "fill" }] }
const minutes: Template = { name: "Board minutes.docx", description: "Fixed wording.", placeholders: [] }

async function render(templates: Template[]) {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "getComputedStyle", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true,
    value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const originalFetch = globalThis.fetch
  globalThis.fetch = (async (input) => String(input).endsWith("/templates")
    ? Response.json({ templates })
    : new Response("not found", { status: 404 })) as typeof fetch
  const drafted: string[] = []
  const batched: string[] = []
  const { createRoot } = await import("react-dom/client")
  const { default: TemplateStartDialog } = await import("../src/components/matter/TemplateStartDialog")
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  await act(async () => {
    root.render(createElement(StrictMode, null, createElement(MemoryRouter, null, createElement(TemplateStartDialog, {
      onDraft: (t: Template) => drafted.push(t.name), onBatch: (t: Template) => batched.push(t.name), onClose: () => {},
    }))))
  })
  const click = async (label: RegExp) => {
    const target = Array.from(document.body.querySelectorAll("button")).find(item => label.test(item.textContent ?? ""))
    if (!target) throw new Error(`no button ${label}`)
    await act(async () => { target.click() })
  }
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
  return { click, drafted, batched, cleanup, text: () => document.body.textContent ?? "" }
}

test("a template with fields asks one or many; one with nothing to fill goes straight to drafting", async () => {
  const view = await render([letter, minutes])
  try {
    expect(view.text()).toContain("1 placeholder")
    await view.click(/^Board minutes/)
    expect(view.drafted).toEqual(["Board minutes.docx"])
    await view.click(/^Consent request letter/)
    expect(view.text()).toContain("Many from a spreadsheet")
    await view.click(/^Many from a spreadsheet/)
    expect(view.batched).toEqual(["Consent request letter.docx"])
    expect(view.drafted).toEqual(["Board minutes.docx"])
  } finally {
    await view.cleanup()
  }
})

test("an empty library points to the Templates page", async () => {
  const view = await render([])
  try {
    expect(view.text()).toContain("No templates yet.")
    expect(document.body.querySelector("a[href='/templates']")).not.toBeNull()
  } finally {
    await view.cleanup()
  }
})
