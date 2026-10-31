import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"
import type { Document } from "../src/api/ingest/documents"
import DocumentUpload from "../src/components/DocumentUpload"
import { WordDocumentsProvider } from "../src/components/WordDocuments"

test("uploaded document menus keep Word visible, dismiss accessibly and require the existing deletion choices", async () => {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, {
    configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window],
  }])))
  const originalFetch = globalThis.fetch
  const { createRoot } = await import("react-dom/client")
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  const changes: string[] = []
  const deletes: string[] = []
  const documents: Document[] = [
    { id: 1, name: "Agreement.docx", doc_path: "/m/Agreement.docx", created_at: 1, role: "source" },
    { id: 2, name: "Thread.eml", doc_path: "/m/Thread.eml", created_at: 2, role: "source", removal: {
      attachments: ["Schedule.docx"], kept: [{ name: "Client.pdf", reason: "changed" }],
    } },
    { id: 3, name: "Schedule.docx", doc_path: "/m/Schedule.docx", created_at: 3, role: "source", attachment_of: { name: "Thread.eml", removed: false } },
  ]
  globalThis.fetch = (async (input, init) => {
    const url = String(input)
    if (init?.method === "DELETE") { deletes.push(url); return Response.json({}) }
    if (url.endsWith("/word/config")) return Response.json({ available: true, origin: "https://localhost:4502", manifestUrl: "https://localhost:4502/manifest.xml" })
    if (url.endsWith("/documents/word")) return Response.json({ documents: [{
      name: "Schedule.docx", open: true, overlap: false, currentVersion: "v-1",
      lease: { state: "other", generation: "lease-1", expiresAt: Date.now() + 60_000 },
    }] })
    throw new Error(`Unexpected synthetic request: ${url}`)
  }) as typeof fetch
  const button = (label: string) => container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)!
  const menu = (name: string) => container.querySelector<HTMLElement>(`summary[aria-label="More actions for ${name}"]`)!.closest("details")!
  const textButton = (label: string) => [...container.querySelectorAll("button")].find(node => node.textContent === label)!
  try {
    await act(async () => root.render(createElement(WordDocumentsProvider, { matterId: "m", onChanged: () => {}, children:
      createElement(DocumentUpload, { matterId: "m", documents, onUploaded: () => {}, onView: () => {}, onCompare: name => changes.push(name) }),
    })))
    expect(button("Open Agreement.docx in Word").closest("details")).toBeNull()
    expect(container.textContent).not.toContain("Disconnect Word")
    expect(textButton("Take over editing")).toBeDefined()
    expect(textButton("Release Word lock")).toBeDefined()
    expect(container.querySelectorAll(".doc-more-actions")).toHaveLength(3)
    const agreement = menu("Agreement.docx")
    expect(agreement.open).toBe(false)
    expect(button("Compare Agreement.docx with another version").closest("details")).toBe(agreement)
    expect(button("Delete Agreement.docx").closest("details")).toBe(agreement)
    // The drawer scrolls: the last card's menu must open upward instead of clipping.
    const scrollArea = container.querySelector<HTMLElement>(".docs-drawer-body")!
    const popup = agreement.querySelector<HTMLElement>(".viewer-more-menu")!
    scrollArea.getBoundingClientRect = () => ({ top: 100, bottom: 500 }) as DOMRect
    agreement.getBoundingClientRect = () => ({ top: 450, bottom: 482 }) as DOMRect
    popup.getBoundingClientRect = () => ({ height: 94 }) as DOMRect
    agreement.open = true
    act(() => agreement.dispatchEvent(new browser.Event("toggle") as unknown as Event))
    expect(agreement.dataset.above).toBe("true")
    expect(popup.style.maxHeight).toBe("344px")
    agreement.getBoundingClientRect = () => ({ top: 110, bottom: 142 }) as DOMRect
    act(() => agreement.dispatchEvent(new browser.Event("toggle") as unknown as Event))
    expect(agreement.dataset.above).toBe("false")
    act(() => scrollArea.dispatchEvent(new browser.Event("scroll") as unknown as Event))
    expect(agreement.open).toBe(false)
    agreement.open = true
    act(() => button("Compare Agreement.docx with another version").click())
    expect(changes).toEqual(["Agreement.docx"])
    expect(agreement.open).toBe(false)
    agreement.open = true
    act(() => agreement.dispatchEvent(new browser.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as KeyboardEvent))
    expect(agreement.open).toBe(false)
    expect(document.activeElement).toBe(agreement.querySelector("summary"))
    agreement.open = true
    act(() => document.body.dispatchEvent(new browser.PointerEvent("pointerdown", { bubbles: true }) as unknown as PointerEvent))
    expect(agreement.open).toBe(false)
    agreement.open = true
    act(() => button("Delete Agreement.docx").click())
    expect(deletes).toEqual([])
    expect(agreement.open).toBe(false)
    await act(async () => button("Confirm deletion of Agreement.docx").click())
    expect(deletes).toEqual([expect.stringContaining("/documents?name=Agreement.docx")])
    act(() => button("Delete Thread.eml").click())
    expect(container.textContent).toContain("Client.pdf is kept: it has changed since it was added")
    expect(textButton("Remove email and 1 attachment")).toBeDefined()
    expect(textButton("Remove email only")).toBeDefined()
    expect(deletes).toHaveLength(1)
    await act(async () => textButton("Remove email only").click())
    expect(deletes.at(-1)).toContain("name=Thread.eml&attachments=keep")
  } finally {
    await act(async () => root.unmount())
    container.remove()
    globalThis.fetch = originalFetch
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
    await browser.happyDOM.close()
  }
})

test("saved Word recovery explains the isolated copy and refreshes the document list", async () => {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, {configurable:true,writable:true,value:name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window]}])))
  const originalFetch = globalThis.fetch
  const { createRoot } = await import("react-dom/client")
  const { WordOpenControl } = await import("../src/components/WordDocuments")
  const container = document.createElement("div"); document.body.append(container)
  const root = createRoot(container)
  let changed = 0, copies = 0
  globalThis.fetch = (async (input, init) => {
    const url = String(input)
    if (url.endsWith("/word/config")) return Response.json({available:true,origin:"https://localhost:4502",manifestUrl:"https://localhost:4502/manifest.xml"})
    if (url.endsWith("/documents/word")) return Response.json({documents:[{name:"Agreement.docx",open:false,overlap:false,currentVersion:"v1",currentNumber:1,documentId:"document",nativeOperation:{state:"pending",sequence:1},lease:{state:"other",expiresAt:null}}]})
    if (url.endsWith("/documents/word/recovery-copy") && init?.method === "POST") { copies++; return Response.json({name:"Agreement (saved copy).docx"}) }
    throw new Error("Unexpected request")
  }) as typeof fetch
  const button = (label:string) => [...container.querySelectorAll("button")].find(node => node.textContent === label)!
  try {
    await act(async () => root.render(createElement(WordDocumentsProvider, {matterId:"m",onChanged:() => {changed++},children:createElement(WordOpenControl,{matterId:"m",name:"Agreement.docx"})})))
    expect(container.querySelector<HTMLButtonElement>('button[aria-label="Open Agreement.docx in Word"]')!.disabled).toBe(true)
    expect(container.textContent).not.toContain("Take over editing")
    await act(async () => button("Create saved copy").click())
    expect(container.textContent).toContain("Unsent local Word edits and the unconfirmed change are not included")
    await act(async () => button("Create copy").click())
    expect(copies).toBe(1)
    expect(changed).toBe(1)
  } finally {
    await act(async () => root.unmount())
    globalThis.fetch = originalFetch
    for (const [name, descriptor] of saved) descriptor ? Object.defineProperty(globalThis,name,descriptor) : delete (globalThis as Record<string, unknown>)[name]
    browser.happyDOM.abort()
  }
})
