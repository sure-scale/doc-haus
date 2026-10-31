import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"
import { WordDocumentsProvider, WordOpenControl } from "../src/components/WordDocuments"
import { ToastProvider } from "../src/components/Toast"

test("Open in Word uses the desktop handoff; failure exposes a deliberate linked-download fallback", async () => {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = [...names, "setTimeout", "clearTimeout"].map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true,
    value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const originalFetch = globalThis.fetch
  const originalClick = browser.HTMLAnchorElement.prototype.click
  const originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL
  const links: Array<{ href: string; download: string }> = []
  browser.HTMLAnchorElement.prototype.click = function () { links.push({ href: this.href, download: this.download }) }
  URL.createObjectURL = () => "blob:synthetic-linked-word"
  URL.revokeObjectURL = () => {}
  const { createRoot } = await import("react-dom/client")
  // Let React retain its native scheduler, then own the toasts' global timers.
  Object.defineProperties(globalThis, {
    setTimeout: { configurable: true, writable: true, value: browser.setTimeout.bind(browser) },
    clearTimeout: { configurable: true, writable: true, value: browser.clearTimeout.bind(browser) },
  })
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  const id = "11111111-2222-4333-8444-555555555555"
  const uri = `doc.haus://word/open?source=${encodeURIComponent(`https://localhost:4511/word/open/${"a".repeat(43)}`)}`
  let state = "failed", downloads = 0, handoffs = 0
  globalThis.fetch = (async (input, init) => {
    const url = String(input)
    if (url.endsWith("/word/config")) return Response.json({ available: true, origin: "https://localhost:4511", manifestUrl: "https://localhost:4511/word-addin/manifest.xml" })
    if (url.includes("/word/open/status?")) return Response.json({ state })
    if (url.includes("/word/open?")) { handoffs++; return Response.json({ uri, id }) }
    if (init?.method === "POST") { downloads++; return new Response(new Uint8Array([80, 75, 3, 4])) }
    return Response.json({ documents: [] })
  }) as typeof fetch
  const button = (label: string) => Array.from(container.querySelectorAll("button")).find(item => item.textContent === label)!
  const render = (name: string) => createElement(ToastProvider, null, createElement(WordDocumentsProvider, { matterId: "m", onChanged: () => {}, children: createElement(WordOpenControl, { matterId: "m", name }) }))
  try {
    await act(async () => { root.render(render("Draft.docx")) })
    await act(async () => { button("Open in Word").click() })
    expect(handoffs).toBe(1)
    expect(downloads).toBe(0)
    expect(links).toEqual([{ href: uri, download: "" }])
    expect(container.textContent).toContain("Word could not be opened automatically")
    expect(button("Download Word copy")).toBeDefined()
    await act(async () => { button("Download Word copy").click() })
    expect(downloads).toBe(1)
    expect(links.at(-1)).toEqual({ href: "blob:synthetic-linked-word", download: "Draft.docx" })
    state = "completed"
    await act(async () => { button("Open in Word").click() })
    expect(handoffs).toBe(2)
    expect(downloads).toBe(1)
    expect(container.textContent).toContain("Sent to Word")
    expect(container.textContent).not.toContain("Saved")
    expect(button("Download Word copy")).toBeUndefined()
    state = "pending"
    await act(async () => { button("Open in Word").click() })
    expect(button("Opening Word…").disabled).toBe(true)
    await act(async () => { root.render(render("Another draft.docx")) })
    expect(button("Open in Word").disabled).toBe(false)
    expect(button("Download Word copy")).toBeUndefined()
  } finally {
    await act(async () => root.unmount())
    container.remove()
    await browser.happyDOM.close()
    globalThis.fetch = originalFetch
    browser.HTMLAnchorElement.prototype.click = originalClick
    URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
})
