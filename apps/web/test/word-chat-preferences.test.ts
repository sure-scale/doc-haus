import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"
import { WordChatDefaults } from "../src/components/settings/WordAddin"

test("Word chat defaults load from installation settings and save deliberate choices without losing a failed edit", async () => {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const originalFetch = globalThis.fetch
  const { createRoot } = await import("react-dom/client")
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container), busy: boolean[] = [], calls: Array<{ url: string; method: string; body?: BodyInit | null }> = []
  let preferences = { defaultScope: "matter", suggestions: false }, fail = false
  globalThis.fetch = (async (url, init) => {
    calls.push({ url: String(url), method: init?.method ?? "GET", body: init?.body })
    if (init?.method === "PUT") {
      if (fail) return Response.json({ error: "Saved settings are temporarily unavailable." }, { status: 503 })
      preferences = JSON.parse(String(init.body))
    }
    return Response.json(preferences)
  }) as typeof fetch
  const button = (text: string) => Array.from(container.querySelectorAll("button")).find(element => element.textContent === text)!
  try {
    await act(async () => root.render(createElement(WordChatDefaults, { onBusy: value => busy.push(value) })))
    expect(calls).toHaveLength(1)
    expect(calls[0]!.url).toEndWith("/word/chat-preferences")
    expect(button("Whole matter").getAttribute("aria-pressed")).toBe("true")
    expect(container.querySelector('[role="switch"]')?.getAttribute("aria-checked")).toBe("false")
    expect(button("Save chat defaults").disabled).toBe(true)
    expect(container.textContent).not.toContain("Suggest about")
    act(() => { button("This document").click(); (container.querySelector('[role="switch"]') as HTMLButtonElement).click() })
    expect(calls).toHaveLength(1)
    fail = true
    await act(async () => button("Save chat defaults").click())
    expect(calls.at(-1)).toMatchObject({ method: "PUT", body: '{"defaultScope":"document","suggestions":true}' })
    expect(container.querySelector('[role="alert"]')?.textContent).toContain("Saved settings are temporarily unavailable")
    expect(button("This document").getAttribute("aria-pressed")).toBe("true")
    expect(button("Save chat defaults").disabled).toBe(false)
    expect(preferences).toEqual({ defaultScope: "matter", suggestions: false })
    fail = false
    await act(async () => button("Save chat defaults").click())
    expect(preferences).toEqual({ defaultScope: "document", suggestions: true })
    expect(container.textContent).toContain("Chat defaults saved.")
    expect(container.querySelector('[role="alert"]')).toBeNull()
    expect(button("Save chat defaults").disabled).toBe(true)
    expect(busy).toContain(true)
    expect(busy.at(-1)).toBe(false)
    await act(async () => root.render(createElement(WordChatDefaults, { key: "reloaded" })))
    expect(button("This document").getAttribute("aria-pressed")).toBe("true")
    expect(container.querySelector('[role="switch"]')?.getAttribute("aria-checked")).toBe("true")
  } finally {
    await act(async () => root.unmount())
    container.remove(); globalThis.fetch = originalFetch
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
    await browser.happyDOM.close()
  }
})
