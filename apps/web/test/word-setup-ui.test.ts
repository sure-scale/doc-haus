import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"
import type { WordSetupStatus } from "../src/api/ingest/word"
import WordAddin from "../src/components/settings/WordAddin"

test("Word setup waits for a deliberate action, holds pending state and confirms a lost response without repeating setup", async () => {
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
  const calls: string[] = []
  const busy: boolean[] = []
  let changes = 0
  window.addEventListener("dochaus:word-setup", () => changes++)
  const initial: WordSetupStatus = {
    platform: "macos", wordInstalled: true, canSetup: true, configured: false, registered: false,
    trusted: false, running: false, available: false, status: "not_configured", origin: null, manifestUrl: null,
  }
  const ready: WordSetupStatus = { ...initial, configured: true, registered: true, trusted: true,
    running: true, available: true, status: "ready", restartWord: true }
  let current = initial
  let finish: ((response: Response) => void) | undefined
  let loseResponse = false
  let failStatus = false
  globalThis.fetch = (async (_url, init) => {
    calls.push(init?.method ?? "GET")
    if (init?.method === "POST") {
      expect(init.body).toBe("{}")
      if (loseResponse) throw new TypeError("Synthetic lost response")
      return new Promise<Response>(resolve => { finish = resolve })
    }
    if (failStatus) throw new TypeError("Synthetic stopped service")
    return Response.json(current)
  }) as typeof fetch
  const findButton = (text: string) => Array.from(container.querySelectorAll("button")).find(button => button.textContent === text)!
  try {
    await act(async () => { root.render(createElement(WordAddin, { onBusy: value => busy.push(value) })) })
    expect(calls).toEqual(["GET"])
    act(() => findButton("Set up Word").click())
    expect(calls).toEqual(["GET", "POST"])
    expect(container.textContent).not.toContain("Check status")
    expect(container.textContent).toContain("Setting up Word")
    expect(busy.at(-1)).toBe(true)
    await act(async () => { current = ready; finish!(Response.json(ready)) })
    expect(container.textContent).toContain("Ready — restart Word")
    expect(changes).toBe(1)
    expect(busy.at(-1)).toBe(false)
    loseResponse = true
    current = initial
    await act(async () => { window.dispatchEvent(new Event("focus")) })
    await act(async () => { findButton("Set up Word").click() })
    expect(calls).toEqual(["GET", "POST", "GET", "POST", "GET"])
    expect(container.querySelector('[role="alert"]')).not.toBeNull()
    expect(container.textContent).not.toContain("Ready")
    // The native operation finishes later. Checking status updates an already
    // mounted matter's configuration without repeating installation or trust.
    current = ready
    await act(async () => { window.dispatchEvent(new Event("focus")) })
    expect(calls).toEqual(["GET", "POST", "GET", "POST", "GET", "GET"])
    expect(changes).toBe(2)
    expect(container.querySelector('[role="alert"]')).toBeNull()
    failStatus = true
    await act(async () => { window.dispatchEvent(new Event("focus")) })
    expect(container.querySelector('[role="alert"]')).not.toBeNull()
    expect(container.textContent).not.toContain("Ready")
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
