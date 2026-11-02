// Run from apps/web: bun test test/background-agents-tab.test.tsx
// Mocked preferences API; no model requests are made.
import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"

type Agents = { chronologySuggestion: boolean; reviewCompare: boolean; narration: boolean }

async function render(options: { load?: Partial<Agents> | "fail"; failSave?: boolean; wordSuggestions?: boolean } = {}) {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "getComputedStyle", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true,
    value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const originalFetch = globalThis.fetch
  const puts: unknown[] = []
  let stored: Partial<Agents> | undefined = options.load === "fail" ? undefined : options.load
  let word = { defaultScope: "matter", suggestions: options.wordSuggestions ?? true }
  globalThis.fetch = (async (input, init) => {
    if (String(input).endsWith("/word/chat-preferences")) {
      if (init?.method === "PUT") {
        const body = JSON.parse(String(init.body))
        puts.push({ word: body })
        if (options.failSave) return new Response("error", { status: 500 })
        word = { ...word, ...body }
      }
      return Response.json(word)
    }
    if (!String(input).endsWith("/preferences")) return new Response("not found", { status: 404 })
    if (init?.method === "PUT") {
      const body = JSON.parse(String(init.body))
      puts.push(body)
      if (options.failSave) return new Response("error", { status: 500 })
      stored = body.backgroundAgents
      return Response.json({ backgroundAgents: stored })
    }
    if (options.load === "fail") return new Response("error", { status: 500 })
    return Response.json(stored ? { backgroundAgents: stored } : {})
  }) as typeof fetch
  const toasts: string[] = []
  const { createRoot } = await import("react-dom/client")
  const { default: BackgroundAgentsTab } = await import("../src/components/settings/BackgroundAgentsTab")
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  await act(async () => { root.render(createElement(BackgroundAgentsTab, { onSaved: (text: string) => toasts.push(text) })) })
  const switches = () => Array.from(document.body.querySelectorAll<HTMLButtonElement>("[role=switch]"))
  const toggle = async (label: string) => {
    const target = switches().find(item => item.getAttribute("aria-label") === label)
    if (!target) throw new Error(`no switch ${label}`)
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
  const state = () => Object.fromEntries(switches().map(item => [item.getAttribute("aria-label"), item.getAttribute("aria-checked") === "true"]))
  return { toggle, state, puts, toasts, cleanup, text: () => document.body.textContent ?? "" }
}

test("missing settings show every agent on; a toggle saves the whole object at once", async () => {
  const view = await render()
  try {
    expect(view.text()).toContain("These run by themselves while you work, without you asking.")
    expect(view.state()).toEqual({ "Suggest a chronology": true, "Compare review answers": true, "Narrate drafting progress": true, "Suggest ideas in Word": true })
    await view.toggle("Compare review answers")
    expect(view.puts).toEqual([{ backgroundAgents: { chronologySuggestion: true, reviewCompare: false, narration: true } }])
    expect(view.state()["Compare review answers"]).toBe(false)
    expect(view.toasts).toEqual(["Compare review answers turned off."])
  } finally {
    await view.cleanup()
  }
})

test("saved switches load as stored", async () => {
  const view = await render({ load: { chronologySuggestion: false, reviewCompare: true, narration: false }, wordSuggestions: false })
  try {
    expect(view.state()).toEqual({ "Suggest a chronology": false, "Compare review answers": true, "Narrate drafting progress": false, "Suggest ideas in Word": false })
  } finally {
    await view.cleanup()
  }
})

test("the Word switch saves only the suggestions setting, leaving firm preferences and the Word chat scope alone", async () => {
  const view = await render()
  try {
    await view.toggle("Suggest ideas in Word")
    expect(view.puts).toEqual([{ word: { suggestions: false } }])
    expect(view.state()["Suggest ideas in Word"]).toBe(false)
    expect(view.toasts).toEqual(["Suggest ideas in Word turned off."])
  } finally {
    await view.cleanup()
  }
})

test("a failed save reverts the switch and shows an error", async () => {
  const view = await render({ failSave: true })
  try {
    await view.toggle("Narrate drafting progress")
    expect(view.state()["Narrate drafting progress"]).toBe(true)
    expect(view.text()).toContain("Could not save this setting. Try again.")
    expect(view.toasts).toEqual([])
  } finally {
    await view.cleanup()
  }
})

test("a failed load offers a retry", async () => {
  const view = await render({ load: "fail" })
  try {
    expect(view.text()).toContain("Could not load background agent settings.")
    expect(view.text()).toContain("Try again")
  } finally {
    await view.cleanup()
  }
})
