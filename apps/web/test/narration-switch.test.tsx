import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement } from "react"
import completed from "../../../dochaus/workflows/lean/fixtures/status-completed.json"
import type { Client } from "../src/api/opencode/client"
import { narratorTitle } from "../src/api/opencode/narrator"

// The "Narrate drafting progress" switch: off, a drafting run's chat opens no
// narrator session and shows its steps without an error; on, the same run looks
// for its narrator to write an update. Recorded lean status fixture; no model.

const owner = (completed as { owner_session_id: string }).owner_session_id

async function render(narration: boolean) {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Node", "Event", "getComputedStyle", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true,
    value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const originalFetch = globalThis.fetch
  globalThis.fetch = (async (input) => {
    const url = String(input)
    if (url.endsWith("/drafting-status")) return Response.json(completed)
    if (url.endsWith("/preferences")) return Response.json({ backgroundAgents: { chronologySuggestion: true, reviewCompare: true, narration } })
    return new Response("not found", { status: 404 })
  }) as typeof fetch
  // Each poll sees time move on, so a checkpoint is past its quiet period at once.
  const originalNow = Date.now
  let clock = originalNow()
  Date.now = () => (clock += 10_000)
  const calls: string[] = []
  const lookups: string[] = []
  const client = {
    session: {
      list: async (input: { search?: string }) => {
        lookups.push(input.search ?? "")
        return { data: [] }
      },
      create: async () => { calls.push("session.create"); throw new Error("no narrator in this test") },
      prompt: async () => { calls.push("session.prompt"); throw new Error("no narrator in this test") },
    },
    message: { list: async () => ({ data: [], cursor: {} }) },
  } as unknown as Client
  const { createRoot } = await import("react-dom/client")
  const { useDraftingNarration } = await import("../src/components/narration/use-drafting-narration")
  const { default: DraftingNarration } = await import("../src/components/narration/DraftingNarration")
  const { MemoryRouter } = await import("react-router-dom")
  function Harness({ startedOwner }: { startedOwner?: string }) {
    const narration = useDraftingNarration({ client, matterId: "matter-switch", matterName: "Switch", startedOwner, sessionID: owner })
    return createElement(MemoryRouter, null,
      createElement(DraftingNarration, { narration, matterId: "matter-switch", showRequest: false, onViewDocument: () => {} }))
  }
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  // The first poll reads any earlier narration; a second, started here rather
  // than after the poll interval, is the first that may write an update.
  await act(async () => { root.render(createElement(Harness, { startedOwner: owner })) })
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 100)) })
  await act(async () => { root.render(createElement(Harness, {})) })
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 300)) })
  const cleanup = async () => {
    await act(async () => root.unmount())
    container.remove()
    globalThis.fetch = originalFetch
    Date.now = originalNow
    await browser.happyDOM.close()
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
  return { calls, lookups, cleanup, text: () => document.body.textContent ?? "" }
}

test("with narration off, a drafting run opens no narrator and still shows its steps", async () => {
  const view = await render(false)
  try {
    expect(view.text()).toContain("Drafting")
    // Only the read of an earlier narration; no narrator is opened or prompted.
    expect(view.lookups).toEqual([narratorTitle(owner)])
    expect(view.calls).toEqual([])
    expect(view.text()).not.toContain("progress summary could not be written")
    expect(view.text()).not.toContain("Summarising progress")
  } finally {
    await view.cleanup()
  }
})

test("with narration on, the same run goes on to open its narrator", async () => {
  const view = await render(true)
  try {
    // The history read, then the lookup that opens the narrator for an update.
    expect(view.lookups).toEqual([narratorTitle(owner), narratorTitle(owner)])
  } finally {
    await view.cleanup()
  }
})
