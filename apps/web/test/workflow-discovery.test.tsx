import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement, type ReactElement } from "react"
import type { Document } from "../src/api/ingest/documents"
import { reviewInputNotice, workflowSearchMatches } from "../src/review-starters"

// Interactions use only synthetic metadata and the local config read. Any
// mutation or provider request fails this fixture immediately.
async function render(element: () => Promise<ReactElement>, config: { model?: string } = { model: "fixture/model" }) {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Element", "Node", "Event", "HTMLInputElement", "getComputedStyle", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map((name) => [name, { configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const originalFetch = globalThis.fetch
  const requests: string[] = []
  globalThis.fetch = (async (input, init) => {
    const request = `${init?.method ?? "GET"} ${String(input)}`
    requests.push(request)
    if (request.startsWith("GET ") && String(input).endsWith("/engine/config")) return Response.json(config)
    throw new Error(`Unexpected request: ${request}`)
  }) as typeof fetch
  const { createRoot } = await import("react-dom/client")
  const container = document.createElement("div")
  document.body.append(container)
  const root = createRoot(container)
  await act(async () => root.render(await element()))
  const button = (name: string) => [...document.querySelectorAll<HTMLButtonElement>("button")].find((node) => node.textContent === name)!
  const click = async (name: string) => { await act(async () => button(name).click()) }
  const search = async (text: string) => {
    const field = document.querySelector<HTMLInputElement>('input[type="search"]')!
    await act(async () => {
      Object.getOwnPropertyDescriptor(browser.HTMLInputElement.prototype, "value")!.set!.call(field, text)
      field.dispatchEvent(new Event("input", { bubbles: true }))
    })
  }
  const cleanup = async () => {
    await act(async () => root.unmount())
    globalThis.fetch = originalFetch
    await browser.happyDOM.close()
    for (const [name, descriptor] of saved) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete (globalThis as Record<string, unknown>)[name]
    }
  }
  return { button, click, search, requests, cleanup, text: () => document.body.textContent ?? "" }
}

test("discovery searches names and purpose without inferring document content", () => {
  expect(workflowSearchMatches("  AGREEMENTS terms ", "Compare agreements", "Key terms side by side")).toBe(true)
  expect(workflowSearchMatches("missing", "Compare agreements", "Key terms side by side")).toBe(false)
  expect(reviewInputNotice([{ name: "Unknown.docx" }])).toBeUndefined()
  expect(reviewInputNotice([{ name: "Scan.pdf", quality: { status: "no_readable_text", signals: [] } }])).toContain("Scan.pdf")
  expect(reviewInputNotice([{ name: "Partial.pdf", quality: { status: "degraded", signals: [] } }])).toContain("may be incomplete")
})

test("chat search cannot launch hidden workflows, and no matches cannot launch the previous selection", async () => {
  const launched: string[] = []
  const view = await render(async () => {
    const { default: WorkflowLauncher } = await import("../src/components/WorkflowLauncher")
    return createElement(WorkflowLauncher, {
      workflows: [
        { name: "review", label: "Full review", description: "Check obligations", scope: "matter", prompt: "Review" },
        { name: "hidden", label: "Hidden workflow", description: "Unavailable", scope: "matter", prompt: "Hidden" },
      ], available: new Set(["review"]), documents: [{}], onLaunch: (name) => launched.push(name),
    })
  })
  try {
    await view.click("Workflows")
    expect(view.text()).not.toContain("Hidden workflow")
    await view.search("obligations")
    expect(view.button("Start").disabled).toBe(false)
    await view.search("Hidden workflow")
    expect(view.text()).toContain("No matching workflows")
    expect(view.button("Start").disabled).toBe(true)
    await view.click("Start")
    expect(launched).toEqual([])
    await view.click("Clear")
    await view.click("Start")
    expect(launched).toEqual(["review"])
    expect(view.requests.every((request) => request.startsWith("GET ") && request.endsWith("/engine/config"))).toBe(true)
  } finally { await view.cleanup() }
})

test("a matter workflow explains missing documents and a missing model before launch", async () => {
  for (const documents of [[], [{}]]) {
    const launched: string[] = []
    const view = await render(async () => {
      const { default: WorkflowLauncher } = await import("../src/components/WorkflowLauncher")
      return createElement(WorkflowLauncher, { workflows: [{ name: "review", label: "Full review", description: "Review", scope: "matter", prompt: "Review" }], available: new Set(["review"]), documents, onLaunch: (name) => launched.push(name) })
    }, {})
    try {
      await view.click("Workflows")
      expect(view.text()).toContain(documents.length ? "Choose a model in Settings" : "Add documents to this matter")
      expect(view.button("Start").disabled).toBe(true)
      expect(launched).toEqual([])
    } finally { await view.cleanup() }
  }
})

test("table chooser filters questions, keeps no-match actions inert, and allows inspection with missing inputs", async () => {
  let added = 0
  const view = await render(async () => {
    const { default: ReviewStarter } = await import("../src/components/ReviewStarter")
    return createElement(ReviewStarter, { documents: [] as Document[], existingQuestions: [], onCancel: () => {}, onConfirm: async () => { added++; return true } })
  })
  try {
    expect(view.text()).toContain("Add documents to this matter")
    await view.search("governs")
    expect(document.querySelectorAll('input[type="radio"]').length).toBe(1)
    expect(view.text()).toContain("Compare agreements")
    await view.search("does not exist")
    expect(view.text()).toContain("No matching starters")
    expect(view.button("Continue").disabled).toBe(true)
    await view.click("Clear")
    await view.click("Continue")
    expect(view.button("Add questions").disabled).toBe(true)
    expect(added).toBe(0)
    expect(view.requests).toEqual([])
  } finally { await view.cleanup() }
})
