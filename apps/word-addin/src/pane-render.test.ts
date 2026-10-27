import { expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { Window } from "happy-dom"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import type { PaneState } from "./sync"
import { createPaneRenderer, setupSaveMenu } from "./pane-render"

test("unchanged poll presentation preserves the pane DOM and focus while real saving updates controls", async () => {
  const page = new Window({ url: `https://localhost:4612/word-addin/pane.html?version=${WORD_ADDIN_VERSION}` })
  const html = readFileSync(new URL("../pane.html", import.meta.url), "utf8")
  page.document.body.innerHTML = html.match(/<body>([\s\S]*)<\/body>/)![1]!
  const document = page.document as unknown as Document
  try {
    const render = createPaneRenderer(document)
    const state: PaneState = {
      status: "saved", busy: false, canSave: true, content: "confirmed", message: "The saved Word document is in doc.haus.",
      acknowledgement: { base: "v1", number: 1, fingerprint: "a".repeat(64), at: 1_700_000_000_000, overlap: false },
      connection: { version: { base: { id: "v1", number: 1 }, current: { id: "v1", number: 1 }, state: "current" }, lease: { state: "owned", generation: "fixture", expiresAt: 90_000 } },
    }
    render(state)
    const button = document.getElementById("save") as HTMLButtonElement
    button.focus()
    const text = document.getElementById("status")!.firstChild
    const mutations: unknown[] = []
    const observer = new page.MutationObserver(records => mutations.push(...records))
    observer.observe(page.document.body, { subtree: true, attributes: true, childList: true, characterData: true })
    try {
      render({ ...state, connection: { ...state.connection!, lease: { ...state.connection!.lease, expiresAt: 110_000 } } })
      render({ ...state })
      await Bun.sleep(0)
      expect(mutations).toHaveLength(0)
      expect(document.getElementById("status")!.firstChild === text).toBe(true)
      expect(document.activeElement === button).toBe(true)
      expect(button.disabled).toBe(false)
      render({ ...state, status: "saving", busy: true, message: "Saving to doc.haus…" })
      await Bun.sleep(0)
      expect(mutations.length).toBeGreaterThan(0)
      expect(button.disabled).toBe(true)
      expect(document.getElementById("status")!.textContent).toBe("Saving…")
      render(state)
      expect(button.disabled).toBe(false)
    } finally { observer.disconnect() }
  } finally {
    page.happyDOM.abort()
  }
})

test("save menu dismisses outside the pane controls without stealing focus", () => {
  const page = new Window()
  page.document.body.innerHTML = '<details id="menu"><summary>More options</summary><button>Save</button></details><button id="outside">Documents</button>'
  const document = page.document as unknown as Document
  const menu = document.getElementById("menu") as HTMLDetailsElement
  const outside = document.getElementById("outside") as HTMLButtonElement
  const summary = menu.querySelector("summary")!
  setupSaveMenu(menu)
  try {
    menu.open = true
    menu.querySelector("button")!.dispatchEvent(new page.Event("pointerdown", { bubbles: true }) as unknown as Event)
    expect(menu.open).toBe(true)
    outside.focus()
    expect(menu.open).toBe(false)
    expect(document.activeElement === outside).toBe(true)
    menu.open = true
    outside.dispatchEvent(new page.Event("pointerdown", { bubbles: true }) as unknown as Event)
    expect(menu.open).toBe(false)
    expect(document.activeElement === outside).toBe(true)
    menu.open = true
    page.dispatchEvent(new page.Event("blur"))
    expect(menu.open).toBe(false)
    expect(document.activeElement === outside).toBe(true)
    menu.open = true
    menu.dispatchEvent(new page.KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }) as unknown as Event)
    expect(menu.open).toBe(false)
    expect(document.activeElement === summary).toBe(true)
  } finally { page.happyDOM.abort() }
})
