import { expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { Window } from "happy-dom"
import { setupPaneLayout } from "./pane-layout"
import { createPaneRenderer } from "./pane-render"
import type { PaneState } from "./sync"

function fixture() {
  const page = new Window()
  page.document.body.innerHTML = readFileSync(new URL("../pane.html", import.meta.url), "utf8").match(/<body>([\s\S]*)<\/body>/)![1]!
  const document = page.document as unknown as Document
  return { page, document, get: <T extends HTMLElement>(id: string) => document.getElementById(id) as T }
}

test("chat scroll compacts the save overview without changing its receipt, controls or focus; other tabs restore normal flow", async () => {
  const { page, document, get } = fixture()
  const render = createPaneRenderer(document)
  const state: PaneState = { status: "saved", busy: false, canSave: true, content: "confirmed", message: "Confirmed",
    acknowledgement: { base: "v2", number: 2, fingerprint: "a".repeat(64), at: 100_000, overlap: false },
    connection: { version: { base: { id: "v2", number: 2 }, current: { id: "v2", number: 2 }, state: "current" }, lease: { state: "owned", generation: "synthetic", expiresAt: 200_000 } } }
  try {
    render(state)
    expect(get("save-panel").dataset.compact).toBe("false")
    const chat = get("pane-chat"); chat.hidden = false
    chat.innerHTML = '<div data-chat="history" tabindex="0">Conversation retained</div><form><textarea>Draft retained</textarea></form>'
    await Bun.sleep(0)
    const history = chat.querySelector<HTMLElement>('[data-chat="history"]')!
    history.focus(); history.scrollTop = 80
    history.dispatchEvent(new page.Event("scroll") as unknown as Event)
    expect(get("save-panel").dataset.compact).toBe("true")
    expect(get("save-panel").parentElement!.dataset.chatActive).toBe("true")
    expect(get("status").textContent).toBe("Saved"); expect(get<HTMLButtonElement>("save").disabled).toBe(false)
    expect(document.activeElement === history).toBe(true)
    render({ ...state, status: "saving", busy: true, message: "Saving to doc.haus…" })
    expect(get("status").textContent).toBe("Saving…"); expect(get<HTMLButtonElement>("save").disabled).toBe(true)
    expect(get("save-panel").dataset.compact).toBe("true")
    history.scrollTop = 0; history.dispatchEvent(new page.Event("scroll") as unknown as Event)
    expect(get("save-panel").dataset.compact).toBe("false")
    history.scrollTop = 80; history.dispatchEvent(new page.Event("scroll") as unknown as Event)
    chat.hidden = true; await Bun.sleep(0)
    expect(get("save-panel").parentElement!.dataset.chatActive).toBe("false")
    expect(get("save-panel").dataset.compact).toBe("false")
    expect(chat.querySelector("textarea")!.value).toBe("Draft retained")
    expect(get("background").previousElementSibling === get("pane-desk")).toBe(true)
  } finally { page.happyDOM.abort() }
})

test("full save advice is disclosed deliberately while unlinked opening instructions stay visible", async () => {
  const { page, document, get } = fixture(), dispose = setupPaneLayout(document)
  try {
    const details = get("save-details"), toggle = get<HTMLButtonElement>("save-details-toggle")
    get("message").textContent = "Keep the current Word copy and review this complete recovery advice."
    expect(details.hidden).toBe(true); expect(toggle.getAttribute("aria-expanded")).toBe("false")
    expect(get("save-panel").querySelector(".save-summary")!.textContent).not.toContain("recovery advice")
    toggle.click(); expect(details.hidden).toBe(false); expect(toggle.getAttribute("aria-expanded")).toBe("true")
    expect(details.textContent).toContain("complete recovery advice")
    get("release").dispatchEvent(new page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as Event)
    expect(details.hidden).toBe(true); expect(document.activeElement === toggle).toBe(true)
    get("actions").hidden = true; await Bun.sleep(0)
    expect(details.hidden).toBe(false); expect(toggle.getAttribute("aria-expanded")).toBe("true")
  } finally { dispose(); page.happyDOM.abort() }
})
