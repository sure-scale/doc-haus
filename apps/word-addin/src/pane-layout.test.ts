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

test("feature views keep tabs first and shared Save last while scrolling preserves draft, focus and confirmed status", () => {
  const { page, document, get } = fixture()
  const render = createPaneRenderer(document)
  const state: PaneState = { status: "saved", busy: false, canSave: true, content: "confirmed", message: "Confirmed",
    acknowledgement: { base: "v2", number: 2, fingerprint: "a".repeat(64), at: 100_000, overlap: false },
    connection: { version: { base: { id: "v2", number: 2 }, current: { id: "v2", number: 2 }, state: "current" }, lease: { state: "owned", generation: "synthetic", expiresAt: 200_000 } } }
  try {
    render(state)
    const bar = get("save-panel"), tabs = get("desk-entry"), chat = get("pane-chat"), desk = get("pane-desk")
    expect(bar.parentElement!.firstElementChild === tabs).toBe(true)
    expect(bar.parentElement!.lastElementChild === bar).toBe(true)
    expect(get("save-details").contains(get("background"))).toBe(true)
    expect(get("background").contains(get("back"))).toBe(true)
    tabs.hidden = false; chat.hidden = false
    chat.innerHTML = '<div data-chat="history" tabindex="0">Conversation retained</div><form><textarea>Draft retained</textarea></form>'
    const history = chat.querySelector<HTMLElement>('[data-chat="history"]')!
    history.focus(); history.scrollTop = 80
    history.dispatchEvent(new page.Event("scroll") as unknown as Event)
    expect(document.activeElement === history).toBe(true)
    expect(get("status").textContent).toBe("Saved"); expect(get("save-version").textContent).toBe("v2")
    expect(get<HTMLButtonElement>("save").disabled).toBe(false)
    for (const feature of ["documents", "compare"]) {
      chat.hidden = true; desk.hidden = false; desk.setAttribute("aria-labelledby", `${feature}-tab`)
      desk.scrollTop = 80; desk.dispatchEvent(new page.Event("scroll") as unknown as Event)
      expect(tabs.hidden).toBe(false); expect(bar.hidden).toBe(false)
      expect(bar.parentElement!.lastElementChild === bar).toBe(true)
      expect(get("status").textContent).toBe("Saved")
    }
    chat.hidden = false; desk.hidden = true
    expect(chat.querySelector("textarea")!.value).toBe("Draft retained")
    render({ ...state, status: "saving", busy: true, message: "Saving to doc.haus…" })
    expect(get("status").textContent).toBe("Saving…"); expect(get<HTMLButtonElement>("save").disabled).toBe(true)
  } finally { page.happyDOM.abort() }
})

test("upward save details fit between tabs and status, support keyboard dismissal, and expose unlinked opening instructions", async () => {
  const { page, document, get } = fixture()
  const bar = get("save-panel"), tabs = get("desk-entry"), details = get("save-details"), toggle = get<HTMLButtonElement>("save-details-toggle")
  let barTop = 530
  bar.getBoundingClientRect = () => ({ top: barTop } as DOMRect)
  tabs.getBoundingClientRect = () => ({ bottom: 44 } as DOMRect)
  tabs.hidden = false
  const dispose = setupPaneLayout(document)
  try {
    get("matter").textContent = "Long matter name"
    get("matter").title = "Long matter name · Services agreement.docx"
    get("message").textContent = "Keep the current Word copy and review this complete recovery advice."
    const chat = get("pane-chat"); chat.hidden = false; chat.innerHTML = '<textarea>Draft retained</textarea>'
    expect(details.hidden).toBe(true); expect(toggle.getAttribute("aria-expanded")).toBe("false")
    expect(bar.querySelector(".save-summary")!.textContent).not.toContain("recovery advice")
    toggle.click(); expect(details.hidden).toBe(false); expect(toggle.getAttribute("aria-expanded")).toBe("true")
    expect(details.textContent).toContain("complete recovery advice")
    expect(get("save-details-matter").textContent).toBe("Long matter name · Services agreement.docx")
    expect(details.style.getPropertyValue("--save-details-space")).toBe("474px")
    barTop = 500; page.dispatchEvent(new page.Event("resize"))
    expect(details.style.getPropertyValue("--save-details-space")).toBe("444px")
    get("save-details-close").focus()
    get("save-details-close").dispatchEvent(new page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as Event)
    expect(details.hidden).toBe(true); expect(document.activeElement === toggle).toBe(true)
    toggle.click(); get("save-details-close").click()
    expect(details.hidden).toBe(true); expect(document.activeElement === toggle).toBe(true)
    toggle.click(); chat.querySelector("textarea")!.click()
    expect(details.hidden).toBe(true); expect(chat.querySelector("textarea")!.value).toBe("Draft retained")
    tabs.hidden = true; chat.hidden = true; get("actions").hidden = true; await Bun.sleep(0)
    expect(details.hidden).toBe(false); expect(toggle.getAttribute("aria-expanded")).toBe("true")
    expect(details.style.getPropertyValue("--save-details-space")).toBe("488px")
    expect(details.textContent).toContain("complete recovery advice")
  } finally { dispose(); page.happyDOM.abort() }
})
