import { expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { Window } from "happy-dom"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import type { PaneState } from "./sync"
import { configureConnectionRecovery, createPaneRenderer } from "./pane-render"

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
    expect(button.textContent).toBe("Save")
    expect(button.hidden).toBe(true)
    expect(document.getElementById("save-menu")).toBeNull()
    expect(document.getElementById("takeover")!.hidden).toBe(true)
    expect(document.getElementById("release")!.hidden).toBe(true)
    expect(document.getElementById("save-version")!.textContent).toBe("v1")
    expect(document.getElementById("save-version")!.title).toBe("This Word copy: version 1")
    const details = document.getElementById("save-details-toggle") as HTMLButtonElement
    expect(details.getAttribute("aria-label")).toBe("Saved, version 1. Show save status and recovery details")
    expect(details.textContent).toBe("v1"); expect(details.textContent).not.toContain("Details")
    expect(document.getElementById("pane-header")!.contains(document.getElementById("status"))).toBe(true)
    let saves = 0; button.addEventListener("click", () => saves++)
    details.click(); expect(saves).toBe(0); expect(document.getElementById("save-details")!.hidden).toBe(false)
    details.click(); expect(saves).toBe(0)
    details.focus()
    const text = document.getElementById("status")!.firstChild
    const mutations: unknown[] = []
    const observer = new page.MutationObserver(records => mutations.push(...records))
    observer.observe(page.document.body, { subtree: true, attributes: true, childList: true, characterData: true })
    try {
      render({ ...state, connection: { ...state.connection!, lease: { ...state.connection!.lease, expiresAt: 110_000 } } })
      render({ ...state })
      await Bun.sleep(0)
      mutations.push(...observer.takeRecords())
      expect(mutations).toHaveLength(0)
      expect(document.getElementById("status")!.firstChild === text).toBe(true)
      expect(document.activeElement === details).toBe(true)
      expect(button.disabled).toBe(false)
      render({ ...state, status: "saving", busy: true, message: "Saving to doc.haus…" })
      await Bun.sleep(0)
      mutations.push(...observer.takeRecords())
      expect(mutations.length).toBeGreaterThan(0)
      expect(button.disabled).toBe(true)
      expect(button.hidden).toBe(false)
      expect(document.getElementById("status")!.textContent).toBe("Saving…")
      render(state)
      expect(button.disabled).toBe(false)
      expect(button.hidden).toBe(true)
      render({ ...state, status: "done" }); expect(button.hidden).toBe(true)
      render({ ...state, status: "saved", content: "unverified" }); expect(button.hidden).toBe(false)
      render({ ...state, status: "saved", content: "changed" }); expect(button.hidden).toBe(false)
      render({ ...state, status: "not-saved", content: "unverified" }); expect(button.hidden).toBe(false)
      render({ ...state, status: "not-saved", content: "changed" }); expect(button.hidden).toBe(false)
      button.click(); expect(saves).toBe(1)
      render({ ...state, status: "saved", problem: true }); expect(button.hidden).toBe(false)
      render({ ...state, status: "not-saved", problem: true, connectionFailed: true })
      expect(button.hidden).toBe(false)
      expect(document.getElementById("connection-recovery")!.hidden).toBe(false)
      render({ ...state, status: "not-saved", problem: true })
      expect(document.getElementById("connection-recovery")!.hidden).toBe(true)
    } finally { observer.disconnect() }
  } finally {
    page.happyDOM.abort()
  }
})

test("blocked saving offers only the recovery action supported by the displayed connection", () => {
  const page = new Window()
  page.document.body.innerHTML = readFileSync(new URL("../pane.html", import.meta.url), "utf8").match(/<body>([\s\S]*)<\/body>/)![1]!
  const document = page.document as unknown as Document
  const render = createPaneRenderer(document)
  const state: PaneState = { status: "not-saved", busy: false, canSave: false, content: "changed", message: "Paused",
    connection: { version: { base: { id: "v1", number: 1 }, current: { id: "v1", number: 1 }, state: "current" }, lease: { state: "other", generation: "other-owner", expiresAt: 90_000 } } }
  const takeover = document.getElementById("takeover") as HTMLButtonElement
  const release = document.getElementById("release") as HTMLButtonElement
  try {
    render(state)
    expect(takeover.hidden).toBe(false)
    expect(takeover.disabled).toBe(false)
    expect(release.hidden).toBe(true)
    render({ ...state, busy: true })
    expect(takeover.disabled).toBe(true)
    render({ ...state, connection: { ...state.connection!, lease: { state: "available", expiresAt: null } } })
    expect(takeover.hidden).toBe(true)
    render({ ...state, problem: true, connection: { ...state.connection!, lease: { state: "owned", generation: "own", expiresAt: 90_000 } } })
    expect(release.hidden).toBe(false)
    expect(takeover.hidden).toBe(true)
    render({ ...state, connection: { ...state.connection!, version: { ...state.connection!.version, state: "stale", current: { id: "v2", number: 2 } } } })
    expect(document.getElementById("save")!.hidden).toBe(true)
    expect(document.getElementById("load-latest")!.hidden).toBe(false)
    expect(document.getElementById("save-version")!.textContent).toBe("v1")
    expect(takeover.hidden).toBe(true)
  } finally { page.happyDOM.abort() }
})

test("connection recovery launches only local doc.haus without navigating away from the live save pane", () => {
  for (const origin of ["https://localhost:4612", "https://word.firm.example"]) {
    const page = new Window({ url: `${origin}/word-addin/pane.html?version=${WORD_ADDIN_VERSION}` })
    page.document.body.innerHTML = readFileSync(new URL("../pane.html", import.meta.url), "utf8").match(/<body>([\s\S]*)<\/body>/)![1]!
    const document = page.document as unknown as Document
    try {
      configureConnectionRecovery(document, origin)
      createPaneRenderer(document)({ status: "not-saved", busy: false, canSave: false, content: "changed", message: "Cannot reach doc.haus.", connectionFailed: true })
      const recovery = document.getElementById("connection-recovery")!
      const launch = document.getElementById("open-dochaus") as HTMLAnchorElement
      expect(recovery.hidden).toBe(false)
      expect(document.getElementById("save-version")!.hidden).toBe(true)
      expect(recovery.querySelectorAll('a[href^="https:"]')).toHaveLength(0)
      expect(recovery.textContent).toContain("Keep this pane open")
      expect(recovery.textContent).toContain("reconnects automatically")
      expect(recovery.textContent).not.toContain("Check status")
      if (origin.includes("localhost")) {
        expect(launch.hidden).toBe(false)
        expect(launch.href).toBe("doc.haus://launch")
        expect(launch.target).toBe("_blank")
        expect(launch.relList.contains("noopener")).toBe(true)
      } else {
        expect(launch.hidden).toBe(true)
        expect(launch.hasAttribute("href")).toBe(false)
        expect(recovery.textContent).toContain("server")
      }
      expect(page.location.href).toBe(`${origin}/word-addin/pane.html?version=${WORD_ADDIN_VERSION}`)
    } finally { page.happyDOM.abort() }
  }
})
