import { expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { Window } from "happy-dom"
import { createPaneRenderer } from "./pane-render"
import { setupLatestVersion } from "./latest-version"
import type { PaneState } from "./sync"

test("stale pane offers a refreshed native handoff without replacing or closing the old document", async () => {
  const page = new Window()
  page.document.body.innerHTML = readFileSync(new URL("../pane.html", import.meta.url), "utf8").match(/<body>([\s\S]*)<\/body>/)![1]!
  const document = page.document as unknown as Document
  let clock = 100_000, requests = 0
  const action = document.getElementById("load-latest") as HTMLAnchorElement
  const state: PaneState = { status: "not-saved", canSave: false, busy: false, content: "changed", message: "Paused",
    connection: { version: { base: { id: "v1", number: 1 }, current: { id: "v2", number: 2 }, state: "stale" }, lease: { state: "available", expiresAt: null } } }
  const render = createPaneRenderer(document)
  const latest = setupLatestVersion(document, async base => { expect(base).toBe("v1"); requests++; return `doc.haus://word/open?source=ticket-${requests}` }, () => clock)
  try {
    render(state); latest(state)
    await Bun.sleep(0)
    expect(document.getElementById("save")!.hidden).toBe(true)
    expect(action.hidden).toBe(false)
    expect(action.textContent).toBe("Load latest version")
    expect(action.getAttribute("aria-disabled")).toBe("false")
    expect(action.target).toBe("_blank")
    expect(action.href).toContain("ticket-1")
    latest(state); await Bun.sleep(0)
    expect(requests).toBe(1)
    clock += 31_000
    latest(state); await Bun.sleep(0)
    expect(action.href).toContain("ticket-2")
    latest({ ...state, connection: { ...state.connection!, version: { ...state.connection!.version, state: "removed" } } })
    render({ ...state, connection: { ...state.connection!, version: { ...state.connection!.version, state: "removed" } } })
    expect(action.hidden).toBe(true)
    expect(action.hasAttribute("href")).toBe(false)
    expect(state.connection!.version.base!.id).toBe("v1")
    expect(state.content).toBe("changed")
  } finally { page.happyDOM.abort() }
})

test("a handoff response arriving after the pane stops being stale cannot restore a launch link", async () => {
  const page = new Window()
  page.document.body.innerHTML = '<a id="load-latest"></a><p id="latest-message"></p>'
  let resolve!: (uri: string) => void
  const latest = setupLatestVersion(page.document as unknown as Document, () => new Promise<string>(done => { resolve = done }))
  try {
    latest({ status: "not-saved", busy: false, canSave: false, content: "unverified", message: "", connection: { version: { state: "stale", base: { id: "v1", number: 1 }, current: { id: "v2", number: 2 } }, lease: { state: "available", expiresAt: null } } })
    latest({ status: "not-saved", busy: false, canSave: false, content: "unverified", message: "" })
    resolve("doc.haus://word/open?source=old-ticket")
    await Bun.sleep(0)
    expect(page.document.getElementById("load-latest")!.hasAttribute("href")).toBe(false)
  } finally { page.happyDOM.abort() }
})
