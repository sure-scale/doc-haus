import { expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { Window } from "happy-dom"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { setupLatestVersion } from "./latest-version"
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
    expect(button.hidden).toBe(true)
    expect(document.getElementById("takeover")!.hidden).toBe(true)
    expect(document.getElementById("release")!.hidden).toBe(true)
    expect(document.getElementById("save-version")!.textContent).toBe("v1")
    const details = document.getElementById("save-details-toggle") as HTMLButtonElement
    let saves = 0; button.addEventListener("click", () => saves++)
    details.click(); expect(saves).toBe(0); expect(document.getElementById("save-details")!.hidden).toBe(false)
    details.click(); expect(saves).toBe(0)
    details.focus()
    const text = document.getElementById("status")!.firstChild
    const unchangedMarkup = document.body.innerHTML
    const mutations: unknown[] = []
    const observer = new page.MutationObserver(records => mutations.push(...records))
    observer.observe(page.document.body, { subtree: true, attributes: true, childList: true, characterData: true })
    try {
      render({ ...state, connection: { ...state.connection!, lease: { ...state.connection!.lease, expiresAt: 110_000 } } })
      render({ ...state })
      await page.happyDOM.waitUntilComplete()
      mutations.push(...observer.takeRecords())
      expect(mutations).toHaveLength(0)
      expect(document.body.innerHTML).toBe(unchangedMarkup)
      expect(document.getElementById("status")!.firstChild === text).toBe(true)
      expect(document.activeElement === details).toBe(true)
      expect(button.disabled).toBe(false)
      render({ ...state, status: "saving", busy: true, message: "Saving to doc.haus…" })
      expect(document.body.innerHTML).not.toBe(unchangedMarkup)
      expect(button.disabled).toBe(true)
      expect(button.hidden).toBe(false)
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
  const state: PaneState = {
    status: "not-saved", busy: false, canSave: false, content: "changed", message: "Paused",
    connection: { version: { base: { id: "v1", number: 1 }, current: { id: "v1", number: 1 }, state: "current" }, lease: { state: "other", generation: "other-owner", expiresAt: 90_000 } },
  }
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
      if (origin.includes("localhost")) {
        expect(launch.hidden).toBe(false)
        expect(launch.href).toBe("doc.haus://launch")
        expect(launch.target).toBe("_blank")
        expect(launch.relList.contains("noopener")).toBe(true)
      } else {
        expect(launch.hidden).toBe(true)
        expect(launch.hasAttribute("href")).toBe(false)
      }
      expect(page.location.href).toBe(`${origin}/word-addin/pane.html?version=${WORD_ADDIN_VERSION}`)
    } finally { page.happyDOM.abort() }
  }
})

const staleState: PaneState = {
  status: "not-saved", canSave: false, busy: false, content: "changed", message: "Paused",
  connection: { version: { base: { id: "v1", number: 1 }, current: { id: "v2", number: 2 }, state: "stale" }, lease: { state: "available", expiresAt: null } },
}
function latestFixture() {
  const page = new Window()
  page.document.body.innerHTML = '<button id="load-latest">Load latest version</button><a id="open-latest-copy" target="_blank"></a><p id="latest-message"></p>'
  return page
}

test("latest refresh is deliberate and serialized while separate-copy fallback is prepared without mutation", async () => {
  const page = latestFixture(), document = page.document as unknown as Document
  let loads = 0, resolve!: (result: boolean) => void
  const render = setupLatestVersion(document, async base => { expect(base).toBe("v1"); return "doc.haus://word/open?source=ticket" }, () => { loads++; return new Promise(done => { resolve = done }) })
  try {
    render(staleState); await Bun.sleep(0)
    const action = document.getElementById("load-latest") as HTMLButtonElement
    const separate = document.getElementById("open-latest-copy") as HTMLAnchorElement
    expect(action.hidden).toBe(false); expect(loads).toBe(0); expect(separate.href).toContain("ticket"); expect(separate.target).toBe("_blank")
    action.click(); action.click(); expect(loads).toBe(1); expect(action.disabled).toBe(true)
    render({ ...staleState, busy: true }); resolve(true); await Bun.sleep(0)
    expect(action.disabled).toBe(true)
    render({ ...staleState, connection: { ...staleState.connection!, version: { ...staleState.connection!.version, state: "current" } } })
    expect(action.hidden).toBe(true); expect(separate.hidden).toBe(true); expect(separate.hasAttribute("href")).toBe(false)
  } finally { page.happyDOM.abort() }
})

test("late separate handoff cannot restore an obsolete action; unsupported hosts offer a separate file", async () => {
  const page = latestFixture(), document = page.document as unknown as Document
  let resolve!: (value: string) => void
  const render = setupLatestVersion(document, () => new Promise(done => { resolve = done }))
  try {
    render(staleState); expect(document.getElementById("load-latest")!.hidden).toBe(true); expect(document.getElementById("open-latest-copy")!.hidden).toBe(false)
    render({ ...staleState, connection: undefined }); resolve("doc.haus://word/open?source=old"); await Bun.sleep(0)
    expect(document.getElementById("open-latest-copy")!.hasAttribute("href")).toBe(false)
  } finally { page.happyDOM.abort() }
})
