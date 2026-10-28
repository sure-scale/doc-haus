import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { setupLatestVersion } from "./latest-version"
import type { PaneState } from "./sync"
const state: PaneState = { status: "not-saved", canSave: false, busy: false, content: "changed", message: "Paused", connection: { version: { base: { id: "v1", number: 1 }, current: { id: "v2", number: 2 }, state: "stale" }, lease: { state: "available", expiresAt: null } } }
function fixture() { const page = new Window(); page.document.body.innerHTML = '<button id="load-latest">Load latest version</button><a id="open-latest-copy" target="_blank"></a><p id="latest-message"></p>'; return page }
test("latest refresh is deliberate and serialized while separate-copy fallback is prepared without mutation", async () => {
  const page = fixture(), document = page.document as unknown as Document
  let loads = 0, resolve!: (result: boolean) => void
  const render = setupLatestVersion(document, async base => { expect(base).toBe("v1"); return "doc.haus://word/open?source=ticket" }, () => { loads++; return new Promise(done => { resolve = done }) })
  try {
    render(state); await Bun.sleep(0)
    const action = document.getElementById("load-latest") as HTMLButtonElement
    const separate = document.getElementById("open-latest-copy") as HTMLAnchorElement
    expect(action.hidden).toBe(false); expect(loads).toBe(0); expect(separate.href).toContain("ticket"); expect(separate.target).toBe("_blank")
    action.click(); action.click(); expect(loads).toBe(1); expect(action.disabled).toBe(true)
    render({ ...state, busy: true }); resolve(true); await Bun.sleep(0)
    expect(action.disabled).toBe(true)
    render({ ...state, connection: { ...state.connection!, version: { ...state.connection!.version, state: "current" } } })
    expect(action.hidden).toBe(true); expect(separate.hidden).toBe(true); expect(separate.hasAttribute("href")).toBe(false)
  } finally { page.happyDOM.abort() }
})
test("late separate handoff cannot restore an obsolete action; unsupported hosts offer a separate file", async () => {
  const page = fixture(), document = page.document as unknown as Document
  let resolve!: (value: string) => void
  const render = setupLatestVersion(document, () => new Promise(done => { resolve = done }))
  try {
    render(state); expect(document.getElementById("load-latest")!.hidden).toBe(true); expect(document.getElementById("open-latest-copy")!.hidden).toBe(false)
    render({ ...state, connection: undefined }); resolve("doc.haus://word/open?source=old"); await Bun.sleep(0)
    expect(document.getElementById("open-latest-copy")!.hasAttribute("href")).toBe(false)
  } finally { page.happyDOM.abort() }
})
