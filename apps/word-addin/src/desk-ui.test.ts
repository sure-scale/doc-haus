import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import type { DeskComparison, DeskSourcePreview, DeskSources, DeskVersions } from "../../../dochaus/lib/word-source-desk-contract"
import { setupSourceDesk } from "./desk-entry"
import type { DeskOffice } from "./desk-host"
import type { DeskRequest } from "./desk-protocol"
import { mountDesk } from "./desk-ui"

const savedAt = 1_700_000_000_000
const preview: DeskSourcePreview = { id: "one", name: "Instructions.docx", role: "source", hash: "a".repeat(64), version: { id: "saved-one", number: 1, createdAt: savedAt }, text: "Payment terms\nA shared provision.\nA shared instruction.", blocks: [{ kind: "heading", level: 1, text: "Payment terms" }, { kind: "paragraph", text: "A shared provision." }, { kind: "paragraph", text: "A shared instruction." }], characters: 55, truncated: false, provenance: "saved-version" }
const second: DeskSourcePreview = { ...preview, id: "two", name: "Letter.docx", hash: "b".repeat(64), text: "A shared passage in the letter.", blocks: undefined }
const sources: DeskSources = { matterName: "Synthetic matter", documentName: "Draft.docx", sources: [preview, second], truncated: false }
const versions: DeskVersions = { document: { id: "draft", name: "Draft.docx" }, matterName: "Synthetic matter", versions: [1, 2, 3].map((number) => ({ id: `v${number}`, number, hash: String(number).repeat(64), parent: number === 1 ? null : `v${number - 1}`, createdAt: savedAt + number * 60_000, bytes: 1024, origin: "word", label: null, current: number === 3, removed: false })) }
function comparison(base = "v1", current = "v2"): DeskComparison {
  return {
    base: { ...versions.versions.find((version) => version.id === base)!, name: "Draft.docx", trackedChanges: { count: 0, authors: [] } },
    current: { ...versions.versions.find((version) => version.id === current)!, name: "Draft.docx", trackedChanges: { count: 1, authors: ["Synthetic reviewer"] } },
    changes: [{ id: "change-one", kind: "modified", location: { scope: "body", paragraph: 2, side: "right", heading: "Payment terms" }, before: "Payment within 14 days.", after: "Payment within 30 days.", untracked: true }, { id: "change-two", kind: "inserted", location: { scope: "body", paragraph: 3, side: "right" }, after: "A new provision.", untracked: false }],
    untracked: [{ id: "untracked", kind: "inserted", location: { scope: "body", paragraph: 3, side: "right" }, untracked: true }], proposals: [], attribution: "label", context: "saved-versions", truncated: false,
  }
}
function install(html = '<main id="desk-root"></main>') {
  const page = new Window({ url: `https://localhost:4612/word-addin/pane.html?version=${WORD_ADDIN_VERSION}` })
  page.document.write(html)
  const names = ["window", "document", "Option", "HTMLButtonElement", "HTMLSelectElement", "HTMLInputElement"] as const
  const previous = names.map((name) => Object.getOwnPropertyDescriptor(globalThis, name))
  function Option(text: string, value: string) { const option = page.document.createElement("option"); option.text = text; option.value = value; return option }
  const values: Record<string, unknown> = { window: page, document: page.document, Option, HTMLButtonElement: page.HTMLButtonElement, HTMLSelectElement: page.HTMLSelectElement, HTMLInputElement: page.HTMLInputElement }
  for (const name of names) Object.defineProperty(globalThis, name, { configurable: true, value: values[name] })
  return { page, root: document.getElementById("desk-root") ?? document.getElementById("pane-desk")!, restore() {
    page.dispatchEvent(new page.Event("pagehide"))
    names.forEach((name, index) => { if (previous[index]) Object.defineProperty(globalThis, name, previous[index]!); else delete (globalThis as unknown as Record<string, unknown>)[name] })
    page.happyDOM.abort()
  } }
}
const element = <T extends HTMLElement>(root: HTMLElement, name: string) => root.querySelector(`[data-el="${name}"]`) as T
const button = (root: HTMLElement, name: string) => element<HTMLButtonElement>(root, name)
async function settle() { await Bun.sleep(0) }
function deferred<T>() { let resolve!: (value: T) => void; const promise = new Promise<T>((done) => { resolve = done }); return { promise, resolve } }
function pollClock() {
  let run: (() => void) | undefined
  return {
    schedulePoll(next: () => void, delay: number) { expect(delay).toBe(1_000); run = next; return () => { if (run === next) run = undefined } },
    async tick() { const next = run; run = undefined; next?.(); await settle() },
    get scheduled() { return Boolean(run) },
  }
}
async function choose(fixture: ReturnType<typeof install>, id: string) {
  if (element(fixture.root, "document-list").hidden) button(fixture.root, "back-documents").click()
  fixture.root.querySelector<HTMLButtonElement>(`[data-source="${id}"]`)!.click()
  await settle()
}
function changeVersion(fixture: ReturnType<typeof install>, name: "base" | "current", id: string) {
  const control = element<HTMLSelectElement>(fixture.root, name)
  control.value = id; control.dispatchEvent(new fixture.page.Event("change") as unknown as Event)
}

test("document list opens one saved preview with a clear way back and no file tabs", async () => {
  const fixture = install(), calls: DeskRequest[] = []
  const desk = mountDesk(fixture.root, { async request(input) { calls.push(input); if (input.action === "sources") return sources; if (input.action === "source") return input.sourceId === preview.id ? preview : second; throw new Error("Unexpected Word access") } }, { close() {} })
  try {
    await desk.ready
    expect(calls.map((call) => call.action)).toEqual(["sources"])
    expect(element(fixture.root, "document-list").textContent).toContain("Instructions.docx")
    expect(element(fixture.root, "document-list").textContent).toContain("Version 1")
    expect(element(fixture.root, "status").hidden).toBe(true)
    expect(fixture.root.querySelector('[data-el="selection-toggle"]')).toBeNull()
    expect(fixture.root.querySelector('[data-el="open-document"]')).toBeNull()
    expect(fixture.root.querySelector('[data-el="close-document"]')).toBeNull()
    expect(element(fixture.root, "document-bar").hidden).toBe(true)
    await choose(fixture, preview.id)
    expect(element(fixture.root, "document-bar").hidden).toBe(false)
    expect(element(fixture.root, "document-name").textContent).toBe("Instructions.docx")
    expect(button(fixture.root, "back-documents").textContent).toContain("Back to documents")
    expect(fixture.root.querySelectorAll('[role="tab"]')).toHaveLength(0)
    expect(element(fixture.root, "right").querySelector(".desk-text h2")?.textContent).toBe("Payment terms")
    expect(element(fixture.root, "right").textContent).toContain("Read-only")
    button(fixture.root, "back-documents").click()
    expect(element(fixture.root, "right").hidden).toBe(true)
    expect(element(fixture.root, "right").textContent).toBe("")
    expect(element(fixture.root, "document-list").hidden).toBe(false)
    expect((document.activeElement as HTMLElement)?.dataset.source).toBe(preview.id)
    await choose(fixture, second.id)
    expect(element(fixture.root, "document-name").textContent).toBe("Letter.docx")
    expect(element(fixture.root, "right").textContent).toContain(second.text)
    expect(element(fixture.root, "right").textContent).not.toContain(preview.text)
    expect(calls.some((call) => call.action === "selection" || call.action === "show")).toBe(false)
  } finally { desk.dispose(); fixture.restore() }
})

test("automatic refresh and view changes retain opened saved snapshots", async () => {
  const fixture = install(), calls: DeskRequest[] = []
  const clock = pollClock()
  let refreshed = false
  const desk = mountDesk(fixture.root, { async request(input) {
    calls.push(input)
    if (input.action === "sources") return refreshed ? { ...sources, sources: [{ ...preview, hash: "d".repeat(64), version: { id: "new", number: 2, createdAt: savedAt + 60_000 } }, second] } : sources
    if (input.action === "source") return refreshed ? { ...preview, hash: "d".repeat(64), version: { id: "new", number: 2, createdAt: savedAt + 60_000 }, text: "New saved wording.", blocks: undefined } : preview
    if (input.action === "versions") return versions
    if (input.action === "compare") return comparison(input.baseVersion, input.currentVersion)
  } }, { embedded: true, close() {}, schedulePoll: clock.schedulePoll })
  try {
    await desk.ready; await choose(fixture, preview.id)
    refreshed = true; await clock.tick()
    expect(element(fixture.root, "right").textContent).toContain("Version 1")
    expect(element(fixture.root, "right").textContent).toContain(preview.hash)
    await desk.setView("compare"); await desk.setView("documents")
    expect(element(fixture.root, "right").textContent).toContain(preview.hash)
    expect(element(fixture.root, "document-name").textContent).toBe(preview.name)
    expect(calls.filter((call) => call.action === "source")).toHaveLength(1)
    expect(calls.some((call) => call.action === "selection" || call.action === "show")).toBe(false)
    expect(fixture.root.querySelector<HTMLElement>(".desk-head")!.hidden).toBe(true)
    button(fixture.root, "back-documents").click()
    expect(element(fixture.root, "right").textContent).toBe("")
    await choose(fixture, preview.id)
    expect(element(fixture.root, "right").textContent).toContain("Version 2")
    expect(element(fixture.root, "right").textContent).toContain("New saved wording.")
    expect(calls.filter(call => call.action === "source")).toHaveLength(2)
    expect(calls.filter(call => call.action === "source").at(-1)).toMatchObject({ expectedHash: "d".repeat(64) })
  } finally { desk.dispose(); fixture.restore() }
})

test("automatic list updates preserve the reader, focus and exact snapshot without repainting unchanged polls", async () => {
  const fixture = install(), clock = pollClock(), calls: DeskRequest[] = []
  let current = sources
  const desk = mountDesk(fixture.root, { async request(input) {
    calls.push(input)
    if (input.action === "sources") return current
    if (input.action === "source") return preview
    if (input.action === "versions") return versions
    if (input.action === "compare") return comparison(input.baseVersion, input.currentVersion)
    throw new Error("Unexpected action")
  } }, { close() {}, schedulePoll: clock.schedulePoll })
  try {
    await desk.ready
    expect(fixture.root.querySelector('[data-el="reload"]')).toBeNull()
    expect(fixture.root.querySelector('[data-el="find-toggle"]')).toBeNull()
    const noDocumentFind = new fixture.page.KeyboardEvent("keydown", { key: "f", metaKey: true, cancelable: true, bubbles: true })
    fixture.root.dispatchEvent(noDocumentFind as unknown as Event)
    expect(noDocumentFind.defaultPrevented).toBe(false)
    await choose(fixture, preview.id)
    const reader = element(fixture.root, "right").firstChild
    const details = element(fixture.root, "right").querySelector("details")!
    details.open = true
    const passage = element(fixture.root, "right").querySelector<HTMLElement>(".desk-text")!
    passage.focus(); passage.scrollTop = 35
    for (const modifiers of [{ metaKey: true }, { ctrlKey: true }]) {
      const nativeFind = new fixture.page.KeyboardEvent("keydown", { key: "f", ...modifiers, bubbles: true, cancelable: true })
      passage.dispatchEvent(nativeFind as unknown as Event)
      expect(nativeFind.defaultPrevented).toBe(false)
    }
    const mutations: unknown[] = []
    const observer = new fixture.page.MutationObserver(records => mutations.push(...records))
    observer.observe(fixture.page.document.getElementById("desk-root")!, { subtree: true, attributes: true, childList: true, characterData: true })
    try {
      await clock.tick()
      expect(mutations).toHaveLength(0)
      const added = { ...second, id: "new", name: "New document.docx" }
      current = { ...sources, sources: [{ ...preview, hash: "f".repeat(64), version: { id: "latest", number: 2, createdAt: savedAt + 60_000 } }, second, added] }
      await clock.tick()
      expect(element(fixture.root, "right").firstChild).toBe(reader)
      expect(details.open).toBe(true)
      expect(document.activeElement).toBe(passage)
      expect(passage.scrollTop).toBe(35)
      expect(element(fixture.root, "right").textContent).toContain(preview.hash)
      expect(calls.filter(input => input.action === "source")).toHaveLength(1)
      button(fixture.root, "back-documents").click()
      expect(fixture.root.querySelector('[data-source="one"]')?.textContent).toContain("Updated")
      expect(fixture.root.querySelector('[data-source="new"]')?.textContent).toContain("New")
      await desk.setView("compare")
      expect(clock.scheduled).toBe(false)
      const count = calls.length; await clock.tick(); expect(calls).toHaveLength(count)
    } finally { observer.disconnect() }
  } finally { desk.dispose(); fixture.restore() }
})

test("background polling serializes requests, pauses while hidden, retries quietly and disposes pending work", async () => {
  const fixture = install(), clock = pollClock(), pending = deferred<DeskSources>()
  let reads = 0, failing = false
  const desk = mountDesk(fixture.root, { async request() {
    reads++
    if (reads === 2) return pending.promise
    if (failing) throw new Error("Disconnected")
    return sources
  } }, { close() {}, schedulePoll: clock.schedulePoll })
  try {
    await desk.ready
    await clock.tick()
    expect(reads).toBe(2)
    expect(clock.scheduled).toBe(false)
    expect(fixture.root.getAttribute("aria-busy")).toBe("false")
    await clock.tick(); expect(reads).toBe(2)
    pending.resolve(sources); await settle()
    failing = true; await clock.tick()
    expect(element(fixture.root, "update-status").textContent).toContain("Reconnecting automatically")
    expect(element(fixture.root, "status").hidden).toBe(true)
    Object.defineProperty(fixture.page.document, "visibilityState", { configurable: true, value: "hidden" })
    fixture.page.document.dispatchEvent(new fixture.page.Event("visibilitychange"))
    expect(clock.scheduled).toBe(false)
    Object.defineProperty(fixture.page.document, "visibilityState", { configurable: true, value: "visible" })
    fixture.page.document.dispatchEvent(new fixture.page.Event("visibilitychange"))
    failing = false; await clock.tick()
    expect(element(fixture.root, "update-status").hidden).toBe(true)
    desk.dispose()
    expect(clock.scheduled).toBe(false)
    const count = reads; await clock.tick(); expect(reads).toBe(count)
  } finally { desk.dispose(); fixture.restore() }
})

test("comparison loads independently, automatically applies valid pairs and retains applied versions on refresh", async () => {
  const fixture = install(), calls: DeskRequest[] = []
  let latest = false
  const desk = mountDesk(fixture.root, { async request(input) {
    calls.push(input)
    if (input.action === "versions") return { ...versions, versions: latest ? versions.versions : versions.versions.slice(0, 2) }
    if (input.action === "compare") return comparison(input.baseVersion, input.currentVersion)
    throw new Error("Document list unavailable")
  } }, { initialView: "compare", close() {} })
  try {
    await desk.ready
    expect(calls.map((call) => call.action)).toEqual(["versions", "compare"])
    expect(element(fixture.root, "right").textContent).toContain("Version 1 → Version 2")
    expect(element(fixture.root, "right").textContent).toContain("Untracked edit")
    expect(element(fixture.root, "right").textContent).toContain("1 untracked edit")
    expect(element(fixture.root, "right").textContent).not.toContain("0 proposal")
    const detail = element(fixture.root, "right").querySelector("details")!
    const technical = detail.querySelector("details")!
    expect(detail.open).toBe(false)
    expect(technical.open).toBe(false)
    expect(technical.querySelector("summary")?.textContent).toBe("Technical details")
    expect(technical.textContent).toContain(comparison().base.hash)
    const earlier = detail.querySelector('[aria-label="Earlier · Version 1"]')!
    const later = detail.querySelector('[aria-label="Later · Version 2"]')!
    expect(earlier.textContent).toContain("Payment within 14 days.")
    expect(later.textContent).toContain("Payment within 30 days.")
    expect(earlier.textContent).toContain("Saved from Word")
    expect(earlier.textContent).not.toContain("0 tracked")
    expect(later.textContent).toContain("1 tracked change in this version")
    expect(earlier.textContent).not.toContain(comparison().base.hash)
    expect(later.querySelector("time")?.dateTime).toBe(new Date(comparison().current.createdAt).toISOString())
    expect(element(fixture.root, "right").querySelector("h3")?.textContent).toBe("Payment terms")
    button(fixture.root, "next").click()
    expect(element(fixture.root, "right").textContent).toContain("2 of 2")
    expect(document.activeElement).toBe(button(fixture.root, "previous"))
    latest = true; button(fixture.root, "reload-versions").click(); await settle()
    expect(calls.filter((call) => call.action === "compare")).toHaveLength(1)
    expect(element(fixture.root, "right").textContent).toContain("Version 1 → Version 2")
    changeVersion(fixture, "current", "v3"); await settle()
    expect(element(fixture.root, "right").textContent).toContain("Version 1 → Version 3")
    changeVersion(fixture, "current", "v2"); await settle()
    expect(calls.filter((call) => call.action === "compare")).toHaveLength(2)
    expect(element(fixture.root, "right").textContent).toContain("Version 1 → Version 2")
    changeVersion(fixture, "base", "v3"); await settle()
    expect(element(fixture.root, "compare-warning").hidden).toBe(false)
    expect(calls.filter((call) => call.action === "compare")).toHaveLength(2)
    expect(element(fixture.root, "status").hidden).toBe(true)
  } finally { desk.dispose(); fixture.restore() }
})

test("late comparison results cannot replace the latest requested pair", async () => {
  const fixture = install(), pending = deferred<DeskComparison>(), requested: string[] = []
  const desk = mountDesk(fixture.root, { async request(input) {
    if (input.action === "versions") return versions
    if (input.action === "compare") { requested.push(`${input.baseVersion}:${input.currentVersion}`); return requested.length === 1 ? pending.promise : comparison(input.baseVersion, input.currentVersion) }
    throw new Error("Unexpected action")
  } }, { initialView: "compare", close() {} })
  try {
    await settle()
    expect(requested).toEqual(["v2:v3"])
    changeVersion(fixture, "base", "v1")
    pending.resolve(comparison("v2", "v3")); await desk.ready; await settle()
    expect(requested).toEqual(["v2:v3", "v1:v3"])
    expect(element(fixture.root, "right").textContent).toContain("Version 1 → Version 3")
    expect(element(fixture.root, "right").textContent).not.toContain("Version 2 → Version 3")
  } finally { desk.dispose(); fixture.restore() }
})

test("setView preserves intent during a pending document request and disposal prevents resurrection", async () => {
  const fixture = install(), documentList = deferred<DeskSources>(), busy: boolean[] = [], changes: string[] = []
  const desk = mountDesk(fixture.root, { async request(input) { if (input.action === "sources") return documentList.promise; if (input.action === "versions") return versions; if (input.action === "compare") return comparison(input.baseVersion, input.currentVersion); throw new Error("Unexpected action") } }, { embedded: true, close() {}, onBusyChange(value) { busy.push(value) }, onViewChange(view) { changes.push(view) } })
  try {
    await desk.setView("compare")
    expect(element(fixture.root, "compare-tools").hidden).toBe(false)
    documentList.resolve(sources); await desk.ready
    expect(element(fixture.root, "compare-tools").hidden).toBe(false)
    expect(changes).toEqual(["compare"])
    expect(busy).toEqual([true, false])
    expect(fixture.root.getAttribute("aria-busy")).toBe("false")
  } finally { desk.dispose(); fixture.restore() }
  const next = install(), previewRead = deferred<DeskSourcePreview>()
  const reader = mountDesk(next.root, { async request(input) { return input.action === "sources" ? sources : previewRead.promise } }, { close() {} })
  try {
    await reader.ready
    next.root.querySelector<HTMLButtonElement>('[data-source="one"]')!.click()
    reader.dispose(); previewRead.resolve(preview); await settle()
    expect(next.root.childNodes).toHaveLength(0)
  } finally { reader.dispose(); next.restore() }
})

test("a failed document read can retry the same list item and success stays quiet", async () => {
  const fixture = install()
  let fail = true
  const desk = mountDesk(fixture.root, { async request(input) { if (input.action === "sources") return sources; if (fail) throw new Error("This saved document is unavailable."); return { ...preview, text: "<script>not markup</script>", blocks: undefined } } }, { close() {} })
  try {
    await desk.ready; await choose(fixture, preview.id)
    expect(element(fixture.root, "status").textContent).toContain("unavailable")
    expect(element(fixture.root, "document-list").hidden).toBe(false)
    fail = false; await choose(fixture, preview.id)
    expect(element(fixture.root, "right").textContent).toContain("<script>not markup</script>")
    expect(element(fixture.root, "right").querySelector("script")).toBe(null)
    expect(element(fixture.root, "status").hidden).toBe(true)
  } finally { desk.dispose(); fixture.restore() }
})

test("unknown and larger saved pairs require one explicit Compare action", async () => {
  for (const bytes of [undefined, 2 * 1024 * 1024]) {
    const fixture = install(), calls: DeskRequest[] = []
    const desk = mountDesk(fixture.root, { async request(input) {
      calls.push(input)
      if (input.action === "versions") return { ...versions, versions: versions.versions.slice(0, 2).map((version) => ({ ...version, bytes })) }
      if (input.action === "compare") return comparison(input.baseVersion, input.currentVersion)
      throw new Error("Unexpected action")
    } }, { initialView: "compare", close() {} })
    try {
      await desk.ready
      expect(calls.map((call) => call.action)).toEqual(["versions"])
      expect(button(fixture.root, "compare").hidden).toBe(false)
      button(fixture.root, "compare").click(); await settle()
      expect(calls.map((call) => call.action)).toEqual(["versions", "compare"])
      expect(element(fixture.root, "right").textContent).toContain("Version 1 → Version 2")
    } finally { desk.dispose(); fixture.restore() }
  }
})

const ORIGIN = "https://localhost:4612"
const link: WordDocumentLink = { version: 1, apiOrigin: ORIGIN, origin: "http://localhost:5274", matter: "synthetic-matter", matterName: "Synthetic matter", document: "synthetic-draft", documentName: "Draft.docx", baseVersion: "v1", token: "synthetic-capability" }
function officeBridge(page: Window, supported: boolean) {
  const handlers = new Map<string, (event: { message: string; origin?: string } | { error: number }) => void>()
  let child: ((event: { message: string; origin?: string }) => void) | undefined
  let openedUrl = ""
  const office: DeskOffice = {
    onReady: async () => {},
    EventType: { DialogMessageReceived: "message", DialogEventReceived: "close", DialogParentMessageReceived: "parent" },
    context: { requirements: { isSetSupported: () => supported }, ui: {
      displayDialogAsync(url, _options, callback) {
        openedUrl = url
        callback({ status: "succeeded", value: { addEventHandler(type, handler) { handlers.set(type, handler) }, messageChild(message) { child?.({ message, origin: ORIGIN }) }, close() {} } })
      },
      addHandlerAsync(_type, handler, callback) { child = handler; callback({ status: "succeeded", value: undefined }) },
      messageParent(message) { handlers.get("message")?.({ message, origin: ORIGIN }) },
    } },
  }
  Object.defineProperty(page, "Office", { value: office })
  return { get openedUrl() { return openedUrl }, disconnect() { handlers.get("close")?.({ error: 12006 }) } }
}

test("real pane and floating child preserve read-only views and stay terminal after a pending-read disconnect", async () => {
  const fixture = install(await Bun.file(new URL("../pane.html", import.meta.url)).text()), initial = deferred<Response>(), pending = deferred<Response>(), calls: string[] = []
  const bridge = officeBridge(fixture.page, true), previousFetch = globalThis.fetch
  let documentReads = 0
  globalThis.fetch = (async (input, init) => {
    const url = String(input); calls.push(url)
    if (url.endsWith("/sources")) return ++documentReads === 1 ? initial.promise : pending.promise
    if (url.endsWith("/versions")) return Response.json(versions)
    if (url.endsWith("/compare")) { const body = JSON.parse(String(init?.body)); return Response.json(comparison(body.baseVersion, body.currentVersion)) }
    throw new Error("Unexpected Word request")
  }) as typeof fetch
  try {
    setupSourceDesk(link)
    expect((document.getElementById("documents-tab") as HTMLButtonElement).disabled).toBe(true)
    expect((document.getElementById("compare-tab") as HTMLButtonElement).disabled).toBe(true)
    initial.resolve(Response.json(sources)); await settle()
    expect(fixture.root.hidden).toBe(false)
    expect(fixture.root.querySelector<HTMLElement>(".desk-head")!.hidden).toBe(true)
    expect(calls).toEqual([`${ORIGIN}/word/desk/sources`])
    document.getElementById("compare-tab")!.click(); await settle()
    expect(document.getElementById("compare-tab")!.getAttribute("aria-selected")).toBe("true")
    expect(calls.slice(1)).toEqual([`${ORIGIN}/word/desk/versions`, `${ORIGIN}/word/desk/compare`])
    expect(element(fixture.root, "right").textContent).toContain("Version 2 → Version 3")
    document.getElementById("open-desk-window")!.click(); await settle()
    const url = new URL(bridge.openedUrl)
    expect(url.searchParams.get("view")).toBe("compare")
    expect(url.searchParams.get("version")).toBe(WORD_ADDIN_VERSION)
    expect(url.href).not.toContain(link.token)
    fixture.page.happyDOM.setURL(url.href)
    const childRoot = document.createElement("main"); childRoot.id = "desk-root"; document.body.append(childRoot)
    await import("./desk"); await settle()
    expect(element(childRoot, "compare-tools").hidden).toBe(false)
    expect(element(childRoot, "right").textContent).toContain("Version 2 → Version 3")
    expect(document.title).toBe("Synthetic matter · doc.haus")
    button(childRoot, "sources-mode").click(); await settle()
    expect(documentReads).toBe(2)
    expect(childRoot.getAttribute("aria-busy")).toBe("true")
    expect(childRoot.querySelectorAll("button").length).toBeGreaterThan(0)
    bridge.disconnect(); await settle()
    const alert = childRoot.querySelector<HTMLElement>('[role="alert"]')!
    expect(alert.textContent).toContain("connection is closed")
    expect(alert.hidden).toBe(false)
    expect(childRoot.hasAttribute("aria-busy")).toBe(false)
    expect(childRoot.querySelectorAll("button, select, input")).toHaveLength(0)
    expect((document.getElementById("open-desk-window") as HTMLButtonElement).disabled).toBe(false)
    pending.resolve(Response.json(sources)); await settle()
    expect(childRoot.querySelector('[role="alert"]')).toBe(alert)
    expect(childRoot.querySelectorAll("button, select, input")).toHaveLength(0)
    expect(childRoot.classList.contains("source-desk")).toBe(false)
  } finally { pending.resolve(Response.json(sources)); fixture.restore(); globalThis.fetch = previousFetch }
})

test("unsupported wider windows preserve the pane reader and its open snapshot", async () => {
  const fixture = install(await Bun.file(new URL("../pane.html", import.meta.url)).text()), previousFetch = globalThis.fetch
  officeBridge(fixture.page, false)
  globalThis.fetch = (async (input) => String(input).endsWith("/sources") ? Response.json(sources) : Response.json(preview)) as typeof fetch
  try {
    setupSourceDesk(link); await settle(); await choose(fixture, preview.id)
    const hash = element(fixture.root, "right").textContent
    document.getElementById("open-desk-window")!.click(); await settle()
    expect(document.getElementById("desk-message")!.textContent).toContain("Continue in this pane")
    expect(element(fixture.root, "right").textContent).toBe(hash)
    expect(fixture.root.hidden).toBe(false)
    expect((document.getElementById("open-desk-window") as HTMLButtonElement).disabled).toBe(false)
  } finally { fixture.restore(); globalThis.fetch = previousFetch }
})
