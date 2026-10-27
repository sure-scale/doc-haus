import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import type { DeskSourcePreview, DeskSources, DeskVersions } from "../../../dochaus/lib/word-source-desk-contract"
import { setupSourceDesk } from "./desk-entry"
import type { DeskOffice } from "./desk-host"
import type { DeskRequest } from "./desk-protocol"
import { mountDesk } from "./desk-ui"

const ORIGIN = "https://localhost:4612"
const link: WordDocumentLink = { version: 1, apiOrigin: ORIGIN, origin: "http://localhost:5274", matter: "synthetic-matter", matterName: "Synthetic matter", document: "synthetic-document", documentName: "Agreement.docx", baseVersion: "version-one", token: "synthetic-capability" }
const preview: DeskSourcePreview = { id: "source-one", name: "Agreement.docx", role: "source", hash: "a".repeat(64), version: { id: "saved-one", number: 1 }, text: "A shared provision in the agreement.", characters: 36, truncated: false, provenance: "saved-version" }
const second: DeskSourcePreview = { ...preview, id: "source-two", name: "Letter.docx", hash: "b".repeat(64), text: "A shared passage in the letter." }
const sources: DeskSources = { matterName: "Synthetic matter", documentName: "Draft.docx", sources: [preview, second], truncated: false }
const versions: DeskVersions = { document: { id: link.document, name: link.documentName }, versions: [1, 2].map((number) => ({ id: `saved-${number}`, number, hash: "c".repeat(64), parent: number === 1 ? null : "saved-1", createdAt: 1_700_000_000_000, origin: "word", label: null, current: number === 2, removed: false })) }

function install(html = '<main id="desk-root"></main>') {
  const page = new Window({ url: `${ORIGIN}/word-addin/pane.html?version=${WORD_ADDIN_VERSION}` })
  page.document.write(html)
  const globals = ["window", "document", "Option", "HTMLButtonElement", "HTMLSelectElement", "HTMLInputElement"] as const
  const previous = globals.map((name) => Object.getOwnPropertyDescriptor(globalThis, name))
  // happy-dom does not expose the browser's Option convenience constructor.
  function Option(text: string, value: string) {
    const option = page.document.createElement("option")
    option.text = text; option.value = value
    return option
  }
  const values: Record<string, unknown> = { window: page, document: page.document, Option, HTMLButtonElement: page.HTMLButtonElement, HTMLSelectElement: page.HTMLSelectElement, HTMLInputElement: page.HTMLInputElement }
  for (const name of globals) Object.defineProperty(globalThis, name, { configurable: true, value: values[name] })
  return {
    page,
    root: document.getElementById("desk-root")!,
    restore() {
      page.dispatchEvent(new page.Event("pagehide"))
      globals.forEach((name, index) => {
        if (previous[index]) Object.defineProperty(globalThis, name, previous[index]!)
        else delete (globalThis as unknown as Record<string, unknown>)[name]
      })
      page.happyDOM.abort()
    },
  }
}

const element = <T extends HTMLElement>(root: HTMLElement, name: string) => root.querySelector(`[data-el="${name}"]`) as T
const button = (root: HTMLElement, name: string) => element<HTMLButtonElement>(root, name)
async function settle() { await Bun.sleep(0) }

test("reading needs no Word selection; manual capture, hiding and list refresh preserve kept saved bytes", async () => {
  const fixture = install()
  const calls: DeskRequest[] = []
  let refreshed = false
  const desk = mountDesk(fixture.root, { async request(input) {
    calls.push(input)
    if (input.action === "sources") return refreshed ? { ...sources, sources: [{ ...preview, hash: "d".repeat(64), version: { id: "new-version", number: 2 } }, second] } : sources
    if (input.action === "source") return input.sourceId === preview.id ? preview : second
    if (input.action === "selection") return { text: "A shared Word passage.", documentHash: "e".repeat(64), saved: false, capturedAt: new Date().toISOString(), navigable: true }
    if (input.action === "show") return null
    throw new Error("Unexpected action")
  } }, { close() {} })
  const choose = async (id: string) => {
    element<HTMLSelectElement>(fixture.root, "source-choice").value = id
    element(fixture.root, "source-choice").dispatchEvent(new fixture.page.Event("change") as unknown as Event)
    await settle()
  }
  try {
    await desk.ready
    expect(calls.map((call) => call.action)).toEqual(["sources"])
    expect(element(fixture.root, "left").hidden).toBe(true)
    expect(element(fixture.root, "columns").classList.contains("has-selection")).toBe(false)
    expect(button(fixture.root, "pin").hidden).toBe(true)
    await choose(preview.id)
    expect(element(fixture.root, "right").textContent).toContain(preview.text)
    expect(button(fixture.root, "pin").hidden).toBe(false)
    button(fixture.root, "pin").click()
    await choose(second.id)
    expect(calls.some((call) => call.action === "selection" || call.action === "show")).toBe(false)
    button(fixture.root, "selection-toggle").click()
    await settle()
    expect(calls.filter((call) => call.action === "selection")).toHaveLength(1)
    expect(element(fixture.root, "left").hidden).toBe(false)
    expect(element(fixture.root, "left").textContent).toContain("Your selected Word text")
    expect(element(fixture.root, "right").textContent).toContain("Reference document")
    button(fixture.root, "find-toggle").click()
    const find = element<HTMLInputElement>(fixture.root, "find")
    find.value = "shared"
    find.dispatchEvent(new fixture.page.Event("input") as unknown as Event)
    expect(fixture.root.querySelectorAll("mark")).toHaveLength(2)
    button(fixture.root, "show").click()
    await settle()
    expect(calls.filter((call) => call.action === "show")).toHaveLength(1)
    button(fixture.root, "selection-toggle").click()
    expect(element(fixture.root, "left").hidden).toBe(true)
    expect(element(fixture.root, "columns").classList.contains("has-selection")).toBe(false)
    expect(element(fixture.root, "right").textContent).toContain(second.text)
    expect(fixture.root.querySelectorAll("mark")).toHaveLength(1)
    fixture.root.querySelector<HTMLButtonElement>(".desk-pin button")!.click()
    expect(element(fixture.root, "right").textContent).toContain(preview.text)
    refreshed = true
    expect(button(fixture.root, "reload").closest("details")?.querySelector("summary")?.textContent).toBe("More options")
    button(fixture.root, "reload").click()
    await settle()
    expect(element(fixture.root, "right").textContent).toContain(preview.text)
    expect(element(fixture.root, "right").textContent).toContain("Version 1 · kept open")
    expect(element(fixture.root, "right").textContent).toContain(preview.hash)
    expect(calls.filter((call) => call.action === "source")).toHaveLength(2)
    expect(calls.filter((call) => call.action === "selection")).toHaveLength(1)
  } finally { desk.dispose(); fixture.restore() }
})

function officeBridge(page: Window, supported: boolean) {
  const handlers = new Map<string, (event: { message: string; origin?: string } | { error: number }) => void>()
  let child: ((event: { message: string; origin?: string }) => void) | undefined
  let openedUrl = ""
  const office: DeskOffice = {
    onReady: async () => {},
    EventType: { DialogMessageReceived: "message", DialogEventReceived: "close", DialogParentMessageReceived: "parent" },
    context: {
      requirements: { isSetSupported: () => supported },
      ui: {
        displayDialogAsync(url, _options, callback) {
          openedUrl = url
          callback({ status: "succeeded", value: {
            addEventHandler(type, handler) { handlers.set(type, handler) },
            messageChild(message) { child?.({ message, origin: ORIGIN }) },
            close() {},
          } })
        },
        addHandlerAsync(_type, handler, callback) { child = handler; callback({ status: "succeeded", value: undefined }) },
        messageParent(message) { handlers.get("message")?.({ message, origin: ORIGIN }) },
      },
    },
  }
  Object.defineProperty(page, "Office", { value: office })
  return { get openedUrl() { return openedUrl } }
}

test("direct Compare versions opens comparison in the floating child without a document-list or selection dependency", async () => {
  const fixture = install(await Bun.file(new URL("../pane.html", import.meta.url)).text())
  const bridge = officeBridge(fixture.page, true)
  const originalFetch = globalThis.fetch
  const calls: string[] = []
  globalThis.fetch = (async (input) => {
    calls.push(String(input))
    return String(input).endsWith("/versions") ? Response.json(versions) : new Response("Document list unavailable", { status: 503 })
  }) as typeof fetch
  try {
    setupSourceDesk(link)
    document.getElementById("compare-desk")!.click()
    await settle()
    const url = new URL(bridge.openedUrl)
    expect(url.searchParams.get("view")).toBe("compare")
    expect(url.searchParams.get("version")).toBe(WORD_ADDIN_VERSION)
    expect(url.href).not.toContain(link.token)
    expect(calls).toHaveLength(0)
    fixture.page.happyDOM.setURL(url.href)
    const childRoot = document.createElement("main")
    childRoot.id = "desk-root"
    document.body.append(childRoot)
    await import("./desk")
    await settle()
    expect(element(childRoot, "compare-tools").hidden).toBe(false)
    expect(element(childRoot, "source-tools").hidden).toBe(true)
    expect(element<HTMLSelectElement>(childRoot, "base").value).toBe("saved-1")
    expect(element<HTMLSelectElement>(childRoot, "current").value).toBe("saved-2")
    expect(button(childRoot, "compare").disabled).toBe(false)
    expect(calls).toEqual([`${ORIGIN}/word/desk/versions`])
    expect(button(childRoot, "reload-versions").closest("details")?.querySelector("summary")?.textContent).toBe("More options")
    button(childRoot, "close").click()
    expect((document.getElementById("compare-desk") as HTMLButtonElement).disabled).toBe(false)
  } finally { fixture.restore(); globalThis.fetch = originalFetch }
})

test("direct comparison preserves its initial mode in automatic and explicit pane fallback", async () => {
  const fixture = install(await Bun.file(new URL("../pane.html", import.meta.url)).text())
  officeBridge(fixture.page, false)
  const originalFetch = globalThis.fetch
  const calls: string[] = []
  globalThis.fetch = (async (input) => { calls.push(String(input)); return Response.json(versions) }) as typeof fetch
  try {
    setupSourceDesk(link)
    const fallback = document.getElementById("pane-desk")!
    document.getElementById("compare-desk")!.click()
    await settle()
    expect(fallback.hidden).toBe(false)
    expect(element(fallback, "title").textContent).toBe("Compare versions")
    expect(element(fallback, "compare-tools").hidden).toBe(false)
    expect(calls).toEqual([`${ORIGIN}/word/desk/versions`])
    button(fallback, "close").click()
    expect(fallback.hidden).toBe(true)
    expect(document.activeElement?.id).toBe("compare-desk")
    document.getElementById("inline-compare-desk")!.click()
    await settle()
    expect(element(fallback, "compare-tools").hidden).toBe(false)
    expect(element(fallback, "left").hidden).toBe(true)
    expect(calls).toEqual([`${ORIGIN}/word/desk/versions`, `${ORIGIN}/word/desk/versions`])
  } finally { fixture.restore(); globalThis.fetch = originalFetch }
})
