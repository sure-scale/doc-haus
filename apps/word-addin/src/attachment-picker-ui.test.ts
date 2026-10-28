import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import type { WordChatMaterial } from "../../../dochaus/lib/word-chat-contract"
import { mountAttachmentPicker } from "./attachment-picker-ui"
import type { DeskRequest } from "./desk-protocol"

const source: WordChatMaterial = { kind: "source", id: "source", hash: "a".repeat(64), name: "Instructions.docx", provenance: "saved-version" }
const playbook: WordChatMaterial = { kind: "playbook", id: "playbook", hash: "b".repeat(64), name: "Review playbook.md", provenance: "instructions" }
const uploaded: WordChatMaterial = { kind: "upload", id: "upload", hash: "c".repeat(64), name: "Notes.txt", provenance: "request-upload" }
function install() {
  const page = new Window({ url: "https://localhost:4612/word-addin/desk.html?view=attachments" })
  page.document.write('<main id="picker"></main>')
  const names = ["window", "document"] as const
  const originals = names.map(name => Object.getOwnPropertyDescriptor(globalThis, name))
  Object.defineProperty(globalThis, "window", { configurable: true, value: page })
  Object.defineProperty(globalThis, "document", { configurable: true, value: page.document })
  return { page, root: document.getElementById("picker")!, restore() {
    names.forEach((name, index) => { if (originals[index]) Object.defineProperty(globalThis, name, originals[index]!); else delete (globalThis as Record<string, unknown>)[name] })
    page.happyDOM.abort()
  } }
}
const element = <T extends HTMLElement = HTMLElement>(root: HTMLElement, name: string) => root.querySelector<T>(`[data-picker="${name}"]`)!
const button = (root: HTMLElement, name: string) => element<HTMLButtonElement>(root, name)
const attach = (root: HTMLElement, id: string) => root.querySelector<HTMLButtonElement>(`[data-material="${id}"]`)!
async function settle() { await Bun.sleep(0) }

test("existing material supports search, clear selection feedback, keyboard tabs and native Close", async () => {
  const fixture = install(), calls: DeskRequest[] = []
  let closed = 0
  const picker = mountAttachmentPicker(fixture.root, { async request(input) {
    calls.push(input)
    if (input.action === "chat-materials") return { materials: [source, playbook], attached: [playbook], matterName: "Synthetic matter", documentName: "Draft.docx" }
    if (input.action === "chat-attach") return { material: source }
    throw new Error("Unexpected action")
  } }, { close() { closed++ } })
  try {
    await picker.ready
    expect(document.activeElement).toBe(element(fixture.root, "search"))
    expect(element(fixture.root, "context").textContent).toBe("Synthetic matter")
    expect(attach(fixture.root, "playbook:playbook").textContent).toContain("Attached")
    const choice = attach(fixture.root, "source:source")
    choice.focus(); choice.click(); choice.click(); await settle()
    expect(calls.filter(input => input.action === "chat-attach")).toHaveLength(1)
    expect(document.activeElement).toBe(choice)
    expect(choice.textContent).toBe("✓ Attached")
    expect(choice.getAttribute("aria-disabled")).toBe("true")
    expect(element(fixture.root, "count").textContent).toBe("2 attached")
    expect(element(fixture.root, "status").textContent).toContain("Instructions.docx attached")
    const search = element<HTMLInputElement>(fixture.root, "search")
    search.value = "playbook"; search.dispatchEvent(new fixture.page.Event("input") as unknown as Event)
    expect(fixture.root.querySelectorAll("[data-material]")).toHaveLength(1)
    button(fixture.root, "existing-tab").dispatchEvent(new fixture.page.KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }) as unknown as Event)
    expect(button(fixture.root, "upload-tab").getAttribute("aria-selected")).toBe("true")
    expect(document.activeElement).toBe(button(fixture.root, "upload-tab"))
    expect(element(fixture.root, "existing").hidden).toBe(true)
    button(fixture.root, "upload-tab").dispatchEvent(new fixture.page.KeyboardEvent("keydown", { key: "Home", bubbles: true }) as unknown as Event)
    expect(document.activeElement).toBe(button(fixture.root, "existing-tab"))
    fixture.root.dispatchEvent(new fixture.page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as Event)
    expect(closed).toBe(1)
    expect(calls.every(input => input.action.startsWith("chat-"))).toBe(true)
  } finally { picker.dispose(); fixture.restore() }
})

test("uploads preserve bytes and destination, reject unsupported files and show attached feedback", async () => {
  const fixture = install(), calls: DeskRequest[] = []
  const picker = mountAttachmentPicker(fixture.root, { async request(input) {
    calls.push(input)
    if (input.action === "chat-materials") return { materials: [], attached: [], matterName: "Synthetic", documentName: "Draft.docx" }
    return { material: uploaded }
  } }, { close() {} })
  try {
    await picker.ready
    button(fixture.root, "upload-tab").click()
    const input = element<HTMLInputElement>(fixture.root, "file")
    let files = [new File(["Synthetic notes"], "Notes.exe")]
    Object.defineProperty(input, "files", { configurable: true, get: () => files })
    input.dispatchEvent(new fixture.page.Event("change") as unknown as Event)
    button(fixture.root, "upload-button").click(); await settle()
    expect(element(fixture.root, "error").textContent).toContain("supported file")
    expect(calls).toHaveLength(1)
    files = [new File(["Synthetic notes"], "Notes.txt")]
    element<HTMLInputElement>(fixture.root, "request-mode").checked = false
    element<HTMLInputElement>(fixture.root, "matter-mode").checked = true
    input.dispatchEvent(new fixture.page.Event("change") as unknown as Event)
    button(fixture.root, "upload-button").click(); await settle()
    expect(calls[1]).toEqual({ action: "chat-upload", name: "Notes.txt", data: btoa("Synthetic notes"), mode: "matter" })
    expect(element(fixture.root, "status").textContent).toBe("Added to the matter and attached.")
    expect(element(fixture.root, "count").textContent).toBe("1 attached")
    expect(attach(fixture.root, "upload:upload").textContent).toContain("Attached")
  } finally { picker.dispose(); fixture.restore() }
})

test("late reads cannot restore a disposed picker and material names render as text", async () => {
  const fixture = install()
  let resolve!: (value: unknown) => void
  const pending = new Promise(done => { resolve = done })
  const picker = mountAttachmentPicker(fixture.root, { request: async () => pending }, { close() {} })
  picker.dispose()
  resolve({ materials: [{ ...source, name: '<img src=x onerror="private()">' }], attached: [], matterName: "Synthetic", documentName: "Draft.docx" })
  await picker.ready
  expect(fixture.root.childNodes).toHaveLength(0)
  const next = mountAttachmentPicker(fixture.root, { request: async () => ({ materials: [{ ...source, name: '<img src=x onerror="private()">' }], attached: [], matterName: "Synthetic", documentName: "Draft.docx" }) }, { close() {} })
  try {
    await next.ready
    expect(fixture.root.querySelector("img")).toBeNull()
    expect(fixture.root.textContent).toContain("<img src=x")
  } finally { next.dispose(); fixture.restore() }
})
