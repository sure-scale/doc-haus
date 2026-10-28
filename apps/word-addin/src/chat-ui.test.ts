import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { mountWordChat, type ChatHost } from "./chat-ui"
import { ChatRequestError, type WordChatApi, type ChatAction } from "./chat-client"
import type { WordChatSend, WordChatTurn } from "../../../dochaus/lib/word-chat-contract"
import type { WordChatProposal, WordChatSnapshot, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"

const passage: WordChatTarget = { id: "target-one", text: "Payment within 14 days.", paragraphs: [{ id: "paragraph-one", text: "Payment within 14 days." }], startOffset: 0, endOffset: 23, capturedAt: "2026-10-28T12:00:00Z", saved: false }
const context: WordChatSnapshot = { text: passage.text, paragraphs: passage.paragraphs, saved: false, capturedAt: passage.capturedAt, truncated: false }
const proposal: WordChatProposal = { id: "proposal-one", documentId: "document-one", before: passage.text, after: "Payment within 30 days.", explanation: "A longer payment period.", openItems: ["Confirm the intended period."], target: passage, createdAt: passage.capturedAt }
function turn(input: Partial<WordChatTurn> = {}): WordChatTurn { return { id: "turn-one", conversationId: "conversation-one", requestId: "request-one", kind: "chat", scope: "document", text: "Explain payment", answer: "Payment is due in 14 days.", status: "complete", sources: [], attachments: [], createdAt: 1, ...input } }
function fixture(overrides: { request?: (action: ChatAction, input: object) => Promise<unknown>; upload?: (conversationId: string, file: File, mode: "request" | "matter") => Promise<unknown>; turns?: WordChatTurn[]; host?: Partial<ChatHost> } = {}) {
  const page = new Window({ url: "https://localhost:4612/word-addin/pane.html" })
  const previous = ["window", "document"].map(key => Object.getOwnPropertyDescriptor(globalThis, key))
  Object.defineProperty(globalThis, "window", { configurable: true, value: page }); Object.defineProperty(globalThis, "document", { configurable: true, value: page.document })
  const root = page.document.createElement("section") as unknown as HTMLElement; page.document.body.append(root as never)
  const calls: { action: ChatAction; input: object }[] = [], scheduled: { run: () => void; at: number; cancelled: boolean }[] = []
  let time = 0
  const conversation = { id: "conversation-one", title: "Document conversation", documentName: "Draft.docx", updatedAt: 1 }
  const api: WordChatApi = {
    async request<T>(action: ChatAction, input: object = {}) {
      calls.push({ action, input }); let result = await overrides.request?.(action, input)
      if (result === undefined) result = action === "conversations" ? { conversations: [conversation] } : action === "conversation" ? { conversation, turns: overrides.turns ?? [] } : action === "materials" ? { materials: [], matterName: "Synthetic matter", documentName: "Draft.docx" } : { turn: turn() }
      return result as T
    },
    async upload<T>(conversationId: string, file: File, mode: "request" | "matter"): Promise<T> { if (overrides.upload) return await overrides.upload(conversationId, file, mode) as T; throw new Error("Not used") },
  }
  const host: ChatHost = { documentName: "Draft.docx", capture: async () => structuredClone(context), captureSelection: async () => structuredClone(passage), apply: async () => {}, now: () => time,
    schedule(run, delay) { const entry = { run, at: time + delay, cancelled: false }; scheduled.push(entry); return () => { entry.cancelled = true } }, ...overrides.host }
  const chat = mountWordChat(root, api, host)
  const get = <T extends HTMLElement>(name: string) => root.querySelector(`[data-chat="${name}"]`) as T
  const click = (name: string) => get<HTMLButtonElement>(name).click()
  return { root, page, chat, calls, get, click,
    async advance(ms: number) { time += ms; const due = scheduled.filter(entry => !entry.cancelled && entry.at <= time); for (const entry of due) { entry.cancelled = true; entry.run() }; await Bun.sleep(0) },
    async send(text: string) { get<HTMLTextAreaElement>("message").value = text; get("message").dispatchEvent(new page.Event("input") as unknown as Event); get("form").dispatchEvent(new page.Event("submit", { cancelable: true }) as unknown as Event); await Bun.sleep(0) },
    restore() { chat.dispose(); ["window", "document"].forEach((key, index) => { if (previous[index]) Object.defineProperty(globalThis, key, previous[index]!); else delete (globalThis as unknown as Record<string, unknown>)[key] }); page.happyDOM.abort() },
  }
}

test("settings and attachments retain typed text and exact optional selected context without generation", async () => {
  const f = fixture()
  try {
    await f.chat.ready
    expect(f.get("settings").hidden).toBe(true); expect(f.get("materials").hidden).toBe(true)
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "Explain this wording"
    f.click("selection"); await Bun.sleep(0); f.click("settings-toggle"); expect(f.get("settings").hidden).toBe(false)
    f.click("attach"); expect(f.get("settings").hidden).toBe(true); expect(f.get("materials").hidden).toBe(false)
    f.click("materials-close"); expect(message.value).toBe("Explain this wording")
    expect(f.get("chips").textContent).toContain(passage.text)
    expect(f.calls.some(call => call.action === "send" || call.action === "suggest")).toBe(false)
    await f.send("Explain this wording")
    const sent = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(sent.scope).toBe("document"); expect(sent.context.saved).toBe(false); expect(sent.context.selection).toEqual(passage)
  } finally { f.restore() }
})

test("answers use compact numbered references and merge repeated source identities without losing distinct quotes", async () => {
  const reference = "playbook:synthetic-engagement-guidance"
  const first = { id: reference, name: "Engagement guidance", quote: "Explain the agreed payment period.", location: "Attached guidance", provenance: "instructions" as const }
  const second = { ...first, quote: "Keep any missing instruction visible.", location: "Open items" }
  const third = { id: "source:42", name: "Instructions.txt", quote: "<img src=x onerror=unsafe()> is literal source text.", location: "Paragraph 2", provenance: "saved-version" as const }
  const original = turn({ answer: `Explain the period [${reference}]. Disclose open items [${reference}]. See instructions [source:42].`, sources: [first, second, first, third] })
  const snapshot = JSON.stringify(original)
  const f = fixture({ turns: [original] })
  try {
    await f.chat.ready
    const displayed = f.root.querySelector(".chat-answer-text")!.textContent
    expect(displayed).toBe("Explain the period [1]. Disclose open items [1]. See instructions [2].")
    expect(displayed).not.toContain(reference)
    const disclosures = f.root.querySelectorAll(".chat-source")
    expect(disclosures).toHaveLength(2)
    expect(disclosures[0]!.querySelector("summary")!.textContent).toBe("[1] Engagement guidance · Playbook")
    expect([...disclosures[0]!.querySelectorAll("p")].map(element => element.textContent)).toEqual([first.quote, second.quote])
    expect(disclosures[0]!.textContent).toContain(first.location); expect(disclosures[0]!.textContent).toContain(second.location)
    expect(disclosures[1]!.textContent).toContain(third.quote); expect(disclosures[1]!.querySelector("img")).toBeNull()
    expect(JSON.stringify(original)).toBe(snapshot)
  } finally { f.restore() }
})

test("Existing and Upload attachment tabs support keyboard focus and retain the composer and chosen context", async () => {
  const source = { kind: "source" as const, id: "source-one", hash: "a".repeat(64), name: "Supporting notes.docx", provenance: "saved-version" as const }
  const uploaded = { ...source, id: "uploaded-one", name: "New notes.txt" }
  const uploads: { conversationId: string; file: string; mode: string }[] = []
  const f = fixture({ request: async action => action === "materials" ? { materials: [source], matterName: "Synthetic matter", documentName: "Draft.docx" } : undefined,
    upload: async (conversationId, file, mode) => { uploads.push({ conversationId, file: file.name, mode }); return { material: uploaded } } })
  try {
    await f.chat.ready; f.click("selection"); await Bun.sleep(0)
    f.get<HTMLTextAreaElement>("message").value = "Keep this question"
    f.click("attach")
    const existing = f.get<HTMLButtonElement>("material-existing-tab"), upload = f.get<HTMLButtonElement>("material-upload-tab")
    expect(f.page.document.activeElement).toBe(existing as never)
    expect(existing.getAttribute("aria-selected")).toBe("true"); expect(f.get("history").hidden).toBe(true)
    existing.dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "End", bubbles: true }) as unknown as Event)
    expect(upload.getAttribute("aria-selected")).toBe("true"); expect(f.page.document.activeElement).toBe(upload as never)
    expect(f.get("material-existing").hidden).toBe(true); expect(f.get("material-upload").hidden).toBe(false)
    upload.dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "Home", bubbles: true }) as unknown as Event)
    f.get("material-list").querySelector<HTMLButtonElement>("button")!.click()
    expect(f.get("materials").hidden).toBe(true); expect(f.get("history").hidden).toBe(false)
    expect(f.get("chips").textContent).toContain(source.name); expect(f.get("chips").textContent).toContain("Selected Word text")
    f.click("attach"); f.click("material-upload-tab")
    f.get<HTMLSelectElement>("upload-mode").value = "matter"
    const input = f.get<HTMLInputElement>("upload")
    Object.defineProperty(input, "files", { configurable: true, value: [new File(["Synthetic notes"], "New notes.txt")] })
    input.dispatchEvent(new f.page.Event("change") as unknown as Event); await Bun.sleep(0)
    expect(uploads).toEqual([{ conversationId: "conversation-one", file: "New notes.txt", mode: "matter" }])
    expect(f.get("chips").textContent).toContain(uploaded.name)
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Keep this question")
    f.get("materials").dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as Event)
    expect(f.get("materials").hidden).toBe(true); expect(f.page.document.activeElement).toBe(f.get("attach") as never)
    expect(f.calls.some(call => call.action === "send" || call.action === "suggest")).toBe(false)
  } finally { f.restore() }
})

test("Stop fences a late foreground poll result or failure so it cannot revive an answer or proposal", async () => {
  for (const fail of [false, true]) {
    let finish!: (value: unknown) => void, reject!: (error: Error) => void
    const response = new Promise((resolve, refusal) => { finish = resolve; reject = refusal })
    const f = fixture({ turns: [turn({ status: "running", answer: "" })], request: async action => action === "poll" ? response : action === "cancel" ? { turn: turn({ status: "cancelled", answer: "" }) } : undefined })
    try {
      await f.chat.ready; await f.advance(1_000)
      expect(f.calls.filter(call => call.action === "poll")).toHaveLength(1)
      f.get<HTMLTextAreaElement>("message").value = "A follow-up draft"
      f.click("stop"); await Bun.sleep(0)
      if (fail) reject(new Error("Outdated poll failure")); else finish({ turn: turn({ status: "complete", proposal, answer: "Outdated answer" }) })
      await Bun.sleep(0); await f.advance(5_000)
      expect(f.root.textContent).not.toContain("Outdated answer"); expect(f.root.textContent).not.toContain("Outdated poll failure")
      expect(f.root.querySelector(".chat-proposal")).toBeNull()
      expect(f.get("activity").textContent).toBe("Stopped"); expect(f.get("stop").hidden).toBe(true)
      expect(f.get<HTMLTextAreaElement>("message").value).toBe("A follow-up draft")
      expect(f.calls.filter(call => call.action === "poll")).toHaveLength(1)
    } finally { finish(undefined); f.restore() }
  }
})

test("lost Send retains its frozen request identifier and context during reconnect", async () => {
  let sends = 0
  const f = fixture({ request: async (action, input) => { if (action !== "send") return; if (++sends === 1) throw new Error("Connection lost"); const body = input as WordChatSend; return { turn: turn({ requestId: body.requestId, text: body.text }) } } })
  try {
    await f.chat.ready; await f.send("Explain payment")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Explain payment")
    await f.send("Explain payment"); expect(sends).toBe(1)
    f.click("reconnect"); await Bun.sleep(0)
    const requests = f.calls.filter(call => call.action === "send")
    expect(requests).toHaveLength(2); expect(requests[1]!.input).toEqual(requests[0]!.input)
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("")
  } finally { f.restore() }
})

test("review shows full wording and open items; deliberate Apply pins the reviewed proposal once", async () => {
  const applied: WordChatProposal[] = []
  const f = fixture({ turns: [turn({ kind: "amendment", proposal })], host: { apply: async value => { applied.push(value) } } })
  try {
    await f.chat.ready
    expect(f.root.textContent).toContain(proposal.before); expect(f.root.textContent).toContain(proposal.after); expect(f.root.textContent).toContain(proposal.openItems[0]!)
    const apply = [...f.root.querySelectorAll("button")].find(button => button.textContent === "Apply") as HTMLButtonElement
    apply.click(); apply.click(); await Bun.sleep(0)
    expect(applied).toEqual([proposal]); expect(f.root.textContent).toContain("Choose Save to publish")
    expect([...f.root.querySelectorAll("button")].some(button => button.textContent === "Apply")).toBe(false)
  } finally { f.restore() }
})

test("a definite admission refusal lets the lawyer correct and resend the draft", async () => {
  let sends = 0
  const f = fixture({ request: async action => { if (action === "send" && ++sends === 1) throw new ChatRequestError("Remove the unavailable attachment.", false) } })
  try {
    await f.chat.ready; await f.send("Explain payment")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Explain payment")
    await f.send("Explain the current payment wording")
    expect(f.calls.filter(call => call.action === "send")).toHaveLength(2)
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("")
  } finally { f.restore() }
})

test("pending Apply fences cached proposal and conversation actions; success clears stale target context without losing the next message", async () => {
  let finish!: () => void
  const waiting = new Promise<void>(resolve => { finish = resolve })
  let changed = false
  const secondProposal = { ...proposal, id: "proposal-two" }
  const f = fixture({ turns: [turn({ kind: "amendment", proposal }), turn({ id: "turn-two", kind: "amendment", proposal: secondProposal })], host: {
    apply: async () => { await waiting; changed = true },
    capture: async () => ({ ...structuredClone(context), text: changed ? proposal.after : proposal.before }),
  } })
  try {
    await f.chat.ready; f.click("selection"); await Bun.sleep(0)
    const actions = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
    actions.find(button => button.textContent === "Refine")!.click()
    expect(f.get("chips").textContent).toContain("Refining amendment")
    const before = f.calls.length
    actions.find(button => button.textContent === "Apply")!.click()
    expect(actions.every(button => button.disabled)).toBe(true)
    expect(f.get<HTMLButtonElement>("new").disabled).toBe(true)
    expect(f.get<HTMLSelectElement>("conversation").disabled).toBe(true)
    for (const action of actions.filter(button => button.textContent === "Decline" || button.textContent === "Refine"))
      action.dispatchEvent(new f.page.MouseEvent("click", { bubbles: true }) as unknown as Event)
    f.click("new")
    await f.send("Keep my next question")
    expect(f.calls).toHaveLength(before)
    expect(f.root.textContent).not.toContain("Declined. Your document is unchanged.")
    finish(); await Bun.sleep(0)
    expect(f.get("chips").textContent).not.toContain("Selected Word text")
    expect(f.get("chips").textContent).not.toContain("Refining amendment")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Keep my next question")
    await f.send("Keep my next question")
    const sent = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(sent.context.text).toBe(proposal.after)
    expect(sent.context.selection).toBeUndefined(); expect(sent.proposal).toBeUndefined()
  } finally { finish(); f.restore() }
})

test("Send freezes selected context and scope before asynchronous Word capture completes", async () => {
  let finish!: (snapshot: WordChatSnapshot) => void
  const captured = new Promise<WordChatSnapshot>(resolve => { finish = resolve })
  const f = fixture({ host: { capture: async () => captured } })
  try {
    await f.chat.ready; f.click("selection"); await Bun.sleep(0)
    await f.send("Explain this passage")
    const scope = f.get<HTMLSelectElement>("scope"); scope.value = "matter"; scope.dispatchEvent(new f.page.Event("change") as unknown as Event)
    f.root.querySelector<HTMLButtonElement>('[aria-label="Remove Selected Word text"]')!.click()
    f.get<HTMLTextAreaElement>("message").value = "A new draft while waiting"
    finish(structuredClone(context)); await Bun.sleep(0)
    const sent = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(sent.scope).toBe("document"); expect(sent.context.selection).toEqual(passage)
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("A new draft while waiting")
  } finally { finish(structuredClone(context)); f.restore() }
})

test("a late selection capture cannot restore the pre-Apply target", async () => {
  let finish!: (target: WordChatTarget) => void
  const captured = new Promise<WordChatTarget>(resolve => { finish = resolve })
  const f = fixture({ turns: [turn({ kind: "amendment", proposal })], host: { captureSelection: async () => captured } })
  try {
    await f.chat.ready; f.click("selection")
    const apply = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Apply")!
    apply.click()
    await Bun.sleep(0); finish(structuredClone(passage)); await Bun.sleep(0)
    expect(f.get("chips").textContent).not.toContain("Selected Word text")
    expect(f.get("activity").textContent).toContain("Applied in Word")
  } finally { finish(structuredClone(passage)); f.restore() }
})

test("verified document replacement clears stale Word context and pauses ideas while retaining draft, attachments and history", async () => {
  const material = { kind: "source" as const, id: "source-one", hash: "a".repeat(64), name: "Supporting notes.docx", provenance: "saved-version" as const }
  const f = fixture({ turns: [turn({ kind: "amendment", proposal })], request: async action => action === "materials" ? { materials: [material], matterName: "Synthetic matter", documentName: "Draft.docx" } : action === "suggest" ? { turn: turn({ id: "idea-one", kind: "suggestions", status: "running", answer: "" }) } : undefined })
  try {
    await f.chat.ready; f.click("selection"); await Bun.sleep(0)
    f.click("attach"); f.get("material-list").querySelector<HTMLButtonElement>("button")!.click()
    const refine = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Refine")!
    refine.click()
    const enabled = f.get<HTMLInputElement>("proactive"); enabled.checked = true; enabled.dispatchEvent(new f.page.Event("change") as unknown as Event)
    await f.advance(2_000); await f.advance(4_000)
    expect(f.root.querySelector(".chat-skeleton")).not.toBeNull()
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "Keep my question"; message.focus()
    f.chat.resetDocumentContext()
    expect(f.get("chips").textContent).not.toContain("Selected Word text"); expect(f.get("chips").textContent).not.toContain("Refining amendment")
    expect(f.get("chips").textContent).toContain(material.name)
    expect(message.value).toBe("Keep my question"); expect(f.page.document.activeElement).toBe(message as never)
    expect(f.get("history").textContent).toContain(proposal.after)
    expect(enabled.checked).toBe(false); expect(f.get("suggestions").hidden).toBe(true)
    expect(f.calls.some(call => call.action === "cancel")).toBe(true)
    const count = f.calls.filter(call => call.action === "suggest").length
    await f.advance(60_000)
    expect(f.calls.filter(call => call.action === "suggest")).toHaveLength(count)
  } finally { f.restore() }
})

test("suggestions require opt-in and stable context; scope changes cancel preparation while Send remains available", async () => {
  const f = fixture({ request: async action => action === "suggest" ? { turn: turn({ kind: "suggestions", status: "running", answer: "" }) } : undefined })
  try {
    await f.chat.ready; await f.advance(5_000)
    expect(f.calls.some(call => call.action === "suggest")).toBe(false)
    const enabled = f.get<HTMLInputElement>("proactive"); enabled.checked = true; enabled.dispatchEvent(new f.page.Event("change") as unknown as Event)
    await f.advance(2_000); expect(f.calls.some(call => call.action === "suggest")).toBe(false)
    await f.advance(4_000); expect(f.calls.filter(call => call.action === "suggest")).toHaveLength(1)
    expect(f.root.querySelector(".chat-skeleton")).not.toBeNull()
    const scope = f.get<HTMLSelectElement>("scope"); scope.value = "matter"; scope.dispatchEvent(new f.page.Event("change") as unknown as Event)
    expect(f.root.querySelector(".chat-skeleton")).toBeNull(); expect(f.calls.some(call => call.action === "cancel")).toBe(true)
    f.click("pause"); expect(f.get("suggestions").hidden).toBe(true)
    await f.send("A foreground question still works"); expect(f.calls.some(call => call.action === "send")).toBe(true)
  } finally { f.restore() }
})

test("a lost suggestion reply does not automatically submit unchanged work under a new identifier", async () => {
  const f = fixture({ request: async action => { if (action === "suggest") throw new Error("Reply lost") } })
  try {
    await f.chat.ready
    const enabled = f.get<HTMLInputElement>("proactive"); enabled.checked = true; enabled.dispatchEvent(new f.page.Event("change") as unknown as Event)
    await f.advance(2_000); await f.advance(4_000); await f.advance(50_000); await f.advance(50_000)
    expect(f.calls.filter(call => call.action === "suggest")).toHaveLength(1)
  } finally { f.restore() }
})
