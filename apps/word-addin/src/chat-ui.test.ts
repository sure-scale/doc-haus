import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { mountWordChat, type ChatHost } from "./chat-ui"
import { ChatRequestError, type WordChatApi, type ChatAction } from "./chat-client"
import type { WordChatSend, WordChatTurn } from "../../../dochaus/lib/word-chat-contract"
import type { WordChatProposal, WordChatSnapshot, WordChatTableRowInsertionProposal, WordChatTableRowDeletionProposal, WordChatTableRowProposal, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import type { V2Event } from "@opencode/client"
import { wordChatApplyEligibility } from "./word-chat-apply"

const passage: WordChatTarget = { id: "target-one", text: "Payment within 14 days.", paragraphs: [{ id: "paragraph-one", text: "Payment within 14 days." }], startOffset: 0, endOffset: 23, capturedAt: "2026-10-28T12:00:00Z", saved: false }
const context: WordChatSnapshot = { text: passage.text, paragraphs: passage.paragraphs, saved: false, capturedAt: passage.capturedAt, truncated: false }
const proposal: WordChatProposal = { id: "proposal-one", documentId: "document-one", before: passage.text, after: "Payment within 30 days.", explanation: "A longer payment period.", openItems: ["Confirm the intended period."], target: passage, createdAt: passage.capturedAt }
const rowProposal: WordChatTableRowInsertionProposal = {kind:"table_row_insert",id:"row-one",documentId:"document-one",position:"end",table:{tableIndex:0,rows:[[{id:"cell-1",text:"Fee"},{id:"cell-2",text:"Ten"}]]},cells:["Tax","Five"],explanation:"Add the agreed item.",openItems:[],createdAt:passage.capturedAt}
const deletedRow: WordChatTableRowDeletionProposal = {kind:"table_row_delete",id:"row-delete",documentId:"document-one",table:{tableIndex:0,rows:[[{id:"cell-1",text:"Fee"},{id:"cell-2",text:"Amount"}],[{id:"cell-3",text:"Retainer"},{id:"cell-4",text:"Twenty"}]]},rowIndex:1,cells:["Retainer","Twenty"],explanation:"Remove the superseded retainer.",openItems:[],createdAt:passage.capturedAt}
function turn(input: Partial<WordChatTurn> = {}): WordChatTurn { return { id: "turn-one", conversationId: "conversation-one", requestId: "request-one", kind: "chat", scope: "document", text: "Explain payment", answer: "Payment is due in 14 days.", status: "complete", sources: [], attachments: [], createdAt: 1, ...input } }
function fixture(overrides: { request?: (action: ChatAction, input: object) => Promise<unknown>; upload?: (conversationId: string, file: File, mode: "request" | "matter") => Promise<unknown>; events?: WordChatApi["events"]; turns?: WordChatTurn[]; host?: Partial<ChatHost>; storage?: Pick<Storage, "getItem" | "setItem"> } = {}) {
  const page = new Window({ url: "https://localhost:4612/word-addin/pane.html" })
  if (overrides.storage) Object.defineProperty(page, "localStorage", { configurable: true, value: overrides.storage })
  const previous = ["window", "document"].map(key => Object.getOwnPropertyDescriptor(globalThis, key))
  Object.defineProperty(globalThis, "window", { configurable: true, value: page }); Object.defineProperty(globalThis, "document", { configurable: true, value: page.document })
  const root = page.document.createElement("section") as unknown as HTMLElement; page.document.body.append(root as never)
  const calls: { action: ChatAction; input: object }[] = [], scheduled: { run: () => void; at: number; cancelled: boolean }[] = []
  let time = 0
  const conversation = { id: "conversation-one", title: "Document conversation", documentName: "Draft.docx", updatedAt: 1 }
  const api: WordChatApi = {
    events: overrides.events,
    async request<T>(action: ChatAction, input: object = {}) {
      calls.push({ action, input }); let result = await overrides.request?.(action, input)
      if (result === undefined) result = action === "conversations" ? { conversations: [conversation] } : action === "conversation" ? { conversation, turns: overrides.turns ?? [] } : action === "materials" ? { materials: [], matterName: "Synthetic matter", documentName: "Draft.docx" } : { turn: turn() }
      return result as T
    },
    async upload<T>(conversationId: string, file: File, mode: "request" | "matter"): Promise<T> { if (overrides.upload) return await overrides.upload(conversationId, file, mode) as T; throw new Error("Not used") },
  }
  // Review lifecycle tests use an explicitly capable fake host. Transport-policy
  // tests below supply the real Word eligibility instead.
  const host: ChatHost = { documentName: "Draft.docx", capture: async () => structuredClone(context), captureSelection: async () => structuredClone(passage), apply: async () => {}, applyEligibility: () => ({ supported: true }), now: () => time,
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

test("row card applies once and disables all review actions while pending", async () => {
  let finish!: () => void, calls = 0
  const waiting = new Promise<void>(resolve => { finish = resolve })
  const f = fixture({turns:[turn({kind:"amendment",rowProposal}),turn({id:"other",proposal})],host:{applyRow:async () => {calls++; await waiting}}})
  try {
    await f.chat.ready
    const card = f.root.querySelector('[aria-label="Table row to review"]')!
    expect(card.textContent).toContain("Add row"); expect(card.textContent).toContain("End of table 1")
    expect(card.textContent).toContain("First row: Fee · Ten")
    expect(card.textContent).toContain("Column 1"); expect(card.textContent).toContain("Tax")
    const apply = [...card.querySelectorAll<HTMLButtonElement>("button")].find(b => b.textContent === "Apply")!
    apply.click(); apply.click()
    expect(calls).toBe(1)
    expect([...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].every(b => b.disabled)).toBe(true)
    finish(); await Bun.sleep(0)
    expect(f.root.textContent).toContain("Applied in Word. Save when you’re ready.")
    expect(f.root.querySelector('[aria-label="Table row to review"] .chat-proposal-actions')).toBeNull()
    apply.click(); expect(calls).toBe(1)
  } finally { finish(); f.restore() }
})

test("row deletion review identifies the existing row and consumes Apply once", async () => {
  const applied: WordChatTableRowProposal[] = []
  const f = fixture({turns:[turn({kind:"amendment",rowProposal:deletedRow})],host:{applyRow:async value => { applied.push(value) }}})
  try {
    await f.chat.ready
    const card = f.root.querySelector('[aria-label="Table row to review"]')!
    expect(card.querySelector("h3")?.textContent).toBe("Delete row")
    expect(card.textContent).toContain("Row 2 · Table 1")
    expect(card.textContent).toContain("Retainer")
    expect(card.textContent).toContain("Twenty")
    expect(card.textContent).not.toContain("End of table")
    const apply = [...card.querySelectorAll<HTMLButtonElement>("button")].find(b => b.textContent === "Apply")!
    apply.click(); apply.click(); await Bun.sleep(0)
    expect(applied).toEqual([deletedRow])
    expect(f.root.querySelector('[aria-label="Table row to review"] .chat-proposal-actions')).toBeNull()
    apply.click(); expect(applied).toHaveLength(1)
  } finally { f.restore() }
})

test("row deletion refinement keeps the reviewed source row when the cursor moves", async () => {
  const f = fixture({turns:[turn({kind:"amendment",rowProposal:deletedRow})],host:{captureSelection:async () => structuredClone(passage)}})
  try {
    await f.chat.ready
    const refine = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(b => b.textContent === "Refine")!
    refine.click(); await f.send("Explain why this entire row is being removed.")
    const sent = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(sent.rowProposal).toEqual(deletedRow)
    expect(sent.proposal).toBeUndefined()
    expect(sent.context.selection).toBeUndefined()
  } finally { f.restore() }
})

test("native paragraph insertion review offers Apply without the legacy import boundary warning", async () => {
  let calls = 0
  const applied: WordChatProposal[] = []
  const inserted = { ...proposal, after: `${proposal.before}\rA new independent obligation.`, formatting:[{paragraph:1,style:"Normal" as const}] }
  const f = fixture({ turns: [turn({ kind: "amendment", proposal: inserted })], host: {
    apply: async value => { calls++; applied.push(value) },
    applyEligibility: p => wordChatApplyEligibility(p.target, p.after, p.formatting, true, p.inlineFormatting),
  } })
  try {
    await f.chat.ready
    expect(f.root.textContent).toContain("A new independent obligation.")
    expect(f.root.textContent).not.toContain("paragraph restructuring")
    const apply = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(b => b.textContent === "Apply")!
    expect(apply).toBeDefined()
    apply.click(); await Bun.sleep(0)
    expect(calls).toBe(1)
    expect(applied[0]?.formatting).toEqual([{paragraph:1,style:"Normal"}])
    expect(f.root.querySelector(".chat-proposal-actions")).toBeNull()
  } finally { f.restore() }
})

test("row uncertainty, prior application, unavailable callback and pending review never offer a retry", async () => {
  let calls = 0
  const uncertain = fixture({turns:[turn({kind:"amendment",rowProposal})],host:{applyRow:async () => {calls++; throw Object.assign(new Error("uncertain"),{mutationState:"uncertain"})}}})
  try {
    await uncertain.chat.ready
    const apply = [...uncertain.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(b => b.textContent === "Apply")!
    apply.click(); await Bun.sleep(0); apply.click()
    expect(calls).toBe(1); expect(uncertain.root.textContent).toContain("Check the change in Word")
    expect(uncertain.root.querySelector(".chat-proposal-actions")).toBeNull()
  } finally { uncertain.restore() }
  const prior = fixture({turns:[turn({kind:"amendment",rowProposal,proposalApplied:true})],host:{applyRow:async () => {calls++}}})
  try { await prior.chat.ready; expect(prior.root.textContent).toContain("Previously applied in Word."); expect(prior.root.querySelector(".chat-proposal-actions")).toBeNull() } finally { prior.restore() }
  const unavailable = fixture({turns:[turn({kind:"amendment",rowProposal})]})
  try { await unavailable.chat.ready; expect(unavailable.root.textContent).toContain("not available in this Word session"); expect([...unavailable.root.querySelectorAll(".chat-proposal-actions button")].some(b => b.textContent === "Apply")).toBe(false) } finally { unavailable.restore() }
  const pending = fixture({turns:[turn({kind:"amendment",status:"running",rowProposal})],host:{applyRow:async () => {calls++}}})
  try { await pending.chat.ready; expect(pending.root.querySelector('[aria-label="Table row to review"]')).toBeNull() } finally { pending.restore() }
  const disposed = fixture({turns:[turn({kind:"amendment",rowProposal})],host:{applyRow:async () => {calls++}}})
  try {
    await disposed.chat.ready
    const apply = [...disposed.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(b => b.textContent === "Apply")!
    const beforeDisposeClick = calls
    disposed.chat.dispose(); apply.click(); expect(calls).toBe(beforeDisposeClick)
  } finally { disposed.restore() }
})

test("row refinement sends the pinned proposal and preserves a draft if capture fails", async () => {
  let fail = true
  const f = fixture({turns:[turn({kind:"amendment",rowProposal})],host:{capture:async () => {if (fail) throw Error("Word unavailable"); return structuredClone(context)}}})
  try {
    await f.chat.ready
    const refine = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(b => b.textContent === "Refine")!
    refine.click(); expect(f.get("chips").textContent).toContain("Refining row")
    expect(f.get<HTMLTextAreaElement>("message").placeholder).toContain("row")
    await f.send("Use a different amount")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Use a different amount")
    expect(f.calls.some(call => call.action === "send")).toBe(false)
    fail = false; await f.send("Use a different amount")
    const sent = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(sent.rowProposal).toEqual(rowProposal); expect(sent.proposal).toBeUndefined(); expect(sent.kind).toBe("amendment")
    expect(sent.context.selection).toBeUndefined(); expect(f.get("chips").textContent).not.toContain("Refining row")
  } finally { f.restore() }
})

test("scope, history and New share the composer row; attachments retain draft and exact optional context without settings", async () => {
  const f = fixture()
  try {
    await f.chat.ready
    expect(f.root.querySelector('[data-chat="settings"]')).toBeNull(); expect(f.root.querySelector('[data-chat="settings-toggle"]')).toBeNull()
    expect(f.get("materials").hidden).toBe(true)
    const contextRow = f.root.querySelector(".chat-composer-context")!
    expect(contextRow.contains(f.get("conversation"))).toBe(true); expect(contextRow.contains(f.get("scope-document"))).toBe(true); expect(contextRow.contains(f.get("new"))).toBe(true)
    expect(f.get("form").contains(contextRow)).toBe(true)
    expect(f.get("conversation").textContent?.trim()).toBe(""); expect(f.get("conversation").getAttribute("aria-label")).toBe("Conversations")
    expect(f.get("new").textContent).toBe("New")
    expect(f.get("scroll").contains(f.get("history"))).toBe(true); expect(f.get("scroll").contains(f.get("suggestions"))).toBe(true)
    expect(f.get("history").getAttribute("role")).toBe("log"); expect(f.get("history").contains(f.get("suggestions"))).toBe(false)
    expect(f.get("scroll").contains(f.get("form"))).toBe(false)
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "Explain this wording"
    f.click("attach"); expect(f.get("materials").hidden).toBe(false)
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
    await f.chat.ready
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

test("Stop fences a late reconciliation result or failure so it cannot revive an answer or proposal", async () => {
  for (const fail of [false, true]) {
    let finish!: (value: unknown) => void, reject!: (error: Error) => void
    const response = new Promise((resolve, refusal) => { finish = resolve; reject = refusal })
    const f = fixture({ turns: [turn({ status: "running", answer: "" })], request: async action => action === "poll" ? response : action === "cancel" ? { turn: turn({ status: "cancelled", answer: "" }) } : undefined })
    try {
      await f.chat.ready; await f.advance(1_000)
      expect(f.calls.filter(call => call.action === "poll")).toHaveLength(0)
      await f.advance(14_000)
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

test("Word shares native event hydration, ignores foreign sessions and keeps a draft through validated updates", async () => {
  let output!: ReadableStreamDefaultController<V2Event>, subscriptions = 0, cancelled = false
  const events: NonNullable<WordChatApi["events"]> = (_conversationId, signal) => ({
    async *[Symbol.asyncIterator]() {
      subscriptions++
      const stream = new ReadableStream<V2Event>({ start(controller) { output = controller }, cancel() { cancelled = true } })
      const reader = stream.getReader(), abort = () => { void reader.cancel() }
      signal.addEventListener("abort", abort, { once: true })
      try { while (!signal.aborted) { const result = await reader.read(); if (result.done) return; yield result.value } }
      finally { signal.removeEventListener("abort", abort); await reader.cancel() }
    },
  })
  const f = fixture({ events, turns: [turn({ status: "running", answer: "" })], request: async action => action === "poll" ? { turn: turn({ answer: "Validated answer", proposal }) } : undefined })
  try {
    await f.chat.ready
    expect(subscriptions).toBe(1)
    f.get<HTMLTextAreaElement>("message").value = "A follow-up draft"
    await f.advance(1_000); expect(f.calls.some(call => call.action === "poll")).toBe(false)
    for (const sessionID of ["other-conversation", "private-reviewer"]) output.enqueue({ id: sessionID, created: 1, type: "session.status", data: { sessionID, status: { type: "idle" } } })
    await Bun.sleep(120); expect(f.calls.some(call => call.action === "poll")).toBe(false)
    output.enqueue({ id: "own-turn", created: 1, type: "session.status", data: { sessionID: "turn-one", status: { type: "idle" } } })
    await Bun.sleep(120)
    expect(f.calls.filter(call => call.action === "poll")).toHaveLength(1)
    expect(f.root.textContent).toContain("Validated answer"); expect(f.root.querySelector(".chat-proposal")).not.toBeNull()
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("A follow-up draft")
    f.chat.dispose(); await Bun.sleep(0); expect(cancelled).toBe(true)
  } finally { f.restore() }
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
    expect(applied).toEqual([proposal]); expect(f.root.querySelector('.chat-local-note[role="status"]')!.textContent).toBe("Applied in Word. Save when you’re ready.")
    expect(f.root.querySelectorAll(".chat-proposal-actions button")).toHaveLength(0)
    expect(f.get("activity").textContent).toBe("")
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
    await f.chat.ready
    const actions = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
    actions.find(button => button.textContent === "Refine")!.click()
    expect(f.get("chips").textContent).toContain("Refining amendment")
    const before = f.calls.length
    actions.find(button => button.textContent === "Apply")!.click()
    expect(actions.every(button => button.disabled)).toBe(true)
    expect(f.get<HTMLButtonElement>("new").disabled).toBe(true)
    expect(f.get<HTMLButtonElement>("conversation").disabled).toBe(true)
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
    await f.chat.ready
    await f.send("Explain this passage")
    f.click("scope-matter")
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
    await f.chat.ready
    const apply = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Apply")!
    apply.click()
    await Bun.sleep(0); finish(structuredClone(passage)); await Bun.sleep(0)
    expect(f.get("chips").textContent).not.toContain("Selected Word text")
    expect(f.root.querySelector('.chat-local-note[role="status"]')!.textContent).toContain("Applied in Word")
  } finally { finish(structuredClone(passage)); f.restore() }
})

test("verified document replacement clears stale Word context and pauses ideas while retaining draft, attachments and history", async () => {
  const material = { kind: "source" as const, id: "source-one", hash: "a".repeat(64), name: "Supporting notes.docx", provenance: "saved-version" as const }
  const f = fixture({ turns: [turn({ kind: "amendment", proposal })], request: async action => action === "materials" ? { materials: [material], matterName: "Synthetic matter", documentName: "Draft.docx" } : action === "suggest" ? { turn: turn({ id: "idea-one", kind: "suggestions", status: "running", answer: "" }) } : undefined })
  try {
    await f.chat.ready
    f.click("attach"); f.get("material-list").querySelector<HTMLButtonElement>("button")!.click()
    const refine = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Refine")!
    refine.click()
    await f.advance(2_000); await f.advance(4_000)
    expect(f.root.querySelector(".chat-skeleton")).toBeNull()
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "Keep my question"; message.focus()
    f.chat.resetDocumentContext()
    expect(f.get("chips").textContent).not.toContain("Selected Word text"); expect(f.get("chips").textContent).not.toContain("Refining amendment")
    expect(f.get("chips").textContent).toContain(material.name)
    expect(message.value).toBe("Keep my question"); expect(f.page.document.activeElement).toBe(message as never)
    expect(f.get("history").textContent).toContain(proposal.after)
    expect(f.get("suggestions").hidden).toBe(true)
    expect(f.calls.some(call => call.action === "suggest")).toBe(false)
    const count = f.calls.filter(call => call.action === "suggest").length
    await f.advance(60_000)
    expect(f.calls.filter(call => call.action === "suggest")).toHaveLength(count)
  } finally { f.restore() }
})

test("suggestions are on by default and wait for stable context; scope changes cancel preparation while Send remains available", async () => {
  const f = fixture({ request: async action => action === "suggest" ? { turn: turn({ kind: "suggestions", status: "running", answer: "" }) } : undefined })
  try {
    await f.chat.ready; await f.advance(5_000)
    expect(f.calls.some(call => call.action === "suggest")).toBe(false)
    await f.advance(2_000); expect(f.calls.some(call => call.action === "suggest")).toBe(false)
    await f.advance(4_000); expect(f.calls.filter(call => call.action === "suggest")).toHaveLength(1)
    const sent = f.calls.find(call => call.action === "suggest")!.input as WordChatSend
    expect(sent.context.text).toBe(context.text); expect(sent.context.paragraphs).toEqual(context.paragraphs); expect(sent.context.selection).toEqual(passage)
    expect(sent).not.toHaveProperty("suggestionTarget"); expect(f.root.querySelector('[data-chat="target"]')).toBeNull()
    expect(f.root.querySelector(".chat-skeleton")).not.toBeNull()
    f.click("scope-matter")
    expect(f.root.querySelector(".chat-skeleton")).toBeNull(); expect(f.calls.some(call => call.action === "cancel")).toBe(true)
    f.click("pause"); expect(f.get("suggestions").hidden).toBe(true)
    await f.send("A foreground question still works"); expect(f.calls.some(call => call.action === "send")).toBe(true)
  } finally { f.restore() }
})

test("a lost suggestion reply does not automatically submit unchanged work under a new identifier", async () => {
  const f = fixture({ request: async action => { if (action === "suggest") throw new Error("Reply lost") } })
  try {
    await f.chat.ready
    await f.advance(2_000); await f.advance(4_000); await f.advance(50_000); await f.advance(50_000)
    expect(f.calls.filter(call => call.action === "suggest")).toHaveLength(1)
  } finally { f.restore() }
})

test("empty chats show three usable prompts immediately; contextual ideas fill the draft directly and supplement fewer than three", async () => {
  const idea = { title: "Clarify the payment period", explanation: "Discuss how the payment wording operates.", passage: passage.text }
  const f = fixture({ request: async action => action === "suggest" ? { turn: turn({ id: "ideas", kind: "suggestions", suggestions: [idea], suggestionCoverage: { reviewed: 1, total: 3 } }) } : undefined })
  try {
    await f.chat.ready
    expect(f.root.querySelectorAll(".chat-empty .chat-starter")).toHaveLength(3)
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "My existing draft"
    f.root.querySelector<HTMLButtonElement>(".chat-starter")!.click()
    expect(message.value).toStartWith("My existing draft\n\n"); expect(message.value).toContain("Summarise")
    expect(f.calls.some(call => call.action === "send" || call.action === "suggest")).toBe(false)
    await f.send("A first question"); expect(f.root.querySelector(".chat-empty")).toBeNull()
    f.click("new"); await Bun.sleep(0)
    expect(f.root.querySelectorAll(".chat-empty .chat-starter")).toHaveLength(3)
    await f.advance(2_000); await f.advance(4_000)
    const prompts = f.root.querySelectorAll<HTMLButtonElement>(".chat-empty .chat-starter")
    expect(prompts).toHaveLength(3); expect(prompts[0]!.textContent).toBe(idea.title)
    expect(f.root.querySelector(".chat-empty")!.textContent).toContain("1 of 3 passage groups")
    const before = f.calls.length; message.value = "Keep this thought"; prompts[0]!.click()
    expect(message.value).toStartWith("Keep this thought\n\n"); expect(message.value).toContain(idea.explanation)
    expect(f.calls).toHaveLength(before); expect(f.root.querySelector(".chat-idea")).toBeNull()
  } finally { f.restore() }
})

test("main-app Word defaults govern new chats across documents without storing or changing live scope in the pane", async () => {
  const values = new Map<string, string>(), storage = { getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => { values.set(key, value) } }
  let preferences = { defaultScope: "document" as "document" | "matter", suggestions: true }
  const request = async (action: ChatAction) => action === "materials" ? { materials: [], matterName: "Synthetic matter", documentName: "Draft.docx", preferences } : undefined
  let f = fixture({ storage, request })
  try {
    await f.chat.ready; f.click("scope-matter")
    expect(f.get("scope-matter").getAttribute("aria-pressed")).toBe("true"); expect(values.size).toBe(0)
    f.get<HTMLTextAreaElement>("message").value = "Private draft should never enter preferences"
    f.click("new"); await Bun.sleep(0)
    expect(f.get("scope-document").getAttribute("aria-pressed")).toBe("true")
    preferences = { defaultScope: "matter", suggestions: false }
    // A focus refresh adopts global defaults, but the explicit scope for this
    // current chat remains unchanged until a new conversation is opened.
    f.click("scope-document"); f.page.dispatchEvent(new f.page.Event("focus")); await Bun.sleep(0)
    expect(f.get("scope-document").getAttribute("aria-pressed")).toBe("true")
    expect(f.root.querySelector<HTMLElement>(".chat-empty-prompts")!.hidden).toBe(true)
    expect(values.size).toBe(0)
    f.click("new"); await Bun.sleep(0)
    expect(f.get("scope-matter").getAttribute("aria-pressed")).toBe("true")
    expect(f.root.querySelector<HTMLElement>(".chat-empty-prompts")!.hidden).toBe(true)
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Private draft should never enter preferences")
    f.restore(); f = fixture({ storage, request, host: { documentName: "Another.docx" } }); await f.chat.ready
    expect(f.get("scope-matter").getAttribute("aria-pressed")).toBe("true")
    expect(f.root.querySelector<HTMLElement>(".chat-empty-prompts")!.hidden).toBe(true)
    expect(values.size).toBe(0)
  } finally { f.restore() }
})

test("selection watching debounces and fences Send, ignores incidental IDs, preserves pane-focus selection and permits deliberate reselection", async () => {
  let selected: WordChatTarget | undefined = structuredClone(passage), changed = () => {}, captures = 0, stopped = 0
  const f = fixture({ host: { captureSelection: async () => { captures++; return selected && structuredClone(selected) }, watchSelection(callback) { changed = callback; return () => { stopped++ } } } })
  try {
    await f.chat.ready; expect(captures).toBe(1); expect(f.get("chips").textContent).toContain(passage.text)
    Object.defineProperty(f.page.document, "hasFocus", { configurable: true, value: () => true })
    selected = { ...passage, id: "new-random-capture", capturedAt: "2026-10-29T12:00:00Z" }
    changed(); changed(); await f.advance(199); expect(captures).toBe(1)
    await f.send("Explain this exact selection"); expect(f.calls.some(call => call.action === "send")).toBe(false)
    expect(f.get<HTMLButtonElement>("send").disabled).toBe(true)
    await f.advance(1); expect(captures).toBe(2)
    await f.send("Explain this exact selection")
    expect((f.calls.find(call => call.action === "send")!.input as WordChatSend).context.selection?.id).toBe(passage.id)
    selected = undefined; changed(); await f.advance(200)
    expect(f.get("chips").textContent).toContain(passage.text)
    f.root.querySelector<HTMLButtonElement>('[aria-label="Remove Selected Word text"]')!.click()
    selected = { ...passage, id: "another-capture" }; changed(); await f.advance(200)
    expect(f.get("chips").textContent).not.toContain("Selected Word text")
    selected = undefined; changed(); await f.advance(200)
    selected = passage; changed(); await f.advance(200)
    expect(f.get("chips").textContent).not.toContain("Selected Word text")
    Object.defineProperty(f.page.document, "hasFocus", { value: () => false })
    selected = undefined; changed(); await f.advance(200)
    selected = passage; changed(); await f.advance(200)
    expect(f.get("chips").textContent).toContain(passage.text)
  } finally { f.restore() }
  expect(stopped).toBe(1)
})

test("late selection captures cannot restore replaced context or change a pinned refinement", async () => {
  let changed = () => {}, finish!: (value: WordChatTarget) => void, pending = false
  const waiting = new Promise<WordChatTarget>(done => { finish = done })
  const f = fixture({ turns: [turn({ kind: "amendment", proposal })], host: { captureSelection: async () => pending ? waiting : structuredClone(passage), watchSelection(callback) { changed = callback; return () => {} } } })
  try {
    await f.chat.ready; pending = true; changed(); await f.advance(200)
    f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")[0]!.click()
    finish({ ...passage, id: "late-capture" }); await Bun.sleep(0)
    expect(f.get("chips").textContent).toContain("Refining amendment")
    await f.send("Keep the meaning but simplify it")
    const sent = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(sent.proposal).toEqual(proposal); expect(sent.context.selection).toEqual(proposal.target)
    f.chat.resetDocumentContext(); changed(); await f.advance(200); await Bun.sleep(0)
    expect(f.get("chips").textContent).not.toContain("Selected Word text")
    expect(f.get("chips").textContent).not.toContain("Refining amendment")
    expect(f.get<HTMLTextAreaElement>("message").placeholder).toBe("Ask doc.haus…")
  } finally { finish(passage); f.restore() }
})

test("fresh buffer validation blocks stale selected text even when the old wording remains elsewhere", async () => {
  const changed = { ...context, text: `Payment within 30 days.\n${passage.text}`, paragraphs: [{ id: "paragraph-one", text: "Payment within 30 days." }, { id: "paragraph-two", text: passage.text }] }
  const f = fixture({ host: { capture: async () => structuredClone(changed) } })
  try {
    await f.chat.ready; await f.send("Explain my selected text")
    expect(f.calls.some(call => call.action === "send")).toBe(false)
    expect(f.get("activity").textContent).toContain("Selected text changed")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Explain my selected text")
    await f.advance(2_000); await f.advance(4_000)
    expect(f.calls.some(call => call.action === "suggest")).toBe(false)
    expect(f.get("chips").textContent).not.toContain("Selected Word text")
    await f.send("Explain my selected text")
    expect((f.calls.find(call => call.action === "send")!.input as WordChatSend).context.selection).toBeUndefined()
  } finally { f.restore() }
})

test("floating attachments bind to one conversation until close, reject late callbacks and fall back inline on unsupported hosts", async () => {
  let options!: Parameters<NonNullable<ChatHost["openAttachments"]>>[0], close!: () => void
  const waiting = new Promise<void>(done => { close = done })
  const material = { kind: "source" as const, id: "notes", hash: "a".repeat(64), name: "Supporting notes.docx", provenance: "saved-version" as const }
  const f = fixture({ host: { openAttachments: async input => { options = input; await waiting } } })
  try {
    await f.chat.ready; f.click("attach")
    expect(options.conversationId).toBe("conversation-one"); expect(options.attached).toEqual([])
    expect(f.get("materials").hidden).toBe(true)
    expect(f.get<HTMLButtonElement>("new").disabled).toBe(true); expect(f.get<HTMLButtonElement>("conversation").disabled).toBe(true)
    const before = f.calls.filter(call => call.action === "conversation").length
    f.get("new").dispatchEvent(new f.page.MouseEvent("click", { bubbles: true }) as unknown as Event)
    await f.send("Preserve my draft")
    expect(f.calls.filter(call => call.action === "conversation")).toHaveLength(before)
    expect(f.calls.some(call => call.action === "send")).toBe(false)
    options.onAttach(material); options.onAttach(material)
    expect([...f.root.querySelectorAll(".chat-chip summary")].filter(row => row.textContent === material.name)).toHaveLength(1)
    close(); await Bun.sleep(0)
    options.onAttach({ ...material, id: "late", name: "Late result.docx" })
    expect(f.get("chips").textContent).not.toContain("Late result")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Preserve my draft")
    expect(f.get<HTMLButtonElement>("new").disabled).toBe(false)
  } finally { close(); f.restore() }
  const fallback = fixture({ host: { openAttachments: async () => { throw new Error("Unsupported") } } })
  try {
    await fallback.chat.ready; fallback.click("attach"); await Bun.sleep(0)
    expect(fallback.get("materials").hidden).toBe(false); expect(fallback.get("activity").textContent).toContain("Choose material here")
  } finally { fallback.restore() }
})

test("generated title refresh keeps the current draft, scope, selection and history", async () => {
  let resolveMetadata!: (value: unknown) => void, reads = 0
  const metadata = new Promise(resolve => { resolveMetadata = resolve })
  const initial = { id: "conversation-one", title: "Word · Draft.docx", documentName: "Draft.docx", updatedAt: 1 }
  const f = fixture({ request: async action => action === "conversation" ? ++reads === 1 ? { conversation: initial, turns: [turn()] } : metadata : undefined })
  try {
    await f.chat.ready; f.click("scope-matter"); await f.advance(200); await f.send("First request")
    f.get<HTMLTextAreaElement>("message").value = "Next draft"
    resolveMetadata({ conversation: { ...initial, title: "Generated title" }, turns: [turn({ answer: "Wrong history" })] })
    await Bun.sleep(0)
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Next draft")
    expect(f.get("scope-matter").getAttribute("aria-pressed")).toBe("true")
    expect(f.root.textContent).toContain(passage.text)
    expect(f.root.textContent).toContain("Payment is due in 14 days.")
    expect(f.root.textContent).not.toContain("Wrong history")
    f.click("conversation")
    expect(f.get("conversation-list").querySelector('[aria-current="true"]')!.textContent).toContain("Generated title")
  } finally { resolveMetadata(undefined); f.restore() }
})

test("completed turns refresh only metadata and discard stale responses after switching chats", async () => {
  let resolveMetadata!: (value: unknown) => void, reads = 0
  const metadata = new Promise(resolve => { resolveMetadata = resolve })
  const first = { id: "conversation-one", title: "Word · Draft.docx", documentName: "Draft.docx", updatedAt: 1 }
  const second = { ...first, id: "conversation-two", title: "Second conversation" }
  const f = fixture({ turns: [turn()], request: async (action, input) => {
    if (action === "conversations") return { conversations: [first, second] }
    if (action === "conversation") {
      const id = (input as { conversationId?: string }).conversationId
      if (id === second.id) return { conversation: second, turns: [] }
      if (++reads === 1) return { conversation: first, turns: [turn()] }
      return metadata
    }
  } })
  try {
    await f.chat.ready
    f.click("scope-matter")
    await f.advance(200)
    await f.send("First request")
    f.get<HTMLTextAreaElement>("message").value = "Next draft"
    expect(f.root.textContent).toContain("Payment is due in 14 days.")
    expect(f.root.textContent).toContain(passage.text)
    expect(f.get("scope-matter").getAttribute("aria-pressed")).toBe("true")
    f.click("conversation")
    f.get("conversation-list").querySelector<HTMLButtonElement>('[data-conversation="conversation-two"]')!.click()
    await Bun.sleep(0)
    resolveMetadata({ conversation: { ...first, title: "Late generated title" }, turns: [turn({ answer: "Stale history" })] })
    await Bun.sleep(0)
    f.click("conversation")
    expect(f.get("conversation-list").querySelector('[aria-current="true"]')!.textContent).toContain("Second conversation")
    expect(f.root.textContent).not.toContain("Late generated title")
    expect(f.root.textContent).not.toContain("Stale history")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Next draft")
  } finally { resolveMetadata(undefined); f.restore() }
})

test("the conversation picker distinguishes dates, refreshes derived titles and keeps the draft while switching", async () => {
  const one = { id: "conversation-one", title: "Repeated question", documentName: "Draft.docx", updatedAt: 1_000 }
  const two = { ...one, id: "conversation-two", updatedAt: 100_000 }
  const third = { ...one, id: "conversation-three", title: "Word · Draft.docx" }
  let thirdOpened = false
  const f = fixture({ request: async (action, input) => action === "conversations" ? { conversations: [one, two] } : action === "conversation" ? (input as { conversationId?: string }).conversationId === one.id ? { conversation: { ...one, title: "Derived first question" }, turns: [] } : (input as { conversationId?: string }).conversationId === two.id ? { conversation: two, turns: [turn({ scope: "matter", conversationId: two.id })] } : { conversation: (input as { conversationId?: string }).conversationId === third.id && thirdOpened ? { ...third, title: "Generated question title" } : (thirdOpened = true, third), turns: [] } : undefined })
  try {
    await f.chat.ready; f.get<HTMLTextAreaElement>("message").value = "Keep the next question"; f.click("conversation")
    const choices = f.get("conversation-list").querySelectorAll<HTMLButtonElement>("button")
    expect(choices).toHaveLength(2); expect(choices[0]!.textContent).toContain("Derived first question")
    expect(choices[0]!.getAttribute("aria-current")).toBe("true")
    expect(choices[0]!.querySelector("small")!.textContent).not.toBe(choices[1]!.querySelector("small")!.textContent)
    choices[1]!.click(); await Bun.sleep(0)
    expect(f.get("conversations").hidden).toBe(true); expect(f.get("scope-matter").getAttribute("aria-pressed")).toBe("true")
    expect(f.get<HTMLTextAreaElement>("message").value).toBe("Keep the next question")
    f.click("new"); await Bun.sleep(0); expect(f.get("scope-document").getAttribute("aria-pressed")).toBe("true")
    await f.send("A meaningful\n first question"); f.click("conversation")
    expect(f.get("conversation-list").querySelector('[aria-current="true"]')!.textContent).toContain("Generated question title")
    f.get("conversations").dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as Event)
    expect(f.get("conversations").hidden).toBe(true); expect(f.page.document.activeElement === f.get("conversation") as never).toBe(true)
    expect(f.root.querySelector('select[data-chat="conversation"]')).toBeNull()
  } finally { f.restore() }
})

test("Apply errors stay beside the retained review without a duplicate composer banner or competing suggestions", async () => {
  const error = `This change needs a clearer paragraph boundary. ${"Synthetic-long-detail-".repeat(30)}`
  const f = fixture({ turns: [turn({ kind: "amendment", proposal })], host: { apply: async () => { throw new Error(error) } },
    request: async action => action === "suggest" ? { turn: turn({ id: "ideas-after-review", kind: "suggestions", status: "running" }) } : undefined })
  try {
    await f.chat.ready
    const actions = () => [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
    actions().find(button => button.textContent === "Refine")!.click()
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "Keep this refinement draft"
    actions().find(button => button.textContent === "Apply")!.click(); await Bun.sleep(0)
    const note = f.root.querySelector(".chat-proposal-error")!
    expect(note.textContent).toBe(error); expect(note.getAttribute("role")).toBe("alert")
    expect(f.get("activity").textContent).toBe("")
    expect(f.root.textContent!.split(error)).toHaveLength(2)
    expect(f.root.querySelector(".chat-proposal")!.textContent).toContain(proposal.before)
    expect(f.root.querySelector(".chat-proposal")!.textContent).toContain(proposal.after)
    expect(f.get("chips").textContent).toContain("Refining amendment")
    expect(message.value).toBe("Keep this refinement draft")
    expect(actions().every(button => !button.disabled)).toBe(true)
    await f.advance(2_000); await f.advance(4_000)
    expect(f.get("suggestions").hidden).toBe(true); expect(f.calls.some(call => call.action === "suggest")).toBe(false)
    actions().find(button => button.textContent === "Decline")!.click()
    expect(f.root.querySelector(".chat-proposal-error")).toBeNull()
    await f.advance(2_000); await f.advance(4_000)
    expect(f.calls.filter(call => call.action === "suggest")).toHaveLength(1)
  } finally { f.restore() }
})

test("unsupported paragraph splits keep the full review and Refine action without offering an Apply that must fail", async () => {
  let applied = false
  const unsupported = { ...proposal, after: "Payment is due.\nWithin the agreed period." }
  const f = fixture({ turns: [turn({ kind: "amendment", proposal: unsupported })], host: { apply: async () => { applied = true } } })
  try {
    await f.chat.ready
    const card = f.root.querySelector(".chat-proposal")!
    expect(card.textContent).toContain(unsupported.after)
    expect(card.querySelector(".chat-proposal-limit")!.textContent).toContain("Refine")
    const actions = [...card.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
    expect(actions.map(button => button.textContent)).toEqual(["Refine", "Decline"])
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "Keep my current draft"
    actions[0]!.click()
    expect(message.value).toBe("Keep my current draft"); expect(message.placeholder).toBe("How should this change?")
    expect(f.get("chips").textContent).toContain("Refining amendment")
    expect(applied).toBe(false)
  } finally { f.restore() }
})

test("native paragraph and soft breaks display as separate review/chip lines while Apply receives the exact original Word text", async () => {
  const before = "Working notes\rFirst item\vMore detail.", after = "Working notes\rRevised item\vMore detail."
  const target: WordChatTarget = { ...passage, text: before, paragraphs: [{ id: "notes", text: "Working notes" }, { id: "item", text: "First item\vMore detail." }], endOffset: "First item\vMore detail.".length }
  const native = { ...proposal, before, after, target }
  const crlf = { ...proposal, id: "crlf-proposal", before: "First paragraph.\r\nSecond paragraph.", after: "Updated first paragraph.\r\nSecond paragraph." }
  const original = JSON.stringify([native, crlf]), applied: WordChatProposal[] = []
  const f = fixture({ turns: [turn({ kind: "amendment", proposal: native }), turn({ id: "crlf-turn", kind: "amendment", proposal: crlf })], host: { captureSelection: async () => target, apply: async value => { applied.push(value) } } })
  try {
    await f.chat.ready
    const cards = f.root.querySelectorAll(".chat-proposal")
    expect(cards[0]!.querySelector(".chat-before")!.textContent).toBe("Working notes\nFirst item\nMore detail.")
    expect(cards[0]!.querySelector(".chat-after")!.textContent).toBe("Working notes\nRevised item\nMore detail.")
    expect(cards[1]!.querySelector(".chat-before")!.textContent).toBe("First paragraph.\nSecond paragraph.")
    expect(cards[1]!.querySelector(".chat-after")!.textContent).toBe("Updated first paragraph.\nSecond paragraph.")
    expect(f.get("chips").querySelector("p")!.textContent).toBe("Working notes\nFirst item\nMore detail.")
    const actions = [...cards[0]!.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
    actions.find(button => button.textContent === "Refine")!.click()
    expect([...f.get("chips").querySelectorAll("p")].map(item => item.textContent)).toEqual(["Working notes\nFirst item\nMore detail.", "Working notes\nFirst item\nMore detail."])
    actions.find(button => button.textContent === "Apply")!.click(); await Bun.sleep(0)
    expect(applied).toHaveLength(1); expect(applied[0]!.before).toBe(before); expect(applied[0]!.after).toBe(after)
    expect(applied[0]!.target).toEqual(target)
    expect(JSON.stringify([native, crlf])).toBe(original)
  } finally { f.restore() }
})

test("a completed amendment refused by review shows the actionable reason beside its answer without offering Apply", async () => {
  const refused = turn({ kind: "amendment", status: "complete", answer: "I prepared a change to the payment wording.", error: "This change did not pass review. Rephrase the request or select a smaller passage.", proposal: undefined })
  const original = JSON.stringify(refused)
  let applied = false
  const f = fixture({ turns: [refused], host: { apply: async () => { applied = true } } })
  try {
    await f.chat.ready
    const answer = f.root.querySelector('.chat-message[data-role="assistant"]')!
    expect(answer.querySelector(".chat-answer-text")!.textContent).toBe(refused.answer)
    const notice = answer.querySelector('[role="alert"]')!
    expect(notice.textContent).toBe(refused.error!)
    expect(notice.textContent).toContain("Rephrase the request or select a smaller passage")
    expect(f.root.querySelector(".chat-proposal")).toBeNull()
    expect([...f.root.querySelectorAll("button")].some(button => button.textContent === "Apply")).toBe(false)
    expect(f.get("activity").textContent).toBe("")
    expect(f.root.textContent!.split(refused.error!)).toHaveLength(2)
    expect(applied).toBe(false); expect(JSON.stringify(refused)).toBe(original)
  } finally { f.restore() }
})

test("document info exposes the exact host filename on hover, focus and tap without changing scope or draft", async () => {
  const name = "Services · agreement <review>.docx"
  const f = fixture({ host: { documentName: name } })
  try {
    await f.chat.ready
    const info = f.get<HTMLButtonElement>("document-info"), tooltip = f.get("document-info-content")
    const message = f.get<HTMLTextAreaElement>("message"); message.value = "Keep this question"
    expect(f.root.querySelector('[data-chat="document"]')).toBeNull()
    expect(f.root.dataset.documentName).toBe(name); expect(tooltip.textContent).toBe(name); expect(tooltip.querySelector("review")).toBeNull()
    expect(tooltip.hidden).toBe(true)
    f.get("document-scope").dispatchEvent(new f.page.MouseEvent("mouseenter") as unknown as Event); expect(tooltip.hidden).toBe(false)
    f.get("document-scope").dispatchEvent(new f.page.MouseEvent("mouseleave") as unknown as Event); expect(tooltip.hidden).toBe(true)
    f.click("scope-matter"); info.focus(); expect(tooltip.hidden).toBe(false)
    info.dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as Event)
    expect(tooltip.hidden).toBe(true); expect(f.page.document.activeElement === info as never).toBe(true)
    f.click("document-info"); expect(tooltip.hidden).toBe(false); expect(info.getAttribute("aria-expanded")).toBe("true")
    f.click("document-info"); expect(tooltip.hidden).toBe(true)
    expect(f.get("scope-matter").getAttribute("aria-pressed")).toBe("true")
    expect(message.value).toBe("Keep this question"); expect(f.calls.some(call => call.action === "send" || call.action === "suggest")).toBe(false)
  } finally { f.restore() }
})

test("response mode picker supports keyboard selection, freezes lost requests and resets new chats to Auto", async () => {
  let sends = 0
  const f = fixture({ request: async (action, input) => {
    if (action !== "send") return
    if (++sends === 1) throw new Error("Connection lost")
    return { turn: turn({ text: (input as WordChatSend).text }) }
  } })
  try {
    await f.chat.ready
    const toggle = f.get<HTMLButtonElement>("mode-toggle")
    expect(f.get("mode-label").textContent).toBe("Auto")
    f.click("mode-toggle"); expect(f.get("mode-picker").hidden).toBe(false)
    f.get("mode-auto").dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "End", bubbles: true }) as unknown as Event)
    expect(f.page.document.activeElement === f.get("mode-redline") as never).toBe(true)
    f.click("mode-redline"); expect(f.get("mode-label").textContent).toBe("Redline"); expect(f.get("mode-picker").hidden).toBe(true)
    expect(f.page.document.activeElement === toggle as never).toBe(true)
    await f.send("Review this wording")
    expect(toggle.disabled).toBe(true); f.click("mode-qa")
    f.click("reconnect"); await Bun.sleep(0)
    const sent = f.calls.filter(call => call.action === "send").map(call => call.input as WordChatSend)
    expect(sent).toHaveLength(2); expect(sent[0]!.mode).toBe("redline"); expect(sent[1]).toEqual(sent[0])
    expect(f.get("mode-label").textContent).toBe("Redline")
    f.click("mode-toggle")
    f.get("mode-picker").dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as Event)
    expect(f.get("mode-picker").hidden).toBe(true); expect(f.page.document.activeElement === toggle as never).toBe(true)
    f.click("new"); await Bun.sleep(0); expect(f.get("mode-label").textContent).toBe("Auto")
    await f.send("Explain this wording")
    expect((f.calls.filter(call => call.action === "send").at(-1)!.input as WordChatSend).mode).toBe("auto")
  } finally { f.restore() }
})

test("formatting review shows one inspectable wording target and exact paragraph changes; Refine and Apply retain its original plan", async () => {
  const following = { id: "unchanged-following-paragraph", text: "This separate clause stays unchanged." }
  const target: WordChatTarget = { ...passage, text: `${passage.text}\r${following.text}`, paragraphs: [...passage.paragraphs, following], endOffset: following.text.length }
  const formatted: WordChatProposal = { ...proposal, before: target.text, after: target.text, target, formatting: [{ paragraph: 0, style: "Heading1", spacingBefore: 0, spacingAfter: 6, lineSpacing: 1.15 }] }
  const original = JSON.stringify(formatted)
  const applied: WordChatProposal[] = []
  const f = fixture({ turns: [turn({ kind: "formatting", proposal: formatted })], host: { capture: async () => ({ ...context, text: target.text, paragraphs: target.paragraphs }), apply: async value => { applied.push(value) } },
    request: async action => action === "send" ? { turn: turn({ kind: "formatting", proposal: formatted }) } : undefined })
  try {
    await f.chat.ready
    const card = () => f.root.querySelector(".chat-proposal")!
    expect(card().querySelector("h3")!.textContent).toBe("Review formatting")
    expect(card().textContent).toContain("Wording unchanged"); expect(card().querySelector(".chat-before")).toBeNull(); expect(card().querySelector(".chat-after")).toBeNull()
    expect(card().querySelectorAll("pre")).toHaveLength(1)
    expect(card().querySelector(".chat-format-target pre")!.textContent).toBe(target.text.replace(/\r/g, "\n"))
    expect(card().querySelector(".chat-format-plan")!.textContent).toContain("Heading 1 · 0 pt before · 6 pt after · 1.15× line spacing")
    const action = (label: string) => [...card().querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === label)!
    action("Refine").click(); f.click("mode-toggle"); f.click("mode-qa"); await f.send("Use more space after the heading")
    const request = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(request.kind).toBe("formatting"); expect(request.mode).toBe("qa"); expect(request.proposal).toEqual(formatted)
    action("Apply").click(); await Bun.sleep(0)
    expect(applied).toEqual([formatted]); expect(JSON.stringify(formatted)).toBe(original)
  } finally { f.restore() }
})

test("formatting review groups adjacent identical operations while preserving unique headings and every original Apply operation", async () => {
  const paragraphs = Array.from({ length: 10 }, (_, index) => ({ id: `synthetic-paragraph-${index}`, text: index === 0 ? "Services agreement" : index === 9 ? "This following clause stays unchanged." : `Synthetic paragraph ${index + 1} wording.` }))
  const text = paragraphs.map(paragraph => paragraph.text).join("\r")
  const target: WordChatTarget = { ...passage, text, paragraphs, endOffset: paragraphs.at(-1)!.text.length }
  const formatting: NonNullable<WordChatProposal["formatting"]> = [
    { paragraph: 0, style: "Title", spacingAfter: 12 },
    ...[2, 1, 3, 4, 5, 6, 7].map(paragraph => ({ paragraph, spacingAfter: 6, lineSpacing: 1.15 })),
    { paragraph: 8, spacingAfter: 12, lineSpacing: 1.15 },
  ]
  const formatted: WordChatProposal = { ...proposal, before: text, after: text, target, formatting }
  const original = JSON.stringify(formatted), applied: WordChatProposal[] = []
  const f = fixture({ turns: [turn({ kind: "formatting", proposal: formatted })], host: { apply: async value => { applied.push(value) } } })
  try {
    await f.chat.ready
    const card = f.root.querySelector(".chat-proposal")!, entries = card.querySelectorAll(".chat-format-plan li")
    expect(entries).toHaveLength(3)
    expect(entries[0]!.querySelector("strong")!.textContent).toBe("Services agreement"); expect(entries[0]!.textContent).toContain("Title · 12 pt after")
    expect(entries[1]!.querySelector("strong")!.textContent).toBe("Paragraphs 2–8"); expect(entries[1]!.textContent).toContain("6 pt after · 1.15× line spacing")
    expect(entries[1]!.textContent).not.toContain(paragraphs[1]!.text)
    expect(entries[2]!.textContent).toContain(paragraphs[8]!.text); expect(entries[2]!.textContent).toContain("12 pt after")
    expect(card.querySelector(".chat-format-target pre")!.textContent).toBe(text.replace(/\r/g, "\n"))
    const apply = [...card.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Apply")!
    apply.click(); await Bun.sleep(0)
    expect(applied).toEqual([formatted]); expect(applied[0]!.formatting).toHaveLength(9); expect(JSON.stringify(formatted)).toBe(original)
  } finally { f.restore() }
})

test("inline formatting review names the exact span and properties, applies once, and removes its actions", async () => {
  const text = passage.paragraphs[0]!.text, selected = "within 14", start = text.indexOf(selected)
  const target: WordChatTarget = { ...passage, text: selected, startOffset: start, endOffset: start + selected.length }
  const formatted: WordChatProposal = { ...proposal, before: selected, after: selected, target, inlineFormatting: { bold: true, italic: false, underline: true } }
  const applied: WordChatProposal[] = []
  const f = fixture({ turns: [turn({ kind: "formatting", proposal: formatted })], host: { apply: async value => { applied.push(value) } } })
  try {
    await f.chat.ready
    const card = f.root.querySelector('[aria-label="Formatting to review"]')!
    expect(card.querySelector("h3")!.textContent).toBe("Review formatting")
    expect(card.textContent).toContain("Wording unchanged")
    expect(card.querySelector(".chat-format-inline-target")!.textContent).toBe(selected)
    expect([...card.querySelectorAll(".chat-format-plan li")].map(item => item.textContent)).toEqual(["Bold on", "Italic off", "Underline on"])
    const apply = [...card.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Apply")!
    apply.click(); apply.click(); await Bun.sleep(0)
    expect(applied).toEqual([formatted])
    expect(f.root.querySelector(".chat-proposal-actions")).toBeNull()
  } finally { f.restore() }
})

test("refining inline formatting retains its exact selection, formatting kind, and requested properties", async () => {
  const text = passage.paragraphs[0]!.text, selected = "within 14", start = text.indexOf(selected)
  const target: WordChatTarget = { ...passage, text: selected, startOffset: start, endOffset: start + selected.length }
  const formatted: WordChatProposal = { ...proposal, before: selected, after: selected, target, inlineFormatting: { bold: false, underline: true } }
  const f = fixture({ turns: [turn({ kind: "formatting", proposal: formatted })] })
  try {
    await f.chat.ready
    const card = f.root.querySelector(".chat-proposal")!
    ;[...card.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Refine")!.click()
    await f.send("Keep the same passage and add italics.")
    const sent = f.calls.find(call => call.action === "send")!.input as WordChatSend
    expect(sent.kind).toBe("formatting")
    expect(sent.proposal).toEqual(formatted)
    expect(sent.proposal?.inlineFormatting).toEqual({ bold: false, underline: true })
    expect(sent.proposal?.target).toEqual(target)
  } finally { f.restore() }
})

test("partial wording review shows the assembled paragraph without expanding the applied span", async () => {
  const text = "Send a notice .", selected = "notice", start = text.indexOf(selected)
  const target: WordChatTarget = { ...passage, text: selected, paragraphs: [{ id: "boundary", text }], startOffset: start, endOffset: start + selected.length }
  const change: WordChatProposal = { ...proposal, before: selected, after: "written notice", target }, applied: WordChatProposal[] = []
  const f = fixture({ turns: [turn({ kind: "amendment", proposal: change })], host: { apply: async value => { applied.push(value) } } })
  try {
    await f.chat.ready
    const card = f.root.querySelector(".chat-proposal")!
    const detail = [...card.querySelectorAll("details")].find(item => item.querySelector("summary")?.textContent === "In context")!
    expect(detail.open).toBe(false)
    expect(detail.querySelector("pre")!.textContent).toBe("Send a written notice .")
    ;[...card.querySelectorAll<HTMLButtonElement>("button")].find(button => button.textContent === "Apply")!.click()
    await Bun.sleep(0)
    expect(applied).toEqual([change])
    expect(applied[0]!.target).toEqual(target)
  } finally { f.restore() }
})

test("context preview is omitted for complete paragraphs, formatting-only and invalid targets", async () => {
  for (const change of [proposal, { ...proposal, after: proposal.before, formatting: [{ paragraph: 0, spacingAfter: 6 }] }, { ...proposal, target: { ...passage, startOffset: 1 } }]) {
    const f = fixture({ turns: [turn({ proposal: change })] })
    try {
      await f.chat.ready
      expect([...f.root.querySelectorAll("summary")].some(item => item.textContent === "In context")).toBe(false)
    } finally { f.restore() }
  }
})

test("mixed review shows the replacement and collapsed layout, and applies the exact combined proposal once", async () => {
  const following = { id: "unchanged-tail", text: "This separate clause stays unchanged." }
  const target: WordChatTarget = { ...passage, text: `${passage.text}\r${following.text}`, paragraphs: [...passage.paragraphs, following], endOffset: following.text.length }
  const mixed: WordChatProposal = { ...proposal, before: target.text, after: `Payment\r${passage.text}\r${following.text}`, target, formatting: [{ paragraph: 0, style: "Heading1", spacingAfter: 6 }] }
  const original = JSON.stringify(mixed), applied: WordChatProposal[] = []
  const f = fixture({ turns: [turn({ kind: "amendment", proposal: mixed })], host: { apply: async value => { applied.push(value) } } })
  try {
    await f.chat.ready
    const card = f.root.querySelector(".chat-proposal")!
    expect(card.getAttribute("aria-label")).toBe("Amendment to review")
    expect(card.querySelector("h3")!.textContent).toBe("Review amendment")
    expect(card.textContent).not.toContain("Wording unchanged")
    expect(card.querySelector(".chat-after")!.textContent).toBe(mixed.after.replace(/\r/g, "\n"))
    const details = [...card.querySelectorAll("details")]
    expect(details.map(item => item.querySelector("summary")!.textContent)).toEqual(["Before", "Layout changes"])
    expect(details.every(item => !item.open)).toBe(true)
    expect(card.querySelector(".chat-format-plan strong")!.textContent).toBe("Payment")
    const apply = [...card.querySelectorAll<HTMLButtonElement>("button")].find(button => button.textContent === "Apply")!
    apply.click(); apply.click(); await Bun.sleep(0)
    expect(applied).toEqual([mixed]); expect(JSON.stringify(mixed)).toBe(original)
    expect(Boolean(f.root.querySelector(".chat-proposal-actions"))).toBe(false)
    expect(f.root.textContent).toContain("Applied in Word")
  } finally { f.restore() }
})

test("Apply uncertainty hides every action and affected context, while pre-import errors with the same text remain retryable", async () => {
  for (const mutationState of ["unchanged", "uncertain"] as const) {
    let calls = 0
    const error = Object.assign(new Error("Word could not confirm the imported change. Inspect it and use Undo if needed."), { mutationState })
    const f = fixture({ turns: [turn({ kind: "amendment", proposal })], host: { apply: async () => { calls++; throw error } } })
    try {
      await f.chat.ready
      const actions = () => [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
      const cached = actions(), apply = cached.find(button => button.textContent === "Apply")!
      cached.find(button => button.textContent === "Refine")!.click()
      const message = f.get<HTMLTextAreaElement>("message"); message.value = "Keep my next message"
      apply.click(); await Bun.sleep(0)
      const notice = f.root.querySelector('.chat-proposal-error[role="alert"]')!
      expect(message.value).toBe("Keep my next message"); expect(f.get("activity").textContent).toBe("")
      if (mutationState === "uncertain") {
        expect(notice.textContent).toBe("Check the change in Word. Use Undo if needed.")
        expect(actions()).toHaveLength(0)
        expect(f.get("chips").textContent).not.toContain("Refining amendment"); expect(f.get("chips").textContent).not.toContain("Selected Word text")
        expect(message.placeholder).toBe("Ask doc.haus…")
        for (const button of cached) { button.disabled = false; button.click() }
        await Bun.sleep(0)
        expect(calls).toBe(1); expect(actions()).toHaveLength(0); expect(f.root.textContent).not.toContain("Declined. Your document is unchanged.")
      } else {
        expect(notice.textContent).toBe(error.message); expect(actions().map(button => button.textContent)).toEqual(["Refine", "Apply", "Decline"])
        expect(actions().every(button => !button.disabled)).toBe(true); expect(f.get("chips").textContent).toContain("Refining amendment")
        actions().find(button => button.textContent === "Apply")!.click(); await Bun.sleep(0); expect(calls).toBe(2)
      }
    } finally { f.restore() }
  }
})

test("formatting the final selected paragraph can be applied without asking for another paragraph", async () => {
  const formatted: WordChatProposal = { ...proposal, before: passage.text, after: passage.text, formatting: [{ paragraph: 0, spacingAfter: 6 }] }
  const original = JSON.stringify(formatted)
  let writes = 0
  const f = fixture({ turns: [turn({ kind: "formatting", proposal: formatted })], host: {
    applyEligibility: p => wordChatApplyEligibility(p.target, p.after, p.formatting), apply: async () => { writes++ },
  } })
  try {
    await f.chat.ready
    const card = f.root.querySelector(".chat-proposal")!
    expect(card.querySelector("h3")!.textContent).toBe("Review formatting")
    expect(card.querySelector(".chat-format-plan")!.textContent).toContain("6 pt after")
    expect(Boolean(card.querySelector(".chat-proposal-limit"))).toBe(false)
    const buttons = [...card.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
    expect(buttons.map(button => button.textContent)).toEqual(["Refine", "Apply", "Decline"])
    expect(card.querySelector(".chat-format-target pre")!.textContent).toBe(passage.text)
    buttons.find(button => button.textContent === "Apply")!.click(); await Bun.sleep(0)
    expect(writes).toBe(1); expect(JSON.stringify(formatted)).toBe(original)
  } finally { f.restore() }
})

test("unavailable Word transports retain review and refinement without offering an impossible Apply", async () => {
  const following = { id: "tail", text: "A separate clause remains." }
  const target: WordChatTarget = { ...passage, text: `${passage.text}\r${following.text}`, paragraphs: [...passage.paragraphs, following], endOffset: following.text.length }
  for (const change of [
    { ...proposal, before: target.text, after: `A new paragraph.\r${target.text}`, target },
    { ...proposal, before: target.text, after: `A new paragraph.\r${target.text}`, target, formatting: [{ paragraph: 0, spacingAfter: 6 }] },
    { ...proposal, before: target.text, after: target.text, target, formatting: [{ paragraph: 0, spacingAfter: 6 }] },
  ]) {
    let writes = 0
    const f = fixture({ turns: [turn({ proposal: change })], host: {
      applyEligibility: p => wordChatApplyEligibility(p.target, p.after, p.formatting), apply: async () => { writes++ },
    } })
    try {
      await f.chat.ready
      const card = f.root.querySelector(".chat-proposal")!
      expect(card.querySelector(".chat-proposal-limit")!.textContent).toContain("directly in Word")
      expect(card.textContent).toContain(change.explanation)
      const actions = [...card.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")]
      expect(actions.map(button => button.textContent)).toEqual(["Refine", "Decline"])
      actions[0]!.click()
      expect(f.get("chips").textContent).toContain("Refining amendment")
      const draft = f.get<HTMLTextAreaElement>("message"); draft.value = "Keep my next instruction"
      actions[1]!.click()
      expect(draft.value).toBe("Keep my next instruction")
      expect(card.querySelector(".chat-after")?.textContent ?? card.querySelector(".chat-format-target pre")!.textContent).toBe(change.after.replace(/\r/g, "\n"))
      expect(writes).toBe(0)
    } finally { f.restore() }
  }
})

test("a cached Apply callback rechecks transport availability before calling the host", async () => {
  let available = true, writes = 0
  const f = fixture({ turns: [turn({ proposal })], host: {
    applyEligibility: () => available ? { supported: true } : { supported: false, reason: "Word can no longer apply this change." },
    apply: async () => { writes++ },
  } })
  try {
    await f.chat.ready
    const cached = [...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Apply")!
    available = false; cached.click(); await Bun.sleep(0)
    expect(writes).toBe(0)
    expect([...f.root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].map(button => button.textContent)).toEqual(["Refine", "Decline"])
    expect(f.get("activity").textContent).toBe("")
  } finally { f.restore() }
})


test("a reconstructed pane shows a confirmed proposal as applied", async () => {
  let writes = 0
  const f = fixture({ turns: [turn({ kind: "amendment", proposal, proposalApplied: true })], host: { apply: async () => { writes++ } } })
  try {
    await f.chat.ready
    expect(f.root.textContent).toContain("Previously applied in Word")
    expect(f.root.querySelectorAll(".chat-proposal-actions button")).toHaveLength(0)
    expect(writes).toBe(0)
  } finally { f.restore() }
})

test("eligible ordinary wording offers Apply in the candidate pane", async () => {
  let writes = 0
  const f = fixture({ turns: [turn({ proposal })], host: {
    applyEligibility: p => wordChatApplyEligibility(p.target, p.after, p.formatting), apply: async () => { writes++ },
  } })
  try {
    await f.chat.ready
    const card = f.root.querySelector(".chat-proposal")!
    expect(card.querySelector(".chat-proposal-limit")).toBeNull()
    const apply = [...card.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")].find(button => button.textContent === "Apply")!
    expect(apply.disabled).toBe(false)
    apply.click(); await Promise.resolve(); await Promise.resolve()
    expect(writes).toBe(1)
  } finally { f.restore() }
})

test("wording eligibility refuses structural, empty and mixed formatting requests", () => {
  expect(wordChatApplyEligibility(passage, "")).toMatchObject({ supported: false })
  expect(wordChatApplyEligibility(passage, "Changed\rparagraph")).toMatchObject({ supported: false })
  expect(wordChatApplyEligibility(passage, "Changed wording", [{ paragraph: 0, spacingAfter: 6 }])).toMatchObject({ supported: false })
  const tabTarget = { ...passage, text: "One\tcell", paragraphs: [{ ...passage.paragraphs[0]!, text: "One\tcell" }], startOffset: 0, endOffset: 8 }
  expect(wordChatApplyEligibility(tabTarget, "Another")).toMatchObject({ supported: false })
})

test("segmented answers preserve literal brackets and render only verified references", async () => {
  const source = { id:"source-one", name:"Notes", quote:"Verified passage", location:"Page 1", provenance:"saved-version" as const }
  const other = { id:"source-two", name:"Other", quote:"Another passage", location:"Page 2", provenance:"saved-version" as const }
  const original = turn({ answer:"untrusted fallback", sources:[source,other], answerFormat:"segments-v1", answerParts:[
    {type:"text",text:'Literal [source-one], [word], [1, 2] and <img src=x> & text. '},
    {type:"reference",sourceId:"source-one",quote:"Verified passage"}, {type:"text",text:" again "},
    {type:"reference",sourceId:"source-one",quote:"Verified passage"}, {type:"text",text:" and "},
    {type:"reference",sourceId:"source-two",quote:"Another passage"}, {type:"text",text:" unknown "},
    {type:"reference",sourceId:"unknown",quote:"forged"}, {type:"reference",sourceId:"source-one",quote:"wrong quote"},
  ] } as unknown as Partial<WordChatTurn>)
  const f = fixture({turns:[original]})
  try {
    await f.chat.ready
    const answer = f.root.querySelector(".chat-answer-text")!
    expect(answer.textContent).toBe('Literal [source-one], [word], [1, 2] and <img src=x> & text. [1] again [1] and [2] unknown [Source unavailable][Source unavailable]')
    expect(answer.querySelector("img")).toBeNull()
    expect([...f.root.querySelectorAll(".chat-source summary")].map(e=>e.textContent)).toEqual(["[1] Notes · Saved source", "[2] Other · Saved source"])
  } finally { f.restore() }
})

test("malformed marked answers never fall back to raw answer or source links", async () => {
  const source = { id:"source-one", name:"Notes", quote:"Verified", location:"Page 1", provenance:"saved-version" as const }
  for (const parts of [undefined, [{type:"reference",sourceId:"source-one"}], [{type:"text",text:"hidden"},{type:"other"}]]) {
    const f = fixture({turns:[turn({answer:"unsafe [source-one]",sources:[source],answerFormat:"segments-v1",answerParts:parts} as unknown as Partial<WordChatTurn>)]})
    try {
      await f.chat.ready
      const answer = f.root.querySelector('.chat-message[data-role="assistant"]')!
      expect(answer.textContent).not.toContain("unsafe")
      expect(answer.textContent).not.toContain("hidden")
      expect(answer.querySelector(".chat-source")).toBeNull()
      expect(answer.querySelector('[role="alert"]')?.textContent).toContain("could not be displayed")
    } finally { f.restore() }
  }
})

test("malformed segmented review cannot offer paragraph or row Apply", async () => {
  for (const review of [{proposal}, {rowProposal}]) {
    const f = fixture({turns:[turn({kind:"amendment", answerFormat:"segments-v1", answerParts:undefined, ...review})], host:{applyRow:async () => ({status:"applied",tracked:true,saved:false,authors:[]})}})
    try {
      await f.chat.ready
      const answer = f.root.querySelector('.chat-message[data-role="assistant"]')!
      expect(answer.querySelector('[role="alert"]')).not.toBeNull()
      expect(answer.querySelector(".chat-proposal")).toBeNull()
      expect(answer.textContent).not.toContain("Apply")
    } finally { f.restore() }
  }
})

test("cleared segmented answers show only their running, stopped or error state", async () => {
  const source = { id:"source-one", name:"Notes", quote:"Stale quote", location:"Page 1", provenance:"saved-version" as const }
  for (const [status, error, expected] of [["running", undefined, "Thinking…"], ["cancelled", undefined, "Stopped"], ["error", "Request failed", "Request failed"], ["complete", "Review failed", "Review failed"]] as const) {
    const f = fixture({turns:[turn({status,error,answer:"stale answer",answerFormat:"segments-v1",answerParts:undefined,sources:[source],proposal})]})
    try {
      await f.chat.ready
      const answer = f.root.querySelector('.chat-message[data-role="assistant"]')!
      expect(answer.textContent).toContain(expected)
      expect(answer.textContent).not.toContain("stale answer")
      expect(answer.textContent).not.toContain("could not be displayed")
      expect(answer.querySelector(".chat-source")).toBeNull()
      expect(answer.querySelector(".chat-proposal")).toBeNull()
    } finally { f.restore() }
  }
})

test("paragraph deletion review shows the text and mark decision then consumes its controls", async () => {
  let calls=0
  const deletion:WordChatProposal={...proposal,after:"",paragraphAction:"delete"}
  const f=fixture({turns:[turn({kind:"amendment",proposal:deletion})],host:{
    apply:async p=>{expect(p.paragraphAction).toBe("delete");calls++},
    applyEligibility:p=>wordChatApplyEligibility(p.target,p.after,p.formatting,true,p.inlineFormatting,p.paragraphAction),
  }})
  try {
    await f.chat.ready
    const card=f.root.querySelector('[aria-label="Paragraph deletion to review"]')!
    expect(card.textContent).toContain(proposal.before)
    expect(card.textContent).toContain("including its paragraph mark")
    expect(card.querySelector(".chat-after")).toBeNull()
    const apply=[...card.querySelectorAll<HTMLButtonElement>("button")].find(b=>b.textContent==="Apply")!
    expect(apply).toBeDefined()
    apply.click();await Bun.sleep(0)
    expect(calls).toBe(1)
    expect(card.isConnected).toBe(false)
    expect(f.root.querySelector(".chat-proposal-actions")).toBeNull()
    expect(f.root.textContent).toContain("Applied in Word")
    apply.click();expect(calls).toBe(1)
  }finally{f.restore()}
})

test("paragraph split review shows the new boundary and consumes Apply once", async () => {
  const text="The supplier shall deliver the services. The customer shall pay the fees."
  const target:WordChatTarget={...passage,text,paragraphs:[{id:"split-source",text}],endOffset:text.length}
  const split:WordChatProposal={...proposal,target,before:text,after:"The supplier shall deliver the services.\rThe customer shall pay the fees.",paragraphAction:"split",explanation:"Separate the obligations without changing any words.",openItems:[]}
  let calls=0
  const f=fixture({turns:[turn({kind:"amendment",proposal:split})],host:{
    apply:async p=>{expect(p).toEqual(split);calls++},
    applyEligibility:p=>wordChatApplyEligibility(p.target,p.after,p.formatting,true,p.inlineFormatting,p.paragraphAction),
  }})
  try {
    await f.chat.ready
    const card=f.root.querySelector('[aria-label="Paragraph split to review"]')!
    expect(card.querySelector("h3")?.textContent).toBe("Split paragraph")
    expect(card.querySelector(".chat-before")?.textContent).toBe(text)
    expect(card.querySelector(".chat-after")?.textContent).toBe(split.after.replace(/\r/g,"\n"))
    const apply=[...card.querySelectorAll<HTMLButtonElement>("button")].find(b=>b.textContent==="Apply")!
    expect(apply.disabled).toBe(false)
    apply.click();await Bun.sleep(0)
    expect(calls).toBe(1)
    expect(f.root.querySelector(".chat-proposal-actions")).toBeNull()
    expect(f.root.textContent).toContain("Applied in Word")
    apply.click();expect(calls).toBe(1)
  }finally{f.restore()}
})

test("paragraph join review offers the native boundary change and consumes Apply once", async () => {
  const paragraphs=[{id:"join-left",text:"The supplier shall deliver the services."},{id:"join-right",text:"The customer shall pay the fees."}]
  const target:WordChatTarget={...passage,paragraphs,text:paragraphs.map(p=>p.text).join("\r"),endOffset:paragraphs[1]!.text.length}
  const merge:WordChatProposal={...proposal,target,before:target.text,after:paragraphs.map(p=>p.text).join(" "),paragraphAction:"merge",explanation:"Join the two paragraphs without changing their words.",openItems:[]}
  let calls=0
  const f=fixture({turns:[turn({kind:"amendment",proposal:merge})],host:{
    apply:async p=>{expect(p).toEqual(merge);calls++},
    applyEligibility:p=>wordChatApplyEligibility(p.target,p.after,p.formatting,true,p.inlineFormatting,p.paragraphAction),
  }})
  try {
    await f.chat.ready
    const card=f.root.querySelector('[aria-label="Paragraph join to review"]')!
    expect(card.querySelector("h3")?.textContent).toBe("Join paragraphs")
    expect(card.querySelector(".chat-before")?.textContent).toBe(target.text.replace(/\r/g,"\n"))
    expect(card.querySelector(".chat-after")?.textContent).toBe(merge.after)
    const apply=[...card.querySelectorAll<HTMLButtonElement>("button")].find(b=>b.textContent==="Apply")!
    expect(apply.disabled).toBe(false)
    apply.click();await Bun.sleep(0)
    expect(calls).toBe(1)
    expect(f.root.querySelector(".chat-proposal-actions")).toBeNull()
    expect(f.root.textContent).toContain("Applied in Word")
    apply.click();expect(calls).toBe(1)
  }finally{f.restore()}
})

test("automatic list review displays the next label and consumes the single approved append", async () => {
  let calls=0
  const append:WordChatProposal={...proposal,after:proposal.before+"\rA further obligation.",paragraphAction:"append_list_item",
    list:{id:17,format:"decimal",members:[{id:passage.paragraphs[0]!.id,text:passage.text,level:0,label:"1."}]}}
  const f=fixture({turns:[turn({kind:"amendment",proposal:append})],host:{
    apply:async p=>{expect(p.list).toEqual(append.list);calls++},
    applyEligibility:p=>wordChatApplyEligibility(p.target,p.after,p.formatting,true,p.inlineFormatting,p.paragraphAction,p.list),
  }})
  try {
    await f.chat.ready
    const card=f.root.querySelector('[aria-label="List item to review"]')!
    expect(card.textContent).toContain("Add list item")
    expect(card.querySelector(".chat-after")?.textContent).toBe("2. A further obligation.")
    expect(card.textContent).toContain("automatic list")
    const apply=[...card.querySelectorAll<HTMLButtonElement>("button")].find(b=>b.textContent==="Apply")!
    apply.click();await Bun.sleep(0)
    expect(calls).toBe(1)
    expect(f.root.querySelector(".chat-proposal-actions")).toBeNull()
    expect(f.root.textContent).toContain("Applied in Word")
    apply.click();expect(calls).toBe(1)
  }finally{f.restore()}
})
