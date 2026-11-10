import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { mountWordConsequences } from "./chat-consequences"
import type { ChatAction, WordChatApi } from "./chat-client"
import type { WordConsequenceAssessment, WordConsequenceFinding } from "../../../dochaus/lib/word-consequence-contract"
import type { WordChatSnapshot } from "../../../dochaus/lib/word-chat-amendment-contract"

const finding: WordConsequenceFinding = { id: "f-one", title: "Timing may differ", explanation: "The related deadline still uses the earlier period.", kind: "effect", basis: "inference", passage: { paragraphId: "p-one", quote: "Payment in 30 days." } }
const review = (input: Partial<WordConsequenceAssessment> = {}): WordConsequenceAssessment => ({ id: "assessment-one", conversationId: "conversation-one", requestId: "request-one", turnId: "amendment-one", status: "running", findings: [finding], coverage: { reviewed: 1, total: 2, candidates: 2, checked: 1 }, gaps: [], createdAt: 1, ...input })
function fixture(request?: (action: ChatAction, input: Record<string, unknown>) => unknown | Promise<unknown>) {
  const page = new Window(), root = page.document.createElement("section"), composer = page.document.createElement("textarea")
  const scroll = page.document.createElement("div"); scroll.scrollTop = 123; scroll.append(root, composer); page.document.body.append(scroll)
  const calls: { action: ChatAction; input: Record<string, unknown> }[] = [], scheduled: { run(): void; at: number; cancelled: boolean }[] = []
  let time = 0, captures = 0, navigations = 0, failCapture = false
  let context: WordChatSnapshot = { text: "Payment in 30 days.", paragraphs: [{ id: "p-one", text: "Payment in 30 days." }], capturedAt: "2026-11-09T12:00:00Z", saved: false, truncated: false }
  const state = { conversationId: "conversation-one", turnId: "amendment-one", revision: "one", enabled: true, busy: false, composing: false }
  const api: WordChatApi = { async request<T>(action: ChatAction, input: object = {}) { calls.push({ action, input: input as Record<string, unknown> }); return (await request?.(action, input as Record<string, unknown>) ?? (action === "assessment" && !("assessmentId" in input) ? { assessment: null } : review())) as T }, async upload<T>(): Promise<T> { throw Error("unused") } }
  const ui = mountWordConsequences(root as unknown as HTMLElement, api, { capture: async () => { captures++; if (failCapture) { failCapture = false; throw Error("Word capture failed") }; return structuredClone(context) }, navigate: async () => { navigations++; return true }, now: () => time, schedule(run, delay) { const item = { run, at: time + delay, cancelled: false }; scheduled.push(item); return () => { item.cancelled = true } } }, () => state)
  return { root, composer, scroll, state, ui, calls, captures: () => captures, failNextCapture() { failCapture = true }, navigations: () => navigations, setContext(value: string) { context = { ...context, text: value, paragraphs: [{ id: "p-one", text: value }] } }, async advance(ms = 3_000) { time += ms; for (const job of scheduled.filter(item => !item.cancelled && item.at <= time)) { job.cancelled = true; job.run() }; await Bun.sleep(0) }, dispose() { ui.dispose(); page.happyDOM.abort() } }
}

test("progressive review keeps typing and scroll intact, deduplicates findings, and consolidates only after idle", async () => {
  let checked = 0
  const f = fixture((action, input) => action === "assessment" && input.assessmentId ? review({ status: ++checked > 1 ? "complete" : "running", findings: [{ ...finding, id: `regenerated-${checked}` }], coverage: { reviewed: 2, total: 2, candidates: 2, checked: 2 } }) : undefined)
  try {
    f.composer.value = "Keep this draft"; f.composer.focus(); f.state.composing = true
    await f.advance(); await f.advance()
    expect(f.root.textContent).toContain(finding.title)
    f.ui.interaction(); await f.advance(1_500); f.ui.interaction(); await f.advance(1_500)
    await f.advance(1_500); f.ui.interaction(); await f.advance(1_500)
    expect([...f.root.querySelectorAll(".chat-proposal h3")].filter(node => node.textContent === finding.title)).toHaveLength(1)
    expect(f.root.textContent).not.toContain("Current assessment")
    expect(f.composer.value).toBe("Keep this draft"); expect(f.composer.ownerDocument.activeElement).toBe(f.composer); expect(f.scroll.scrollTop).toBe(123)
    await f.advance(3_000)
    expect(f.composer.ownerDocument.activeElement).toBe(f.composer)
    expect(f.root.textContent).toContain("Current assessment")
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(1)
  } finally { f.dispose() }
})

test("changed wording cancels obsolete work and suppresses old navigation", async () => {
  const f = fixture()
  try {
    await f.advance(); await f.advance()
    const show = [...f.root.querySelectorAll("button")].find(node => node.textContent === "Show in Word")!
    f.setContext("Payment in 45 days."); show.click(); await Bun.sleep(0)
    expect(f.navigations()).toBe(0); expect(f.root.textContent).toContain("This wording changed")
    show.blur(); await f.advance()
    expect(f.calls.some(call => call.action === "assessment-cancel")).toBe(true)
    expect(f.root.textContent).toContain("Superseded")
    await f.advance()
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(2)
  } finally { f.dispose() }
})

test("disable cancels late admission, retains existing history, and starts no further work", async () => {
  let finish!: (value: unknown) => void
  const f = fixture(action => action === "assess" ? new Promise(resolve => { finish = resolve }) : undefined)
  try {
    await f.advance(); await f.advance()
    f.state.enabled = false; f.ui.refresh(); finish(review()); await Bun.sleep(0)
    expect(f.calls.some(call => call.action === "assessment-cancel")).toBe(true)
    const count = f.captures(); await f.advance(10_000)
    expect(f.captures()).toBe(count); expect(f.calls.filter(call => call.action === "assess")).toHaveLength(1)
  } finally { f.dispose() }
})

test("disabled reload restores previous findings without document capture or model admission", async () => {
  const f = fixture(action => action === "assessment" ? review({ status: "complete" }) : undefined)
  try {
    f.state.enabled = false; await f.advance()
    expect(f.root.textContent).toContain(finding.title); expect(f.root.textContent).toContain("Consequence review is off")
    expect(f.captures()).toBe(0); expect(f.calls.some(call => call.action === "assess")).toBe(false)
  } finally { f.dispose() }
})

test("reconnect checks completed persisted assessments against fresh wording before consolidation", async () => {
  const f = fixture((action, input) => action === "assessment" ? review({ status: input.assessmentId ? "stale" : "complete" }) : undefined)
  try {
    f.setContext("Payment in 45 days."); await f.advance(); await f.advance()
    expect(f.root.textContent).not.toContain("Current assessment")
    expect(f.calls.some(call => call.action === "assessment" && call.input.context)).toBe(true)
    await f.advance(); await f.advance()
    expect(f.calls.some(call => call.action === "assess")).toBe(true)
  } finally { f.dispose() }
})

test("new material chat instructions recheck unchanged Word wording and retain superseded history", async () => {
  const f = fixture()
  try {
    await f.advance(); await f.advance()
    f.state.revision = "changed instructions"; f.ui.refresh()
    expect(f.calls.some(call => call.action === "assessment-cancel")).toBe(true)
    expect(f.root.textContent).toContain("Superseded")
    await f.advance(); await f.advance()
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(2)
  } finally { f.dispose() }
})

test("reload restores archived findings and meaning, with prior checks kept distinct", async () => {
  const old = { ...finding, id: "old", title: "Earlier deadline consequence" }
  const f = fixture(action => action === "assessment" ? { ...review({ status: "complete" }), meaning: "The proposed wording gives more time to pay.", history: [review({ id: "previous", status: "stale", findings: [old] })] } : undefined)
  try {
    await f.advance(); await f.advance()
    expect(f.root.textContent).toContain("What changes")
    expect(f.root.textContent).toContain("The proposed wording gives more time to pay.")
    expect(f.root.textContent).toContain("Earlier deadline consequence")
    expect(f.root.textContent).toContain("Superseded by a later review")
  } finally { f.dispose() }
})

test("a conversation change discards a late assessment and cancels its admitted job", async () => {
  let finish!: (value: unknown) => void
  const f = fixture(action => action === "assess" ? new Promise(resolve => { finish = resolve }) : undefined)
  try {
    await f.advance(); await f.advance()
    f.state.conversationId = "conversation-two"; f.ui.reset()
    finish(review()); await Bun.sleep(0)
    expect(f.root.textContent).not.toContain(finding.title)
    expect(f.calls.some(call => call.action === "assessment-cancel" && call.input.conversationId === "conversation-one")).toBe(true)
  } finally { f.dispose() }
})

test("a chat without an amendment stops review polling and wakes when an amendment arrives", async () => {
  const f = fixture()
  try {
    f.state.turnId = ""; await f.advance()
    const requests = f.calls.length
    await f.advance(30_000)
    expect(f.captures()).toBe(0); expect(f.calls).toHaveLength(requests)
    f.state.turnId = "amendment-one"; f.ui.refresh()
    await f.advance(); await f.advance()
    expect(f.calls.some(call => call.action === "assess")).toBe(true)
  } finally { f.dispose() }
})

test("slow admission cannot publish findings or meaning after Word changes during the response", async () => {
  let finish!: (value: unknown) => void
  const f = fixture(action => action === "assess" ? new Promise(resolve => { finish = resolve }) : undefined)
  try {
    await f.advance(); await f.advance()
    f.setContext("Payment in 45 days.")
    finish({ ...review(), meaning: "Obsolete proposed meaning" }); await Bun.sleep(0)
    expect(f.root.textContent).not.toContain(finding.title)
    expect(f.root.textContent).not.toContain("Obsolete proposed meaning")
    expect(f.calls.some(call => call.action === "assessment-cancel")).toBe(true)
    expect(f.root.textContent).not.toContain("Current assessment")
  } finally { f.dispose() }
})

test("slow polling revalidates Word before replacing existing findings with current results", async () => {
  let finish!: (value: unknown) => void
  const f = fixture((action, input) => action === "assessment" && input.assessmentId ? new Promise(resolve => { finish = resolve }) : undefined)
  try {
    await f.advance(); await f.advance(); await f.advance()
    f.setContext("Payment in 45 days.")
    finish(review({ status: "complete", findings: [{ ...finding, title: "Obsolete follow-up" }] })); await Bun.sleep(0)
    expect(f.root.textContent).not.toContain("Obsolete follow-up")
    expect(f.root.textContent).toContain("Superseded")
    expect(f.root.textContent).not.toContain("Current assessment")
  } finally { f.dispose() }
})

test("discovery keeps persisted findings historical until a fresh-context reconciliation succeeds", async () => {
  const f = fixture((action, input) => action === "assessment" ? { ...review({ status: input.assessmentId ? "stale" : "complete" }), meaning: "Old meaning" } : undefined)
  try {
    f.setContext("Payment in 45 days."); await f.advance()
    expect(f.root.textContent).toContain(finding.title)
    expect(f.root.textContent).toContain("Superseded")
    expect(f.root.textContent).not.toContain("Possible consequence")
    expect(f.root.textContent).not.toContain("What changes")
    await f.advance()
    expect(f.root.textContent).not.toContain("Current assessment")
    expect(f.root.textContent).not.toContain("Possible consequence")
    expect(f.root.textContent).not.toContain("What changes")
  } finally { f.dispose() }
})

test("focused findings relabel on disable without losing the focused source button, and reenable starts fresh work", async () => {
  const f = fixture()
  try {
    await f.advance(); await f.advance()
    const show = [...f.root.querySelectorAll("button")].find(node => node.textContent === "Show in Word")!
    show.focus()
    f.state.enabled = false; f.ui.refresh()
    expect(f.root.textContent).toContain("Earlier finding · review is off")
    expect(f.root.textContent).not.toContain("Possible consequence")
    expect(show.ownerDocument.activeElement).toBe(show)
    expect(f.root.textContent).not.toContain("Current assessment")
    f.state.enabled = true; f.ui.refresh(); await f.advance(); await f.advance()
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(2)
    expect(show.ownerDocument.activeElement).toBe(show)
  } finally { f.dispose() }
})

test("a disabled persisted review restarts after enabling instead of remaining disabled", async () => {
  const f = fixture((action, input) => action === "assessment" && !input.assessmentId ? review({ status: "disabled" }) : undefined)
  try {
    f.state.enabled = false; await f.advance()
    expect(f.root.textContent).toContain("Earlier finding · review is off")
    f.state.enabled = true; f.ui.refresh(); await f.advance(); await f.advance()
    expect(f.calls.some(call => call.action === "assess")).toBe(true)
    expect(f.root.textContent).toContain("Checking consequences")
  } finally { f.dispose() }
})

test("removing the final amendment target cancels the running review and retains findings only as history", async () => {
  const f = fixture()
  try {
    await f.advance(); await f.advance()
    f.state.turnId = ""; f.state.revision = "declined"; f.ui.refresh()
    expect(f.calls.some(call => call.action === "assessment-cancel")).toBe(true)
    expect(f.root.textContent).toContain("Earlier finding · amendment declined or unavailable")
    const count = f.calls.filter(call => call.action === "assess").length
    await f.advance(30_000)
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(count)
    expect(f.root.textContent).not.toContain("Current assessment")
  } finally { f.dispose() }
})

test("an empty review mounts no source or proposal shells", async () => {
  const f = fixture()
  try {
    f.state.turnId = ""; await f.advance()
    expect(f.root.querySelector(".chat-source")).toBeNull()
    expect(f.root.querySelector(".chat-proposal")).toBeNull()
  } finally { f.dispose() }
})

test("failed fresh capture cancels an admitted job before dispose can lose its identity", async () => {
  let finish!: (value: unknown) => void
  const f = fixture(action => action === "assess" ? new Promise(resolve => { finish = resolve }) : undefined)
  try {
    await f.advance(); await f.advance()
    f.failNextCapture(); finish(review()); await Bun.sleep(0)
    f.ui.dispose()
    expect(f.calls.some(call => call.action === "assessment-cancel" && call.input.assessmentId === "assessment-one")).toBe(true)
    expect(f.root.textContent).not.toContain(finding.title)
  } finally { f.dispose() }
})

test("empty conversation discovery failures stay hidden and keep retrying every eight seconds", async () => {
  const f = fixture(action => { if (action === "assessment") throw Error("Disconnected") })
  try {
    f.state.turnId = ""; await f.advance(1_500)
    expect(f.root.hidden).toBe(true); expect(f.calls).toHaveLength(1)
    await f.advance(7_999)
    expect(f.calls).toHaveLength(1); expect(f.root.hidden).toBe(true)
    await f.advance(1)
    expect(f.calls).toHaveLength(2); expect(f.root.hidden).toBe(true)
    await f.advance(8_000); await f.advance(8_000)
    expect(f.calls).toHaveLength(4); expect(f.root.hidden).toBe(true)
    expect(f.captures()).toBe(0)
  } finally { f.dispose() }
})

test("amendment connection warnings require three consecutive failures lasting at least fifteen seconds", async () => {
  const f = fixture(action => { if (action === "assess") throw Error("Disconnected") })
  try {
    await f.advance(1_500); await f.advance(3_000)
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(1)
    expect(f.root.hidden).toBe(true)
    await f.advance(7_999)
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(1)
    await f.advance(1)
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(2)
    expect(f.root.hidden).toBe(true)
    await f.advance(7_000)
    // Fifteen seconds alone is insufficient while only two requests have failed.
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(2)
    expect(f.root.hidden).toBe(true)
    await f.advance(1_000)
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(3)
    expect(f.root.hidden).toBe(false)
    expect(f.root.textContent).toContain("unavailable")
    expect(f.captures()).toBeGreaterThanOrEqual(4)
  } finally { f.dispose() }
})

for (const transientFailures of [1, 2]) test(`${transientFailures} transient amendment connection failure${transientFailures === 1 ? "" : "s"} recover without a warning`, async () => {
  let failures = 0
  const f = fixture(action => {
    if (action === "assess" && failures++ < transientFailures) throw Error("Disconnected")
    return undefined
  })
  try {
    await f.advance(1_500); await f.advance(3_000)
    expect(f.root.hidden).toBe(true)
    for (let index = 1; index < transientFailures; index++) { await f.advance(8_000); expect(f.root.hidden).toBe(true) }
    await f.advance(8_000)
    expect(f.root.textContent).toContain(finding.title)
    expect(f.root.textContent).not.toContain("unavailable")
    await f.advance(3_000)
    expect(f.root.textContent).not.toContain("unavailable")
  } finally { f.dispose() }
})

test("API recovery clears a visible warning and starts a fresh failure streak", async () => {
  let failing = true
  const f = fixture((action, input) => {
    if (action === "assessment" && input.assessmentId && failing) throw Error("Disconnected")
    return undefined
  })
  try {
    await f.advance(1_500); await f.advance(3_000)
    await f.advance(3_000); await f.advance(8_000); await f.advance(8_000)
    expect(f.root.textContent).toContain("unavailable")
    failing = false; await f.advance(8_000)
    expect(f.root.textContent).not.toContain("unavailable")
    expect(f.root.textContent).toContain(finding.title)
    failing = true; await f.advance(3_000)
    expect(f.root.textContent).not.toContain("unavailable")
    await f.advance(8_000)
    expect(f.root.textContent).not.toContain("unavailable")
    await f.advance(8_000)
    expect(f.root.textContent).toContain("unavailable")
  } finally { f.dispose() }
})

test("successful null discovery clears failed discovery retry state in an empty conversation", async () => {
  let failing = true
  const f = fixture(action => {
    if (action === "assessment" && failing) throw Error("Disconnected")
    return undefined
  })
  try {
    f.state.turnId = ""; await f.advance(1_500); await f.advance(8_000)
    failing = false; await f.advance(8_000)
    expect(f.root.hidden).toBe(true)
    const count = f.calls.length
    await f.advance(30_000)
    expect(f.calls).toHaveLength(count)
    expect(f.captures()).toBe(0)
  } finally { f.dispose() }
})

for (const transition of ["conversation", "decline", "disable"] as const) test(`${transition} clears the previous review's connection warning and failure streak`, async () => {
  const f = fixture(action => { if (action === "assess") throw Error("Disconnected") })
  try {
    await f.advance(1_500); await f.advance(3_000); await f.advance(8_000); await f.advance(8_000)
    expect(f.root.textContent).toContain("unavailable")
    if (transition === "conversation") { f.state.conversationId = "conversation-two"; f.ui.reset() }
    else if (transition === "decline") { f.state.turnId = ""; f.state.revision = "declined"; f.ui.refresh() }
    else { f.state.enabled = false; f.ui.refresh() }
    expect(f.root.hidden).toBe(true)
    await f.advance(8_000)
    expect(f.root.hidden).toBe(true)
    if (transition === "decline") { f.state.turnId = "amendment-two"; f.state.revision = "new proposal"; f.ui.refresh() }
    else if (transition === "disable") { f.state.enabled = true; f.ui.refresh() }
    await f.advance(3_000); await f.advance(3_000)
    expect(f.root.textContent).not.toContain("unavailable")
    await f.advance(8_000)
    expect(f.root.textContent).not.toContain("unavailable")
  } finally { f.dispose() }
})

test("disabled discovery retries remain silent and still restore saved findings after recovery", async () => {
  let failing = true
  const f = fixture(action => {
    if (action === "assessment") {
      if (failing) throw Error("Disconnected")
      return review({ status: "complete" })
    }
    return undefined
  })
  try {
    f.state.enabled = false; await f.advance(1_500); await f.advance(8_000); await f.advance(8_000)
    expect(f.root.hidden).toBe(true)
    expect(f.calls.filter(call => call.action === "assessment")).toHaveLength(3)
    failing = false; await f.advance(8_000)
    expect(f.root.textContent).toContain(finding.title)
    expect(f.root.textContent).toContain("Consequence review is off")
    expect(f.root.textContent).not.toContain("unavailable")
    expect(f.captures()).toBe(0)
    expect(f.calls.some(call => call.action === "assess")).toBe(false)
  } finally { f.dispose() }
})

test("a completed assessment recovers its summary after a temporary Word read failure", async () => {
  const f = fixture((action, input) => action === "assess" || action === "assessment" && input.assessmentId ? review({ status: "complete" }) : undefined)
  try {
    await f.advance(1_500); await f.advance(3_000)
    expect(f.root.textContent).toContain("Current assessment")
    f.failNextCapture(); await f.advance(3_000)
    expect(f.root.textContent).not.toContain("Current assessment")
    expect(f.root.textContent).not.toContain("unavailable")
    await f.advance(8_000)
    expect(f.calls.some(call => call.action === "assessment" && call.input.assessmentId)).toBe(true)
    expect(f.root.textContent).toContain("Current assessment")
    expect(f.root.textContent).not.toContain("unavailable")
    expect(f.calls.filter(call => call.action === "assess")).toHaveLength(1)
  } finally { f.dispose() }
})
