import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import type { ClauseHistoryResult, ClauseHistorySourceResult } from "../../../dochaus/lib/clause-history-contract"
import { mountClauseHistory } from "../../../dochaus/lib/clause-history-view"
import { act, createElement } from "react"

const version = { id: "saved-one", number: 1, hash: "a".repeat(64), parent: null, createdAt: 1, actor: "Avery", origin: "upload", label: "Sent draft" }
const source = { versionId: version.id, hash: version.hash, kind: "comment" as const, id: "4" }
const reason = { kind: "comment" as const, text: "Please explain the period <script>bad()</script>.", author: "Morgan", date: null, source }
const result: ClauseHistoryResult = {
  document: { id: "doc-one", name: "Agreement.docx", currentVersion: "saved-one" }, anchorVersion: version.id, currentVersion: version.id,
  anchor: { status: "verified", paragraphs: [{ anchorId: "body:0", scope: "body", paragraph: 1, text: "Pay within fourteen days." }] },
  rounds: [{ version, status: "verified", paragraphs: [{ anchorId: "body:0", scope: "body", paragraph: 1, text: "Pay within fourteen days." }], evidence: "anchor-version", reasons: [reason], changes: [], reasonStatus: "recorded" }],
  wordingView: "current-text", limitations: ["Only retained sources can be checked."],
}
function fixture() {
  const page = new Window(), root = page.document.createElement("section") as unknown as HTMLElement
  page.document.body.append(root as never)
  return { page, root, button(text: string) { return [...root.querySelectorAll<HTMLButtonElement>("button")].find(item => item.textContent === text)! } }
}

test("the shared Word/web history shows exact round wording and qualified reasons, and follows the immutable source", async () => {
  const f = fixture(), calls: unknown[] = [], opened: string[] = []
  const view = mountClauseHistory(f.root, { load: async () => result, source: async input => { calls.push(input); return { version, source, reason } }, onOpenVersion: id => opened.push(id) })
  try {
    await view.ready
    expect(f.root.textContent).toContain("Pay within fourteen days.")
    expect(f.root.textContent).toContain("Morgan · Date not recorded")
    expect(f.root.textContent).toContain("Comments are recorded notes; they do not establish an agreed concession.")
    expect(f.root.querySelector("script")).toBeNull()
    expect(f.root.querySelector("blockquote")?.textContent).toBe("Pay within fourteen days.")
    expect(calls).toEqual([]); expect(opened).toEqual([])
    f.button("View recorded source").click(); await Bun.sleep(0)
    expect(calls).toEqual([source])
    expect(f.root.querySelector('[aria-label="Recorded source"]')?.textContent).toContain(`SHA-256 ${version.hash}`)
    expect(f.root.querySelector('[aria-label="Recorded source"]')?.textContent).toContain(reason.text)
    f.button("Open saved version").click(); expect(opened).toEqual([version.id])
  } finally { view.dispose(); await f.page.happyDOM.abort() }
})

test("unresolved and removed rounds show metadata and limitations without presenting candidate wording or reasons", async () => {
  const f = fixture(), data = structuredClone(result)
  data.anchor = { status: "unresolved", reason: "The current unsaved paragraph does not match the saved base.", paragraphs: [] }
  data.rounds.push({ ...structuredClone(data.rounds[0]!), version: { ...version, id: "saved-two", number: 2, parent: version.id }, status: "unresolved", reason: "Repeated wording cannot be linked uniquely." })
  data.rounds.push({ ...structuredClone(data.rounds[0]!), version: { ...version, id: "saved-three", number: 3, parent: "saved-two" }, status: "unavailable", reason: "Removed by redaction." })
  const view = mountClauseHistory(f.root, { load: async () => data, source: async () => { throw new Error("unused") } })
  try {
    await view.ready
    const rounds = f.root.querySelectorAll(".clause-history-round")
    expect(rounds[1]!.textContent).toContain("Repeated wording cannot be linked uniquely.")
    expect(rounds[1]!.textContent).toContain("Based on Version 1")
    expect(rounds[2]!.textContent).toContain("Removed by redaction.")
    for (const round of [rounds[1]!, rounds[2]!]) {
      expect(round.querySelector("blockquote")).toBeNull(); expect(round.querySelector("button")).toBeNull()
    }
    expect(f.root.textContent).toContain("current unsaved paragraph")
  } finally { view.dispose(); await f.page.happyDOM.abort() }
})

test("a decision without a rationale stays factual and keeps the missing negotiation reason visible", async () => {
  const f = fixture(), data = structuredClone(result)
  data.rounds[0]!.reasonStatus = "not-recorded"
  data.rounds[0]!.reasons = [{ kind: "decision", text: "", author: null, date: null, source: { ...source, kind: "decision", id: "event:1" }, decision: "accept", outcome: "accepted" }]
  const view = mountClauseHistory(f.root, { load: async () => data, source: async () => { throw new Error("unused") } })
  try {
    await view.ready
    expect(f.root.textContent).toContain("Recorded decision · accept")
    expect(f.root.textContent).toContain("No negotiation reason recorded for this clause in this version.")
    expect(f.root.querySelectorAll("blockquote")).toHaveLength(1)
  } finally { view.dispose(); await f.page.happyDOM.abort() }
})

test("source redaction or identity mismatch clears all cached history authority and permits an explicit refresh", async () => {
  for (const mode of ["redacted", "mismatched"] as const) {
    const f = fixture(); let loads = 0
    const view = mountClauseHistory(f.root, { load: async () => { loads++; return result }, source: async () => {
      if (mode === "redacted") throw new Error("This source was removed by redaction.")
      return { version, source: { ...source, hash: "b".repeat(64) }, reason }
    } })
    try {
      await view.ready; f.button("View recorded source").click(); await Bun.sleep(0)
      expect(f.root.querySelector(".clause-history-body")?.textContent).toBe("")
      expect(f.root.textContent).not.toContain(reason.text)
      expect(f.root.textContent).not.toContain("Pay within fourteen days.")
      expect(f.root.textContent).toContain(mode === "redacted" ? "removed by redaction" : "no longer matches")
      expect(f.button("Refresh").disabled).toBe(false)
      f.button("Refresh").click(); await Bun.sleep(0)
      expect(loads).toBe(2); expect(f.root.textContent).toContain("Pay within fourteen days.")
    } finally { view.dispose(); await f.page.happyDOM.abort() }
  }
})

test("close and refresh cancel pending reads; late history or source responses cannot repaint the surface", async () => {
  const f = fixture(); let finish!: (value: ClauseHistoryResult) => void, firstSignal: AbortSignal | undefined
  let loads = 0, closes = 0
  const view = mountClauseHistory(f.root, { load: signal => { loads++; if (loads === 1) { firstSignal = signal; return new Promise(resolve => { finish = resolve }) }; return Promise.resolve(result) }, source: async () => { throw new Error("unused") }, onClose: () => { closes++ } })
  try {
    await view.refresh(); expect(firstSignal?.aborted).toBe(true)
    finish({ ...result, document: { ...result.document, name: "Old document" } }); await view.ready
    expect(f.root.textContent).not.toContain("Old document")
    view.focus(); expect(f.page.document.activeElement?.textContent).toBe("Refresh")
    f.root.dispatchEvent(new f.page.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as KeyboardEvent)
    expect(closes).toBe(1)
    view.dispose(); expect(f.root.childNodes.length).toBe(0)
  } finally { view.dispose(); await f.page.happyDOM.abort() }
})

test("a disposed source read is aborted and its late result stays discarded", async () => {
  const f = fixture(); let finish!: (value: ClauseHistorySourceResult) => void, signal: AbortSignal | undefined
  const view = mountClauseHistory(f.root, { load: async () => result, source: (_source, requestSignal) => { signal = requestSignal; return new Promise(resolve => { finish = resolve }) } })
  await view.ready; f.button("View recorded source").click(); view.dispose()
  expect(signal?.aborted).toBe(true)
  finish({ version, source, reason }); await Bun.sleep(0)
  expect(f.root.childNodes.length).toBe(0)
  await f.page.happyDOM.abort()
})

test("linked-return rounds name their physical source version and source reads retain the receipt qualification", async () => {
  for (const mismatch of [false, true]) {
    const f = fixture(), data = structuredClone(result), received = { ...version, id: "other-doc-v1", hash: "b".repeat(64), linkedReturn: { exchangeId: 44, sourceDocumentId: "other-doc", sourceDocumentName: "Separate returned copy.docx", targetDocumentId: data.document.id, baseVersionId: version.id } }
    const qualified = { ...source, versionId: received.id, hash: received.hash, exchangeId: 44 }, note = { ...reason, source: qualified }
    data.rounds[0]!.reasons = []
    data.rounds.push({ ...structuredClone(data.rounds[0]!), version: received, reasons: [note] })
    const calls: unknown[] = []
    const view = mountClauseHistory(f.root, { load: async () => data, source: async request => { calls.push(request); return { version: received, source: mismatch ? { ...qualified, exchangeId: 45 } : qualified, reason: note } } })
    try {
      await view.ready
      expect(f.root.textContent).toContain("Linked return · Separate returned copy.docx version 1")
      expect(f.root.querySelectorAll(".clause-history-round")[1]!.textContent).toContain("Based on Version 1")
      f.button("View recorded source").click(); await Bun.sleep(0)
      expect(calls).toEqual([qualified])
      if (mismatch) { expect(f.root.querySelector(".clause-history-body")?.textContent).toBe(""); expect(f.root.textContent).toContain("no longer matches") }
      else expect(f.root.querySelector('[aria-label="Recorded source"]')?.textContent).toContain("Linked receipt 44")
    } finally { view.dispose(); await f.page.happyDOM.abort() }
  }
})

test("the web History component cancels an old selected paragraph and uses the same exact source route", async () => {
  const page = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Element", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const saved = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : page[name as keyof Window] }])))
  const originalFetch = globalThis.fetch, calls: { url: string; init?: RequestInit }[] = []
  let old!: (response: Response) => void
  globalThis.fetch = (async (url, init) => {
    calls.push({ url: String(url), init })
    if (calls.length === 1) return new Promise<Response>(resolve => { old = resolve })
    return Response.json(String(url).includes("/source?") ? { version, source, reason } : result)
  }) as typeof fetch
  const { ClauseHistory } = await import("../src/components/ClauseHistory")
  const { createRoot } = await import("react-dom/client")
  const container = document.createElement("div"); document.body.append(container)
  const root = createRoot(container)
  const request = { anchorVersion: version.id, paragraph: { scope: "body" as const, text: "Full paragraph." }, expectedCurrentHash: version.hash }
  try {
    await act(async () => { root.render(createElement(ClauseHistory, { matterId: "matter/one", name: "Draft & terms.docx", request })) })
    await act(async () => { root.render(createElement(ClauseHistory, { matterId: "matter/one", name: "Draft & terms.docx", request: { ...request, paragraph: { ...request.paragraph, text: "Another full paragraph." } } })) })
    expect(calls[0]!.init?.signal?.aborted).toBe(true)
    await act(async () => { old(Response.json({ ...result, document: { ...result.document, name: "Old paragraph document" } })) })
    expect(container.textContent).not.toContain("Old paragraph document")
    expect(container.textContent).toContain("Agreement.docx")
    const sourceButton = [...container.querySelectorAll<HTMLButtonElement>("button")].find(button => button.textContent === "View recorded source")!
    await act(async () => { sourceButton.click() })
    expect(calls[2]!.url).toContain("/matters/matter%2Fone/documents/clause-history/source?name=Draft%20%26%20terms.docx")
    expect(JSON.parse(calls[2]!.init!.body as string)).toEqual(source)
    expect(JSON.parse(calls[1]!.init!.body as string)).toEqual({ ...request, paragraph: { ...request.paragraph, text: "Another full paragraph." } })
    expect(container.querySelector('[aria-label="Recorded source"]')?.textContent).toContain(reason.text)
  } finally {
    await act(async () => root.unmount()); globalThis.fetch = originalFetch; await page.happyDOM.close()
    for (const [name, descriptor] of saved) { if (descriptor) Object.defineProperty(globalThis, name, descriptor); else delete (globalThis as Record<string, unknown>)[name] }
  }
})

test("saved proposal rationale and wording retain their author separately from the human decision, including source metadata pins", async () => {
  const f = fixture(), data = structuredClone(result), calls: unknown[] = []
  const pinned = { ...source, kind: "proposal" as const, id: "redline:12:7", recordHash: "c".repeat(64) }
  const proposal = { kind: "proposal" as const, text: "Allow time for invoice approval.", author: "Drafting agent", date: "2026-11-10T12:00:00Z", source: pinned, proposal: { id: 7, scope: "phrase", oldText: "fourteen days", newText: "thirty days" } }
  data.rounds[0]!.reasons = [proposal, { kind: "decision", text: "", author: "Avery", date: "2026-11-10T12:05:00Z", source: { ...pinned, kind: "decision" }, decision: "accept", outcome: "accepted" }]
  data.rounds[0]!.reasonStatus = "not-recorded"
  const view = mountClauseHistory(f.root, { load: async () => data, source: async request => { calls.push(request); return { version, source: pinned, reason: proposal } } })
  try {
    await view.ready
    const records = f.root.querySelectorAll(".clause-history-record")
    expect(records[0]!.textContent).toContain("Proposed reasonDrafting agent")
    expect(records[0]!.textContent).toContain("Original wordingfourteen daysProposed replacementthirty days")
    expect(records[1]!.textContent).toContain("Recorded decision · acceptAvery")
    expect(Boolean(records[1]!.querySelector("blockquote"))).toBe(false)
    expect(f.root.textContent).toContain("No negotiation reason recorded")
    expect(f.root.textContent).toContain("does not establish a human rationale or an agreed concession")
    f.button("View recorded source").click(); await Bun.sleep(0)
    expect(calls).toEqual([pinned])
    expect(f.root.querySelector('[aria-label="Recorded source"]')?.textContent).toContain(`Record SHA-256 ${pinned.recordHash}`)
    expect(f.root.querySelector('[aria-label="Recorded source"]')?.textContent).toContain("Proposed replacementthirty days")
  } finally { view.dispose(); await f.page.happyDOM.abort() }
})

test("changed or missing proposal record pins discard cached history for either source response identity", async () => {
  for (const mismatch of ["response", "record"] as const) {
    const f = fixture(), data = structuredClone(result)
    const pinned = { ...source, kind: "proposal" as const, id: "redline:12:7", recordHash: "c".repeat(64) }
    const proposal = { ...reason, kind: "proposal" as const, source: pinned }
    data.rounds[0]!.reasons = [proposal]
    const view = mountClauseHistory(f.root, { load: async () => data, source: async () => ({ version, source: mismatch === "response" ? { ...pinned, recordHash: "d".repeat(64) } : pinned, reason: mismatch === "record" ? { ...proposal, source: { ...pinned, recordHash: undefined } } : proposal }) })
    try {
      await view.ready; f.button("View recorded source").click(); await Bun.sleep(0)
      expect(f.root.querySelector(".clause-history-body")?.textContent).toBe("")
      expect(f.root.textContent).toContain("no longer matches")
    } finally { view.dispose(); await f.page.happyDOM.abort() }
  }
})
