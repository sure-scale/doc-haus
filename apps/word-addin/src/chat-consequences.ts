import type { WordConsequenceAssessment, WordConsequenceFinding } from "../../../dochaus/lib/word-consequence-contract"
import type { WordChatSnapshot, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { wordChatTargetCurrent } from "../../../dochaus/lib/word-chat-amendment-contract"
import type { WordChatApi } from "./chat-client"

type ReviewState = {
  conversationId?: string
  turnId?: string
  revision: string
  enabled: boolean
  busy: boolean
}
type LatestAssessment = WordConsequenceAssessment

type ReviewHost = {
  capture(): Promise<WordChatSnapshot>
  navigate?(target: WordChatTarget): Promise<boolean>
  now(): number
  schedule(run: () => void, delay: number): () => void
}

/** Background review has its own lifecycle: it never owns foreground controls or scroll. */
export function mountWordConsequences(root: HTMLElement, api: WordChatApi, host: ReviewHost, state: () => ReviewState) {
  const document = root.ownerDocument
  const line = (tag: string, value: string, className?: string) => {
    const node = document.createElement(tag); node.textContent = value
    if (className) node.className = className
    return node
  }
  root.setAttribute("aria-label", "Amendment consequences")
  root.hidden = true
  const heading = line("strong", "Consequences · doc.haus", "chat-author")
  const progress = line("p", "", "chat-message-meta")
  progress.setAttribute("role", "status"); progress.setAttribute("aria-live", "polite")
  const meaning = document.createElement("section"), findings = document.createElement("div"), summary = document.createElement("div"), earlier = document.createElement("details")
  meaning.className = summary.className = "chat-proposal"
  earlier.className = "chat-source"; earlier.append(line("summary", "Earlier findings"))
  const earlierBody = document.createElement("div"); earlier.append(earlierBody)
  root.append(heading, progress, findings)
  const records = new Map<string, { finding: WordConsequenceFinding; context?: WordChatSnapshot; active: boolean }>()
  let conversationId: string | undefined, epoch = 0, disposed = false, busy = false, discovered = false, needsReconcile = false
  let assessment: LatestAssessment | undefined
  let observedRevision = "", observedEnabled = state().enabled, key = "", stableAt = 0, interactionAt = host.now(), requestKey = "", requestId = "", error = ""
  const cards = new Map<string, { node: HTMLElement; status: HTMLElement; stamp: string }>()
  let cancelTimer: (() => void) | undefined, pollAbort: AbortController | undefined
  const changed = () => { interactionAt = host.now() }
  function cancel(value = assessment) {
    if (value?.status === "running") void api.request("assessment-cancel", { conversationId: value.conversationId, assessmentId: value.id }).catch(() => {})
  }
  function invalidate() {
    epoch++; pollAbort?.abort(); pollAbort = undefined; cancel(); assessment = undefined
    requestKey = ""; requestId = ""; error = ""
    for (const record of records.values()) record.active = false
  }
  function findingKey(finding: WordConsequenceFinding) {
    // Model-generated identifiers can change across rechecks. Unchanged evidence and meaning do not make another card.
    return JSON.stringify([finding.kind, finding.title, finding.explanation, finding.passage, finding.related])
  }
  function accept(value: LatestAssessment | undefined, context?: WordChatSnapshot, historical = false) {
    if (!value || value.conversationId !== conversationId) return
    assessment = value; error = ""
    for (const record of records.values()) record.active = false
    for (const previous of value.history ?? []) for (const finding of previous.findings) {
      const id = findingKey(finding)
      if (!records.has(id)) records.set(id, { finding, active: false })
    }
    for (const finding of value.findings) records.set(findingKey(finding), { finding, context, active: !historical && ["running", "complete", "partial"].includes(value.status) })
  }
  function card(record: { finding: WordConsequenceFinding; context?: WordChatSnapshot; active: boolean }) {
    const { finding, context } = record
    const card = document.createElement("section"); card.className = "chat-proposal"
    card.append(line("h3", finding.title), line("p", finding.explanation))
    const status = line("p", "", "chat-message-meta"); card.append(status)
    for (const passage of [finding.passage, ...(finding.related ?? [])]) {
      const detail = document.createElement("details"); detail.className = "chat-source"
      detail.append(line("summary", passage === finding.passage ? "Affected wording" : "Related wording"), line("p", passage.quote))
      const paragraph = context?.paragraphs.find(item => item.id === passage.paragraphId)
      if (host.navigate && context && paragraph && paragraph.text.includes(passage.quote)) {
        const target: WordChatTarget = { id: `consequence-${paragraph.id}`, text: paragraph.text, paragraphs: [paragraph], startOffset: 0, endOffset: paragraph.text.length, capturedAt: context.capturedAt, saved: context.saved }
        const show = document.createElement("button"); show.type = "button"; show.className = "chat-subtle"; show.textContent = "Show in Word"
        const notice = line("p", "", "chat-message-meta"); notice.setAttribute("role", "status")
        show.onclick = async () => {
          show.disabled = true
          try {
            const current = await host.capture()
            if (disposed) return
            if (!wordChatTargetCurrent(current, target) || !await host.navigate!(target)) notice.textContent = "This wording changed. Review the latest finding before navigating."
          } catch { notice.textContent = "Word could not locate this wording. Check the document directly." }
          finally { show.disabled = false }
        }
        detail.append(show, notice)
      }
      card.append(detail)
    }
    return { node: card, status }
  }
  function render() {
    if (disposed) return
    const current = state(), complete = assessment && ["complete", "partial"].includes(assessment.status)
    const settled = complete && current.enabled && Boolean(current.turnId) && !needsReconcile && !current.busy && host.now() - Math.max(stableAt, interactionAt) >= 3_000
    root.hidden = !assessment && !records.size && !error
    root.className = root.hidden ? "" : "chat-message"
    progress.textContent = !current.enabled ? "Consequence review is off. Earlier findings remain here." : error || (assessment?.status === "running" ? "Checking consequences alongside your conversation…" : assessment?.status === "stale" ? "The wording changed. Checking the latest document…" : assessment?.status === "cancelled" ? "Review stopped." : !assessment && records.size && current.turnId ? "Rechecking the latest wording and instructions…" : "")
    meaning.replaceChildren()
    if (assessment?.meaning && ["running", "complete", "partial"].includes(assessment.status) && !needsReconcile && current.enabled && current.turnId) {
      meaning.append(line("h3", "What changes"), line("p", assessment.meaning), line("p", "Inferred from the reviewed before and after wording", "chat-message-meta"))
      if (!meaning.isConnected) root.insertBefore(meaning, findings)
    } else meaning.remove()
    for (const [id, cached] of cards) if (!records.has(id)) { cached.node.remove(); cards.delete(id) }
    for (const [id, record] of records) {
      const stamp = JSON.stringify(record.context?.paragraphs)
      let cached = cards.get(id)
      // Keep an interacted-with source/button mounted. Status still updates immediately,
      // and its existing navigation closure continues to verify its original wording.
      if (!cached || cached.stamp !== stamp && !cached.node.contains(document.activeElement)) {
        const next = { ...card(record), stamp }
        cached?.node.replaceWith(next.node); cached = next; cards.set(id, cached)
      }
      cached.status.textContent = record.active ? record.finding.kind === "uncertain" ? "Needs checking · inferred from the wording" : "Possible consequence · inferred from the wording" : !current.enabled ? "Earlier finding · review is off" : !current.turnId ? "Earlier finding · amendment declined or unavailable" : "Superseded by a later review"
      const destination = record.active ? findings : earlierBody
      if (cached.node.parentElement !== destination && !cached.node.contains(document.activeElement)) destination.append(cached.node)
    }
    if (earlierBody.children.length) { if (!earlier.isConnected) root.append(earlier) } else earlier.remove()
    summary.replaceChildren()
    if (settled && assessment) {
      const { coverage } = assessment
      summary.append(line("h3", "Current assessment"), line("p", assessment.findings.length ? `${assessment.findings.length} consequence${assessment.findings.length === 1 ? "" : "s"} to consider. Your document is unchanged by this review.` : "No consequences were identified in the checked passages."))
      summary.append(line("p", `${coverage.checked} of ${coverage.candidates} candidate passages checked · ${coverage.reviewed} of ${coverage.total} paragraphs searched`, "chat-message-meta"))
      for (const gap of assessment.gaps) summary.append(line("p", gap, "chat-message-meta"))
      if (!summary.isConnected) root.insertBefore(summary, earlier.isConnected ? earlier : null)
    } else summary.remove()
  }
  const contextKey = (current: ReviewState, context: WordChatSnapshot) => JSON.stringify([current.turnId, current.revision, context.text, context.paragraphs, context.truncated])
  function syncEnabled() {
    const enabled = state().enabled
    if (!enabled || observedEnabled !== enabled) { invalidate(); needsReconcile = false; key = ""; stableAt = host.now() }
    observedEnabled = enabled
  }
  async function publish(value: LatestAssessment, context: WordChatSnapshot, current: ReviewState, operationEpoch: number) {
    const valid = () => !disposed && operationEpoch === epoch && state().enabled && state().conversationId === current.conversationId && state().turnId === current.turnId && state().revision === current.revision
    if (!valid()) { cancel(value); return false }
    // Word edits do not change chat revision. A fresh host read after the response
    // is required before any finding or meaning can become current.
    let fresh: WordChatSnapshot
    try { fresh = await host.capture() }
    catch (error) { cancel(value); throw error }
    if (!valid()) { cancel(value); return false }
    if (contextKey(current, fresh) !== contextKey(current, context)) {
      cancel(value); invalidate(); key = contextKey(current, fresh); stableAt = host.now(); render(); return false
    }
    needsReconcile = false; accept(value, fresh); return true
  }
  async function tick() {
    if (disposed || busy) return
    busy = true
    let operationEpoch = epoch
    try {
      const current = state()
      observedRevision = current.revision
      syncEnabled()
      if (current.conversationId !== conversationId) {
        invalidate(); conversationId = current.conversationId; records.clear(); key = ""; discovered = false; needsReconcile = false
      }
      operationEpoch = epoch
      if (!conversationId) { render(); return }
      // A hidden pane needs no fresh captures unless it has unfinished work.
      if ((document.hidden || root.parentElement?.closest("[hidden]")) && assessment?.status !== "running") return
      const id = conversationId
      if (!discovered) {
        const result = await api.request<LatestAssessment | { assessment: null }>("assessment", { conversationId: id })
        if (disposed || operationEpoch !== epoch || state().conversationId !== id) return
        discovered = true
        if ("id" in result) { accept(result, undefined, true); needsReconcile = true }
      }
      if (!current.enabled || !state().enabled) { invalidate(); render(); return }
      if (!current.turnId || !state().turnId) { invalidate(); render(); return }
      const context = await host.capture()
      if (disposed || operationEpoch !== epoch || state().conversationId !== id || !state().enabled) return
      const nextKey = contextKey(current, context)
      if (key !== nextKey) {
        const restored = !key && assessment?.turnId === current.turnId
        if (!restored) invalidate()
        key = nextKey; stableAt = host.now(); operationEpoch = epoch
        if (restored && assessment) needsReconcile = true
        render(); return
      }
      if (state().revision !== current.revision) { invalidate(); render(); return }
      if (host.now() - stableAt < 1_500) { render(); return }
      if (assessment && needsReconcile) {
        const result = await api.request<WordConsequenceAssessment>("assessment", { conversationId: id, assessmentId: assessment.id, context })
        if (disposed || operationEpoch !== epoch || !state().enabled || state().conversationId !== id || state().revision !== current.revision) return
        await publish(result, context, current, operationEpoch)
      } else if (!assessment) {
        if (requestKey !== key) { requestKey = key; requestId = crypto.randomUUID() }
        const result = await api.request<WordConsequenceAssessment>("assess", { conversationId: id, requestId, turnId: current.turnId, context })
        if (disposed || operationEpoch !== epoch || !state().enabled || state().conversationId !== id || state().revision !== current.revision) { cancel(result); return }
        await publish(result, context, current, operationEpoch)
      } else if (assessment.status === "running") {
        pollAbort = new AbortController()
        const result = await api.request<WordConsequenceAssessment>("assessment", { conversationId: id, assessmentId: assessment.id, context }, pollAbort.signal)
        if (disposed || operationEpoch !== epoch || !state().enabled || state().conversationId !== id || state().revision !== current.revision) return
        await publish(result, context, current, operationEpoch)
      } else if (["stale", "cancelled", "disabled"].includes(assessment.status)) { invalidate() }
      render()
    } catch {
      if (!disposed && operationEpoch === epoch) { error = "Consequence review is unavailable. Your chat and amendments are ready; the review will reconnect."; render() }
    } finally { busy = false; if (!disposed && (state().turnId || assessment?.status === "running" || error)) cancelTimer = host.schedule(() => { cancelTimer = undefined; void tick() }, error ? 8_000 : 3_000) }
  }
  function refresh() {
    const current = state()
    syncEnabled()
    if (!current.turnId || observedRevision && observedRevision !== current.revision) { invalidate(); observedRevision = current.revision }
    render()
    if (!disposed && !busy && !cancelTimer) cancelTimer = host.schedule(() => { cancelTimer = undefined; void tick() }, 1_500)
  }
  // Feature tabs stay mounted. Wake a stopped, empty review when its pane opens.
  const Observer = document.defaultView?.MutationObserver
  const visibility = Observer ? new Observer(refresh) : undefined
  for (let parent = root.parentElement; parent; parent = parent.parentElement) visibility?.observe(parent, { attributes: true, attributeFilter: ["hidden"] })
  cancelTimer = host.schedule(() => { cancelTimer = undefined; void tick() }, 1_500)
  return {
    interaction: changed,
    refresh,
    reset() { invalidate(); records.clear(); conversationId = undefined; key = ""; discovered = false; render() },
    dispose() { disposed = true; epoch++; visibility?.disconnect(); pollAbort?.abort(); cancelTimer?.(); cancel() },
  }
}
