import type { ClauseHistoryChange, ClauseHistoryReason, ClauseHistoryResult, ClauseHistoryRound, ClauseHistorySource, ClauseHistorySourceResult, ClauseHistoryVersion } from "./clause-history-contract"

export type ClauseHistoryViewOptions = {
  load(signal: AbortSignal): Promise<ClauseHistoryResult>
  source(source: ClauseHistorySource, signal: AbortSignal): Promise<ClauseHistorySourceResult>
  onClose?(): void
  onOpenVersion?(versionId: string): void
}

/** One read-only history surface in Word and the web. Source failures discard
 * the cached view: removed or changed bytes must not retain visible authority. */
export function mountClauseHistory(root: HTMLElement, options: ClauseHistoryViewOptions) {
  const doc = root.ownerDocument
  const node = <K extends keyof HTMLElementTagNameMap>(tag: K, text?: string, className?: string) => {
    const item = doc.createElement(tag)
    if (text !== undefined) item.textContent = text
    if (className) item.className = className
    return item
  }
  const button = (text: string, run: () => void) => {
    const item = node("button", text); item.type = "button"; item.onclick = run; return item
  }
  root.classList.add("clause-history")
  root.setAttribute("aria-label", "Clause history")
  const heading = node("div", undefined, "clause-history-heading"), title = node("h2", "Clause history")
  const refresh = button("Refresh", () => { void load() })
  heading.append(title, refresh)
  if (options.onClose) heading.append(button("Done", options.onClose))
  const status = node("p", "", "clause-history-status")
  status.setAttribute("role", "status"); status.setAttribute("aria-live", "polite")
  const body = node("div", undefined, "clause-history-body"); body.tabIndex = 0
  body.setAttribute("aria-label", "Saved clause history")
  root.replaceChildren(heading, status, body)
  let epoch = 0, disposed = false, controller: AbortController | undefined
  const sourceControllers = new Set<AbortController>()
  function cancelSources() { for (const controller of sourceControllers) controller.abort(); sourceControllers.clear() }
  const date = (value: string | number | null) => {
    if (value === null) return "Date not recorded"
    const at = new Date(value)
    return Number.isNaN(at.getTime()) ? "Date not recorded" : at.toLocaleString()
  }
  function versionLabel(version: ClauseHistoryVersion) {
    if (version.linkedReturn) return `Linked return · ${version.linkedReturn.sourceDocumentName} version ${version.number}`
    return `Version ${version.number}${version.label ? ` · ${version.label}` : ""}`
  }
  function metadata(version: ClauseHistoryVersion) {
    return `${version.actor || "Author not recorded"} · ${date(version.createdAt)}${version.party ? ` · ${version.party}` : ""} · ${version.origin}`
  }
  function sourceIdentity(source: ClauseHistorySource) {
    return node("p", `Source ${source.kind} ${source.id} · Version ${source.versionId} · SHA-256 ${source.hash}${source.recordHash ? ` · Record SHA-256 ${source.recordHash}` : ""}${source.exchangeId !== undefined ? ` · Linked receipt ${source.exchangeId}` : ""}`, "clause-history-source-id")
  }
  function recorded(item: ClauseHistoryReason | ClauseHistoryChange) {
    const section = node("div", undefined, "clause-history-record")
    const reason = item as ClauseHistoryReason
    const label = item.kind === "comment" ? `${reason.parent ? "Reply" : "Comment"}${reason.resolved === true ? " · Resolved" : ""}` : item.kind === "proposal" ? "Proposed reason" : item.kind === "decision" ? `Recorded decision${reason.decision ? ` · ${reason.decision}` : ""}` : `Tracked change · ${item.kind}`
    section.append(node("h4", label), node("p", `${item.author || "Author not recorded"} · ${date(item.date)}`, "clause-history-meta"))
    if (item.text) section.append(node("blockquote", item.text))
    if (reason.proposal) {
      section.append(node("h4", "Proposed wording"), node("p", `Proposal ${reason.proposal.id} · ${reason.proposal.scope}`, "clause-history-meta"), node("p", "Original wording"), node("blockquote", reason.proposal.oldText), node("p", "Proposed replacement"), node("blockquote", reason.proposal.newText))
    }
    if (reason.parent) section.append(node("p", `Reply to comment ${reason.parent}`, "clause-history-meta"))
    if (reason.outcome) section.append(node("p", reason.outcome))
    return section
  }
  function appendSource(parent: HTMLElement, source: ClauseHistorySource, viewEpoch: number) {
    const output = node("div", undefined, "clause-history-source")
    const open = button("View recorded source", async () => {
      if (disposed || epoch !== viewEpoch || open.disabled) return
      const request = new AbortController(); sourceControllers.add(request)
      open.disabled = true; output.replaceChildren(node("p", "Checking saved source…"))
      try {
        const result = await options.source(source, request.signal)
        if (disposed || epoch !== viewEpoch || request.signal.aborted) return
        // A mismatched source cannot be presented as the referenced record.
        if (result.source.versionId !== source.versionId || result.source.hash !== source.hash || result.source.kind !== source.kind || result.source.id !== source.id || result.source.recordHash !== source.recordHash || result.source.exchangeId !== source.exchangeId || result.version.id !== source.versionId || result.version.hash !== source.hash || result.version.linkedReturn?.exchangeId !== source.exchangeId) throw new Error("The saved source no longer matches this history. Refresh to check it again.")
        const item = result.reason ?? result.change
        if (!item) throw new Error("The recorded source is unavailable. Refresh to check it again.")
        if (item.source.versionId !== source.versionId || item.source.hash !== source.hash || item.source.kind !== source.kind || item.source.id !== source.id || item.source.recordHash !== source.recordHash || item.source.exchangeId !== source.exchangeId) throw new Error("The saved source no longer matches this history. Refresh to check it again.")
        output.replaceChildren(node("h4", `Recorded source · ${versionLabel(result.version)}`), recorded(item), sourceIdentity(result.source))
        output.setAttribute("role", "region"); output.setAttribute("aria-label", "Recorded source")
        status.textContent = "Saved source checked."
      } catch (error) {
        if (disposed || epoch !== viewEpoch || request.signal.aborted) return
        epoch++; cancelSources(); body.replaceChildren()
        status.textContent = error instanceof Error ? error.message : "The saved source is unavailable. Refresh to check it again."
      } finally { sourceControllers.delete(request); open.disabled = false }
    })
    parent.append(open, output)
  }
  function roundView(round: ClauseHistoryRound, result: ClauseHistoryResult, viewEpoch: number) {
    const item = node("li", undefined, "clause-history-round")
    item.dataset.status = round.status
    item.append(node("h3", `${versionLabel(round.version)}${round.version.id === result.currentVersion ? " · Current saved version" : ""}${round.version.id === result.anchorVersion ? " · Selected version" : ""}`), node("p", metadata(round.version), "clause-history-meta"))
    const base = round.version.linkedReturn?.baseVersionId ?? round.version.parent
    if (base) {
      const parent = result.rounds.find(candidate => candidate.version.id === base)
      item.append(node("p", `Based on ${parent ? `Version ${parent.version.number}` : `saved version ${base}`}`, "clause-history-meta"))
    }
    const labels = { verified: "Wording linked", unresolved: "Link unresolved", unavailable: "Version unavailable" }
    item.append(node("p", labels[round.status], "clause-history-state"))
    if (round.reason) item.append(node("p", round.reason))
    if (round.status === "verified") {
      const evidence = { "anchor-version": "Exact selected wording in this saved version", "word-paragraph-id": "Linked by saved Word paragraph identity", "exact-wording": "Linked by unique exact wording", "exact-split-merge": "Linked by exact wording across split or merged paragraphs" }
      if (round.evidence) item.append(node("p", evidence[round.evidence], "clause-history-meta"))
      for (const paragraph of round.paragraphs) {
        const wording = node("div", undefined, "clause-history-wording")
        wording.append(node("p", `${paragraph.scope} · Paragraph ${paragraph.paragraph}`, "clause-history-meta"), node("blockquote", paragraph.text))
        item.append(wording)
      }
      if (round.reasonStatus !== "recorded") item.append(node("p", round.reasonStatus === "not-recorded" ? "No negotiation reason recorded for this clause in this version." : "Recorded reasons are unavailable for this version.", "clause-history-note"))
      if (round.reasons.some(reason => reason.kind === "comment")) item.append(node("p", "Comments are recorded notes; they do not establish an agreed concession.", "clause-history-note"))
      if (round.reasons.some(reason => reason.kind === "proposal")) item.append(node("p", "Proposed reasons belong to their recorded proposal author; accepting or rejecting wording does not establish a human rationale or an agreed concession.", "clause-history-note"))
      for (const record of [...round.reasons, ...round.changes]) {
        const section = recorded(record); appendSource(section, record.source, viewEpoch); item.append(section)
      }
    }
    if (options.onOpenVersion && round.status !== "unavailable") item.append(button("Open saved version", () => options.onOpenVersion?.(round.version.id)))
    return item
  }
  function render(result: ClauseHistoryResult, viewEpoch: number) {
    body.replaceChildren(node("p", result.document.name, "clause-history-document"), node("p", "Wording shows the current text saved in each version. Pending tracked edits are not treated as approved.", "clause-history-note"))
    if (result.anchor.status === "unresolved") body.append(node("p", result.anchor.reason || "This paragraph could not be linked to the selected saved version.", "clause-history-state"))
    const list = node("ol", undefined, "clause-history-rounds")
    for (const round of result.rounds) list.append(roundView(round, result, viewEpoch))
    body.append(list)
    if (!result.rounds.length) body.append(node("p", "No saved clause history is available."))
    if (result.limitations.length) {
      const notes = node("ul", undefined, "clause-history-limitations")
      for (const text of result.limitations) notes.append(node("li", text))
      body.append(notes)
    }
  }
  async function load() {
    const viewEpoch = ++epoch
    controller?.abort(); cancelSources(); controller = new AbortController()
    body.replaceChildren(); status.textContent = "Loading saved clause history…"; refresh.disabled = true
    try {
      const result = await options.load(controller.signal)
      if (disposed || viewEpoch !== epoch) return
      render(result, viewEpoch)
      status.textContent = result.anchor.status === "verified" ? "Saved clause history loaded." : "Paragraph link unresolved."
    } catch (error) {
      if (disposed || viewEpoch !== epoch) return
      status.textContent = error instanceof Error ? error.message : "Could not load clause history. Try Refresh."
    } finally { if (!disposed && viewEpoch === epoch) refresh.disabled = false }
  }
  const escape = (event: KeyboardEvent) => {
    if (event.key !== "Escape" || !options.onClose) return
    event.preventDefault(); event.stopPropagation(); options.onClose()
  }
  root.addEventListener("keydown", escape)
  const ready = load()
  return { ready, refresh: load, focus: () => (heading.querySelector<HTMLElement>("button") ?? body).focus(), dispose() {
    disposed = true; epoch++; controller?.abort(); cancelSources(); root.removeEventListener("keydown", escape); root.replaceChildren()
  } }
}
