import type { DeskComparison, DeskSource, DeskSources, DeskVersions } from "../../../dochaus/lib/word-source-desk-contract"
import { wordDiff, type Piece } from "../../../dochaus/lib/redline"
import { CHAT_UPLOAD_MAX_BYTES, isChatUploadName, type DeskRequest, type DeskView } from "./desk-protocol"

export type DeskTransport = { request(input: DeskRequest): Promise<unknown> }
export type DeskContext = { view: DeskView; documentName?: string; matterName?: string }
type DeskOptions = {
  close(): void
  floating?: boolean
  embedded?: boolean
  initialView?: DeskView
  onViewChange?(view: DeskView): void
  onContextChange?(context: DeskContext): void
  onBusyChange?(busy: boolean): void
  schedulePoll?(run: () => void, delayMs: number): () => void
  now?(): number
}

/** Matter documents open in desktop Word; comparison remains a read-only pane view. */
export function mountDesk(root: HTMLElement, transport: DeskTransport, options: DeskOptions) {
  root.classList.add("source-desk")
  root.classList.toggle("is-embedded", Boolean(options.embedded))
  root.innerHTML = `
    <header class="desk-head" ${options.embedded ? "hidden" : ""}>
      <p class="desk-identity" data-el="identity"></p>
      <nav class="desk-modes" aria-label="Document views"><button data-el="sources-mode" type="button" aria-pressed="true">Documents</button><button data-el="compare-mode" type="button" aria-pressed="false">Compare</button></nav>
      <button type="button" class="desk-icon" data-el="close" aria-label="${options.floating ? "Close window" : "Back to saving"}" title="${options.floating ? "Close window" : "Back to saving"}">×</button>
    </header>
    <section data-el="source-tools">
      <div class="desk-document-actions"><button data-el="upload-document" type="button"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 16V3m-5 5 5-5 5 5M4 16v4h16v-4"/></svg><span data-el="upload-label">Upload document</span></button><input data-el="upload-file" type="file" accept=".docx,.pdf,.txt,.md,.csv,.xlsx,.rtf" aria-label="Choose a document to add to this matter" hidden></div>
      <p class="desk-meta desk-upload-status" data-el="upload-status" role="status" hidden></p>
      <div class="desk-document-list" data-el="document-list"></div>
      <p class="desk-meta desk-update-status" data-el="update-status" role="status" hidden></p>
    </section>
    <section data-el="compare-tools" hidden><div class="desk-compare-controls"><label>Earlier<select data-el="base" aria-label="Earlier version"></select></label><span aria-hidden="true">→</span><label>Later<select data-el="current" aria-label="Later version"></select></label><button type="button" data-el="compare" hidden>Compare</button><button type="button" class="desk-icon" data-el="reload-versions" aria-label="Refresh version list" title="Refresh version list">↻</button></div><p class="desk-warning" data-el="compare-warning" hidden></p></section>
    <div class="desk-columns"><section class="desk-reader" aria-label="Saved document" data-el="right"></section></div>
    <p class="desk-error" role="alert" data-el="status" hidden></p>`
  const el = <T extends HTMLElement>(name: string) => root.querySelector(`[data-el="${name}"]`) as T
  const button = (name: string) => el<HTMLButtonElement>(name)
  const select = (name: string) => el<HTMLSelectElement>(name)
  const node = <K extends keyof HTMLElementTagNameMap>(tag: K, text = "", className?: string) => {
    const result = document.createElement(tag)
    result.textContent = text
    if (className) result.className = className
    return result
  }
  const action = (name: string, label: string, run: () => void, icon?: string) => {
    const result = node("button", icon ?? label, icon ? "desk-icon" : undefined)
    result.type = "button"; result.dataset.el = name; result.onclick = run
    if (icon) { result.setAttribute("aria-label", label); result.title = label }
    return result
  }
  const date = (value?: number | string) => value === undefined ? "" : new Date(value).toLocaleString([], { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit", hourCycle: "h23" })
  const originLabel: Record<string, string> = {
    word: "Saved from Word", upload: "Uploaded", replace: "Replaced", generated: "Created in doc.haus",
    revision: "Revised in doc.haus", "agent-edit": "Edited in doc.haus", redaction: "Redacted",
    accept: "Changes accepted", conversion: "Converted", copy: "Copied", restore: "Restored",
    system: "Saved by doc.haus", unknown: "Saved in doc.haus",
  }
  let disposed = false
  let pending = 0
  let uploading = false
  let mode: DeskView = options.initialView === "compare" ? "compare" : "documents"
  let sources: DeskSources | undefined
  let versions: DeskVersions | undefined
  let sourcesJob: Promise<void> | undefined
  let versionsJob: Promise<void> | undefined
  let comparisonJob: Promise<void> | undefined
  const updates = new Map<string, "New" | "Updated">()
  const schedulePoll = options.schedulePoll ?? ((run: () => void, delay: number) => { const timer = setTimeout(run, delay); return () => clearTimeout(timer) })
  let cancelPoll: (() => void) | undefined
  const now = options.now ?? Date.now
  type Handoff = { key: string; uri?: string; at?: number; retryAt?: number; error?: boolean; openingUntil?: number }
  const openLinks = new Map<string, Handoff>()
  const visible = new Set<Element>()
  let preparing = false
  const observer = typeof window.IntersectionObserver === "function" ? new window.IntersectionObserver(entries => {
    for (const entry of entries) { if (entry.isIntersecting) visible.add(entry.target); else visible.delete(entry.target) }
    void prepareOpenLinks()
  }, { rootMargin: "80px" }) : undefined
  let versionsRefreshed: number | undefined
  let comparison: DeskComparison | undefined
  let desiredPair = ""
  let retryPair = ""
  let change = 0
  const comparisons = new Map<string, DeskComparison>()
  const busy = () => pending > 0
  function error(message = "") {
    if (disposed) return
    el("status").textContent = message
    el("status").hidden = !message
  }
  function context() {
    if (disposed) return
    const documentName = sources?.documentName ?? versions?.document.name
    const matterName = sources?.matterName ?? versions?.matterName
    el("identity").textContent = [documentName, matterName].filter(Boolean).join(" · ")
    options.onContextChange?.({ view: mode, documentName, matterName })
  }
  function chosenPair() {
    const earlier = versions?.versions.find((version) => version.id === select("base").value && !version.removed)
    const later = versions?.versions.find((version) => version.id === select("current").value && !version.removed)
    return earlier && later && earlier.number < later.number ? { earlier, later, key: `${earlier.id}:${later.id}` } : undefined
  }
  function automaticPair(pair: NonNullable<ReturnType<typeof chosenPair>>) {
    return typeof pair.earlier.bytes === "number" && typeof pair.later.bytes === "number" && pair.earlier.bytes >= 0 && pair.later.bytes >= 0 && pair.earlier.bytes + pair.later.bytes <= 2 * 1024 * 1024
  }
  function controls() {
    if (disposed) return
    root.setAttribute("aria-busy", String(busy()))
    for (const control of root.querySelectorAll<HTMLButtonElement | HTMLSelectElement>("button, select")) {
      if (control.dataset.el !== "close") control.disabled = busy()
    }
    el<HTMLInputElement>("upload-file").disabled = busy()
    el("upload-label").textContent = uploading ? "Uploading…" : "Upload document"
    const count = comparison?.changes.length ?? 0
    const previous = root.querySelector<HTMLButtonElement>('[data-el="previous"]')
    const next = root.querySelector<HTMLButtonElement>('[data-el="next"]')
    if (previous) previous.disabled = busy() || change <= 0 || !count
    if (next) next.disabled = busy() || change >= count - 1 || !count
    const earlier = versions?.versions.find((version) => version.id === select("base").value)
    const later = versions?.versions.find((version) => version.id === select("current").value)
    const invalid = !!earlier && !!later && earlier.number >= later.number
    const pair = chosenPair()
    button("compare").hidden = !pair || (automaticPair(pair) && retryPair !== pair.key)
    button("compare").disabled = busy() || !pair
    el("compare-warning").hidden = !invalid
    el("compare-warning").textContent = invalid ? "Choose a later version with a higher version number." : ""
  }
  async function work(run: () => Promise<void>) {
    if (disposed) return
    if (pending++ === 0) options.onBusyChange?.(true)
    error(); controls()
    try { await run() } catch (cause) { error(cause instanceof Error ? cause.message : "This action could not be completed.") }
    finally {
      pending--
      if (!disposed) { controls(); if (!pending) options.onBusyChange?.(false) }
    }
  }
  function renderList() {
    const list = el("document-list")
    if (!sources) { if (sourcesJob && !list.firstChild) list.append(node("p", "Loading documents…", "desk-empty")); return }
    for (const note of list.querySelectorAll(".desk-empty, .desk-warning")) note.remove()
    // Replacing a matter file may change its registry ID. Its unique filename
    // keeps the row stable; the request still uses the latest ID and hash.
    const rows = new Map(Array.from(list.querySelectorAll<HTMLAnchorElement>("[data-source-name]"), row => [row.dataset.sourceName!, row]))
    const retained = new Set(sources.sources.map(source => source.name))
    for (const [name, row] of rows) if (!retained.has(name)) { observer?.unobserve(row); visible.delete(row); openLinks.delete(name); row.remove() }
    if (!sources.sources.length) list.append(node("p", "No saved documents are available.", "desk-empty"))
    for (const [index, source] of sources.sources.entries()) {
      const row = rows.get(source.name) ?? node("a")
      if (!row.dataset.sourceName) {
        row.dataset.sourceName = source.name; row.className = "desk-document-choice"; row.target = "_blank"; row.rel = "noopener noreferrer"; row.tabIndex = 0
        row.setAttribute("role", "link")
        row.append(node("span"), node("span", "", "desk-meta"), node("span", "", "desk-open-status"))
        observer?.observe(row)
      }
      if (row.dataset.source !== source.id) row.dataset.source = source.id
      row.onclick = event => activateDocument(event, row, source)
      row.onkeydown = event => { if (event.key === "Enter" && !row.hasAttribute("href")) activateDocument(event, row, source) }
      const name = row.children[0]!, meta = row.children[1]!
      if (name.textContent !== source.name) name.textContent = source.name
      const label = [`Version ${source.version?.number ?? "unavailable"}`, date(source.version?.createdAt)].filter(Boolean).join(" · ")
      if (meta.textContent !== label) meta.textContent = label
      const update = updates.get(source.name)
      let badge = row.querySelector(".desk-update-badge")
      if (update && !badge) { badge = node("span", update, "desk-update-badge"); row.append(badge) }
      else if (!update && badge) badge.remove()
      else if (badge && badge.textContent !== update) badge.textContent = update!
      row.classList.toggle("has-update", Boolean(update))
      updateOpenLink(row, source)
      const at = list.children[index]
      if (at !== row) list.insertBefore(row, at ?? null)
    }
    if (sources.truncated) list.append(node("p", "Some documents are unavailable or exceed the list limit.", "desk-warning"))
    void prepareOpenLinks()
  }
  const sourceKey = (source: DeskSource) => `${source.id}:${source.hash}`
  function patchLink(row: HTMLAnchorElement, href: string | undefined, label: string) {
    if (href) { if (row.getAttribute("href") !== href) row.href = href }
    else if (row.hasAttribute("href")) row.removeAttribute("href")
    const disabled = String(!href)
    if (row.getAttribute("aria-disabled") !== disabled) row.setAttribute("aria-disabled", disabled)
    const state = row.querySelector(".desk-open-status")!
    if (state.textContent !== label) state.textContent = label
  }
  function updateOpenLink(row: HTMLAnchorElement, source: DeskSource) {
    if (!/\.docx$/i.test(source.name)) {
      let href: string | undefined
      try { const value = new URL(sources?.matterUrl ?? ""); if (/^https?:$/.test(value.protocol) && !value.username && !value.password) href = value.href } catch { /* The parent supplies a validated matter URL. */ }
      patchLink(row, href, href ? "Open in doc.haus" : "Use Open matter in doc.haus below")
      return
    }
    let state = openLinks.get(source.name)
    if (!state || state.key !== sourceKey(source)) { state = { key: sourceKey(source) }; openLinks.set(source.name, state) }
    const ready = state.uri && state.at !== undefined && now() - state.at < 45_000
    patchLink(row, ready ? state.uri : undefined, state.openingUntil && state.openingUntil > now() ? "Opening in Word…" : state.error ? "Could not prepare Word. Retrying…" : ready ? "Open in Word" : "Preparing Word…")
  }
  function visibleSources() {
    const rows = Array.from(el("document-list").querySelectorAll<HTMLAnchorElement>("[data-source-name]"))
    const available = observer ? rows.filter(row => visible.has(row)) : rows.filter(row => { const bounds = row.getBoundingClientRect(); return !bounds.height || (bounds.bottom >= 0 && bounds.top < window.innerHeight + 80) })
    const names = new Set(available.slice(0, 12).map(row => row.dataset.sourceName))
    return (sources?.sources ?? []).filter(source => /\.docx$/i.test(source.name) && names.has(source.name))
  }
  async function prepareOpenLinks() {
    if (preparing || !pollingAllowed() || !sources) return
    preparing = true
    try {
      for (const source of visibleSources()) {
        if (!pollingAllowed()) break
        const row = Array.from(el("document-list").querySelectorAll<HTMLAnchorElement>("[data-source-name]")).find(item => item.dataset.sourceName === source.name)
        if (!row) continue
        updateOpenLink(row, source)
        const state = openLinks.get(source.name)!
        if ((state.at !== undefined && now() - state.at < 30_000) || (state.retryAt && now() < state.retryAt)) continue
        const started = now()
        try {
          const handoff = await transport.request({ action: "open", sourceId: source.id, expectedHash: source.hash }) as { uri: string }
          if (disposed || openLinks.get(source.name) !== state || sources?.sources.find(item => item.name === source.name)?.id !== source.id || sources?.sources.find(item => item.name === source.name)?.hash !== source.hash) continue
          if (typeof handoff?.uri !== "string" || !handoff.uri.startsWith("doc.haus://word/open?") || now() - started >= 45_000) throw new Error("Expired Word handoff")
          state.uri = handoff.uri; state.at = started; state.error = false; state.retryAt = undefined
        } catch {
          if (disposed || openLinks.get(source.name) !== state) continue
          state.uri = undefined; state.at = undefined; state.error = true; state.retryAt = now() + 5_000
        }
        if (!disposed) updateOpenLink(row, source)
      }
    } finally { preparing = false }
  }
  function activateDocument(event: Event, row: HTMLAnchorElement, source: DeskSource) {
    if (!/\.docx$/i.test(source.name)) { if (!row.hasAttribute("href")) event.preventDefault(); return }
    updateOpenLink(row, source)
    const state = openLinks.get(source.name)!
    if (!row.hasAttribute("href")) { event.preventDefault(); state.retryAt = undefined; void prepareOpenLinks(); return }
    updates.delete(source.name)
    row.querySelector(".desk-update-badge")?.remove(); row.classList.remove("has-update")
    state.openingUntil = now() + 5_000
    // Clear the single-use link only after the real anchor's default navigation.
    setTimeout(() => {
      if (disposed || openLinks.get(source.name) !== state) return
      state.uri = undefined; state.at = undefined; updateOpenLink(row, source); void prepareOpenLinks()
    }, 0)
  }
  function details(area: HTMLElement, lines: string[]) {
    const detail = node("details", "", "desk-details")
    detail.append(node("summary", "Details"))
    for (const line of lines) detail.append(node("p", line))
    area.append(detail)
    return detail
  }
  function render() {
    if (disposed) return
    el("source-tools").hidden = mode !== "documents"
    el("compare-tools").hidden = mode !== "compare"
    button("sources-mode").setAttribute("aria-pressed", String(mode === "documents"))
    button("compare-mode").setAttribute("aria-pressed", String(mode === "compare"))
    renderList()
    const right = el("right")
    right.replaceChildren()
    const textBlock = (text: string) => {
      const block = node("div", "", "desk-text")
      block.tabIndex = 0
      for (const paragraph of text.split(/\n+/)) { const p = node("p"); p.textContent = paragraph; block.append(p) }
      return block
    }
    if (mode === "documents") {
      right.hidden = true
    } else {
      right.hidden = false
      const paper = node("article", "", "desk-paper")
      if (comparison) {
        const hunk = comparison.changes[change]
        const head = node("div", "", "desk-change-head")
        const heading = node("div")
        const scopeLabel = hunk && ({ body: "Text", header: "Header", footer: "Footer", footnote: "Footnote", endnote: "Endnote" }[hunk.location.scope])
        heading.append(node("h3", hunk?.location.heading ?? (hunk ? `${scopeLabel} change` : "No text changes")), node("p", `Version ${comparison.base.number} → Version ${comparison.current.number} · Saved versions · Read-only`, "desk-meta"))
        if (hunk?.untracked) heading.append(node("span", "Untracked edit", "desk-tag"))
        const nav = node("div", "", "desk-change-nav")
        nav.append(action("previous", "Previous change", () => moveChange(-1), "‹"), node("span", comparison.changes.length ? `${change + 1} of ${comparison.changes.length}` : "0 changes"), action("next", "Next change", () => moveChange(1), "›"))
        head.append(heading, nav); paper.append(head)
        const counts = [comparison.untracked.length ? `${comparison.untracked.length} untracked ${comparison.untracked.length === 1 ? "edit" : "edits"}` : "", comparison.proposals.length ? `${comparison.proposals.length} proposal ${comparison.proposals.length === 1 ? "outcome" : "outcomes"}` : ""].filter(Boolean)
        if (counts.length) paper.append(node("p", counts.join(" · "), "desk-meta"))
        if (comparison.truncated) paper.append(node("p", "Partial comparison · open Versions in doc.haus for the full review.", "desk-warning"))
        if (hunk) {
          const redline = node("div", "", "desk-text compare-redline")
          redline.tabIndex = 0
          const pieces: Piece[] = hunk.kind === "moved" && hunk.before === hunk.after ? [{ kind: "move", text: hunk.after ?? "" }] : wordDiff(hunk.before ?? "", hunk.after ?? "")
          for (const piece of pieces) {
            if (piece.kind === "same") { redline.append(document.createTextNode(piece.text)); continue }
            const mark = node(piece.kind === "del" ? "del" : "ins", "", piece.kind === "move" ? "is-move" : undefined)
            mark.append(node("span", `${piece.kind === "del" ? "Deleted" : piece.kind === "move" ? "Moved" : "Inserted"}: `, "sr-only"), document.createTextNode(piece.text)); redline.append(mark)
          }
          paper.append(redline)
          const legend = node("div", "", "desk-legend")
          legend.append(node("span", "Inserted", "legend-inserted"), node("span", "Deleted", "legend-deleted"), node("span", "Moved", "legend-moved")); paper.append(legend)
        }
        const detail = details(paper, [])
        detail.classList.add("desk-comparison-details")
        const pair = node("div", "", "desk-version-pair")
        for (const [index, side] of [comparison.base, comparison.current].entries()) {
          const section = node("section", "", "desk-version-section")
          const title = `${index === 0 ? "Earlier" : "Later"} · Version ${side.number}`
          section.setAttribute("aria-label", title)
          section.append(node("h4", title))
          const meta = node("p", "", "desk-meta")
          const time = node("time", date(side.createdAt))
          time.dateTime = new Date(side.createdAt).toISOString()
          meta.append(time, document.createTextNode(` · ${originLabel[side.origin] ?? "Saved in doc.haus"}`))
          section.append(meta)
          if (side.trackedChanges.count) section.append(node("p", `${side.trackedChanges.count} tracked ${side.trackedChanges.count === 1 ? "change" : "changes"} in this version`, "desk-meta"))
          if (hunk) section.append(textBlock((index === 0 ? hunk.before : hunk.after) || "No text."))
          pair.append(section)
        }
        detail.append(pair)
        const technical = node("details", "", "desk-technical")
        technical.append(node("summary", "Technical details"))
        if (hunk) technical.append(node("p", `${hunk.kind} · ${hunk.location.scope} paragraph ${hunk.location.paragraph} in the ${hunk.location.side === "left" ? "earlier" : "later"} version`, "desk-meta"))
        for (const side of [comparison.base, comparison.current]) {
          technical.append(node("h5", `Version ${side.number}`))
          const fields = node("dl")
          for (const [label, value] of [["Version ID", side.id], ["SHA-256", side.hash]]) {
            const valueElement = node("dd")
            valueElement.append(node("code", value))
            fields.append(node("dt", label), valueElement)
          }
          technical.append(fields)
        }
        if (versionsRefreshed) technical.append(node("p", `Version list checked ${date(versionsRefreshed)}. Comparing the selected saved versions.`, "desk-meta"))
        detail.append(technical)
      } else if (versions && versions.versions.filter((version) => !version.removed).length < 2) paper.append(node("p", "Two saved versions are needed to compare.", "desk-empty"))
      right.append(paper)
    }
    context(); controls()
  }
  function moveChange(direction: -1 | 1) {
    if (busy() || !comparison) return
    change = Math.max(0, Math.min(comparison.changes.length - 1, change + direction)); render()
    const preferred = button(direction === 1 ? "next" : "previous")
    ;(preferred.disabled ? button(direction === 1 ? "previous" : "next") : preferred).focus()
  }
  function pollingAllowed() { return !disposed && mode === "documents" && document.visibilityState !== "hidden" }
  function scheduleNextPoll() {
    cancelPoll?.(); cancelPoll = undefined
    if (!pollingAllowed()) return
    cancelPoll = schedulePoll(() => {
      cancelPoll = undefined
      if (!pollingAllowed()) return
      if (busy()) { scheduleNextPoll(); return }
      void loadSources(true).finally(scheduleNextPoll)
    }, 1_000)
  }
  async function loadSources(background = false) {
    if (sourcesJob) return sourcesJob
    const read = async () => {
      const value = await transport.request({ action: "sources" }) as DeskSources
      if (disposed) return
      const changed = JSON.stringify(value) !== JSON.stringify(sources)
      const identityChanged = value.documentName !== sources?.documentName || value.matterName !== sources?.matterName
      const initial = !sources
      if (sources && changed) for (const source of value.sources) {
        const previous = sources.sources.find(item => item.name === source.name)
        if (!previous) updates.set(source.name, "New")
        else if (JSON.stringify(previous) !== JSON.stringify(source)) updates.set(source.name, "Updated")
      }
      sources = value
      if (background && initial) error()
      const status = el("update-status")
      if (!status.hidden) status.hidden = true
      if (!background) render()
      else if (changed) { renderList(); if (identityChanged) context() }
      void prepareOpenLinks()
    }
    sourcesJob = background ? read().catch(() => {
      if (disposed) return
      const status = el("update-status")
      const message = "Document updates paused. Reconnecting automatically…"
      if (status.textContent !== message) status.textContent = message
      if (status.hidden) status.hidden = false
    }) : work(read)
    if (!background) renderList()
    try { await sourcesJob } finally { sourcesJob = undefined; if (!disposed && !background) renderList() }
  }
  async function loadVersions(autoCompare: boolean) {
    if (versionsJob) return versionsJob
    versionsJob = work(async () => {
      const value = await transport.request({ action: "versions" }) as DeskVersions
      if (disposed) return
      versions = value; versionsRefreshed = Date.now()
      const available = versions.versions.filter((version) => !version.removed).sort((a, b) => b.number - a.number)
      for (const name of ["base", "current"]) {
        const control = select(name), previous = control.value
        control.replaceChildren()
        for (const version of versions.versions) {
          const option = new Option(`v${version.number} · ${date(version.createdAt)}${version.current ? " · latest known" : ""}${version.removed ? " · removed" : ""}${version.label ? ` · ${version.label}` : ""}`, version.id)
          option.disabled = version.removed; control.add(option)
        }
        control.value = available.some((version) => version.id === previous) ? previous : available[name === "base" ? 1 : 0]?.id ?? ""
      }
      render()
    })
    try { await versionsJob } finally { versionsJob = undefined }
    if (!disposed && autoCompare && !comparison) await compareSelected()
  }
  async function compareSelected(manual = false): Promise<void> {
    if (disposed) return
    const pair = chosenPair()
    desiredPair = pair?.key ?? ""
    if (!pair) { controls(); return }
    const cached = comparisons.get(pair.key)
    if (cached) { comparison = cached; change = 0; retryPair = ""; render(); return }
    if (!manual && !automaticPair(pair)) { controls(); return }
    if (comparisonJob) return comparisonJob
    const requested = pair.key
    comparisonJob = work(async () => {
      let value: DeskComparison
      try { value = await transport.request({ action: "compare", baseVersion: pair.earlier.id, currentVersion: pair.later.id }) as DeskComparison }
      catch (cause) { if (desiredPair !== requested) return; retryPair = requested; throw cause }
      if (disposed) return
      comparisons.set(requested, value)
      if (comparisons.size > 8) comparisons.delete(comparisons.keys().next().value!)
      if (desiredPair !== requested || chosenPair()?.key !== requested) return
      comparison = value; change = 0; retryPair = ""; render()
    })
    try { await comparisonJob } finally { comparisonJob = undefined }
    if (!disposed && desiredPair && desiredPair !== requested && chosenPair()?.key === desiredPair) await compareSelected()
  }
  button("close").onclick = options.close
  button("reload-versions").onclick = () => { if (!busy()) void loadVersions(false) }
  select("base").onchange = select("current").onchange = () => { void compareSelected() }
  button("compare").onclick = () => { if (!busy()) void compareSelected(true) }
  button("sources-mode").onclick = () => { void setView("documents") }
  button("compare-mode").onclick = () => { void setView("compare") }
  const uploadFile = el<HTMLInputElement>("upload-file")
  button("upload-document").onclick = () => { if (!busy() && !disposed) uploadFile.click() }
  uploadFile.onchange = () => {
    const file = uploadFile.files?.[0]
    if (!file || busy() || disposed) return
    uploadFile.value = ""
    el("upload-status").hidden = true
    if (!isChatUploadName(file.name) || file.size < 1 || file.size > CHAT_UPLOAD_MAX_BYTES) {
      error("Choose a Word, PDF, text, spreadsheet or RTF file up to 16 MB."); return
    }
    uploading = true
    void work(async () => {
      const bytes = new Uint8Array(await file.arrayBuffer())
      if (disposed) return
      const chunks: string[] = []
      for (let start = 0; start < bytes.length; start += 32_768)
        chunks.push(String.fromCharCode(...bytes.subarray(start, start + 32_768)))
      const result = await transport.request({ action: "document-upload", name: file.name, data: btoa(chunks.join("")) }) as { name: string }
      if (disposed) return
      // Finish any earlier background read before refreshing the saved list.
      await sourcesJob
      await loadSources()
      if (disposed) return
      el("upload-status").textContent = `${result.name} added to this matter.`
      el("upload-status").hidden = false
    }).finally(() => { uploading = false; if (!disposed) controls() })
  }
  async function setView(view: DeskView) {
    if (disposed) return
    const next = view === "compare" ? "compare" : "documents"
    if (mode !== next) { mode = next; error(); options.onViewChange?.(mode) }
    render()
    if (mode === "documents" && !sources) await loadSources()
    if (mode === "compare" && !versions) await loadVersions(true)
    scheduleNextPoll()
  }
  const visibilityChanged = () => scheduleNextPoll()
  const viewportChanged = () => { void prepareOpenLinks() }
  document.addEventListener("visibilitychange", visibilityChanged)
  document.addEventListener("scroll", viewportChanged, true)
  window.addEventListener("resize", viewportChanged)
  render()
  const ready = mode === "compare" ? loadVersions(true) : loadSources()
  void ready.then(scheduleNextPoll)
  return { ready, setView, dispose() {
    if (disposed) return
    disposed = true; cancelPoll?.(); document.removeEventListener("visibilitychange", visibilityChanged)
    document.removeEventListener("scroll", viewportChanged, true); window.removeEventListener("resize", viewportChanged)
    desiredPair = ""; comparisons.clear(); openLinks.clear(); observer?.disconnect()
    if (pending) options.onBusyChange?.(false)
    root.replaceChildren(); root.removeAttribute("aria-busy"); root.classList.remove("source-desk", "is-embedded")
  } }
}
