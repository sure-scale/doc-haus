import type { DeskComparison, DeskSourcePreview, DeskSources, DeskVersions } from "../../../dochaus/lib/word-source-desk-contract"
import { wordDiff, type Piece } from "../../../dochaus/lib/redline"
import type { DeskRequest } from "./desk-protocol"
import type { WordSelectionSnapshot } from "./selection"

export type DeskTransport = { request(input: DeskRequest): Promise<unknown> }

/** Open previews and kept sources retain their exact received saved snapshot. */
export function mountDesk(root: HTMLElement, transport: DeskTransport, options: { close(): void; floating?: boolean }) {
  root.classList.add("source-desk")
  root.innerHTML = `
    <header class="desk-head"><div class="desk-brand"><img src="./assets/icon.png" width="32" height="32" alt="" />
      <div><h2>Source Desk</h2><p class="desk-subtitle" data-el="identity">Connecting to this matter…</p></div></div>
      <button type="button" data-el="close">${options.floating ? "Close desk" : "Back to saving"}</button></header>
    <nav class="desk-modes" aria-label="Desk views"><button class="desk-mode" data-el="sources-mode" aria-pressed="true" type="button">Sources</button>
      <button class="desk-mode" data-el="compare-mode" aria-pressed="false" type="button">Compare versions</button></nav>
    <div data-el="source-tools"><div class="desk-toolbar">
      <label>Open a source<select data-el="source-choice" aria-label="Open a source"><option value="">Loading sources…</option></select></label>
      <button type="button" data-el="pin" disabled>Keep open</button><button type="button" data-el="reload">Refresh list</button>
      <button type="button" data-el="find-toggle" aria-expanded="false">Find</button></div>
      <div class="desk-find" data-el="find-bar" hidden><label>Find exact text<input data-el="find" type="search" placeholder="Find in these passages" maxlength="255" /></label>
        <span data-el="find-count" role="status"></span><button type="button" data-el="find-close">Close find</button></div>
      <div class="desk-pins" data-el="pins" role="group" aria-label="Kept sources"></div></div>
    <div data-el="compare-tools" hidden><div class="desk-toolbar">
      <label>Earlier version<select data-el="base" aria-label="Earlier version"></select></label><label>Later version<select data-el="current" aria-label="Later version"></select></label>
      <button type="button" data-el="compare" class="desk-primary" disabled>Compare</button><button type="button" data-el="reload-versions">Refresh versions</button></div><p class="desk-warning" data-el="compare-warning" hidden></p></div>
    <div class="desk-columns" data-el="columns"><section class="desk-column desk-context" aria-label="Word selection" data-el="left"></section>
      <section class="desk-column desk-reader" aria-label="Source or comparison" data-el="right"></section></div>
    <footer class="desk-footer"><p role="status" aria-live="polite" data-el="status">Opening Source Desk…</p><p class="desk-note" data-el="scope">Read-only · saved source previews</p></footer>`
  const el = <T extends HTMLElement>(name: string) => root.querySelector(`[data-el="${name}"]`) as T
  const button = (name: string) => el<HTMLButtonElement>(name)
  const select = (name: string) => el<HTMLSelectElement>(name)
  let disposed = false
  let busy = false
  let mode: "sources" | "compare" = "sources"
  let findOpen = false
  let selection: WordSelectionSnapshot | undefined
  let sources: DeskSources | undefined
  let versions: DeskVersions | undefined
  const pins = new Map<string, DeskSourcePreview>()
  let active: DeskSourcePreview | undefined
  let comparison: DeskComparison | undefined
  let change = 0
  const key = (source: DeskSourcePreview) => `${source.id}:${source.hash}`
  const say = (text: string) => { if (!disposed) el("status").textContent = text }
  const node = <K extends keyof HTMLElementTagNameMap>(tag: K, text: string, className?: string) => {
    const result = document.createElement(tag)
    result.textContent = text
    if (className) result.className = className
    return result
  }
  const action = (name: string, label: string, run: () => void) => {
    const result = node("button", label)
    result.type = "button"; result.dataset.el = name; result.onclick = run
    return result
  }
  const textBlock = (text: string, query = "") => {
    const result = node("div", "", "desk-text")
    result.tabIndex = 0
    if (!query) { result.textContent = text; return result }
    let from = 0
    let count = 0
    while (count++ < 200) {
      const at = text.indexOf(query, from)
      if (at < 0) break
      result.append(document.createTextNode(text.slice(from, at)), node("mark", query))
      from = at + query.length
    }
    result.append(document.createTextNode(text.slice(from)))
    return result
  }
  const column = (name: string, label: string, title: string) => {
    const area = el(name)
    area.replaceChildren(node("p", label, "desk-eyebrow"), node("h3", title))
    return area
  }
  const metadata = (area: HTMLElement, label: string, lines: string[], hash?: string) => {
    const block = node("details", "", "desk-provenance")
    block.append(node("summary", label))
    for (const line of lines) block.append(node("p", line))
    if (hash) block.append(node("code", `SHA-256 · ${hash}`))
    area.append(block)
  }
  function controls() {
    for (const control of root.querySelectorAll<HTMLButtonElement | HTMLSelectElement>("button, select")) {
      if (!["close", "sources-mode", "compare-mode", "find-toggle", "find-close"].includes(control.dataset.el ?? "")) control.disabled = busy
    }
    button("pin").disabled = busy || !active || pins.has(key(active)) || pins.size >= 8
    button("pin").textContent = active && pins.has(key(active)) ? "Kept open" : "Keep open"
    const earlier = versions?.versions.find((version) => version.id === select("base").value)
    const later = versions?.versions.find((version) => version.id === select("current").value)
    const reversed = !!earlier && !!later && earlier.number >= later.number
    button("compare").disabled = busy || !earlier || !later || reversed
    el("compare-warning").hidden = !reversed
    el("compare-warning").textContent = reversed ? "Choose a later version with a higher version number." : ""
    const show = root.querySelector<HTMLButtonElement>('[data-el="show"]')
    if (show) show.disabled = busy || !selection?.navigable
    const count = comparison?.changes.length ?? 0
    const previous = root.querySelector<HTMLButtonElement>('[data-el="previous"]')
    const next = root.querySelector<HTMLButtonElement>('[data-el="next"]')
    if (previous) previous.disabled = busy || !count || change <= 0
    if (next) next.disabled = busy || !count || change >= count - 1
  }
  function renderPins() {
    el("pins").replaceChildren()
    el("pins").hidden = !pins.size
    for (const [id, source] of pins) {
      const group = node("div", "", "desk-pin")
      const pick = node("button", `${source.name} · v${source.version?.number ?? "?"}`)
      pick.type = "button"; pick.dataset.pin = id
      pick.setAttribute("aria-pressed", String(active && key(active) === id))
      pick.onclick = () => {
        active = source; select("source-choice").value = source.id; renderPins(); render(); say("Kept source opened at its saved version.")
        Array.from(el("pins").querySelectorAll<HTMLButtonElement>("button")).find((item) => item.dataset.pin === id)?.focus()
      }
      const remove = node("button", "×", "desk-unpin")
      remove.type = "button"; remove.setAttribute("aria-label", `Stop keeping ${source.name} open`)
      remove.onclick = () => { pins.delete(id); renderPins(); render(); button("pin").focus(); say("Source removed from kept sources. Its preview is still open.") }
      group.append(pick, remove); el("pins").append(group)
    }
    controls()
  }
  function render() {
    if (disposed) return
    el("source-tools").hidden = mode !== "sources"
    el("compare-tools").hidden = mode !== "compare"
    el("left").hidden = mode !== "sources"
    el("columns").classList.toggle("is-compare", mode === "compare")
    button("sources-mode").setAttribute("aria-pressed", String(mode === "sources"))
    button("compare-mode").setAttribute("aria-pressed", String(mode === "compare"))
    el("find-bar").hidden = !findOpen
    button("find-toggle").setAttribute("aria-expanded", String(findOpen))
    if (mode === "sources") {
      const query = findOpen ? el<HTMLInputElement>("find").value : ""
      const left = column("left", "In Word", "Selected passage")
      left.append(node("p", selection ? (selection.saved ? "Saved in Word" : "Unsaved in Word") : "Select text in your Word document.", selection && !selection.saved ? "badge is-unsaved" : "badge"))
      const actions = node("div", "", "desk-actions")
      actions.append(action("capture", "Use current selection", () => { void perform(captureSelection).then(() => { root.querySelector<HTMLButtonElement>('[data-el="capture"]')?.focus() }) }), action("show", "Show in Word", () => { void perform(async () => {
        if (!selection?.navigable) return
        await transport.request({ action: "show", selection })
        say("The captured passage is selected in Word.")
      }) }))
      left.append(actions)
      if (selection) {
        left.append(textBlock(selection.text || "No text is selected.", query))
        left.append(node("p", `Captured at ${new Date(selection.capturedAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })} · use current selection to update`, "desk-note"))
        if (selection.reason) left.append(node("p", selection.reason, "desk-warning"))
        metadata(left, "Selection details", ["Live Word context, separate from saved matter versions.", "Show in Word requires one exact, unchanged body passage. Formatting and markup are not checked."], selection.documentHash)
      }
      const right = column("right", active ? "Saved matter source" : "Matter sources", active?.name ?? "Open a source to start reading")
      if (active) {
        right.append(node("p", `Version ${active.version?.number ?? "unavailable"}${pins.has(key(active)) ? " · kept open" : ""}`, "desk-version"))
        if (active.truncated) right.append(node("p", `Partial preview · showing ${active.text.length.toLocaleString()} of ${active.characters.toLocaleString()} characters.`, "desk-warning"))
        right.append(textBlock(active.text || "This saved source has no previewable text.", query))
        metadata(right, "Source details", [`${active.role === "deliverable" ? "Matter output" : "Matter source"} · saved version ${active.version?.number ?? "unavailable"}`, `Version ID · ${active.version?.id ?? "unavailable"}`, `${active.characters.toLocaleString()} characters`, "This snapshot stays unchanged when your Word selection or the source list changes."], active.hash)
      } else right.append(node("p", "Choose a document above. Keep useful sources open to switch between them while you work in Word.", "desk-empty"))
      const matches = root.querySelectorAll("mark").length
      el("find-count").textContent = query ? `${matches >= 200 ? "At least " : ""}${matches} exact ${matches === 1 ? "match" : "matches"}` : "In the selection and source"
      el("scope").textContent = "Read-only · saved source previews"
    } else {
      const hunk = comparison?.changes[change]
      const right = column("right", "Saved version comparison", comparison ? comparison.current.name : "Compare saved versions")
      if (comparison) {
        right.append(node("p", `Version ${comparison.base.number} → Version ${comparison.current.number}`, "desk-compare-heading"))
        const nav = node("div", "", "desk-change-nav")
        nav.append(action("previous", "← Previous", () => moveChange(-1)), node("span", comparison.changes.length ? `Change ${change + 1} of ${comparison.changes.length}` : "No text changes"), action("next", "Next →", () => moveChange(1)))
        right.append(nav)
        if (comparison.truncated) right.append(node("p", "Partial comparison · open Versions in doc.haus for the full review.", "desk-warning"))
        if (hunk) {
          const legend = node("div", "", "desk-legend")
          legend.append(node("ins", "Inserted"), node("del", "Deleted"), node("ins", "Moved", "is-move"))
          right.append(legend, node("p", `${hunk.kind[0].toUpperCase() + hunk.kind.slice(1)} · ${hunk.location.heading ? `${hunk.location.heading} · ` : ""}${hunk.location.scope} paragraph ${hunk.location.paragraph} in ${hunk.location.side === "left" ? "earlier" : "later"} version${hunk.untracked ? " · edit made without tracking" : ""}`, "desk-change-label"))
          const pieces: Piece[] = hunk.kind === "moved" && hunk.before === hunk.after ? [{ kind: "move", text: hunk.after ?? "" }] : wordDiff(hunk.before ?? "", hunk.after ?? "")
          const redline = node("div", "", "desk-text compare-redline")
          redline.tabIndex = 0
          for (const piece of pieces) {
            if (piece.kind === "same") { redline.append(document.createTextNode(piece.text)); continue }
            const mark = node(piece.kind === "del" ? "del" : "ins", "", piece.kind === "move" ? "is-move" : undefined)
            mark.append(node("span", `${piece.kind === "del" ? "Deleted" : piece.kind === "move" ? "Moved" : "Inserted"}: `, "sr-only"), document.createTextNode(piece.text))
            redline.append(mark)
          }
          right.append(redline)
          const plain = node("details", "", "desk-before-after")
          plain.append(node("summary", "Read earlier and later text separately"), node("h4", `Earlier · version ${comparison.base.number}`), textBlock(hunk.before || "No text in the earlier version."), node("h4", `Later · version ${comparison.current.number}`), textBlock(hunk.after || "No text in the later version."))
          right.append(plain)
        } else right.append(node("p", "No textual changes were reported between these saved versions.", "desk-empty"))
        for (const side of [comparison.base, comparison.current]) metadata(right, `Version ${side.number} details`, [`Saved ${new Date(side.createdAt).toLocaleString()} · ${side.origin}`, `Version ID · ${side.id}`, `${side.trackedChanges.count} input tracked changes`], side.hash)
        el("scope").textContent = `Saved versions only · ${comparison.untracked.length} edits without tracking · ${comparison.proposals.length} earlier proposal outcomes`
      } else {
        right.append(node("p", "Choose two versions, then Compare to see their changed words. Unsaved Word edits are separate.", "desk-empty"))
        el("scope").textContent = "Read-only · saved versions only"
      }
    }
    controls()
  }
  function moveChange(direction: -1 | 1) {
    change += direction; render()
    const preferred = button(direction === 1 ? "next" : "previous")
    const focusTarget = preferred.disabled ? button(direction === 1 ? "previous" : "next") : preferred
    focusTarget.focus()
  }
  async function perform(work: () => Promise<void>) {
    if (busy || disposed) return
    busy = true; controls()
    try { await work() } catch (error) { say(error instanceof Error ? error.message : "The desk could not complete this action. Try again.") }
    finally { busy = false; if (!disposed) controls() }
  }
  async function captureSelection() {
    const value = await transport.request({ action: "selection" }) as WordSelectionSnapshot
    if (disposed) return
    selection = value; render(); say(value.reason ?? "Word selection updated. Your sources stay open.")
  }
  async function loadSources() {
    const value = await transport.request({ action: "sources" }) as DeskSources
    if (disposed) return
    sources = value
    el("identity").textContent = `${sources.documentName} · ${sources.matterName}`
    const choice = select("source-choice"), previous = choice.value
    choice.replaceChildren(new Option("Choose a saved source…", ""))
    for (const source of sources.sources) choice.add(new Option(`${source.name} · v${source.version?.number ?? "?"}`, source.id))
    if (sources.sources.some((source) => source.id === previous)) choice.value = previous
    say(sources.truncated ? "Some sources have no saved preview or exceed the list limit. Open previews are unchanged." : `${sources.sources.length} sources available.`)
  }
  async function loadVersions() {
    const value = await transport.request({ action: "versions" }) as DeskVersions
    if (disposed) return
    versions = value
    for (const name of ["base", "current"]) {
      const control = select(name), previous = control.value
      control.replaceChildren()
      for (const version of versions.versions) {
        const date = new Date(version.createdAt).toLocaleDateString([], { day: "numeric", month: "short" })
        const option = new Option(`Version ${version.number} · ${date}${version.current ? " · latest at refresh" : ""}${version.removed ? " · removed" : ""}${version.label ? ` · ${version.label}` : ""}`, version.id)
        option.disabled = version.removed; control.add(option)
      }
      const available = versions.versions.filter((version) => !version.removed).sort((a, b) => b.number - a.number)
      control.value = available.some((version) => version.id === previous) ? previous : available[name === "base" ? 1 : 0]?.id ?? ""
    }
    say("Version list updated. Choose Compare to update the comparison.")
  }
  function toggleFind(open: boolean) {
    findOpen = open; render()
    if (open) el<HTMLInputElement>("find").focus()
    else button("find-toggle").focus()
  }
  const onKey = (event: KeyboardEvent) => {
    if (mode !== "sources") return
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "f") { event.preventDefault(); toggleFind(true) }
    if (event.key === "Escape" && findOpen) { event.preventDefault(); toggleFind(false) }
  }
  root.addEventListener("keydown", onKey)
  button("close").onclick = options.close
  button("reload").onclick = () => { void perform(loadSources) }
  button("reload-versions").onclick = () => { void perform(loadVersions) }
  select("source-choice").onchange = () => { void perform(async () => {
    const source = sources?.sources.find((item) => item.id === select("source-choice").value)
    if (!source) return
    const retained = pins.get(`${source.id}:${source.hash}`)
    let preview: DeskSourcePreview
    try { preview = retained ?? await transport.request({ action: "source", sourceId: source.id, expectedHash: source.hash }) as DeskSourcePreview }
    catch (error) { if (!disposed) select("source-choice").value = ""; throw error }
    if (disposed) return
    active = preview; renderPins(); render(); say(`${preview.name} opened at version ${preview.version?.number ?? "unavailable"}.`)
  }) }
  select("base").onchange = select("current").onchange = () => { controls(); say("Choose Compare to show the selected versions.") }
  el<HTMLInputElement>("find").oninput = render
  button("find-toggle").onclick = () => toggleFind(!findOpen)
  button("find-close").onclick = () => toggleFind(false)
  button("pin").onclick = () => {
    if (!active || pins.size >= 8 || busy) return
    pins.set(key(active), active); renderPins(); render(); say("Source kept open at this saved version.")
  }
  button("sources-mode").onclick = () => { mode = "sources"; render(); say("Sources and your captured Word selection.") }
  button("compare-mode").onclick = () => { mode = "compare"; render(); if (!versions) void perform(loadVersions) }
  button("compare").onclick = () => { void perform(async () => {
    const value = await transport.request({ action: "compare", baseVersion: select("base").value, currentVersion: select("current").value }) as DeskComparison
    if (disposed) return
    comparison = value; change = 0; render(); say("Comparison ready. Insertions are underlined; deletions are struck through.")
  }) }
  render(); renderPins()
  const ready = perform(async () => {
    await loadSources()
    try { await captureSelection() } catch { say("Sources are ready. Select text in Word, then choose Use current selection.") }
  })
  return { ready, dispose() { disposed = true; pins.clear(); root.removeEventListener("keydown", onKey); root.replaceChildren(); root.classList.remove("source-desk") } }
}
