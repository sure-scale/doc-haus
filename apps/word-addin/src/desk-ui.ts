import type { DeskComparison, DeskSourcePreview, DeskSources, DeskVersions } from "../../../dochaus/lib/word-source-desk-contract"
import type { DeskRequest } from "./desk-protocol"
import type { WordSelectionSnapshot } from "./selection"

export type DeskTransport = { request(input: DeskRequest): Promise<unknown> }

/** Each pin is the exact received snapshot, never an implicit latest-file read. */
export function mountDesk(root: HTMLElement, transport: DeskTransport, options: { close(): void; floating?: boolean }) {
  root.classList.add("source-desk")
  root.innerHTML = `
    <header class="desk-head"><div class="desk-brand"><img src="./assets/icon.png" width="38" height="38" alt="" />
      <div><p class="desk-eyebrow">doc.haus · Word</p><h2>Source Desk</h2><p class="desk-subtitle" data-el="identity">Connecting to this matter…</p></div></div>
      <button type="button" data-el="close">${options.floating ? "Close desk" : "Back to saving"}</button></header>
    <nav class="desk-modes" aria-label="Desk views"><button class="desk-mode" data-el="sources-mode" aria-pressed="true" type="button">Sources</button>
      <button class="desk-mode" data-el="compare-mode" aria-pressed="false" type="button">Compare versions</button></nav>
    <div data-el="source-tools"><div class="desk-toolbar">
      <label>Matter source<select data-el="source-choice" aria-label="Matter source"><option value="">Loading sources…</option></select></label>
      <button type="button" data-el="pin" disabled>Pin source</button><button type="button" data-el="reload">Refresh list</button>
      <label>Find exact text in both passages<input data-el="find" type="search" placeholder="A phrase to connect" maxlength="255" /></label></div>
      <div class="desk-pins" data-el="pins" role="group" aria-label="Pinned sources"></div></div>
    <div data-el="compare-tools" hidden><div class="desk-toolbar">
      <label>Base · saved version<select data-el="base" aria-label="Base saved version"></select></label><label>Compared · saved version<select data-el="current" aria-label="Compared saved version"></select></label>
      <button type="button" data-el="compare" class="desk-primary" disabled>Compare</button><button type="button" data-el="reload-versions">Refresh versions</button></div>
      <div class="desk-change-nav"><button type="button" data-el="previous">← Previous</button><span data-el="change-count"></span><button type="button" data-el="next">Next →</button></div></div>
    <div class="desk-columns"><section class="desk-column" aria-label="Word context or base version" data-el="left"></section>
      <section class="desk-column" aria-label="Pinned source or compared version" data-el="right"></section></div>
    <footer class="desk-footer"><div><p role="status" aria-live="polite" data-el="status">Opening Source Desk…</p><p class="desk-note" data-el="scope">Read-only workspace · captures refresh when you ask</p></div>
      <button type="button" class="desk-primary" data-el="show" disabled>Show passage in Word</button></footer>`
  const el = <T extends HTMLElement>(name: string) => root.querySelector(`[data-el="${name}"]`) as T
  const button = (name: string) => el<HTMLButtonElement>(name)
  const select = (name: string) => el<HTMLSelectElement>(name)
  let disposed = false
  let busy = false
  let mode: "sources" | "compare" = "sources"
  let selection: WordSelectionSnapshot | undefined
  let sources: DeskSources | undefined
  let versions: DeskVersions | undefined
  const pins = new Map<string, DeskSourcePreview>()
  let active: string | undefined
  let comparison: DeskComparison | undefined
  let change = 0
  const say = (text: string) => { if (!disposed) el("status").textContent = text }
  const node = <K extends keyof HTMLElementTagNameMap>(tag: K, text: string, className?: string) => {
    const result = document.createElement(tag)
    result.textContent = text
    if (className) result.className = className
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
  const metadata = (area: HTMLElement, lines: string[], hash?: string) => {
    const block = node("div", "", "desk-provenance")
    for (const line of lines) block.append(node("p", line))
    if (hash) block.append(node("code", `SHA-256 · ${hash}`))
    area.append(block)
  }
  function controls() {
    for (const name of ["pin", "reload", "reload-versions", "compare", "show", "capture", "previous", "next"]) {
      const control = root.querySelector(`[data-el="${name}"]`) as HTMLButtonElement | null
      if (control) control.disabled = busy
    }
    button("pin").disabled = busy || !select("source-choice").value || pins.size >= 8
    button("compare").disabled = busy || !select("base").value || !select("current").value || select("base").value === select("current").value
    button("show").disabled = busy || mode !== "sources" || !selection?.navigable
    const count = comparison?.changes.length ?? 0
    button("previous").disabled = busy || !count || change <= 0
    button("next").disabled = busy || !count || change >= count - 1
  }
  function renderPins() {
    el("pins").replaceChildren()
    for (const [key, source] of pins) {
      const group = node("div", "", "desk-pin")
      const pick = node("button", `${source.name} · v${source.version?.number ?? "?"}`)
      pick.type = "button"
      pick.setAttribute("aria-pressed", String(active === key))
      pick.dataset.pin = key
      pick.onclick = () => {
        active = key; renderPins(); render(); say("Pinned source selected.")
        const replacement = Array.from(el("pins").querySelectorAll<HTMLButtonElement>("button")).find((item) => item.dataset.pin === key)
        replacement?.focus()
      }
      const remove = node("button", "×", "desk-unpin")
      remove.type = "button"
      remove.setAttribute("aria-label", `Unpin ${source.name}`)
      remove.onclick = () => { pins.delete(key); if (active === key) active = pins.keys().next().value; renderPins(); render(); button("pin").focus(); say("Source unpinned.") }
      group.append(pick, remove)
      el("pins").append(group)
    }
    controls()
  }
  function render() {
    if (disposed) return
    el("source-tools").hidden = mode !== "sources"
    el("compare-tools").hidden = mode !== "compare"
    button("sources-mode").setAttribute("aria-pressed", String(mode === "sources"))
    button("compare-mode").setAttribute("aria-pressed", String(mode === "compare"))
    button("show").hidden = mode !== "sources"
    if (mode === "sources") {
      const query = el<HTMLInputElement>("find").value
      const left = column("left", "In Word · captured selection", "Your current passage")
      const capture = node("button", "Refresh selection")
      capture.type = "button"
      capture.dataset.el = "capture"
      capture.onclick = () => { void perform(captureSelection) }
      left.append(node("p", selection ? `${selection.saved ? "Saved in Word" : "Unsaved Word context"} · not a confirmed matter version` : "Select a passage in Word to connect it to your sources.", "desk-note"), capture)
      if (selection) {
        left.append(textBlock(selection.text || "No text is selected in Word.", query))
        metadata(left, [`Captured ${new Date(selection.capturedAt).toLocaleTimeString()}. Refresh to use a different selection.`, selection.reason ?? "Navigation checks that the document still matches this capture."])
      }
      const source = active ? pins.get(active) : undefined
      const right = column("right", "Pinned matter source · saved bytes", source?.name ?? "Keep your evidence beside you")
      if (source) {
        right.append(textBlock(source.text || "This saved source has no previewable text.", query))
        metadata(right, [`${source.role === "deliverable" ? "Matter output" : "Matter source"} · saved version ${source.version?.number ?? "unavailable"}`, `Version ID · ${source.version?.id ?? "unavailable"}`, source.truncated ? `Preview shows the first ${source.text.length.toLocaleString()} of ${source.characters.toLocaleString()} characters.` : `${source.characters.toLocaleString()} characters · complete text preview`, "Pinned snapshot stays unchanged when your Word selection changes."], source.hash)
        if (query) right.append(node("p", source.text.includes(query) ? "Underlined highlights show the exact phrase in this preview." : "The exact phrase is absent from this preview. This does not assess legal support.", "desk-note"))
      } else right.append(node("p", "Choose a matter source and pin it here. Switch between your pins while keeping your Word passage in view.", "desk-empty"))
    } else {
      const hunk = comparison?.changes[change]
      for (const [name, side, text] of [["left", comparison?.base, hunk?.before], ["right", comparison?.current, hunk?.after]] as const) {
        const area = column(name, name === "left" ? "Base · confirmed saved version" : "Compared · confirmed saved version", side ? `${side.name} · version ${side.number}` : "Choose two saved versions")
        if (side) {
          area.append(textBlock(text ?? (hunk ? "No text on this side of the change." : "No textual changes reported.")))
          metadata(area, [`Version ID · ${side.id}`, `Saved ${new Date(side.createdAt).toLocaleString()} · ${side.origin}`, `${side.trackedChanges.count} input tracked changes`, hunk ? `${hunk.kind} · location in ${hunk.location.side === "left" ? "base" : "compared"} version: ${hunk.location.scope} paragraph ${hunk.location.paragraph}${hunk.location.heading ? ` · ${hunk.location.heading}` : ""}` : "Saved-byte comparison; your open Word context is separate."], side.hash)
        } else area.append(node("p", "Compare retained versions of the linked document. The open Word selection is never substituted for saved bytes.", "desk-empty"))
      }
      el("change-count").textContent = comparison ? `${comparison.changes.length ? `Change ${change + 1} of ${comparison.changes.length}` : "No textual changes"}${hunk?.untracked ? " · includes an edit made without tracking" : ""}${comparison.truncated ? " · partial result" : ""}` : "Saved versions only"
      el("scope").textContent = comparison ? `${comparison.untracked.length} edits without tracking · ${comparison.proposals.length} earlier proposal outcomes${comparison.truncated ? " · result truncated; use Versions in doc.haus for full review" : ""}` : "Read-only workspace · saved versions only"
    }
    if (mode === "sources") el("scope").textContent = "Read-only workspace · captures refresh when you ask"
    controls()
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
    selection = value
    render()
    say(value.reason ?? "Selection captured. Pinned sources are unchanged.")
  }
  async function loadSources() {
    const value = await transport.request({ action: "sources" }) as DeskSources
    if (disposed) return
    sources = value
    el("identity").textContent = `${sources.matterName} · ${sources.documentName}`
    const choice = select("source-choice")
    const previous = choice.value
    choice.replaceChildren(new Option("Choose a saved matter source…", ""))
    for (const source of sources.sources) choice.add(new Option(`${source.name} · v${source.version?.number ?? "?"}`, source.id))
    if (sources.sources.some((source) => source.id === previous)) choice.value = previous
    say(sources.truncated ? "Some material has no confirmed saved preview or exceeds the list limit. Existing pins are unchanged." : `${sources.sources.length} saved sources available. Existing pins are unchanged.`)
  }
  async function loadVersions() {
    const value = await transport.request({ action: "versions" }) as DeskVersions
    if (disposed) return
    versions = value
    for (const name of ["base", "current"]) {
      const control = select(name)
      const previous = control.value
      control.replaceChildren()
      for (const version of versions.versions) {
        const option = new Option(`Version ${version.number}${version.current ? " · current in doc.haus" : ""}${version.removed ? " · removed" : ""}${version.label ? ` · ${version.label}` : ""}`, version.id)
        option.disabled = version.removed
        control.add(option)
      }
      const available = versions.versions.filter((version) => !version.removed).sort((a, b) => b.number - a.number)
      control.value = available.some((version) => version.id === previous) ? previous : available[name === "base" ? 1 : 0]?.id ?? ""
    }
    say("Saved version list refreshed. Existing comparison remains the same until you choose Compare.")
  }
  button("close").onclick = options.close
  button("reload").onclick = () => { void perform(loadSources) }
  button("reload-versions").onclick = () => { void perform(loadVersions) }
  select("source-choice").onchange = controls
  select("base").onchange = controls
  select("current").onchange = controls
  el<HTMLInputElement>("find").oninput = render
  button("pin").onclick = () => { void perform(async () => {
    const source = sources?.sources.find((item) => item.id === select("source-choice").value)
    if (!source || pins.size >= 8) return
    const preview = await transport.request({ action: "source", sourceId: source.id, expectedHash: source.hash }) as DeskSourcePreview
    if (disposed) return
    active = `${preview.id}:${preview.hash}`
    pins.set(active, preview)
    renderPins(); render(); say("Source pinned at its displayed saved version.")
  }) }
  button("show").onclick = () => { void perform(async () => {
    if (!selection?.navigable) return
    await transport.request({ action: "show", selection })
    say("The exact captured passage is selected in Word.")
  }) }
  button("sources-mode").onclick = () => { mode = "sources"; render(); say("Pinned source snapshots and captured Word passage. Refresh the selection to use a different passage.") }
  button("compare-mode").onclick = () => { mode = "compare"; render(); say("Compare confirmed saved versions. Your open Word context is separate."); if (!versions) void perform(loadVersions) }
  button("compare").onclick = () => { void perform(async () => {
    const value = await transport.request({ action: "compare", baseVersion: select("base").value, currentVersion: select("current").value }) as DeskComparison
    if (disposed) return
    comparison = value; change = 0; render(); say("Saved versions compared. Use Previous and Next to inspect both sides together.")
  }) }
  button("previous").onclick = () => { change--; render() }
  button("next").onclick = () => { change++; render() }
  render()
  const ready = perform(async () => {
    await loadSources()
    try { await captureSelection() } catch { say("Sources are ready. Select a passage in Word, then choose Refresh selection.") }
  })
  return { ready, dispose() { disposed = true; pins.clear(); root.replaceChildren(); root.classList.remove("source-desk") } }
}
