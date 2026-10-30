import type { WordChatConversation, WordChatConversationResult, WordChatMaterials, WordChatMaterial, WordChatScope, WordChatSend, WordChatTurn, WordChatSuggestion, WordChatSource, WordChatPreferences } from "../../../dochaus/lib/word-chat-contract"
import type { WordChatProposal, WordChatSnapshot, WordChatTableRowInsertionProposal, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { wordChatApplyPlan, wordChatTargetCurrent } from "../../../dochaus/lib/word-chat-amendment-contract"
import { ChatRequestError, type WordChatApi } from "./chat-client"
import { nativeEventSession, subscribeNativeChatStream } from "../../../dochaus/lib/native-chat-stream"
import type { WordChatApplyEligibility } from "./word-chat-apply"
import { nativeWordChatParagraphInsertionEligible } from "./word-chat-apply"

export type ChatHost = {
  documentName: string
  capture(includeSelection?: boolean): Promise<WordChatSnapshot>
  captureSelection(): Promise<WordChatTarget | undefined>
  watchSelection?(onChange: () => void): () => void
  openAttachments?(options: { conversationId: string; attached: WordChatMaterial[]; onAttach(material: WordChatMaterial): void }): Promise<void>
  apply(proposal: WordChatProposal, request: { conversationId: string; turnId: string }): Promise<unknown>
  applyRow?(proposal: WordChatTableRowInsertionProposal, request: { conversationId: string; turnId: string }): Promise<unknown>
  applyEligibility(proposal: WordChatProposal): WordChatApplyEligibility
  sourceUrl?(source: WordChatSource): string | undefined
  now?: () => number
  schedule?: (run: () => void, delay: number) => () => void
}

/** Mount once. Hiding a feature tab never remounts or submits the conversation. */
export function mountWordChat(root: HTMLElement, api: WordChatApi, host: ChatHost) {
  root.dataset.documentName = host.documentName
  root.innerHTML = `<section data-chat="conversations" id="chat-conversations" class="chat-conversations" aria-label="Conversations" hidden>
    <div class="chat-panel-heading"><h2>Conversations</h2><button data-chat="conversations-close" class="chat-subtle" type="button">Done</button></div><div data-chat="conversation-list" class="chat-conversation-list"></div>
  </section>
  <section data-chat="materials" class="chat-materials" aria-label="Attach material" hidden>
    <div class="chat-materials-heading"><h2>Attach material</h2><button data-chat="materials-close" class="chat-subtle" type="button">Done</button></div>
    <div class="chat-material-tabs" role="tablist" aria-label="Attachment source">
      <button data-chat="material-existing-tab" id="chat-material-existing-tab" type="button" role="tab" aria-selected="true" aria-controls="chat-material-existing">Existing</button>
      <button data-chat="material-upload-tab" id="chat-material-upload-tab" type="button" role="tab" aria-selected="false" aria-controls="chat-material-upload" tabindex="-1">Upload</button>
    </div>
    <div data-chat="material-body" class="chat-materials-body">
      <section data-chat="material-existing" id="chat-material-existing" role="tabpanel" aria-labelledby="chat-material-existing-tab" tabindex="0"><div data-chat="material-list" class="chat-material-list"></div></section>
      <section data-chat="material-upload" id="chat-material-upload" role="tabpanel" aria-labelledby="chat-material-upload-tab" tabindex="0" hidden>
        <label>File destination<select data-chat="upload-mode"><option value="request">For this chat only</option><option value="matter">Add to this matter</option></select></label>
        <p>Attached material is context for this chat.</p>
        <label class="chat-file-label">Choose a file<input data-chat="upload" type="file" aria-label="Choose a supported document" accept=".docx,.pdf,.txt,.md,.rtf,.csv,.xlsx,.eml,.msg"></label>
        <p data-chat="upload-message" role="status" aria-live="polite"></p>
      </section>
    </div>
  </section>
  <div data-chat="scroll" class="chat-scroll"><div data-chat="history" class="chat-history" role="log" aria-label="Conversation" aria-live="off"></div>
  <section data-chat="suggestions" class="chat-suggestions" aria-label="Suggestions" hidden>
    <div class="chat-suggestions-heading"><h2>Suggestions</h2><button data-chat="pause" class="chat-subtle" type="button">Pause</button></div>
    <div data-chat="suggestion-content"></div>
  </section></div>
  <div class="chat-recovery"><p data-chat="activity" class="chat-activity" role="status" aria-live="polite"></p>
    <button data-chat="reconnect" class="chat-subtle" type="button" hidden>Reconnect</button>
    <button data-chat="stop" class="chat-subtle" type="button" hidden>Stop</button></div>
  <form data-chat="form" class="chat-composer">
    <div data-chat="chips" class="chat-chips"></div>
    <div class="chat-composer-context">
      <div class="chat-scope" role="group" aria-label="Chat context"><span data-chat="document-scope" class="chat-document-scope"><button data-chat="scope-document" type="button" aria-pressed="true" title="Use this Word document and attached material">This document</button><button data-chat="document-info" class="chat-document-info" type="button" aria-label="Document information" aria-describedby="chat-document-info" aria-expanded="false" aria-controls="chat-document-info"><svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="7.25"/><path d="M10 9v5m0-8v.2"/></svg></button><span data-chat="document-info-content" id="chat-document-info" class="chat-document-tooltip" role="tooltip" hidden></span></span><button data-chat="scope-matter" type="button" aria-pressed="false" title="Also use the matter’s sources and playbook">Whole matter</button></div>
      <button data-chat="conversation" class="chat-subtle chat-history-toggle" type="button" title="Conversations" aria-label="Conversations" aria-expanded="false" aria-controls="chat-conversations"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 12a9 9 0 1 0 2.7-6.4L3 8m0-5v5h5m4-1v5l3 2"/></svg></button>
      <button data-chat="new" class="chat-subtle chat-new" type="button" title="New conversation" aria-label="New conversation"><svg viewBox="0 0 20 20" aria-hidden="true"><path d="M10 4v12M4 10h12"/></svg><span>New</span></button>
    </div>
    <textarea data-chat="message" aria-label="Message doc.haus" placeholder="Ask doc.haus…" rows="2"></textarea>
    <div class="chat-composer-actions">
      <button data-chat="attach" class="chat-subtle" type="button" aria-expanded="false">＋ Attach</button>
      <div data-chat="mode" class="chat-mode"><button data-chat="mode-toggle" class="chat-subtle chat-mode-toggle" type="button" aria-label="Response mode: Auto" aria-haspopup="menu" aria-expanded="false" aria-controls="chat-mode-picker"><span data-chat="mode-label">Auto</span><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="m6 9 6 6 6-6"/></svg></button>
        <div data-chat="mode-picker" id="chat-mode-picker" class="chat-mode-picker" role="menu" aria-label="Response mode" hidden>
          <button data-chat="mode-auto" type="button" role="menuitemradio" aria-checked="true" tabindex="0"><span>Auto</span><small>Choose for this request</small></button>
          <button data-chat="mode-qa" type="button" role="menuitemradio" aria-checked="false" tabindex="-1"><span>Q&amp;A</span><small>Ask about the document</small></button>
          <button data-chat="mode-redline" type="button" role="menuitemradio" aria-checked="false" tabindex="-1"><span>Redline</span><small>Revise wording and layout</small></button>
        </div>
      </div>
      <button data-chat="send" class="chat-send" type="submit">Send <span aria-hidden="true">↑</span></button>
    </div>
  </form>`
  const get = <T extends HTMLElement>(name: string) => root.querySelector(`[data-chat="${name}"]`) as T
  const button = (name: string) => get<HTMLButtonElement>(name)
  const message = get<HTMLTextAreaElement>("message")
  const history = get("history"), scroller = get("scroll"), activity = get("activity")
  const now = host.now ?? Date.now
  const schedule = host.schedule ?? ((run: () => void, delay: number) => { const id = setTimeout(run, delay); return () => clearTimeout(id) })
  let preferences: WordChatPreferences = { defaultScope: "document", suggestions: true }
  let conversations: WordChatConversation[] = [], conversation: WordChatConversation | undefined
  let turns: WordChatTurn[] = [], materials: WordChatMaterial[] = [], attachments: WordChatMaterial[] = []
  let selection: WordChatTarget | undefined, scope: WordChatScope = preferences.defaultScope
  let defaultScope = preferences.defaultScope, scopeChanged = false
  let mode: "auto" | "qa" | "redline" = "auto"
  const modeLabels = { auto: "Auto", qa: "Q&A", redline: "Redline" }
  let disposed = false, connecting = false, sending = false, uploading = false, applying = false, stopping = false, pickerOpen = false, suggestionsPaused = false
  let pending: WordChatSend | undefined, foreground: WordChatTurn | undefined, suggestion: WordChatTurn | undefined
  let refinement: WordChatProposal | undefined, suggestionEpoch = 0, suggestionCalls = 0
  let rowRefinement: WordChatTableRowInsertionProposal | undefined
  let selectionCaptureEpoch = 0, selectionPending = false
  let suppressedSelectionKey = "", pickerEpoch = 0
  let cancelSelection: (() => void) | undefined, stopWatchingSelection: (() => void) | undefined
  let foregroundEpoch = 0
  let materialTab: "existing" | "upload" = "existing"
  let suggestionKey = "", stableKey = "", stableAt = 0, lastSuggestion = -Infinity, lastMaterials = 0
  let cancelMonitor: (() => void) | undefined, cancelReconcile: (() => void) | undefined, monitorBusy = false
  let stream: AbortController | undefined, streamEpoch = 0, conversationEpoch = 0, metadataRequest = 0
  const dismissed = new Set<string>(), applied = new Set<string>(), declined = new Set<string>(), uncertain = new Set<string>()
  const proposalErrors = new Map<string, string>()
  const nodes = new Map<string, { stamp: string; node: HTMLElement }>()
  get("document-info-content").textContent = host.documentName
  let documentInfoPinned = false
  function showDocumentInfo(open: boolean) {
    get("document-info-content").hidden = !open
    button("document-info").setAttribute("aria-expanded", String(open))
  }
  const dismissDocumentInfo = (event: MouseEvent) => {
    if (!get("document-scope").contains(event.target as Node)) { documentInfoPinned = false; showDocumentInfo(false) }
    if (!get("mode").contains(event.target as Node)) showModePicker(false)
  }
  get("document-scope").onmouseenter = () => showDocumentInfo(true)
  get("document-scope").onmouseleave = () => { if (!documentInfoPinned && root.ownerDocument.activeElement !== button("document-info")) showDocumentInfo(false) }
  button("document-info").onfocus = () => showDocumentInfo(true)
  button("document-info").onblur = () => { documentInfoPinned = false; showDocumentInfo(false) }
  button("document-info").onclick = () => { documentInfoPinned = !documentInfoPinned; showDocumentInfo(documentInfoPinned) }
  button("document-info").onkeydown = event => {
    if (event.key === "Escape") { documentInfoPinned = false; showDocumentInfo(false); event.preventDefault(); event.stopPropagation() }
  }
  root.ownerDocument.addEventListener("click", dismissDocumentInfo)
  function showModePicker(open: boolean, restoreFocus = false) {
    get("mode-picker").hidden = !open
    button("mode-toggle").setAttribute("aria-expanded", String(open))
    if (open) button(`mode-${mode}`).focus()
    else if (restoreFocus) button("mode-toggle").focus()
  }

  function status(text = "", error = false) {
    if (disposed) return
    activity.textContent = text; activity.dataset.error = String(error)
  }
  const suggestionsEnabled = () => preferences.suggestions && !suggestionsPaused
  const reviewingAmendment = () => turns.some(turn => { const proposal = turn.proposal ?? turn.rowProposal; return turn.status === "complete" && proposal && !applied.has(proposal.id) && !declined.has(proposal.id) })
  const proposalFinished = (id: string) => applied.has(id) || declined.has(id) || uncertain.has(id)
  const conversationLocked = () => connecting || sending || applying || uploading || pickerOpen || Boolean(pending) || foreground?.status === "running"
  function controls() {
    if (disposed) return
    button("send").disabled = conversationLocked() || selectionPending || !conversation || !message.value.trim()
    button("new").disabled = conversationLocked() || !conversation
    button("conversation").disabled = conversationLocked()
    for (const choice of root.querySelectorAll<HTMLButtonElement>(".chat-conversation-list button")) choice.disabled = conversationLocked()
    button("stop").hidden = foreground?.status !== "running"
    button("stop").disabled = stopping
    for (const action of root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")) action.disabled = applying
    button("attach").disabled = !conversation || connecting || uploading || pickerOpen || applying
    button("mode-toggle").disabled = conversationLocked()
    for (const value of ["auto", "qa", "redline"] as const) {
      const choice = button(`mode-${value}`)
      choice.disabled = conversationLocked(); choice.setAttribute("aria-checked", String(mode === value)); choice.tabIndex = mode === value ? 0 : -1
    }
    get("mode-label").textContent = modeLabels[mode]
    button("mode-toggle").setAttribute("aria-label", `Response mode: ${modeLabels[mode]}`)
    if (conversationLocked()) showModePicker(false)
    for (const value of ["document", "matter"] as const) button(`scope-${value}`).setAttribute("aria-pressed", String(scope === value))
  }
  function options() {
    const list = get("conversation-list"); list.replaceChildren()
    if (!conversations.length) list.append(line("p", "No earlier conversations.", "chat-message-meta"))
    for (const item of conversations) {
      const choice = document.createElement("button"); choice.type = "button"; choice.dataset.conversation = item.id
      choice.setAttribute("aria-current", String(item.id === conversation?.id))
      const title = line("span", item.title || "Conversation", "chat-conversation-title")
      const date = new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }).format(new Date(item.updatedAt))
      const detail = line("small", `${item.id === conversation?.id ? "Current · " : ""}${date}${item.documentName !== host.documentName ? ` · ${item.documentName}` : ""}`)
      choice.append(title, detail); choice.disabled = conversationLocked()
      choice.onclick = () => {
        if (conversationLocked()) return
        setPanel("conversations", false)
        if (item.id !== conversation?.id) changeConversation(item.id)
        else message.focus()
      }
      list.append(choice)
    }
  }
  function setPanel(name: "materials" | "conversations", open: boolean) {
    const toggle = (panel: typeof name) => panel === "materials" ? "attach" : "conversation"
    if (open) {
      for (const other of ["materials", "conversations"] as const) if (other !== name) {
        get(other).hidden = true; button(toggle(other)).setAttribute("aria-expanded", "false")
      }
    }
    get(name).hidden = !open
    button(toggle(name)).setAttribute("aria-expanded", String(open))
    // Attachments use the conversation's space, with one scrolling body.
    // The composer and its draft remain mounted underneath.
    scroller.hidden = history.hidden = !get("materials").hidden || !get("conversations").hidden
    renderSuggestions()
    if (open) {
      if (name === "conversations") (get("conversation-list").querySelector<HTMLElement>("button") ?? button("conversations-close")).focus()
      else get(`material-${materialTab}-tab`).focus()
    }
  }
  function setMaterialTab(next: "existing" | "upload", focus = false) {
    materialTab = next
    for (const name of ["existing", "upload"] as const) {
      const tab = button(`material-${name}-tab`)
      tab.setAttribute("aria-selected", String(name === next)); tab.tabIndex = name === next ? 0 : -1
      get(`material-${name}`).hidden = name !== next
    }
    get("material-body").scrollTop = 0
    if (focus) button(`material-${next}-tab`).focus()
  }
  function line(tag: string, text: string, className?: string) {
    const element = document.createElement(tag); element.textContent = text
    if (className) element.className = className
    return element
  }
  const displayWordText = (text: string) => text.replace(/\r\n?|\v/g, "\n")
  function renderChips() {
    const chips = get("chips"); chips.replaceChildren()
    function chip(label: string, detail: string, remove: () => void) {
      const row = document.createElement("div"); row.className = "chat-chip"
      const details = document.createElement("details"); details.append(line("summary", label), line("p", displayWordText(detail)))
      const close = document.createElement("button"); close.type = "button"; close.textContent = "×"; close.setAttribute("aria-label", `Remove ${label}`)
      close.onclick = remove; row.append(details, close); chips.append(row)
    }
    if (selection) chip("Selected Word text", selection.text, () => { clearSelection(); refinement = undefined; rowRefinement = undefined; message.placeholder = "Ask doc.haus…"; renderChips(); invalidateSuggestions() })
    for (const material of attachments) chip(material.name, `${material.kind === "playbook" ? "Playbook · " : ""}${material.provenance === "request-upload" ? "For this chat only" : "Saved matter material"}${material.description ? `\n${material.description}` : ""}`, () => { attachments = attachments.filter(item => item.id !== material.id || item.kind !== material.kind); renderChips(); invalidateSuggestions() })
    if (refinement) chip("Refining amendment", refinement.before, () => { refinement = undefined; message.placeholder = "Ask doc.haus…"; renderChips() })
    if (rowRefinement) chip("Refining row", rowRefinement.cells.join(" · "), () => { rowRefinement = undefined; message.placeholder = "Ask doc.haus…"; renderChips() })
  }
  function selectionKey(selected: WordChatTarget): string {
    return JSON.stringify([selected.text, selected.paragraphs.map(({ id, text }) => [id, text]), selected.startOffset, selected.endOffset])
  }
  function cancelSelectionCapture() {
    selectionCaptureEpoch++; cancelSelection?.(); cancelSelection = undefined; selectionPending = false
  }
  function clearSelection() {
    cancelSelectionCapture()
    if (selection) suppressedSelectionKey = selectionKey(selection)
    selection = undefined
  }
  async function captureSelection() {
    if (disposed || applying) return
    const epoch = ++selectionCaptureEpoch
    selectionPending = true; controls()
    try {
      const selected = await host.captureSelection()
      if (disposed || applying || epoch !== selectionCaptureEpoch) return
      if (!selected?.text.trim()) {
        if (!document.hasFocus()) suppressedSelectionKey = ""
        return
      }
      const key = selectionKey(selected)
      // Pane focus can collapse the native range. Keep its last exact capture;
      // removing the chip suppresses that range until a different one is chosen.
      if (key === suppressedSelectionKey || selection && key === selectionKey(selection)) return
      suppressedSelectionKey = ""; selection = selected
      renderChips(); invalidateSuggestions()
    } catch { /* Selection is optional, including on older Word hosts. */ }
    finally { if (epoch === selectionCaptureEpoch) { selectionPending = false; controls() } }
  }
  function selectionChanged() {
    cancelSelectionCapture()
    if (disposed || applying) return
    selectionPending = true; controls()
    const epoch = selectionCaptureEpoch
    cancelSelection = schedule(() => { if (epoch === selectionCaptureEpoch) void captureSelection() }, 200)
  }
  function attachMaterial(material: WordChatMaterial) {
    attachments = attachments.filter(item => item.kind !== material.kind || item.id !== material.id)
    attachments.push(material); materialList(); renderChips(); invalidateSuggestions()
  }
  async function openAttachments() {
    if (disposed || !conversation || connecting || applying || pickerOpen || uploading) return
    if (!host.openAttachments) { setPanel("materials", get("materials").hidden); materialList(); return }
    const conversationId = conversation.id, epoch = ++pickerEpoch
    setPanel("conversations", false); setPanel("materials", false)
    pickerOpen = true; controls(); button("attach").setAttribute("aria-expanded", "true")
    try {
      await host.openAttachments({ conversationId, attached: attachments.map(material => ({ ...material })), onAttach(material) {
        if (!disposed && pickerOpen && epoch === pickerEpoch && conversation?.id === conversationId) attachMaterial(material)
      } })
    } catch {
      if (!disposed && epoch === pickerEpoch && conversation?.id === conversationId) {
        setPanel("materials", true); materialList(); status("Attachment window unavailable. Choose material here.")
      }
    } finally {
      if (!disposed && epoch === pickerEpoch) {
        pickerOpen = false; pickerEpoch++; controls()
        button("attach").setAttribute("aria-expanded", String(!get("materials").hidden))
      }
    }
  }
  function materialList() {
    const list = get("material-list"); list.replaceChildren()
    if (!materials.length) list.append(line("p", "No material available. You can upload a file."))
    for (const material of materials) {
      const choice = document.createElement("button"); choice.type = "button"
      choice.append(line("span", material.name), line("small", material.kind === "playbook" ? "Playbook" : "Saved source"))
      choice.disabled = attachments.some(item => item.kind === material.kind && item.id === material.id)
      choice.onclick = () => {
        attachments.push(material); materialList(); renderChips(); invalidateSuggestions(); setPanel("materials", false); message.focus()
      }
      list.append(choice)
    }
  }
  function clearProposalContext(proposal: WordChatProposal) {
    const affected = new Set(proposal.target.paragraphs.map(paragraph => paragraph.id))
    if (selection && selection.paragraphs.some(paragraph => affected.has(paragraph.id))) clearSelection()
    if (refinement?.id === proposal.id || refinement?.target.paragraphs.some(paragraph => affected.has(paragraph.id))) {
      refinement = undefined; message.placeholder = "Ask doc.haus…"
    }
    if (rowRefinement) { rowRefinement = undefined; message.placeholder = "Ask doc.haus…" }
    renderChips()
  }
  function proposalCard(proposal: WordChatProposal, turn: WordChatTurn): HTMLElement {
    const formatting = proposal.formatting
    const inlineFormatting = proposal.inlineFormatting
    const wordingChanged = proposal.before !== proposal.after
    const plan = wordChatApplyPlan(proposal.target, proposal.after, proposal.formatting)
    const eligibility = host.applyEligibility(proposal)
    const insertion = nativeWordChatParagraphInsertionEligible(proposal.target, proposal.after, formatting, inlineFormatting)
    const unavailable = !eligibility.supported ? eligibility.reason : !plan.supported && !insertion ? plan.reason : undefined
    const card = document.createElement("section"); card.className = "chat-proposal"; card.setAttribute("aria-label", (formatting || inlineFormatting) && !wordingChanged ? "Formatting to review" : "Amendment to review")
    card.append(line("h3", wordingChanged ? "Review amendment" : "Review formatting"))
    if (wordingChanged) {
      const before = document.createElement("details"); before.className = "chat-format-target"
      before.append(line("summary", "Before"), line("pre", displayWordText(proposal.before), "chat-before"))
      card.append(before, line("h4", "After"), line("pre", displayWordText(proposal.after), "chat-after"))
      const target = proposal.target, first = target.paragraphs[0], last = target.paragraphs.at(-1)
      const partial = first && last && (target.startOffset !== 0 || target.endOffset !== last.text.length)
      if (partial && plan.supported && proposal.before === target.text && wordChatTargetCurrent({ text: target.paragraphs.map(p => p.text).join("\r"), paragraphs: target.paragraphs, saved: target.saved, capturedAt: target.capturedAt, truncated: false }, target)) {
        const context = document.createElement("details"); context.className = "chat-format-target"
        context.append(line("summary", "In context"), line("pre", displayWordText(plan.paragraphs.join("\r"))))
        card.append(context)
      }
    }
    if (formatting) {
      if (!wordingChanged) card.append(line("p", "Wording unchanged", "chat-message-meta"))
      const list = document.createElement("ul"); list.className = "chat-format-plan"
      const groups: { first: number; last: number; change: typeof formatting[number]; key: string }[] = []
      for (const change of [...formatting].sort((a, b) => a.paragraph - b.paragraph)) {
        const key = JSON.stringify([change.style, change.spacingBefore, change.spacingAfter, change.lineSpacing])
        const previous = groups.at(-1)
        if (previous && previous.last + 1 === change.paragraph && previous.key === key) previous.last = change.paragraph
        else groups.push({ first: change.paragraph, last: change.paragraph, change, key })
      }
      for (const { first, last, change } of groups) {
        const item = document.createElement("li")
        const text = displayWordText(proposal.after.split("\r")[change.paragraph] ?? "").replace(/\s+/g, " ").trim()
        const description = [change.style?.replace(/Heading(\d)/, "Heading $1"), change.spacingBefore === undefined ? "" : `${change.spacingBefore} pt before`, change.spacingAfter === undefined ? "" : `${change.spacingAfter} pt after`, change.lineSpacing === undefined ? "" : `${change.lineSpacing}× line spacing`].filter(Boolean).join(" · ")
        const label = first !== last ? `Paragraphs ${first + 1}–${last + 1}` : text.length > 90 ? `${text.slice(0, 90)}…` : text || `Paragraph ${first + 1}`
        item.append(line("strong", label), line("p", description)); list.append(item)
      }
      const details = document.createElement("details"); details.className = "chat-format-target"
      if (wordingChanged) {
        details.append(line("summary", "Layout changes"), list)
        card.append(details)
      } else {
        details.append(line("summary", "Selected text"), line("pre", displayWordText(proposal.before)))
        card.append(list, details)
      }
    }
    if (inlineFormatting) {
      card.append(line("p", "Wording unchanged", "chat-message-meta"), line("pre", displayWordText(proposal.before), "chat-format-inline-target"))
      const list = document.createElement("ul"); list.className = "chat-format-plan"
      for (const [property, label] of [["bold", "Bold"], ["italic", "Italic"], ["underline", "Underline"]] as const) {
        if (inlineFormatting[property] !== undefined) list.append(line("li", `${label} ${inlineFormatting[property] ? "on" : "off"}`))
      }
      card.append(list)
    }
    if (proposal.explanation) card.append(line("p", proposal.explanation))
    if (proposal.openItems.length) {
      card.append(line("h4", "Open items")); const items = document.createElement("ul"); items.className = "chat-open-items"
      for (const text of proposal.openItems) items.append(line("li", text)); card.append(items)
    }
    if (applied.has(proposal.id)) {
      const note = line("p", turn.proposalApplied ? "Previously applied in Word." : "Applied in Word. Save when you’re ready.", "chat-local-note"); note.setAttribute("role", "status"); card.append(note); return card
    }
    if (uncertain.has(proposal.id)) {
      const note = line("p", "Check the change in Word. Use Undo if needed.", "chat-proposal-error"); note.setAttribute("role", "alert"); card.append(note); return card
    }
    if (declined.has(proposal.id)) { card.append(line("p", "Declined. Your document is unchanged.", "chat-message-meta")); return card }
    const error = proposalErrors.get(proposal.id)
    if (error) { const note = line("p", error, "chat-proposal-error"); note.setAttribute("role", "alert"); card.append(note) }
    if (unavailable) card.append(line("p", unavailable, "chat-proposal-limit"))
    const actions = document.createElement("div"); actions.className = "chat-proposal-actions"
    const refine = document.createElement("button"); refine.type = "button"; refine.className = "chat-subtle"; refine.textContent = "Refine"
    refine.disabled = applying
    refine.onclick = () => { if (applying || proposalFinished(proposal.id)) return; cancelSelectionCapture(); refinement = proposal; rowRefinement = undefined; renderChips(); controls(); message.placeholder = "How should this change?"; message.focus() }
    const apply = document.createElement("button"); apply.type = "button"; apply.className = "chat-send"; apply.textContent = "Apply"; apply.disabled = applying
    apply.onclick = async () => {
      if (applying || proposalFinished(proposal.id)) return
      const eligibility = host.applyEligibility(proposal)
      if (!eligibility.supported) { proposalErrors.delete(proposal.id); renderTurns(); return }
      applying = true; cancelSelectionCapture(); controls(); status("Applying in Word…")
      try {
        await host.apply(proposal, { conversationId: turn.conversationId, turnId: turn.id })
        if (disposed) return
        applied.add(proposal.id); proposalErrors.delete(proposal.id)
        clearProposalContext(proposal); status(); invalidateSuggestions(); renderTurns()
      }
      catch (error) {
        if (disposed) return
        if (error && typeof error === "object" && "mutationState" in error && error.mutationState === "uncertain") {
          uncertain.add(proposal.id); proposalErrors.delete(proposal.id); clearProposalContext(proposal)
        } else proposalErrors.set(proposal.id, error instanceof Error ? error.message : "Word could not apply this change. Refine it and try again.")
        status(); renderTurns()
      }
      finally { applying = false; if (!disposed) apply.disabled = false; controls() }
    }
    const decline = document.createElement("button"); decline.type = "button"; decline.className = "chat-subtle"; decline.textContent = "Decline"
    decline.disabled = applying
    decline.onclick = () => { if (applying || proposalFinished(proposal.id)) return; declined.add(proposal.id); proposalErrors.delete(proposal.id); if (refinement?.id === proposal.id) { refinement = undefined; renderChips() }; renderTurns() }
    actions.append(refine)
    if (!unavailable) actions.append(apply)
    actions.append(decline); card.append(actions); return card
  }
  function rowProposalCard(proposal: WordChatTableRowInsertionProposal, turn: WordChatTurn): HTMLElement {
    const card = document.createElement("section"); card.className = "chat-proposal"; card.setAttribute("aria-label", "Table row to review")
    const firstRow = proposal.table.rows[0]?.map(cell => cell.text.replace(/\s+/g, " ").trim()).join(" · ") ?? ""
    const excerpt = firstRow.length > 120 ? `${firstRow.slice(0, 120)}…` : firstRow
    card.append(line("h3", "Add row"), line("p", `End of table ${proposal.table.tableIndex + 1}`, "chat-message-meta"))
    if (excerpt) card.append(line("p", `First row: ${excerpt}`, "chat-message-meta"))
    const list = document.createElement("ul"); list.className = "chat-format-plan"
    proposal.cells.forEach((text, index) => { const item = document.createElement("li"); item.append(line("strong", `Column ${index + 1}`), line("p", text)); list.append(item) })
    card.append(list)
    if (proposal.explanation) card.append(line("p", proposal.explanation))
    if (proposal.openItems.length) {
      card.append(line("h4", "Open items")); const items = document.createElement("ul"); items.className = "chat-open-items"
      for (const text of proposal.openItems) items.append(line("li", text)); card.append(items)
    }
    if (applied.has(proposal.id)) { const note = line("p", turn.proposalApplied ? "Previously applied in Word." : "Applied in Word. Save when you’re ready.", "chat-local-note"); note.setAttribute("role", "status"); card.append(note); return card }
    if (uncertain.has(proposal.id)) { const note = line("p", "Check the change in Word. Use Undo if needed.", "chat-proposal-error"); note.setAttribute("role", "alert"); card.append(note); return card }
    if (declined.has(proposal.id)) { card.append(line("p", "Declined. Your document is unchanged.", "chat-message-meta")); return card }
    const error = proposalErrors.get(proposal.id)
    if (error) { const note = line("p", error, "chat-proposal-error"); note.setAttribute("role", "alert"); card.append(note) }
    if (!host.applyRow) card.append(line("p", "Table row changes are not available in this Word session.", "chat-proposal-limit"))
    const actions = document.createElement("div"); actions.className = "chat-proposal-actions"
    const refine = document.createElement("button"); refine.type = "button"; refine.className = "chat-subtle"; refine.textContent = "Refine"; refine.disabled = applying
    refine.onclick = () => { if (disposed || applying || proposalFinished(proposal.id)) return; clearSelection(); refinement = undefined; rowRefinement = proposal; renderChips(); controls(); message.placeholder = "How should this row change?"; message.focus() }
    const apply = document.createElement("button"); apply.type = "button"; apply.className = "chat-send"; apply.textContent = "Apply"; apply.disabled = applying
    apply.onclick = async () => {
      if (disposed || applying || proposalFinished(proposal.id) || !host.applyRow) return
      applying = true; cancelSelectionCapture(); controls(); status("Applying in Word…")
      try {
        await host.applyRow(proposal, {conversationId:turn.conversationId,turnId:turn.id})
        if (disposed) return
        applied.add(proposal.id); proposalErrors.delete(proposal.id)
        refinement = undefined; rowRefinement = undefined; message.placeholder = "Ask doc.haus…"
        clearSelection(); renderChips(); status(); invalidateSuggestions(); renderTurns()
      } catch (error) {
        if (disposed) return
        if (error && typeof error === "object" && "mutationState" in error && error.mutationState === "uncertain") {
          uncertain.add(proposal.id); proposalErrors.delete(proposal.id)
          refinement = undefined; rowRefinement = undefined; message.placeholder = "Ask doc.haus…"
          clearSelection(); renderChips()
        } else proposalErrors.set(proposal.id, error instanceof Error ? error.message : "Word could not apply this change. Refine it and try again.")
        status(); renderTurns()
      } finally { applying = false; if (!disposed) apply.disabled = false; controls() }
    }
    const decline = document.createElement("button"); decline.type = "button"; decline.className = "chat-subtle"; decline.textContent = "Decline"; decline.disabled = applying
    decline.onclick = () => { if (disposed || applying || proposalFinished(proposal.id)) return; declined.add(proposal.id); proposalErrors.delete(proposal.id); if (rowRefinement?.id === proposal.id) { rowRefinement = undefined; message.placeholder = "Ask doc.haus…"; renderChips() }; renderTurns() }
    actions.append(refine); if (host.applyRow) actions.append(apply); actions.append(decline); card.append(actions); return card
  }
  function renderTurns(forceScroll = false) {
    if (disposed) return
    const nearBottom = scroller.scrollHeight - scroller.scrollTop - scroller.clientHeight < 90
    const visible = turns.filter(turn => turn.kind !== "suggestions")
    const live = new Set(visible.map(turn => turn.id))
    for (const [id, cached] of nodes) if (!live.has(id)) { cached.node.remove(); nodes.delete(id) }
    history.querySelector(".chat-empty")?.remove()
    if (!visible.length) {
      const empty = document.createElement("div"); empty.className = "chat-empty"
      const prompts = document.createElement("div"); prompts.className = "chat-empty-prompts"; prompts.setAttribute("aria-label", "Suggested prompts")
      empty.append(line("h2", "What are you working on?"), prompts); history.append(empty)
    }
    for (const turn of visible) {
      const proposal = turn.status === "complete" ? turn.proposal as WordChatProposal | undefined : undefined
      const row = turn.status === "complete" ? turn.rowProposal : undefined
      const selectedProposal = proposal ?? row
      const stamp = JSON.stringify([turn, selectedProposal && applied.has(selectedProposal.id), selectedProposal && declined.has(selectedProposal.id), selectedProposal && uncertain.has(selectedProposal.id), selectedProposal && proposalErrors.get(selectedProposal.id), proposal && host.applyEligibility(proposal), Boolean(host.applyRow)])
      let cached = nodes.get(turn.id)
      if (cached?.stamp === stamp) continue
      const wrapper = document.createElement("div"); wrapper.dataset.turn = turn.id
      const user = document.createElement("article"); user.className = "chat-message"; user.dataset.role = "user"
      user.append(line("p", turn.text), line("div", `${turn.scope === "document" ? "This document" : "Whole matter"}${turn.selection ? " · selected Word text" : ""}${turn.attachments.length ? ` · ${turn.attachments.length} attached` : ""}`, "chat-message-meta"))
      wrapper.append(user)
      const answer = document.createElement("article"); answer.className = "chat-message"; answer.dataset.role = "assistant"
      answer.append(line("strong", "doc.haus", "chat-author"))
      const sources = new Map<string, { source: WordChatSource; excerpts: WordChatSource[] }>()
      for (const source of turn.sources) {
        const group = sources.get(source.id)
        if (!group) sources.set(source.id, { source, excerpts: [source] })
        else if (!group.excerpts.some(excerpt => excerpt.quote === source.quote && excerpt.location === source.location && excerpt.provenance === source.provenance && excerpt.name === source.name)) group.excerpts.push(source)
      }
      const references = new Map([...sources.keys()].map((id, index) => [id, index + 1]))
      // Marked answers carry explicit references; literal text is never parsed
      // for bracket tokens. Legacy persisted turns retain their old renderer.
      const format: unknown = turn.answerFormat, parts: unknown = turn.answerParts
      const cited = new Set<string>()
      let malformed = false
      const suppressed = format !== undefined && (turn.status !== "complete" || Boolean(turn.error?.trim()))
      if (format !== undefined) {
        if (suppressed) { /* The service clears stale answer content while not complete. */ }
        else if (format !== "segments-v1" || !Array.isArray(parts)) malformed = true
        else {
          const paragraph = document.createElement("p"); paragraph.className = "chat-answer-text"
          for (const part of parts) {
            if (!part || typeof part !== "object" || Array.isArray(part)) { malformed = true; break }
            if (part.type === "text" && typeof part.text === "string" && Object.keys(part).every(key => ["type", "text"].includes(key))) paragraph.append(document.createTextNode(part.text))
            else if (part.type === "reference" && typeof part.sourceId === "string" && typeof part.quote === "string" && Object.keys(part).every(key => ["type", "sourceId", "quote"].includes(key))) {
              const group = sources.get(part.sourceId)
              const valid = group?.excerpts.some(excerpt => excerpt.quote === part.quote)
              if (valid) cited.add(part.sourceId)
              paragraph.append(document.createTextNode(valid ? `[${references.get(part.sourceId)}]` : "[Source unavailable]"))
            } else { malformed = true; break }
          }
          if (!malformed && paragraph.textContent) answer.append(paragraph)
        }
        if (malformed) { const notice = line("p", "This answer could not be displayed. Please try again.", "chat-message-meta"); notice.setAttribute("role", "alert"); answer.append(notice) }
      } else {
        const displayedAnswer = turn.answer.replace(/\[([^\]\r\n]+)\]/g, (token, id: string) => references.has(id) ? `[${references.get(id)}]` : token)
        if (displayedAnswer) answer.append(line("p", displayedAnswer, "chat-answer-text"))
      }
      if (turn.status === "running") answer.append(line("p", "Thinking…", "chat-message-meta"))
      if (turn.status === "error" || turn.status === "complete" && turn.error?.trim()) {
        const notice = line("p", turn.error || "This request could not be completed. Your Word document is unchanged.", "chat-message-meta")
        notice.setAttribute("role", "alert"); answer.append(notice)
      }
      if (turn.status === "cancelled") answer.append(line("p", "Stopped", "chat-message-meta"))
      const provenance = { "word-buffer": "Word text", "saved-version": "Saved source", instructions: "Playbook", "request-upload": "Chat attachment" }
      for (const [index, { source, excerpts }] of (malformed || suppressed ? [] : [...sources.values()]).entries()) {
        if (format !== undefined && !cited.has(source.id)) continue
        const detail = document.createElement("details"); detail.className = "chat-source"
        detail.append(line("summary", `[${index + 1}] ${source.name} · ${provenance[source.provenance]}`))
        for (const excerpt of excerpts) {
          detail.append(line("p", excerpt.quote), line("div", [excerpt.name !== source.name ? excerpt.name : "", provenance[excerpt.provenance], excerpt.location].filter(Boolean).join(" · "), "chat-message-meta"))
        }
        const url = host.sourceUrl?.(source)
        if (url) {
          const open = document.createElement("a"); open.href = url; open.target = "_blank"; open.rel = "noopener noreferrer"; open.textContent = "Open sources ↗"; open.setAttribute("aria-label", "Open current matter sources in doc.haus")
          detail.append(open)
        }
        answer.append(detail)
      }
      if (!malformed && !suppressed && proposal) answer.append(proposalCard(proposal, turn))
      else if (!malformed && !suppressed && row) answer.append(rowProposalCard(row, turn))
      wrapper.append(answer)
      if (cached) cached.node.replaceWith(wrapper); else history.append(wrapper)
      cached = { stamp, node: wrapper }; nodes.set(turn.id, cached)
    }
    renderSuggestions()
    if (nearBottom || forceScroll) scroller.scrollTop = scroller.scrollHeight
  }
  function updateTurn(turn: WordChatTurn) {
    if (turn.proposalApplied && turn.proposal) applied.add(turn.proposal.id)
    if (turn.proposalApplied && turn.rowProposal) applied.add(turn.rowProposal.id)
    const index = turns.findIndex(item => item.id === turn.id)
    if (index < 0) turns.push(turn); else turns[index] = turn
    if (turn.kind !== "suggestions") { foreground = turn.status === "running" ? turn : undefined; renderTurns() }
    controls()
  }
  async function refreshMaterials() {
    const conversationId = conversation?.id
    const result = await api.request<WordChatMaterials>("materials", { conversationId })
    if (disposed || conversation?.id !== conversationId) return
    const changed = JSON.stringify(materials) !== JSON.stringify(result.materials)
    materials = result.materials
    if (result.preferences) {
      const previousSuggestions = preferences.suggestions
      preferences = { defaultScope: result.preferences.defaultScope === "matter" ? "matter" : "document", suggestions: result.preferences.suggestions !== false }
      defaultScope = preferences.defaultScope
      if (!scopeChanged && !turns.some(turn => turn.kind !== "suggestions")) scope = defaultScope
      if (previousSuggestions !== preferences.suggestions) invalidateSuggestions()
      renderSuggestions(); controls()
    }
    if (changed) {
      attachments = attachments.map(item => item.kind === "upload" ? item : materials.find(next => next.kind === item.kind && next.id === item.id) ?? item)
      renderChips(); materialList(); invalidateSuggestions()
    }
    lastMaterials = now()
  }
  async function openConversation(id?: string) {
    if (disposed || conversationLocked()) return
    connecting = true; controls(); status("Connecting…")
    try {
      const result = await api.request<WordChatConversationResult>("conversation", id ? { conversationId: id } : {})
      if (disposed) return
      if (conversation?.id !== result.conversation.id) {
        history.querySelectorAll(".chat-idea").forEach(node => node.remove())
        scope = id ? [...result.turns].reverse().find(turn => turn.kind !== "suggestions")?.scope ?? defaultScope : defaultScope
      }
      conversation = result.conversation; turns = result.turns; for (const turn of turns) if (turn.proposalApplied && (turn.proposal || turn.rowProposal)) applied.add((turn.proposal ?? turn.rowProposal)!.id); foreground = turns.find(turn => turn.kind !== "suggestions" && turn.status === "running")
      const existing = conversations.findIndex(item => item.id === conversation!.id)
      if (existing < 0) conversations.unshift(conversation); else conversations[existing] = conversation
      options(); renderTurns(true); status(); button("reconnect").hidden = true
      startStream(); await refreshMaterials(); scheduleReconciliation()
    } catch (error) { failure(error) }
    finally { connecting = false; controls() }
  }
  function failure(error: unknown) {
    if (disposed) return
    status(error instanceof Error ? error.message : "Chat is unavailable. Reconnect to continue.", true); button("reconnect").hidden = false
  }
  async function connect() {
    if (disposed) return
    try {
      const result = await api.request<{ conversations: WordChatConversation[] }>("conversations")
      if (disposed) return
      conversations = result.conversations; options()
      const current = conversation?.id ?? conversations.find(item => item.documentName === host.documentName)?.id
      await openConversation(current)
    } catch (error) { failure(error) }
  }
  async function refreshConversationMetadata() {
    if (disposed || !conversation) return
    const id = conversation.id, epoch = conversationEpoch, request = ++metadataRequest
    try {
      const result = await api.request<WordChatConversationResult>("conversation", { conversationId: id })
      if (disposed || conversation?.id !== id || result.conversation.id !== id || epoch !== conversationEpoch || request !== metadataRequest) return
      conversation = result.conversation
      const index = conversations.findIndex(item => item.id === id)
      if (index < 0) conversations.unshift(conversation); else conversations[index] = conversation
      options()
    } catch { /* Metadata refresh is best-effort; the current chat stays usable. */ }
  }
  function scheduleReconciliation() {
    cancelReconcile?.()
    if (disposed || !foreground && suggestion?.status !== "running") return
    // SSE is the normal path. A missed event, suspended pane or unavailable
    // streaming host still reconciles against native session history.
    cancelReconcile = schedule(() => { void reconcileTurns() }, 15_000)
  }
  async function readTurn(id: string, conversationId: string, signal?: AbortSignal) {
    const epoch = foregroundEpoch, ideasEpoch = suggestionEpoch
    const result = await api.request<{ turn: WordChatTurn }>("poll", { conversationId, turnId: id }, signal)
    return { turn: result.turn, conversationId, epoch, ideasEpoch }
  }
  function receiveTurn(result: Awaited<ReturnType<typeof readTurn>>) {
    const { turn, conversationId, epoch, ideasEpoch } = result
    if (disposed || conversation?.id !== conversationId || turn.conversationId !== conversationId) return
    if (turn.kind === "suggestions") {
      if (suggestion?.id !== turn.id || ideasEpoch !== suggestionEpoch || !suggestionsEnabled()) return
      suggestion = turn; renderSuggestions(); controls()
    } else {
      if (stopping || foreground?.id !== turn.id || epoch !== foregroundEpoch) return
      updateTurn(turn)
      if (turn.status !== "running") { void refreshConversationMetadata(); status(turn.status === "error" ? turn.error : "", turn.status === "error"); button("reconnect").hidden = true }
    }
    scheduleReconciliation()
  }
  async function reconcileTurns() {
    if (disposed || !conversation) return
    const conversationId = conversation.id, epoch = foregroundEpoch, ideasEpoch = suggestionEpoch
    const active = [!stopping ? foreground : undefined, suggestion?.status === "running" ? suggestion : undefined].filter((turn): turn is WordChatTurn => Boolean(turn))
    await Promise.all(active.map(async current => {
      try { receiveTurn(await readTurn(current.id, conversationId)) }
      catch (error) { if (!disposed && conversation?.id === conversationId && foreground?.id === current.id && epoch === foregroundEpoch) failure(error) }
    }))
    if (!disposed && conversation?.id === conversationId && epoch === foregroundEpoch && ideasEpoch === suggestionEpoch) scheduleReconciliation()
  }
  function startStream() {
    stream?.abort(); streamEpoch++
    if (disposed || !conversation || !api.events) return
    const controller = new AbortController(), conversationId = conversation.id, epoch = streamEpoch
    stream = controller
    const current = () => !disposed && conversation?.id === conversationId && epoch === streamEpoch
    void subscribeNativeChatStream({
      connect: signal => api.events!(conversationId, signal), signal: controller.signal,
      acceptsSession: id => current() && (id === conversationId || !stopping && foreground?.id === id || suggestion?.id === id && suggestion.status === "running"),
      needsRefresh: event => event.type === "session.status" && nativeEventSession(event) !== conversationId,
      read: async (id, signal) => {
        const turnEpoch = foregroundEpoch
        try { return await readTurn(id, conversationId, signal) }
        catch (error) {
          if (current() && !stopping && foreground?.id === id && turnEpoch === foregroundEpoch) failure(error)
          throw error
        }
      },
      onSnapshot: (_id, result) => { if (current()) receiveTurn(result) },
      onReconnect: () => { if (current()) void reconcileTurns() },
      onDisconnect: () => { if (current() && foreground && !sending && !stopping && !applying) status("Reconnecting…") },
      onEvent: event => {
        if (event.type === "session.deleted") {
          if (nativeEventSession(event) === conversationId) { controller.abort(); failure(new Error("This conversation is no longer available. Start a new chat.")) }
          else if (foreground?.id === nativeEventSession(event)) { foregroundEpoch++; foreground = undefined; scheduleReconciliation(); controls(); status("This response is no longer available.") }
        } else if (nativeEventSession(event) === conversationId) { void reconcileTurns(); void refreshConversationMetadata() }
      },
    })
  }
  async function pollForeground() {
    if (disposed || !foreground || !conversation) return
    const id = foreground.id, conversationId = conversation.id, epoch = foregroundEpoch
    try {
      const result = await readTurn(id, conversationId)
      if (disposed || conversation?.id !== conversationId || foreground?.id !== id || epoch !== foregroundEpoch) return
      receiveTurn(result)
    } catch (error) { if (!disposed && foreground?.id === id && epoch === foregroundEpoch) failure(error) }
  }
  async function submit(retry = false) {
    if (disposed || connecting || sending || applying || uploading || pickerOpen || selectionPending || foreground?.status === "running" || !conversation || pending && !retry) return
    const text = message.value.trim()
    if (!retry && !text) return
    const requestScope = scope, requestSelection = rowRefinement ? undefined : refinement?.target ?? selection, requestProposal = refinement, requestRow = rowRefinement, requestMode = mode
    const requestAttachments = attachments.map(({ kind, id, hash }) => ({ kind, id, hash }))
    sending = true; controls(); status("Reading current Word text…"); invalidateSuggestions()
    try {
      if (!retry || !pending) {
        const context = await host.capture(false)
        if (disposed) return
        if (requestSelection && !wordChatTargetCurrent(context, requestSelection)) throw new Error("Selected text changed. Select it again or remove it before sending.")
        if (requestSelection) context.selection = requestSelection
        pending = { conversationId: conversation.id, requestId: crypto.randomUUID(), text, scope: requestScope, mode: requestMode, attachments: requestAttachments, context, ...(requestRow ? { kind:"amendment" as const, rowProposal:requestRow } : requestProposal ? { kind: requestProposal.before === requestProposal.after && (requestProposal.formatting || requestProposal.inlineFormatting) ? "formatting" : "amendment", proposal: requestProposal } : {}) }
      }
      status("Sending…")
      const submitted = pending
      const result = await api.request<{ turn: WordChatTurn }>("send", submitted)
      if (disposed) return
      conversation.updatedAt = now(); options()
      if (message.value.trim() === submitted.text) message.value = ""
      if (refinement?.id === submitted.proposal?.id) { refinement = undefined; message.placeholder = "Ask doc.haus…" }
      if (rowRefinement?.id === submitted.rowProposal?.id) { rowRefinement = undefined; message.placeholder = "Ask doc.haus…" }
      pending = undefined; renderChips()
      updateTurn(result.turn); renderTurns(true); status(); button("reconnect").hidden = true; scheduleReconciliation()
      if (result.turn.status !== "running") void refreshConversationMetadata()
    } catch (error) {
      // A definite admission refusal can be corrected. An uncertain/lost
      // response keeps its identifier until Reconnect reconciles it.
      if (error instanceof ChatRequestError && !error.retryable) pending = undefined
      failure(error)
    }
    finally { sending = false; controls() }
  }
  function invalidateSuggestions() {
    suggestionEpoch++; suggestionKey = ""; stableKey = ""; stableAt = 0
    const old = suggestion; suggestion = undefined; get("suggestion-content").replaceChildren()
    if (old?.status === "running") void api.request("cancel", { conversationId: old.conversationId, turnId: old.id }).catch(() => {})
    renderSuggestions(); controls()
  }
  function renderSuggestions() {
    const area = get("suggestions"), content = get("suggestion-content")
    const empty = history.querySelector<HTMLElement>(".chat-empty-prompts")
    const ideas = suggestion?.status === "complete" ? suggestion.suggestions?.map((idea, index) => ({ idea, index })).filter(({ index }) => !dismissed.has(`${suggestion!.id}:${index}`)) ?? [] : []
    if (empty) {
      empty.replaceChildren(); empty.hidden = !suggestionsEnabled()
      if (suggestionsEnabled()) {
        const starters = [
          { title: "Summarise this document", prompt: "Summarise this document and its key points." },
          { title: selection ? "Explain selected text" : "Explain the wording", prompt: selection ? "Explain the selected wording in plain language." : "Explain the key wording in this document in plain language." },
          { title: "Check for ambiguities", prompt: "Check this document for ambiguities, potential issues and missing information." },
        ]
        for (const { idea } of ideas.slice(0, 3)) {
          const prompt = document.createElement("button"); prompt.type = "button"; prompt.className = "chat-starter"; prompt.textContent = idea.title
          prompt.onclick = () => fillPrompt(`Let’s discuss: ${idea.title}\n${idea.explanation}`); empty.append(prompt)
        }
        for (const starter of starters.slice(0, 3 - Math.min(3, ideas.length))) {
          const prompt = document.createElement("button"); prompt.type = "button"; prompt.className = "chat-starter"; prompt.textContent = starter.title
          prompt.onclick = () => fillPrompt(starter.prompt); empty.append(prompt)
        }
        const coverage = suggestion?.suggestionCoverage
        if (ideas.length && coverage && coverage.reviewed < coverage.total) {
          const note = line("p", `${coverage.reviewed} of ${coverage.total} passage groups`, "chat-message-meta"); note.title = "These suggestions sample the document. This is not a complete document review."; empty.append(note)
        }
      }
    }
    area.hidden = !suggestionsEnabled() || reviewingAmendment() || history.hidden || !suggestion || Boolean(empty && suggestion.status !== "running")
    if (!suggestionsEnabled()) return
    content.replaceChildren()
    if (suggestion?.status === "running") {
      const skeleton = document.createElement("div"); skeleton.className = "chat-skeleton"; skeleton.setAttribute("role", "status"); skeleton.setAttribute("aria-label", "Preparing suggestions")
      content.append(skeleton); return
    }
    if (suggestion?.status === "error") { content.append(line("p", suggestion.error || "Suggestions are unavailable. Chat is ready.")); return }
    const coverage = suggestion?.suggestionCoverage
    if (coverage && coverage.reviewed < coverage.total) {
      const note = line("p", `${coverage.reviewed} of ${coverage.total} passage groups`); note.title = "This batch samples the document. It is not a complete document review."; content.append(note)
    }
    const suggestions = ideas
    if (!suggestions.length) { content.append(line("p", suggestion ? "No further suggestions for now." : "Waiting for a quiet moment…")); return }
    for (const { idea, index } of suggestions.slice(0, 3)) {
      const row = document.createElement("div"); row.className = "chat-suggestion-row"
      const card = document.createElement("button"); card.type = "button"; card.className = "chat-suggestion-card"; card.textContent = idea.title
      card.onclick = () => discussSuggestion(idea)
      const dismiss = document.createElement("button"); dismiss.type = "button"; dismiss.className = "chat-icon"; dismiss.textContent = "×"; dismiss.setAttribute("aria-label", `Dismiss ${idea.title}`)
      dismiss.onclick = () => { dismissed.add(`${suggestion!.id}:${index}`); renderSuggestions() }
      row.append(card, dismiss); content.append(row)
    }
  }
  function fillPrompt(text: string) {
    message.value = message.value.trim() ? `${message.value}\n\n${text}` : text
    message.focus(); controls()
  }
  function discussSuggestion(idea: WordChatSuggestion) {
    const card = document.createElement("article"); card.className = "chat-message chat-idea"
    card.append(line("strong", "Suggestion · doc.haus", "chat-author"), line("p", idea.explanation), line("div", suggestion?.scope === "matter" ? "Whole matter" : "This document", "chat-message-meta"))
    if (idea.passage) card.append(line("p", idea.passage, "chat-message-meta"))
    const ask = document.createElement("button"); ask.type = "button"; ask.className = "chat-subtle"; ask.textContent = "Discuss this"
    ask.onclick = () => fillPrompt(`Let’s discuss: ${idea.title}\n${idea.explanation}`)
    card.append(ask); history.querySelector(".chat-empty")?.remove(); history.append(card)
    // Opening an idea never applies wording or makes another model call.
    scroller.scrollTop = scroller.scrollHeight
  }
  async function monitor() {
    if (disposed) return
    const epoch = suggestionEpoch
    try {
      if (!conversation || connecting || applying || pickerOpen || uploading || monitorBusy) return
      monitorBusy = true
      if (now() - lastMaterials > 30_000) await refreshMaterials()
      if (!suggestionsEnabled() || reviewingAmendment() || sending || foreground?.status === "running") return
      const context = await host.capture(false)
      if (disposed || !suggestionsEnabled() || epoch !== suggestionEpoch) return
      if (selection) {
        if (!wordChatTargetCurrent(context, selection)) {
          clearSelection(); renderChips(); invalidateSuggestions(); return
        }
        context.selection = selection
      }
      const selected = context.selection
      const key = JSON.stringify([scope, attachments.map(({ kind, id, hash }) => [kind, id, hash]), context.text, context.paragraphs, selected ? [selected.text, selected.paragraphs, selected.startOffset, selected.endOffset] : null])
      if (key !== stableKey) { stableKey = key; stableAt = now(); if (suggestionKey && suggestionKey !== key) invalidateSuggestions(); stableKey = key; stableAt = now(); return }
      if (suggestion?.status === "running") return
      if (suggestionCalls >= 6 || key === suggestionKey || now() - stableAt < 3_000 || now() - lastSuggestion < 45_000) return
      const requestId = crypto.randomUUID()
      // Count admission attempts too. A lost reply must not start paid work
      // again for unchanged inputs under a fresh identifier.
      suggestionCalls++; lastSuggestion = now(); suggestionKey = key; controls()
      const result = await api.request<{ turn: WordChatTurn }>("suggest", { conversationId: conversation.id, requestId, text: "Prepare a few focused suggestions.", scope, attachments: attachments.map(({ kind, id, hash }) => ({ kind, id, hash })), context, kind: "suggestions" })
      if (disposed || epoch !== suggestionEpoch || !suggestionsEnabled()) {
        if (result.turn.status === "running") void api.request("cancel", { conversationId: result.turn.conversationId, turnId: result.turn.id }).catch(() => {})
        return
      }
      suggestion = result.turn; renderSuggestions(); controls(); scheduleReconciliation()
    } catch (error) {
      if (!disposed && suggestionsEnabled() && epoch === suggestionEpoch) { get("suggestion-content").replaceChildren(line("p", error instanceof Error ? error.message : "Suggestions are unavailable. Chat is ready.")); lastSuggestion = now() }
    } finally {
      monitorBusy = false
      if (!disposed) cancelMonitor = schedule(() => { void monitor() }, 2_000)
    }
  }

  get<HTMLFormElement>("form").onsubmit = event => { event.preventDefault(); void submit() }
  message.oninput = controls
  message.onkeydown = event => { if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) { event.preventDefault(); void submit() } }
  button("conversation").onclick = () => { if (!conversationLocked()) { options(); setPanel("conversations", get("conversations").hidden) } }
  button("conversations-close").onclick = () => { setPanel("conversations", false); button("conversation").focus() }
  button("attach").onclick = () => { void openAttachments() }
  button("mode-toggle").onclick = () => { if (!conversationLocked()) showModePicker(get("mode-picker").hidden) }
  button("mode-toggle").onkeydown = event => {
    if (["ArrowDown", "ArrowUp"].includes(event.key) && !conversationLocked()) { event.preventDefault(); showModePicker(true) }
  }
  for (const value of ["auto", "qa", "redline"] as const) {
    button(`mode-${value}`).onclick = () => { if (!conversationLocked()) { mode = value; controls(); showModePicker(false, true) } }
    button(`mode-${value}`).onkeydown = event => {
      const values = ["auto", "qa", "redline"] as const, at = values.indexOf(value)
      if (["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) {
        event.preventDefault()
        const next = event.key === "Home" ? 0 : event.key === "End" ? 2 : (at + (event.key === "ArrowDown" ? 1 : 2)) % 3
        button(`mode-${values[next]}`).focus()
      }
    }
  }
  button("materials-close").onclick = () => { setPanel("materials", false); message.focus() }
  for (const name of ["existing", "upload"] as const) {
    const tab = button(`material-${name}-tab`)
    tab.onclick = () => setMaterialTab(name, true)
    tab.onkeydown = event => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return
      event.preventDefault()
      setMaterialTab(event.key === "Home" ? "existing" : event.key === "End" ? "upload" : materialTab === "existing" ? "upload" : "existing", true)
    }
  }
  function changeConversation(id?: string) {
    if (disposed || conversationLocked() || !conversation) return
    if (!id) mode = "auto"
    conversationEpoch++; foregroundEpoch++; clearSelection(); invalidateSuggestions(); attachments = []; refinement = undefined; rowRefinement = undefined; message.placeholder = "Ask doc.haus…"
    suggestionsPaused = false; suggestionCalls = 0; lastSuggestion = -Infinity
    scopeChanged = false
    if (!id) scope = defaultScope
    renderChips(); setPanel("conversations", false); void openConversation(id)
  }
  button("new").onclick = () => changeConversation()
  for (const value of ["document", "matter"] as const) button(`scope-${value}`).onclick = () => { scope = value; scopeChanged = true; invalidateSuggestions(); controls() }
  const pause = () => { suggestionsPaused = true; invalidateSuggestions() }
  button("pause").onclick = pause
  button("reconnect").onclick = () => { if (pending) void submit(true); else if (foreground) void pollForeground(); else void connect() }
  button("stop").onclick = async () => {
    if (!foreground || stopping) return
    const current = foreground
    stopping = true; foregroundEpoch++; cancelReconcile?.(); controls(); status("Stopping…")
    try { const result = await api.request<{ turn: WordChatTurn }>("cancel", { conversationId: current.conversationId, turnId: current.id }); updateTurn(result.turn); status("Stopped") }
    catch (error) { failure(error) }
    finally { stopping = false; controls(); scheduleReconciliation() }
  }
  get<HTMLInputElement>("upload").onchange = async () => {
    const input = get<HTMLInputElement>("upload"), file = input.files?.[0]
    if (!file || !conversation || uploading) return
    const conversationId = conversation.id
    uploading = true; controls(); get("upload-message").textContent = "Uploading…"
    try {
      const mode = get<HTMLSelectElement>("upload-mode").value === "matter" ? "matter" : "request"
      const result = await api.upload<{ material: WordChatMaterial }>(conversationId, file, mode)
      if (disposed || conversation?.id !== conversationId) return
      attachMaterial(result.material); get("upload-message").textContent = mode === "matter" ? "Added to the matter and attached." : "Attached to this chat."
      if (mode === "matter") await refreshMaterials()
    } catch (error) { if (!disposed) get("upload-message").textContent = error instanceof Error ? error.message : "Upload failed. Try a supported file." }
    finally { uploading = false; input.value = ""; controls() }
  }
  root.addEventListener("keydown", event => {
    if (event.key !== "Escape") return
    if (!get("mode-picker").hidden) { showModePicker(false, true); event.preventDefault(); event.stopPropagation() }
    else if (!get("materials").hidden) { setPanel("materials", false); button("attach").focus() }
    else if (!get("conversations").hidden) { setPanel("conversations", false); button("conversation").focus() }
  })
  const onFocus = () => {
    if (!root.ownerDocument.hidden) {
      void reconcileTurns()
      if (conversation && !connecting && !applying) void refreshMaterials().catch(() => {})
    }
  }
  root.ownerDocument.addEventListener("visibilitychange", onFocus)
  root.ownerDocument.defaultView?.addEventListener("focus", onFocus)
  renderTurns(); controls()
  const ready = connect().then(() => {
    if (disposed) return
    try { stopWatchingSelection = host.watchSelection?.(selectionChanged) } catch { /* Older hosts still get the initial capture. */ }
    void captureSelection()
  })
  cancelMonitor = schedule(() => { void monitor() }, 2_000)
  return { ready, resetDocumentContext() {
    if (disposed) return
    clearSelection()
    refinement = undefined; rowRefinement = undefined; suggestionsPaused = true
    message.placeholder = "Ask doc.haus…"
    renderChips(); invalidateSuggestions()
  }, dispose() {
    root.ownerDocument.removeEventListener("click", dismissDocumentInfo)
    root.ownerDocument.removeEventListener("visibilitychange", onFocus)
    root.ownerDocument.defaultView?.removeEventListener("focus", onFocus)
    disposed = true; streamEpoch++; stream?.abort(); foregroundEpoch++; suggestionEpoch++; selectionCaptureEpoch++; pickerEpoch++; cancelSelection?.(); stopWatchingSelection?.(); cancelMonitor?.(); cancelReconcile?.(); root.replaceChildren()
  } }
}
