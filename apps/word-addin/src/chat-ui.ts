import type { WordChatConversation, WordChatConversationResult, WordChatMaterials, WordChatMaterial, WordChatScope, WordChatSend, WordChatTurn, WordChatSuggestion, WordChatSource } from "../../../dochaus/lib/word-chat-contract"
import type { WordChatProposal, WordChatSnapshot, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { wordChatApplyPlan, wordChatTargetCurrent } from "../../../dochaus/lib/word-chat-amendment-contract"
import { ChatRequestError, type WordChatApi } from "./chat-client"
import { readChatPreferences, writeChatPreferences } from "./chat-preferences"

export type ChatHost = {
  documentName: string
  capture(includeSelection?: boolean): Promise<WordChatSnapshot>
  captureSelection(): Promise<WordChatTarget | undefined>
  watchSelection?(onChange: () => void): () => void
  openAttachments?(options: { conversationId: string; attached: WordChatMaterial[]; onAttach(material: WordChatMaterial): void }): Promise<void>
  apply(proposal: WordChatProposal, request: { conversationId: string; turnId: string }): Promise<unknown>
  sourceUrl?(source: WordChatSource): string | undefined
  now?: () => number
  schedule?: (run: () => void, delay: number) => () => void
}

/** Mount once. Hiding a feature tab never remounts or submits the conversation. */
export function mountWordChat(root: HTMLElement, api: WordChatApi, host: ChatHost) {
  root.innerHTML = `<p data-chat="document" class="chat-document"></p>
  <section data-chat="settings" id="chat-settings" class="chat-settings" aria-label="Chat settings" hidden>
    <div class="chat-panel-heading"><h2>Chat settings</h2><button data-chat="settings-close" class="chat-subtle" type="button">Done</button></div>
    <label>New chats use<select data-chat="default-scope"><option value="document">This document</option><option value="matter">Whole matter</option></select></label>
    <label class="chat-toggle"><input data-chat="proactive" type="checkbox"> Suggestions by default</label>
    <div data-chat="suggestion-settings"><label>Suggest about<select data-chat="target"><option value="document">Whole document</option><option value="selection">Selected text</option></select></label>
      <button data-chat="pause-setting" class="chat-subtle" type="button">Pause in this chat</button><button data-chat="resume" class="chat-subtle" type="button" hidden>Resume in this chat</button></div>
    <p>Defaults are saved for all Word documents.</p>
  </section>
  <section data-chat="conversations" id="chat-conversations" class="chat-conversations" aria-label="Conversations" hidden>
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
  <div data-chat="history" class="chat-history" role="log" aria-label="Conversation" aria-live="off"></div>
  <section data-chat="suggestions" class="chat-suggestions" aria-label="Suggestions" hidden>
    <div class="chat-suggestions-heading"><h2>Suggestions</h2><button data-chat="pause" class="chat-subtle" type="button">Pause</button></div>
    <div data-chat="suggestion-content"></div>
  </section>
  <div class="chat-recovery"><p data-chat="activity" class="chat-activity" role="status" aria-live="polite"></p>
    <button data-chat="reconnect" class="chat-subtle" type="button" hidden>Reconnect</button>
    <button data-chat="stop" class="chat-subtle" type="button" hidden>Stop</button></div>
  <div class="chat-toolbar">
    <button data-chat="conversation" class="chat-subtle chat-history-toggle" type="button" aria-label="Conversations" aria-expanded="false" aria-controls="chat-conversations"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 12a9 9 0 1 0 2.7-6.4L3 8m0-5v5h5m4-1v5l3 2"/></svg><span>Chats</span></button>
    <button data-chat="new" class="chat-subtle" type="button" title="New conversation" aria-label="New conversation">＋ New chat</button>
    <button data-chat="settings-toggle" class="chat-icon" type="button" title="Chat settings" aria-label="Chat settings" aria-expanded="false" aria-controls="chat-settings"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="m9.5 3-.6 2.2-1.4.8-2.2-.6-2.5 4.2 1.6 1.6v1.6L2.8 14l2.5 4.3 2.2-.6 1.4.8.6 2.5h5l.6-2.5 1.4-.8 2.2.6 2.5-4.3-1.6-1.2v-1.6l1.6-1.6-2.5-4.2-2.2.6-1.4-.8L14.5 3Z"/><circle cx="12" cy="12" r="3"/></svg><span>Settings</span></button>
  </div>
  <form data-chat="form" class="chat-composer">
    <div data-chat="chips" class="chat-chips"></div>
    <div class="chat-scope" role="group" aria-label="Chat context"><button data-chat="scope-document" type="button" aria-pressed="true" title="Use this Word document and attached material">This document</button><button data-chat="scope-matter" type="button" aria-pressed="false" title="Also use the matter’s sources and playbook">Whole matter</button></div>
    <textarea data-chat="message" aria-label="Message doc.haus" placeholder="Ask doc.haus…" rows="2"></textarea>
    <div class="chat-composer-actions">
      <button data-chat="attach" class="chat-subtle" type="button" aria-expanded="false">＋ Attach</button>
      <button data-chat="send" class="chat-send" type="submit">Send <span aria-hidden="true">↑</span></button>
    </div>
  </form>`
  const get = <T extends HTMLElement>(name: string) => root.querySelector(`[data-chat="${name}"]`) as T
  const button = (name: string) => get<HTMLButtonElement>(name)
  const message = get<HTMLTextAreaElement>("message")
  const history = get("history"), activity = get("activity")
  const proactive = get<HTMLInputElement>("proactive")
  const target = get<HTMLSelectElement>("target")
  const now = host.now ?? Date.now
  const schedule = host.schedule ?? ((run: () => void, delay: number) => { const id = setTimeout(run, delay); return () => clearTimeout(id) })
  const storage = (() => { try { return root.ownerDocument.defaultView?.localStorage } catch { return undefined } })()
  const preferences = readChatPreferences(storage)
  proactive.checked = preferences.suggestions; target.value = preferences.target
  get<HTMLSelectElement>("default-scope").value = preferences.scope
  let conversations: WordChatConversation[] = [], conversation: WordChatConversation | undefined
  let turns: WordChatTurn[] = [], materials: WordChatMaterial[] = [], attachments: WordChatMaterial[] = []
  let selection: WordChatTarget | undefined, scope: WordChatScope = preferences.scope
  let defaultScope = preferences.scope
  let disposed = false, connecting = false, sending = false, uploading = false, applying = false, stopping = false, pickerOpen = false, suggestionsPaused = false
  let pending: WordChatSend | undefined, foreground: WordChatTurn | undefined, suggestion: WordChatTurn | undefined
  let refinement: WordChatProposal | undefined, suggestionEpoch = 0, suggestionCalls = 0
  let selectionCaptureEpoch = 0, selectionPending = false
  let suppressedSelectionKey = "", pickerEpoch = 0
  let cancelSelection: (() => void) | undefined, stopWatchingSelection: (() => void) | undefined
  let foregroundEpoch = 0
  let materialTab: "existing" | "upload" = "existing"
  let suggestionKey = "", stableKey = "", stableAt = 0, lastSuggestion = -Infinity, lastMaterials = 0
  let cancelMonitor: (() => void) | undefined, cancelPoll: (() => void) | undefined, monitorBusy = false
  const dismissed = new Set<string>(), applied = new Set<string>(), declined = new Set<string>()
  const proposalErrors = new Map<string, string>()
  const nodes = new Map<string, { stamp: string; node: HTMLElement }>()
  get("document").textContent = host.documentName

  function status(text = "", error = false) {
    if (disposed) return
    activity.textContent = text; activity.dataset.error = String(error)
  }
  const suggestionsEnabled = () => proactive.checked && !suggestionsPaused
  const reviewingAmendment = () => turns.some(turn => turn.status === "complete" && turn.proposal && !applied.has(turn.proposal.id) && !declined.has(turn.proposal.id))
  const conversationLocked = () => connecting || sending || applying || uploading || pickerOpen || Boolean(pending) || foreground?.status === "running"
  function savePreferences() {
    writeChatPreferences(storage, { scope: defaultScope, suggestions: proactive.checked, target: target.value === "selection" ? "selection" : "document" })
  }
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
    for (const value of ["document", "matter"] as const) button(`scope-${value}`).setAttribute("aria-pressed", String(scope === value))
    target.disabled = !proactive.checked
    get("suggestion-settings").hidden = !proactive.checked
    button("pause-setting").hidden = suggestionsPaused
    button("resume").hidden = !suggestionsPaused
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
  function setPanel(name: "settings" | "materials" | "conversations", open: boolean) {
    const toggle = (panel: typeof name) => panel === "settings" ? "settings-toggle" : panel === "materials" ? "attach" : "conversation"
    if (open) {
      for (const other of ["settings", "materials", "conversations"] as const) if (other !== name) {
        get(other).hidden = true; button(toggle(other)).setAttribute("aria-expanded", "false")
      }
    }
    get(name).hidden = !open
    button(toggle(name)).setAttribute("aria-expanded", String(open))
    // Attachments use the conversation's space, with one scrolling body.
    // The composer and its draft remain mounted underneath.
    history.hidden = !get("materials").hidden || !get("conversations").hidden
    renderSuggestions()
    if (open) {
      if (name === "conversations") (get("conversation-list").querySelector<HTMLElement>("button") ?? button("conversations-close")).focus()
      else get(name === "settings" ? "default-scope" : `material-${materialTab}-tab`).focus()
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
    if (selection) chip("Selected Word text", selection.text, () => { clearSelection(); refinement = undefined; message.placeholder = "Ask doc.haus…"; renderChips(); invalidateSuggestions() })
    for (const material of attachments) chip(material.name, `${material.kind === "playbook" ? "Playbook · " : ""}${material.provenance === "request-upload" ? "For this chat only" : "Saved matter material"}${material.description ? `\n${material.description}` : ""}`, () => { attachments = attachments.filter(item => item.id !== material.id || item.kind !== material.kind); renderChips(); invalidateSuggestions() })
    if (refinement) chip("Refining amendment", refinement.before, () => { refinement = undefined; renderChips() })
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
    setPanel("settings", false); setPanel("conversations", false); setPanel("materials", false)
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
  function proposalCard(proposal: WordChatProposal, turn: WordChatTurn): HTMLElement {
    const card = document.createElement("section"); card.className = "chat-proposal"; card.setAttribute("aria-label", "Amendment to review")
    card.append(line("h3", "Review amendment"), line("h4", "Before"), line("pre", displayWordText(proposal.before), "chat-before"), line("h4", "After"), line("pre", displayWordText(proposal.after), "chat-after"))
    if (proposal.explanation) card.append(line("p", proposal.explanation))
    if (proposal.openItems.length) {
      card.append(line("h4", "Open items")); const items = document.createElement("ul"); items.className = "chat-open-items"
      for (const text of proposal.openItems) items.append(line("li", text)); card.append(items)
    }
    if (applied.has(proposal.id)) { card.append(line("p", "Applied in Word as tracked changes. Choose Save to publish it to the matter.", "chat-local-note")); return card }
    if (declined.has(proposal.id)) { card.append(line("p", "Declined. Your document is unchanged.", "chat-message-meta")); return card }
    const error = proposalErrors.get(proposal.id)
    if (error) { const note = line("p", error, "chat-proposal-error"); note.setAttribute("role", "alert"); card.append(note) }
    const plan = wordChatApplyPlan(proposal.target, proposal.after)
    if (!plan.supported) card.append(line("p", plan.reason, "chat-proposal-limit"))
    const actions = document.createElement("div"); actions.className = "chat-proposal-actions"
    const refine = document.createElement("button"); refine.type = "button"; refine.className = "chat-subtle"; refine.textContent = "Refine"
    refine.disabled = applying
    refine.onclick = () => { if (applying) return; cancelSelectionCapture(); refinement = proposal; renderChips(); controls(); message.placeholder = "How should this change?"; message.focus() }
    const apply = document.createElement("button"); apply.type = "button"; apply.className = "chat-send"; apply.textContent = "Apply"; apply.disabled = applying
    apply.onclick = async () => {
      if (applying || applied.has(proposal.id)) return
      applying = true; cancelSelectionCapture(); controls(); status("Applying in Word…")
      try {
        await host.apply(proposal, { conversationId: turn.conversationId, turnId: turn.id })
        if (disposed) return
        applied.add(proposal.id); proposalErrors.delete(proposal.id)
        const affected = new Set(proposal.target.paragraphs.map(paragraph => paragraph.id))
        if (selection && selection.paragraphs.some(paragraph => affected.has(paragraph.id))) clearSelection()
        if (refinement?.id === proposal.id || refinement?.target.paragraphs.some(paragraph => affected.has(paragraph.id))) {
          refinement = undefined; message.placeholder = "Ask doc.haus…"
        }
        renderChips(); status("Applied in Word. Save when you’re ready."); invalidateSuggestions(); renderTurns()
      }
      catch (error) {
        if (disposed) return
        proposalErrors.set(proposal.id, error instanceof Error ? error.message : "Word could not apply this change. Refine it and try again.")
        status(); renderTurns()
      }
      finally { applying = false; if (!disposed) apply.disabled = false; controls() }
    }
    const decline = document.createElement("button"); decline.type = "button"; decline.className = "chat-subtle"; decline.textContent = "Decline"
    decline.disabled = applying
    decline.onclick = () => { if (applying) return; declined.add(proposal.id); proposalErrors.delete(proposal.id); if (refinement?.id === proposal.id) { refinement = undefined; renderChips() }; renderTurns() }
    actions.append(refine)
    if (plan.supported) actions.append(apply)
    actions.append(decline); card.append(actions); return card
  }
  function renderTurns(forceScroll = false) {
    if (disposed) return
    const nearBottom = history.scrollHeight - history.scrollTop - history.clientHeight < 90
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
      const stamp = JSON.stringify([turn, proposal && applied.has(proposal.id), proposal && declined.has(proposal.id), proposal && proposalErrors.get(proposal.id)])
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
      const displayedAnswer = turn.answer.replace(/\[([^\]\r\n]+)\]/g, (token, id: string) => references.has(id) ? `[${references.get(id)}]` : token)
      if (displayedAnswer) answer.append(line("p", displayedAnswer, "chat-answer-text"))
      if (turn.status === "running") answer.append(line("p", "Thinking…", "chat-message-meta"))
      if (turn.status === "error" || turn.status === "complete" && turn.error?.trim()) {
        const notice = line("p", turn.error || "This request could not be completed. Your Word document is unchanged.", "chat-message-meta")
        notice.setAttribute("role", "alert"); answer.append(notice)
      }
      if (turn.status === "cancelled") answer.append(line("p", "Stopped", "chat-message-meta"))
      const provenance = { "word-buffer": "Word text", "saved-version": "Saved source", instructions: "Playbook", "request-upload": "Chat attachment" }
      for (const [index, { source, excerpts }] of [...sources.values()].entries()) {
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
      if (proposal) answer.append(proposalCard(proposal, turn))
      wrapper.append(answer)
      if (cached) cached.node.replaceWith(wrapper); else history.append(wrapper)
      cached = { stamp, node: wrapper }; nodes.set(turn.id, cached)
    }
    renderSuggestions()
    if (nearBottom || forceScroll) history.scrollTop = history.scrollHeight
  }
  function updateTurn(turn: WordChatTurn) {
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
      conversation = result.conversation; turns = result.turns; foreground = turns.find(turn => turn.kind !== "suggestions" && turn.status === "running")
      const existing = conversations.findIndex(item => item.id === conversation!.id)
      if (existing < 0) conversations.unshift(conversation); else conversations[existing] = conversation
      options(); renderTurns(true); status(); button("reconnect").hidden = true
      await refreshMaterials(); startPolling()
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
  function startPolling() {
    cancelPoll?.()
    if (disposed || !foreground) return
    cancelPoll = schedule(() => { void pollForeground() }, 1_000)
  }
  async function pollForeground() {
    if (disposed || !foreground || !conversation) return
    const id = foreground.id, conversationId = conversation.id, epoch = foregroundEpoch
    try {
      const result = await api.request<{ turn: WordChatTurn }>("poll", { conversationId, turnId: id })
      if (disposed || conversation?.id !== conversationId || foreground?.id !== id || epoch !== foregroundEpoch) return
      updateTurn(result.turn)
      if (result.turn.status !== "running") { status(result.turn.status === "error" ? result.turn.error : "", result.turn.status === "error"); button("reconnect").hidden = true }
      startPolling()
    } catch (error) { if (!disposed && foreground?.id === id && epoch === foregroundEpoch) failure(error) }
  }
  async function submit(retry = false) {
    if (disposed || connecting || sending || applying || uploading || pickerOpen || selectionPending || foreground?.status === "running" || !conversation || pending && !retry) return
    const text = message.value.trim()
    if (!retry && !text) return
    const requestScope = scope, requestSelection = refinement?.target ?? selection, requestProposal = refinement
    const requestAttachments = attachments.map(({ kind, id, hash }) => ({ kind, id, hash }))
    sending = true; controls(); status("Reading current Word text…"); invalidateSuggestions()
    try {
      if (!retry || !pending) {
        const context = await host.capture(false)
        if (disposed) return
        if (requestSelection && !wordChatTargetCurrent(context, requestSelection)) throw new Error("Selected text changed. Select it again or remove it before sending.")
        if (requestSelection) context.selection = requestSelection
        pending = { conversationId: conversation.id, requestId: crypto.randomUUID(), text, scope: requestScope, attachments: requestAttachments, context, ...(requestProposal ? { kind: "amendment", proposal: requestProposal } : {}) }
      }
      status("Sending…")
      const submitted = pending
      const result = await api.request<{ turn: WordChatTurn }>("send", submitted)
      if (disposed) return
      if ([`Word · ${conversation.documentName}`, "Word conversation"].includes(conversation.title)) conversation.title = submitted.text.replace(/\s+/g, " ").slice(0, 100)
      conversation.updatedAt = now(); options()
      if (message.value.trim() === submitted.text) message.value = ""
      if (refinement?.id === submitted.proposal?.id) { refinement = undefined; message.placeholder = "Ask doc.haus…" }
      pending = undefined; renderChips()
      updateTurn(result.turn); renderTurns(true); status(); button("reconnect").hidden = true; startPolling()
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
    if (!suggestions.length) { content.append(line("p", suggestion ? "No further suggestions for now." : suggestionCalls >= 6 ? "Batch limit reached. Pause and re-enable to prepare more." : target.value === "selection" ? "Waiting for a stable selection in Word." : "Waiting for a quiet moment…")); return }
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
    history.scrollTop = history.scrollHeight
  }
  async function monitor() {
    if (disposed) return
    const epoch = suggestionEpoch
    try {
      if (!suggestionsEnabled() || reviewingAmendment() || !conversation || sending || foreground?.status === "running" || connecting || applying || pickerOpen || uploading || monitorBusy) return
      monitorBusy = true
      if (now() - lastMaterials > 30_000) await refreshMaterials()
      const context = await host.capture(false)
      if (disposed || !suggestionsEnabled() || epoch !== suggestionEpoch) return
      if (target.value === "selection" && selection) {
        if (!wordChatTargetCurrent(context, selection)) {
          clearSelection(); renderChips(); invalidateSuggestions(); return
        }
        context.selection = selection
      }
      if (target.value === "selection" && !context.selection?.text.trim()) {
        if (suggestion || suggestionKey) invalidateSuggestions()
        renderSuggestions(); return
      }
      const selected = context.selection
      const key = JSON.stringify([scope, target.value, attachments.map(({ kind, id, hash }) => [kind, id, hash]), context.text, context.paragraphs, target.value === "selection" && selected ? [selected.text, selected.paragraphs, selected.startOffset, selected.endOffset] : null])
      if (key !== stableKey) { stableKey = key; stableAt = now(); if (suggestionKey && suggestionKey !== key) invalidateSuggestions(); stableKey = key; stableAt = now(); return }
      if (suggestion?.status === "running") {
        const result = await api.request<{ turn: WordChatTurn }>("poll", { conversationId: suggestion.conversationId, turnId: suggestion.id })
        if (disposed || epoch !== suggestionEpoch || !suggestionsEnabled()) return
        suggestion = result.turn; renderSuggestions(); return
      }
      if (suggestionCalls >= 6 || key === suggestionKey || now() - stableAt < 3_000 || now() - lastSuggestion < 45_000) return
      const requestId = crypto.randomUUID()
      // Count admission attempts too. A lost reply must not start paid work
      // again for unchanged inputs under a fresh identifier.
      suggestionCalls++; lastSuggestion = now(); suggestionKey = key; controls()
      const result = await api.request<{ turn: WordChatTurn }>("suggest", { conversationId: conversation.id, requestId, text: "Prepare a few focused suggestions.", scope, attachments: attachments.map(({ kind, id, hash }) => ({ kind, id, hash })), context, kind: "suggestions", suggestionTarget: target.value === "selection" ? "selection" : "document" })
      if (disposed || epoch !== suggestionEpoch || !suggestionsEnabled()) {
        if (result.turn.status === "running") void api.request("cancel", { conversationId: result.turn.conversationId, turnId: result.turn.id }).catch(() => {})
        return
      }
      suggestion = result.turn; renderSuggestions(); controls()
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
  button("settings-toggle").onclick = () => setPanel("settings", get("settings").hidden)
  button("settings-close").onclick = () => { setPanel("settings", false); message.focus() }
  button("conversation").onclick = () => { if (!conversationLocked()) { options(); setPanel("conversations", get("conversations").hidden) } }
  button("conversations-close").onclick = () => { setPanel("conversations", false); button("conversation").focus() }
  button("attach").onclick = () => { void openAttachments() }
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
    foregroundEpoch++; clearSelection(); invalidateSuggestions(); attachments = []; refinement = undefined
    suggestionsPaused = false; suggestionCalls = 0; lastSuggestion = -Infinity
    if (!id) scope = defaultScope
    renderChips(); setPanel("conversations", false); void openConversation(id)
  }
  button("new").onclick = () => changeConversation()
  for (const value of ["document", "matter"] as const) button(`scope-${value}`).onclick = () => { scope = value; invalidateSuggestions(); controls() }
  get<HTMLSelectElement>("default-scope").onchange = () => { defaultScope = get<HTMLSelectElement>("default-scope").value === "matter" ? "matter" : "document"; savePreferences() }
  proactive.onchange = () => { suggestionsPaused = false; suggestionCalls = 0; lastSuggestion = -Infinity; savePreferences(); invalidateSuggestions() }
  target.onchange = () => { savePreferences(); invalidateSuggestions() }
  const pause = () => { suggestionsPaused = true; invalidateSuggestions() }
  button("pause").onclick = pause; button("pause-setting").onclick = pause
  button("resume").onclick = () => { suggestionsPaused = false; suggestionCalls = 0; lastSuggestion = -Infinity; invalidateSuggestions() }
  button("reconnect").onclick = () => { if (pending) void submit(true); else if (foreground) void pollForeground(); else void connect() }
  button("stop").onclick = async () => {
    if (!foreground || stopping) return
    const current = foreground
    stopping = true; foregroundEpoch++; cancelPoll?.(); controls(); status("Stopping…")
    try { const result = await api.request<{ turn: WordChatTurn }>("cancel", { conversationId: current.conversationId, turnId: current.id }); updateTurn(result.turn); status("Stopped"); cancelPoll?.() }
    catch (error) { failure(error) }
    finally { stopping = false; controls() }
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
    if (!get("settings").hidden) { setPanel("settings", false); button("settings-toggle").focus() }
    else if (!get("materials").hidden) { setPanel("materials", false); button("attach").focus() }
    else if (!get("conversations").hidden) { setPanel("conversations", false); button("conversation").focus() }
  })
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
    refinement = undefined; suggestionsPaused = true
    message.placeholder = "Ask doc.haus…"
    renderChips(); invalidateSuggestions()
  }, dispose() {
    disposed = true; suggestionEpoch++; selectionCaptureEpoch++; pickerEpoch++; cancelSelection?.(); stopWatchingSelection?.(); cancelMonitor?.(); cancelPoll?.(); root.replaceChildren()
  } }
}
