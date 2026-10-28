import type { WordChatConversation, WordChatConversationResult, WordChatMaterials, WordChatMaterial, WordChatScope, WordChatSend, WordChatTurn, WordChatSuggestion, WordChatSource } from "../../../dochaus/lib/word-chat-contract"
import type { WordChatProposal, WordChatSnapshot, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { ChatRequestError, type WordChatApi } from "./chat-client"

export type ChatHost = {
  documentName: string
  capture(includeSelection?: boolean): Promise<WordChatSnapshot>
  captureSelection(): Promise<WordChatTarget | undefined>
  apply(proposal: WordChatProposal, request: { conversationId: string; turnId: string }): Promise<unknown>
  sourceUrl?(source: WordChatSource): string | undefined
  now?: () => number
  schedule?: (run: () => void, delay: number) => () => void
}

/** Mount once. Hiding a feature tab never remounts or submits the conversation. */
export function mountWordChat(root: HTMLElement, api: WordChatApi, host: ChatHost) {
  root.innerHTML = `<div class="chat-toolbar">
    <select data-chat="conversation" aria-label="Conversation"><option>Connecting…</option></select>
    <button data-chat="new" class="chat-icon" type="button" title="New conversation" aria-label="New conversation">＋</button>
    <button data-chat="settings-toggle" class="chat-icon" type="button" title="Chat settings" aria-label="Chat settings" aria-expanded="false" aria-controls="chat-settings">⚙</button>
  </div>
  <p data-chat="document" class="chat-document"></p>
  <section data-chat="settings" id="chat-settings" class="chat-settings" aria-label="Chat settings" hidden>
    <h2>Chat settings</h2>
    <label>Chat context<select data-chat="scope"><option value="document">This document</option><option value="matter">Whole matter</option></select></label>
    <p>This document uses the open Word text and material you attach. Whole matter can also use the matter’s sources and playbook.</p>
    <label class="chat-toggle"><input data-chat="proactive" type="checkbox"> Prepare suggestions</label>
    <label>Suggestions about<select data-chat="target"><option value="document">Whole document</option><option value="selection">Current selection</option></select></label>
    <p>Off by default. When enabled, prepares up to 6 small batches for this document. Nothing changes until you review and Apply.</p>
    <button data-chat="settings-close" class="chat-subtle" type="button">Done</button>
  </section>
  <section data-chat="materials" class="chat-materials" aria-label="Attach material" hidden>
    <h2>Attach material</h2><p>Context for your next requests. The matter’s default playbook stays as it is.</p>
    <div data-chat="material-list" class="chat-material-list"></div>
    <div class="chat-upload"><label>Upload<select data-chat="upload-mode"><option value="request">For this chat only</option><option value="matter">Add to this matter</option></select>
    <input data-chat="upload" type="file" aria-label="Choose a supported document" accept=".docx,.pdf,.txt,.md,.rtf,.csv,.xlsx,.eml,.msg"></label>
    <p data-chat="upload-message"></p></div>
    <button data-chat="materials-close" class="chat-subtle" type="button">Done</button>
  </section>
  <div data-chat="history" class="chat-history" role="log" aria-label="Conversation" aria-live="off"></div>
  <section data-chat="suggestions" class="chat-suggestions" aria-label="Suggestions" hidden>
    <div class="chat-suggestions-heading"><h2>Suggestions</h2><button data-chat="pause" class="chat-subtle" type="button">Pause</button></div>
    <div data-chat="suggestion-content"></div>
  </section>
  <div class="chat-recovery"><p data-chat="activity" class="chat-activity" role="status" aria-live="polite"></p>
    <button data-chat="reconnect" class="chat-subtle" type="button" hidden>Reconnect</button>
    <button data-chat="stop" class="chat-subtle" type="button" hidden>Stop</button></div>
  <form data-chat="form" class="chat-composer">
    <div data-chat="chips" class="chat-chips"></div>
    <textarea data-chat="message" aria-label="Message doc.haus" placeholder="Ask doc.haus…" rows="2"></textarea>
    <div class="chat-composer-actions">
      <button data-chat="selection" class="chat-subtle" type="button" title="Include the passage selected in Word">Use selected text</button>
      <button data-chat="attach" class="chat-subtle" type="button" aria-expanded="false">＋ Attach</button>
      <button data-chat="send" class="chat-send" type="submit">Send <span aria-hidden="true">↑</span></button>
    </div>
  </form>
  <div class="chat-scope"><button data-chat="scope-summary" class="chat-subtle" type="button">This document</button><span data-chat="preparation-summary"></span></div>`
  const get = <T extends HTMLElement>(name: string) => root.querySelector(`[data-chat="${name}"]`) as T
  const button = (name: string) => get<HTMLButtonElement>(name)
  const message = get<HTMLTextAreaElement>("message")
  const history = get("history"), activity = get("activity")
  const select = get<HTMLSelectElement>("conversation")
  const scopeSelect = get<HTMLSelectElement>("scope")
  const proactive = get<HTMLInputElement>("proactive")
  const target = get<HTMLSelectElement>("target")
  const now = host.now ?? Date.now
  const schedule = host.schedule ?? ((run: () => void, delay: number) => { const id = setTimeout(run, delay); return () => clearTimeout(id) })
  let conversations: WordChatConversation[] = [], conversation: WordChatConversation | undefined
  let turns: WordChatTurn[] = [], materials: WordChatMaterial[] = [], attachments: WordChatMaterial[] = []
  let selection: WordChatTarget | undefined, scope: WordChatScope = "document"
  let disposed = false, connecting = false, sending = false, uploading = false, applying = false
  let pending: WordChatSend | undefined, foreground: WordChatTurn | undefined, suggestion: WordChatTurn | undefined
  let refinement: WordChatProposal | undefined, suggestionEpoch = 0, suggestionCalls = 0
  let selectionCaptureEpoch = 0
  let suggestionKey = "", stableKey = "", stableAt = 0, lastSuggestion = -Infinity, lastMaterials = 0
  let cancelMonitor: (() => void) | undefined, cancelPoll: (() => void) | undefined, monitorBusy = false
  const dismissed = new Set<string>(), applied = new Set<string>(), declined = new Set<string>()
  const nodes = new Map<string, { stamp: string; node: HTMLElement }>()
  get("document").textContent = host.documentName

  function status(text = "", error = false) {
    if (disposed) return
    activity.textContent = text; activity.dataset.error = String(error)
  }
  function controls() {
    if (disposed) return
    button("send").disabled = connecting || sending || applying || uploading || Boolean(pending) || !conversation || Boolean(foreground?.status === "running") || !message.value.trim()
    button("new").disabled = connecting || sending || applying || Boolean(pending) || Boolean(foreground?.status === "running")
    select.disabled = connecting || sending || applying || Boolean(pending) || Boolean(foreground?.status === "running")
    button("stop").hidden = foreground?.status !== "running"
    button("selection").disabled = sending || applying
    for (const action of root.querySelectorAll<HTMLButtonElement>(".chat-proposal-actions button")) action.disabled = applying
    button("attach").disabled = !conversation || connecting
    get("preparation-summary").textContent = proactive.checked ? `${suggestionCalls}/6 batches` : ""
    get("scope-summary").textContent = scope === "document" ? "This document" : "Whole matter"
    target.disabled = !proactive.checked
  }
  function options() {
    select.replaceChildren()
    for (const item of conversations) {
      const option = document.createElement("option"); option.value = item.id; option.textContent = item.title || "Conversation"
      select.append(option)
    }
    if (conversation) select.value = conversation.id
  }
  function setPanel(name: "settings" | "materials", open: boolean) {
    if (open) {
      const other = name === "settings" ? "materials" : "settings"
      get(other).hidden = true
      button(other === "settings" ? "settings-toggle" : "attach").setAttribute("aria-expanded", "false")
    }
    get(name).hidden = !open
    button(name === "settings" ? "settings-toggle" : "attach").setAttribute("aria-expanded", String(open))
    if (open) get(name === "settings" ? "scope" : "material-list").focus()
  }
  function line(tag: string, text: string, className?: string) {
    const element = document.createElement(tag); element.textContent = text
    if (className) element.className = className
    return element
  }
  function renderChips() {
    const chips = get("chips"); chips.replaceChildren()
    function chip(label: string, detail: string, remove: () => void) {
      const row = document.createElement("div"); row.className = "chat-chip"
      const details = document.createElement("details"); details.append(line("summary", label), line("p", detail))
      const close = document.createElement("button"); close.type = "button"; close.textContent = "×"; close.setAttribute("aria-label", `Remove ${label}`)
      close.onclick = remove; row.append(details, close); chips.append(row)
    }
    if (selection) chip("Selected Word text", selection.text, () => { selection = undefined; refinement = undefined; renderChips(); invalidateSuggestions() })
    for (const material of attachments) chip(material.name, `${material.kind === "playbook" ? "Playbook · " : ""}${material.provenance === "request-upload" ? "For this chat only" : "Saved matter material"}${material.description ? `\n${material.description}` : ""}`, () => { attachments = attachments.filter(item => item.id !== material.id || item.kind !== material.kind); renderChips(); invalidateSuggestions() })
    if (refinement) chip("Refining amendment", refinement.before, () => { refinement = undefined; renderChips() })
  }
  function materialList() {
    const list = get("material-list"); list.replaceChildren()
    if (!materials.length) list.append(line("p", "No material available. You can upload a file."))
    for (const material of materials) {
      const choice = document.createElement("button"); choice.type = "button"
      choice.append(line("span", material.name), line("small", material.kind === "playbook" ? "Playbook · instructions" : "Matter source · saved version"))
      choice.disabled = attachments.some(item => item.kind === material.kind && item.id === material.id)
      choice.onclick = () => {
        attachments.push(material); materialList(); renderChips(); invalidateSuggestions(); setPanel("materials", false); message.focus()
      }
      list.append(choice)
    }
  }
  function proposalCard(proposal: WordChatProposal, turn: WordChatTurn): HTMLElement {
    const card = document.createElement("section"); card.className = "chat-proposal"; card.setAttribute("aria-label", "Amendment to review")
    card.append(line("h3", "Review amendment"), line("h4", "Before"), line("pre", proposal.before, "chat-before"), line("h4", "After"), line("pre", proposal.after, "chat-after"))
    if (proposal.explanation) card.append(line("p", proposal.explanation))
    if (proposal.openItems.length) {
      card.append(line("h4", "Open items")); const items = document.createElement("ul"); items.className = "chat-open-items"
      for (const text of proposal.openItems) items.append(line("li", text)); card.append(items)
    }
    if (applied.has(proposal.id)) { card.append(line("p", "Applied in Word as tracked changes. Choose Save to publish it to the matter.", "chat-local-note")); return card }
    if (declined.has(proposal.id)) { card.append(line("p", "Declined. Your document is unchanged.", "chat-message-meta")); return card }
    const actions = document.createElement("div"); actions.className = "chat-proposal-actions"
    const refine = document.createElement("button"); refine.type = "button"; refine.className = "chat-subtle"; refine.textContent = "Refine"
    refine.disabled = applying
    refine.onclick = () => { if (applying) return; refinement = proposal; renderChips(); message.placeholder = "How should this change?"; message.focus() }
    const apply = document.createElement("button"); apply.type = "button"; apply.className = "chat-send"; apply.textContent = "Apply"; apply.disabled = applying
    apply.onclick = async () => {
      if (applying || applied.has(proposal.id)) return
      applying = true; selectionCaptureEpoch++; controls(); status("Applying in Word…")
      try {
        await host.apply(proposal, { conversationId: turn.conversationId, turnId: turn.id })
        if (disposed) return
        applied.add(proposal.id)
        const affected = new Set(proposal.target.paragraphs.map(paragraph => paragraph.id))
        if (selection && selection.paragraphs.some(paragraph => affected.has(paragraph.id))) selection = undefined
        if (refinement?.id === proposal.id || refinement?.target.paragraphs.some(paragraph => affected.has(paragraph.id))) {
          refinement = undefined; message.placeholder = "Ask doc.haus…"
        }
        renderChips(); status("Applied in Word. Save when you’re ready."); invalidateSuggestions(); renderTurns()
      }
      catch (error) { status(error instanceof Error ? error.message : "Word could not apply this amendment. Refresh it before trying again.", true) }
      finally { applying = false; if (!disposed) apply.disabled = false; controls() }
    }
    const decline = document.createElement("button"); decline.type = "button"; decline.className = "chat-subtle"; decline.textContent = "Decline"
    decline.disabled = applying
    decline.onclick = () => { if (applying) return; declined.add(proposal.id); if (refinement?.id === proposal.id) { refinement = undefined; renderChips() }; renderTurns() }
    actions.append(refine, apply, decline); card.append(actions); return card
  }
  function renderTurns(forceScroll = false) {
    if (disposed) return
    const nearBottom = history.scrollHeight - history.scrollTop - history.clientHeight < 90
    const visible = turns.filter(turn => turn.kind !== "suggestions")
    const live = new Set(visible.map(turn => turn.id))
    for (const [id, cached] of nodes) if (!live.has(id)) { cached.node.remove(); nodes.delete(id) }
    history.querySelector(".chat-empty")?.remove()
    if (!visible.length) {
      const empty = document.createElement("div"); empty.className = "chat-empty"; empty.append(line("h2", "What are you working on?"), line("p", "Ask about this document, explore wording or request a change.")); history.append(empty)
    }
    for (const turn of visible) {
      const proposal = turn.proposal as WordChatProposal | undefined
      const stamp = JSON.stringify([turn, proposal && applied.has(proposal.id), proposal && declined.has(proposal.id)])
      let cached = nodes.get(turn.id)
      if (cached?.stamp === stamp) continue
      const wrapper = document.createElement("div"); wrapper.dataset.turn = turn.id
      const user = document.createElement("article"); user.className = "chat-message"; user.dataset.role = "user"
      user.append(line("p", turn.text), line("div", `${turn.scope === "document" ? "This document" : "Whole matter"}${turn.selection ? " · selected Word text" : ""}${turn.attachments.length ? ` · ${turn.attachments.length} attached` : ""}`, "chat-message-meta"))
      wrapper.append(user)
      const answer = document.createElement("article"); answer.className = "chat-message"; answer.dataset.role = "assistant"
      answer.append(line("strong", "doc.haus", "chat-author"))
      if (turn.answer) answer.append(line("p", turn.answer))
      if (turn.status === "running") answer.append(line("p", "Thinking…", "chat-message-meta"))
      if (turn.status === "error") answer.append(line("p", turn.error || "This request could not be completed. Your Word document is unchanged.", "chat-message-meta"))
      if (turn.status === "cancelled") answer.append(line("p", "Stopped", "chat-message-meta"))
      for (const source of turn.sources) {
        const detail = document.createElement("details"); detail.className = "chat-source"
        detail.append(line("summary", `${source.name} · ${source.provenance === "word-buffer" ? "Word text" : source.location || "Source"}`), line("p", source.quote))
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
    if (disposed || connecting || sending || applying || foreground?.status === "running") return
    connecting = true; controls(); status("Connecting…")
    try {
      const result = await api.request<WordChatConversationResult>("conversation", id ? { conversationId: id } : {})
      if (disposed) return
      if (conversation?.id !== result.conversation.id) history.querySelectorAll(".chat-idea").forEach(node => node.remove())
      conversation = result.conversation; turns = result.turns; foreground = turns.find(turn => turn.kind !== "suggestions" && turn.status === "running")
      if (!conversations.some(item => item.id === conversation!.id)) conversations.unshift(conversation)
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
    const id = foreground.id, conversationId = conversation.id
    try {
      const result = await api.request<{ turn: WordChatTurn }>("poll", { conversationId, turnId: id })
      if (disposed || conversation?.id !== conversationId) return
      updateTurn(result.turn)
      if (result.turn.status !== "running") { status(result.turn.status === "error" ? result.turn.error : "", result.turn.status === "error"); button("reconnect").hidden = true }
      startPolling()
    } catch (error) { failure(error) }
  }
  async function submit(retry = false) {
    if (disposed || sending || applying || foreground?.status === "running" || !conversation || pending && !retry) return
    const text = message.value.trim()
    if (!retry && !text) return
    const requestScope = scope, requestSelection = selection, requestProposal = refinement
    const requestAttachments = attachments.map(({ kind, id, hash }) => ({ kind, id, hash }))
    sending = true; controls(); status("Reading current Word text…"); invalidateSuggestions()
    try {
      if (!retry || !pending) {
        const context = await host.capture(false)
        if (disposed) return
        if (requestSelection) context.selection = requestSelection
        pending = { conversationId: conversation.id, requestId: crypto.randomUUID(), text, scope: requestScope, attachments: requestAttachments, context, ...(requestProposal ? { kind: "amendment", proposal: requestProposal } : {}) }
      }
      status("Sending…")
      const submitted = pending
      const result = await api.request<{ turn: WordChatTurn }>("send", submitted)
      if (disposed) return
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
    area.hidden = !proactive.checked
    if (!proactive.checked) return
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
    const suggestions = suggestion?.suggestions?.map((idea, index) => ({ idea, index })).filter(({ index }) => !dismissed.has(`${suggestion!.id}:${index}`)) ?? []
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
  function discussSuggestion(idea: WordChatSuggestion) {
    const card = document.createElement("article"); card.className = "chat-message chat-idea"
    card.append(line("strong", "Suggestion · doc.haus", "chat-author"), line("p", idea.explanation), line("div", suggestion?.scope === "matter" ? "Whole matter" : "This document", "chat-message-meta"))
    if (idea.passage) card.append(line("p", idea.passage, "chat-message-meta"))
    const ask = document.createElement("button"); ask.type = "button"; ask.className = "chat-subtle"; ask.textContent = "Discuss this"
    ask.onclick = () => { message.value = `Let’s discuss this suggestion: ${idea.title}\n${idea.explanation}`; message.focus(); controls() }
    card.append(ask); history.querySelector(".chat-empty")?.remove(); history.append(card)
    // Opening an idea never applies wording or makes another model call.
    history.scrollTop = history.scrollHeight
  }
  async function monitor() {
    if (disposed) return
    const epoch = suggestionEpoch
    try {
      if (!proactive.checked || !conversation || sending || foreground?.status === "running" || connecting || applying || monitorBusy) return
      monitorBusy = true
      if (now() - lastMaterials > 30_000) await refreshMaterials()
      const context = await host.capture(target.value === "selection")
      if (disposed || !proactive.checked || epoch !== suggestionEpoch) return
      if (target.value === "selection" && !context.selection?.text.trim()) {
        if (suggestion || suggestionKey) invalidateSuggestions()
        renderSuggestions(); return
      }
      const selected = context.selection
      const key = JSON.stringify([scope, target.value, attachments.map(({ kind, id, hash }) => [kind, id, hash]), context.text, context.paragraphs, target.value === "selection" && selected ? [selected.text, selected.paragraphs, selected.startOffset, selected.endOffset] : null])
      if (key !== stableKey) { stableKey = key; stableAt = now(); if (suggestionKey && suggestionKey !== key) invalidateSuggestions(); stableKey = key; stableAt = now(); return }
      if (suggestion?.status === "running") {
        const result = await api.request<{ turn: WordChatTurn }>("poll", { conversationId: suggestion.conversationId, turnId: suggestion.id })
        if (disposed || epoch !== suggestionEpoch || !proactive.checked) return
        suggestion = result.turn; renderSuggestions(); return
      }
      if (suggestionCalls >= 6 || key === suggestionKey || now() - stableAt < 3_000 || now() - lastSuggestion < 45_000) return
      const requestId = crypto.randomUUID()
      // Count admission attempts too. A lost reply must not start paid work
      // again for unchanged inputs under a fresh identifier.
      suggestionCalls++; lastSuggestion = now(); suggestionKey = key; controls()
      const result = await api.request<{ turn: WordChatTurn }>("suggest", { conversationId: conversation.id, requestId, text: "Prepare a few focused suggestions.", scope, attachments: attachments.map(({ kind, id, hash }) => ({ kind, id, hash })), context, kind: "suggestions", suggestionTarget: target.value === "selection" ? "selection" : "document" })
      if (disposed || epoch !== suggestionEpoch || !proactive.checked) {
        if (result.turn.status === "running") void api.request("cancel", { conversationId: result.turn.conversationId, turnId: result.turn.id }).catch(() => {})
        return
      }
      suggestion = result.turn; renderSuggestions(); controls()
    } catch (error) {
      if (!disposed && proactive.checked && epoch === suggestionEpoch) { get("suggestion-content").replaceChildren(line("p", error instanceof Error ? error.message : "Suggestions are unavailable. Chat is ready.")); lastSuggestion = now() }
    } finally {
      monitorBusy = false
      if (!disposed) cancelMonitor = schedule(() => { void monitor() }, 2_000)
    }
  }

  get<HTMLFormElement>("form").onsubmit = event => { event.preventDefault(); void submit() }
  message.oninput = controls
  message.onkeydown = event => { if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) { event.preventDefault(); void submit() } }
  button("settings-toggle").onclick = () => setPanel("settings", get("settings").hidden)
  button("scope-summary").onclick = () => setPanel("settings", true)
  button("settings-close").onclick = () => { setPanel("settings", false); message.focus() }
  button("attach").onclick = () => { setPanel("materials", get("materials").hidden); materialList() }
  button("materials-close").onclick = () => { setPanel("materials", false); message.focus() }
  button("selection").onclick = async () => {
    if (applying) return
    const captureEpoch = ++selectionCaptureEpoch
    button("selection").disabled = true
    try {
      const selected = await host.captureSelection()
      if (disposed || applying || captureEpoch !== selectionCaptureEpoch) return
      if (!selected?.text.trim()) throw new Error("Select a passage in Word first.")
      selection = selected; renderChips(); invalidateSuggestions(); status(); message.focus()
    }
    catch (error) { if (!disposed && !applying && captureEpoch === selectionCaptureEpoch) status(error instanceof Error ? error.message : "Select a passage in Word first.", true) }
    finally { controls() }
  }
  button("new").onclick = () => { if (applying) return; selectionCaptureEpoch++; invalidateSuggestions(); selection = undefined; attachments = []; refinement = undefined; renderChips(); void openConversation() }
  select.onchange = () => { if (applying) { if (conversation) select.value = conversation.id; return }; selectionCaptureEpoch++; invalidateSuggestions(); selection = undefined; attachments = []; refinement = undefined; renderChips(); void openConversation(select.value) }
  scopeSelect.onchange = () => { scope = scopeSelect.value === "matter" ? "matter" : "document"; invalidateSuggestions(); controls() }
  proactive.onchange = () => { suggestionCalls = 0; lastSuggestion = -Infinity; invalidateSuggestions(); renderSuggestions(); controls() }
  target.onchange = () => invalidateSuggestions()
  button("pause").onclick = () => { proactive.checked = false; invalidateSuggestions() }
  button("reconnect").onclick = () => { if (pending) void submit(true); else if (foreground) void pollForeground(); else void connect() }
  button("stop").onclick = async () => {
    if (!foreground) return
    const current = foreground
    try { const result = await api.request<{ turn: WordChatTurn }>("cancel", { conversationId: current.conversationId, turnId: current.id }); updateTurn(result.turn); status("Stopped"); cancelPoll?.() }
    catch (error) { failure(error) }
  }
  get<HTMLInputElement>("upload").onchange = async () => {
    const input = get<HTMLInputElement>("upload"), file = input.files?.[0]
    if (!file || !conversation || uploading) return
    uploading = true; controls(); get("upload-message").textContent = "Uploading…"
    try {
      const mode = get<HTMLSelectElement>("upload-mode").value === "matter" ? "matter" : "request"
      const result = await api.upload<{ material: WordChatMaterial }>(conversation.id, file, mode)
      if (disposed) return
      attachments.push(result.material); renderChips(); invalidateSuggestions(); get("upload-message").textContent = mode === "matter" ? "Added to the matter and attached." : "Attached to this chat."
      if (mode === "matter") await refreshMaterials()
    } catch (error) { if (!disposed) get("upload-message").textContent = error instanceof Error ? error.message : "Upload failed. Try a supported file." }
    finally { uploading = false; input.value = ""; controls() }
  }
  root.addEventListener("keydown", event => {
    if (event.key !== "Escape") return
    if (!get("settings").hidden) { setPanel("settings", false); button("settings-toggle").focus() }
    else if (!get("materials").hidden) { setPanel("materials", false); button("attach").focus() }
  })
  renderTurns(); controls()
  const ready = connect()
  cancelMonitor = schedule(() => { void monitor() }, 2_000)
  return { ready, dispose() {
    disposed = true; suggestionEpoch++; cancelMonitor?.(); cancelPoll?.(); root.replaceChildren()
  } }
}
