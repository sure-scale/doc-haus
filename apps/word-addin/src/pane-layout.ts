/** Size Chat from the actual header/footer, and retain Save while its history
 * scrolls. No native document, draft or feature-tab state is changed here. */
export function setupPaneLayout(document: Document): () => void {
  const main = document.getElementById("save-panel")?.parentElement
  const header = document.getElementById("save-panel")
  const chat = document.getElementById("pane-chat")
  const toggle = document.getElementById("save-details-toggle")
  const details = document.getElementById("save-details")
  const actions = document.getElementById("actions")
  if (!main || !header || !chat || !toggle || !details || !actions) return () => {}
  let compact = false
  function showDetails(open: boolean) {
    details!.hidden = !open
    toggle!.setAttribute("aria-expanded", String(open))
  }
  function update() {
    const active = !chat!.hidden
    const position = chat!.querySelector<HTMLElement>('[data-chat="history"]')?.scrollTop ?? 0
    if (!active || position <= 0) compact = false
    else if (position > 24) compact = true
    const mode = String(active), reduced = String(compact)
    if (main!.dataset.chatActive !== mode) main!.dataset.chatActive = mode
    if (header!.dataset.compact !== reduced) header!.dataset.compact = reduced
    // Cached/outdated and unlinked panes have no Save action. Their opening
    // instructions must remain visible without asking the user to find them.
    if (actions!.hidden) showDetails(true)
  }
  const click = () => showDetails(details!.hidden)
  const keydown = (event: KeyboardEvent) => {
    if (event.key === "Escape" && !details!.hidden && header!.contains(event.target as Node)) {
      showDetails(false); toggle!.focus(); event.preventDefault()
    }
  }
  toggle.addEventListener("click", click)
  header.addEventListener("keydown", keydown)
  chat.addEventListener("scroll", update, true)
  const Observer = document.defaultView?.MutationObserver
  const observer = Observer && new Observer(update)
  observer?.observe(chat, { attributes: true, attributeFilter: ["hidden"], childList: true })
  observer?.observe(actions, { attributes: true, attributeFilter: ["hidden"] })
  update()
  return () => {
    observer?.disconnect(); toggle.removeEventListener("click", click); header.removeEventListener("keydown", keydown); chat.removeEventListener("scroll", update, true)
  }
}
