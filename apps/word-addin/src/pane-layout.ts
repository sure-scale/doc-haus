/** Keep recovery above the breadcrumb/version row and below the persistent
 * tabs/Save row. Native document, conversation and draft state stay intact. */
export function setupPaneLayout(document: Document): () => void {
  const bar = document.getElementById("save-panel")
  const tabs = document.getElementById("desk-entry")
  const header = document.getElementById("pane-header")
  const toggle = document.getElementById("save-details-toggle")
  const close = document.getElementById("save-details-close")
  const details = document.getElementById("save-details")
  const actions = document.getElementById("actions")
  const matter = document.getElementById("matter")
  const back = document.getElementById("back")
  const chat = document.getElementById("pane-chat")
  const documentName = document.getElementById("save-document")
  const separator = document.getElementById("matter-path-separator")
  const fullMatter = document.getElementById("save-details-matter")
  if (!bar?.parentElement || !tabs || !toggle || !close || !details || !actions || !matter || !fullMatter) return () => {}
  const view = document.defaultView
  function updateIdentity() {
    const name = chat?.dataset.documentName ?? ""
    if (documentName) {
      if (documentName.textContent !== name) documentName.textContent = name
      if (documentName.title !== name) documentName.title = name
      if (documentName.hidden !== !name) documentName.hidden = !name
    }
    if (separator && separator.hidden !== !name) separator.hidden = !name
    const href = back?.getAttribute("href")
    if (href && matter!.getAttribute("href") !== href) matter!.setAttribute("href", href)
    else if (!href && matter!.hasAttribute("href")) matter!.removeAttribute("href")
    const fullName = name ? `${matter!.textContent} / ${name}` : matter!.title || matter!.textContent || "doc.haus"
    if (fullMatter!.textContent !== fullName) fullMatter!.textContent = fullName
  }
  function positionDetails() {
    if (details!.hidden) return
    updateIdentity()
    const top = header ? header.getBoundingClientRect().bottom : tabs!.hidden ? 0 : tabs!.getBoundingClientRect().bottom
    const space = Math.max(0, bar!.getBoundingClientRect().top - top - 12)
    const value = `${space}px`
    if (details!.style.getPropertyValue("--save-details-space") !== value) details!.style.setProperty("--save-details-space", value)
  }
  function showDetails(open: boolean, restoreFocus = false) {
    details!.hidden = !open
    toggle!.setAttribute("aria-expanded", String(open))
    if (open) positionDetails()
    else if (restoreFocus) toggle!.focus()
  }
  const click = () => showDetails(details!.hidden)
  const dismiss = () => showDetails(false, true)
  const keydown = (event: KeyboardEvent) => {
    if (event.key === "Escape" && !details!.hidden) {
      dismiss(); event.preventDefault()
    }
  }
  const outsideClick = (event: MouseEvent) => {
    if (!details!.hidden && !bar!.contains(event.target as Node)) showDetails(false)
  }
  // Cached/outdated and unlinked panes have no Save action. Show their opening
  // instructions immediately, even when the feature tabs are unavailable.
  const openingInstructions = () => { if (actions!.hidden) showDetails(true) }
  toggle.addEventListener("click", click)
  close.addEventListener("click", dismiss)
  document.addEventListener("keydown", keydown)
  document.addEventListener("click", outsideClick)
  view?.addEventListener("resize", positionDetails)
  const observer = view && new view.MutationObserver(() => { openingInstructions(); positionDetails() })
  observer?.observe(actions, { attributes: true, attributeFilter: ["hidden"] })
  const matterObserver = view && new view.MutationObserver(() => { updateIdentity(); positionDetails() })
  matterObserver?.observe(matter, { attributes: true, attributeFilter: ["title"], childList: true, characterData: true, subtree: true })
  if (chat) matterObserver?.observe(chat, { attributes: true, attributeFilter: ["data-document-name"] })
  if (back) matterObserver?.observe(back, { attributes: true, attributeFilter: ["href"] })
  const resizeObserver = view?.ResizeObserver && new view.ResizeObserver(positionDetails)
  resizeObserver?.observe(bar); resizeObserver?.observe(header ?? tabs)
  updateIdentity(); openingInstructions()
  return () => {
    observer?.disconnect(); matterObserver?.disconnect(); resizeObserver?.disconnect()
    toggle.removeEventListener("click", click); close.removeEventListener("click", dismiss)
    document.removeEventListener("keydown", keydown); document.removeEventListener("click", outsideClick)
    view?.removeEventListener("resize", positionDetails)
  }
}
