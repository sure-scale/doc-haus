/** Keep Save details above the persistent bottom bar, within the space below
 * the feature tabs. Native document, conversation and draft state stay intact. */
export function setupPaneLayout(document: Document): () => void {
  const bar = document.getElementById("save-panel")
  const tabs = document.getElementById("desk-entry")
  const toggle = document.getElementById("save-details-toggle")
  const close = document.getElementById("save-details-close")
  const details = document.getElementById("save-details")
  const actions = document.getElementById("actions")
  const matter = document.getElementById("matter")
  const fullMatter = document.getElementById("save-details-matter")
  if (!bar?.parentElement || !tabs || !toggle || !close || !details || !actions || !matter || !fullMatter) return () => {}
  const view = document.defaultView
  function positionDetails() {
    if (details!.hidden) return
    const name = matter!.title || matter!.textContent || "doc.haus"
    if (fullMatter!.textContent !== name) fullMatter!.textContent = name
    const space = Math.max(0, bar!.getBoundingClientRect().top - (tabs!.hidden ? 0 : tabs!.getBoundingClientRect().bottom) - 12)
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
  const matterObserver = view && new view.MutationObserver(positionDetails)
  matterObserver?.observe(matter, { attributes: true, attributeFilter: ["title"], childList: true, characterData: true, subtree: true })
  const resizeObserver = view?.ResizeObserver && new view.ResizeObserver(positionDetails)
  resizeObserver?.observe(bar); resizeObserver?.observe(tabs)
  openingInstructions()
  return () => {
    observer?.disconnect(); matterObserver?.disconnect(); resizeObserver?.disconnect()
    toggle.removeEventListener("click", click); close.removeEventListener("click", dismiss)
    document.removeEventListener("keydown", keydown); document.removeEventListener("click", outsideClick)
    view?.removeEventListener("resize", positionDetails)
  }
}
