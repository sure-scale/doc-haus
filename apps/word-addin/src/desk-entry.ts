import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { deskRequestHandler } from "./desk-client"
import { openDeskDialog } from "./desk-host"
import { mountDesk } from "./desk-ui"
import type { DeskView } from "./desk-protocol"

/** Read-only views stay in the pane; a separate window is an optional wider view. */
export function setupSourceDesk(link: WordDocumentLink): void {
  const entry = document.getElementById("desk-entry")
  const documents = document.getElementById("documents-tab") as HTMLButtonElement | null
  const compare = document.getElementById("compare-tab") as HTMLButtonElement | null
  const launch = document.getElementById("open-desk-window") as HTMLButtonElement | null
  const panel = document.getElementById("pane-desk")
  const message = document.getElementById("desk-message")
  if (!entry || !documents || !compare || !launch || !panel || !message) return

  const request = deskRequestHandler(link)
  const tabs = { documents, compare }
  let view: DeskView = "documents"
  let busy = false
  let opening = false
  let disposed = false
  let dialog: { close(): void } | undefined
  function controls() {
    for (const [name, tab] of Object.entries(tabs)) {
      tab.disabled = busy
      tab.setAttribute("aria-selected", String(view === name))
      tab.tabIndex = view === name ? 0 : -1
    }
    panel!.setAttribute("aria-labelledby", tabs[view].id)
    launch!.disabled = busy || opening || Boolean(dialog)
    launch!.setAttribute("aria-label", `Open ${view === "documents" ? "documents" : "comparison"} in a separate window`)
    launch!.title = launch!.getAttribute("aria-label")!
  }
  function error(text = "") {
    message!.textContent = text
    message!.hidden = !text
  }
  entry.hidden = false
  panel.hidden = false
  panel.setAttribute("role", "tabpanel")
  const reader = mountDesk(panel, { request }, {
    embedded: true,
    initialView: view,
    close() {},
    onViewChange(next) { view = next; controls() },
    onBusyChange(value) { busy = value; controls() },
  })
  controls()

  async function show(next: DeskView) {
    if (disposed || busy) return
    error()
    await reader.setView(next)
  }
  for (const [name, tab] of Object.entries(tabs)) {
    tab.onclick = () => { void show(name as DeskView) }
    tab.onkeydown = event => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key) || busy) return
      event.preventDefault()
      const next = event.key === "Home" ? "documents" : event.key === "End" ? "compare" : view === "documents" ? "compare" : "documents"
      tabs[next].focus()
      void show(next)
    }
  }
  launch.onclick = async () => {
    if (disposed || opening || dialog || busy) return
    opening = true
    error()
    controls()
    try {
      const url = new URL("./desk.html", window.location.href)
      url.search = new URLSearchParams({ version: WORD_ADDIN_VERSION, view }).toString()
      dialog = await openDeskDialog({ url: url.href, handleRequest: request, onClose() {
        dialog = undefined
        if (disposed) return
        controls()
        launch.focus()
      } })
      if (disposed) dialog.close()
    } catch {
      if (!disposed) error("The wider window is unavailable. Continue in this pane.")
    } finally {
      opening = false
      if (!disposed) controls()
    }
  }
  window.addEventListener("pagehide", () => {
    disposed = true
    dialog?.close()
    reader.dispose()
  }, { once: true })
}
