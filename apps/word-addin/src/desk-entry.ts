import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { deskRequestHandler } from "./desk-client"
import { openDeskDialog } from "./desk-host"
import { mountDesk } from "./desk-ui"
import type { DeskView } from "./desk-protocol"
import { readFeatureTab, rememberFeatureTab, type FeatureTab } from "./feature-tabs"

/** Read-only views stay in the pane; a separate window is an optional wider view. */
export function setupSourceDesk(link: WordDocumentLink, chat?: { panel: HTMLElement; storage?: Storage }): void {
  const entry = document.getElementById("desk-entry")
  const documents = document.getElementById("documents-tab") as HTMLButtonElement | null
  const compare = document.getElementById("compare-tab") as HTMLButtonElement | null
  const chatTab = document.getElementById("chat-tab") as HTMLButtonElement | null
  const launch = document.getElementById("open-desk-window") as HTMLButtonElement | null
  const panel = document.getElementById("pane-desk")
  const message = document.getElementById("desk-message")
  if (!entry || !documents || !compare || !launch || !panel || !message) return

  const request = deskRequestHandler(link)
  const tabs: Partial<Record<FeatureTab, HTMLButtonElement>> = { documents, compare, ...(chat && chatTab ? { chat: chatTab } : {}) }
  const remembered = chat ? readFeatureTab(chat.storage) : { tab: "documents" as const, obsolete: false }
  let selected: FeatureTab = remembered.tab
  let view: DeskView = selected === "compare" ? "compare" : "documents"
  let busy = false
  let opening = false
  let disposed = false
  let dialog: { close(): void } | undefined
  function controls() {
    for (const [name, tab] of Object.entries(tabs)) {
      tab.disabled = busy && name !== "chat"
      tab.setAttribute("aria-selected", String(selected === name))
      tab.tabIndex = selected === name ? 0 : -1
    }
    panel!.setAttribute("aria-labelledby", tabs[view]!.id)
    panel!.hidden = selected === "chat"
    if (chat) chat.panel.hidden = selected !== "chat"
    launch!.hidden = selected === "chat"
    launch!.disabled = busy || opening || Boolean(dialog)
    launch!.setAttribute("aria-label", `Open ${view === "documents" ? "documents" : "comparison"} in a separate window`)
    launch!.title = launch!.getAttribute("aria-label")!
  }
  function error(text = "") {
    message!.textContent = text
    message!.hidden = !text
  }
  entry.hidden = false
  panel.hidden = selected === "chat"
  panel.setAttribute("role", "tabpanel")
  const reader = mountDesk(panel, { request }, {
    embedded: true,
    initialView: view,
    close() {},
    onViewChange(next) { view = next; if (selected !== "chat") selected = next; controls() },
    onBusyChange(value) { busy = value; controls() },
  })
  controls()
  if (remembered.obsolete) error("Your previous view is unavailable. Chat is ready.")

  async function show(next: FeatureTab) {
    if (disposed || (busy && next !== "chat")) return
    error()
    selected = next
    if (chat) rememberFeatureTab(next, chat.storage)
    controls()
    if (next !== "chat") await reader.setView(next)
  }
  for (const [name, tab] of Object.entries(tabs)) {
    tab.onclick = () => { void show(name as FeatureTab) }
    tab.onkeydown = event => {
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return
      event.preventDefault()
      const order: FeatureTab[] = chat && chatTab ? ["chat", "documents", "compare"] : ["documents", "compare"]
      const available = order.filter(value => !tabs[value]!.disabled)
      const index = available.indexOf(selected)
      const next = event.key === "Home" ? available[0]! : event.key === "End" ? available.at(-1)! : available[(index + (event.key === "ArrowRight" ? 1 : -1) + available.length) % available.length]!
      tabs[next]!.focus()
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
