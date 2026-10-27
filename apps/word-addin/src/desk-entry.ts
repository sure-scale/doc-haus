import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { deskRequestHandler } from "./desk-client"
import { openDeskDialog } from "./desk-host"
import { mountDesk } from "./desk-ui"
import type { DeskView } from "./desk-protocol"

export function setupSourceDesk(link: WordDocumentLink): void {
  const entry = document.getElementById("desk-entry")!
  const launch = document.getElementById("open-desk") as HTMLButtonElement
  const compare = document.getElementById("compare-desk") as HTMLButtonElement
  const usePane = document.getElementById("inline-desk") as HTMLButtonElement
  const compareInPane = document.getElementById("inline-compare-desk") as HTMLButtonElement
  const message = document.getElementById("desk-message")!
  const fallback = document.getElementById("pane-desk")!
  if (!entry || !launch || !compare || !usePane || !compareInPane || !message || !fallback) return
  const request = deskRequestHandler(link)
  let dialog: { close(): void } | undefined
  let inline: ReturnType<typeof mountDesk> | undefined
  let opening = false
  const label = (view: DeskView) => view === "compare" ? "Version comparison" : "Matter documents"
  function controls() {
    launch.disabled = compare.disabled = opening || Boolean(dialog || inline)
    usePane.disabled = compareInPane.disabled = opening || Boolean(inline)
  }
  function showInline(view: DeskView, trigger: HTMLButtonElement) {
    if (inline) return
    dialog?.close()
    dialog = undefined
    message.textContent = `${label(view)} is open below. Saving continues above.`
    fallback.hidden = false
    inline = mountDesk(fallback, { request }, { initialView: view, close() {
      inline?.dispose()
      inline = undefined
      fallback.hidden = true
      controls()
      message.textContent = `${label(view)} closed. Saving continues in this pane.`
      trigger.focus()
    } })
    controls()
    fallback.querySelector<HTMLButtonElement>("button")?.focus()
  }
  usePane.onclick = () => showInline("documents", usePane)
  compareInPane.onclick = () => showInline("compare", compareInPane)
  entry.hidden = false
  async function open(view: DeskView, trigger: HTMLButtonElement) {
    if (opening || dialog || inline) return
    opening = true
    controls()
    try {
      const url = new URL("./desk.html", window.location.href)
      url.search = new URLSearchParams({ version: WORD_ADDIN_VERSION, view }).toString()
      dialog = await openDeskDialog({ url: url.href, handleRequest: request, onClose: () => {
        dialog = undefined
        controls()
        message.textContent = `${label(view)} closed. Saving continues in this pane.`
        trigger.focus()
      } })
      message.textContent = `${label(view)} is open. Keep this pane open for its connection and saving.`
    } catch {
      showInline(view, trigger)
      message.textContent = `The floating window is unavailable in this Word window. ${label(view)} is open below.`
    } finally {
      opening = false
      controls()
    }
  }
  launch.onclick = () => { void open("documents", launch) }
  compare.onclick = () => { void open("compare", compare) }
  window.addEventListener("pagehide", () => { dialog?.close(); inline?.dispose() }, { once: true })
}
