import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { deskRequestHandler } from "./desk-client"
import { openDeskDialog } from "./desk-host"
import { mountDesk } from "./desk-ui"

export function setupSourceDesk(link: WordDocumentLink): void {
  const entry = document.getElementById("desk-entry")!
  const launch = document.getElementById("open-desk") as HTMLButtonElement
  const usePane = document.getElementById("inline-desk") as HTMLButtonElement
  const message = document.getElementById("desk-message")!
  const fallback = document.getElementById("pane-desk")!
  if (!entry || !launch || !usePane || !message || !fallback) return
  const request = deskRequestHandler(link)
  let dialog: { close(): void } | undefined
  let inline: ReturnType<typeof mountDesk> | undefined
  function showInline() {
    if (inline) return
    dialog?.close()
    dialog = undefined
    launch.disabled = true
    usePane.disabled = true
    message.textContent = "Source Desk is open below. Saving continues above."
    fallback.hidden = false
    inline = mountDesk(fallback, { request }, { close() {
      inline?.dispose()
      inline = undefined
      fallback.hidden = true
      launch.disabled = false
      usePane.disabled = false
      message.textContent = "Source Desk closed. Saving continues in this pane."
      usePane.focus()
    } })
    fallback.querySelector<HTMLButtonElement>("button")?.focus()
  }
  usePane.onclick = showInline
  entry.hidden = false
  launch.onclick = async () => {
    if (dialog || inline) return
    launch.disabled = true
    usePane.disabled = true
    try {
      const url = new URL("./desk.html", window.location.href)
      url.search = new URLSearchParams({ version: WORD_ADDIN_VERSION }).toString()
      dialog = await openDeskDialog({ url: url.href, handleRequest: request, onClose: () => {
        dialog = undefined
        launch.disabled = false
        message.textContent = "Source Desk closed. Saving continues in this pane."
        launch.focus()
      } })
      message.textContent = "Source Desk is open. Keep this pane open for its connection and saving."
    } catch {
      showInline()
      message.textContent = "The floating desk is unavailable in this Word window. Source Desk is open below."
    } finally {
      usePane.disabled = Boolean(inline)
    }
  }
  window.addEventListener("pagehide", () => { dialog?.close(); inline?.dispose() }, { once: true })
}
