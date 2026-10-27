import { isCurrentPaneVersion, WORD_UPDATE_MESSAGE } from "./client"
import { connectDeskDialog } from "./desk-host"
import { mountDesk } from "./desk-ui"
import { isDeskView } from "./desk-protocol"

async function main() {
  const root = document.getElementById("desk-root")!
  if (!isCurrentPaneVersion(window.location.href)) { root.textContent = WORD_UPDATE_MESSAGE; return }
  try {
    const connection = await connectDeskDialog({ onDisconnect() {
      const status = root.querySelector('[data-el="status"]')
      if (status) status.textContent = "The Word pane connection is closed. Reopen matter documents or comparison from Word."
      root.querySelectorAll("button, select, input").forEach((control) => {
        if (control instanceof HTMLButtonElement || control instanceof HTMLSelectElement || control instanceof HTMLInputElement) control.disabled = true
      })
    } })
    const view = new URL(window.location.href).searchParams.get("view")
    const initialView = isDeskView(view) ? view : "documents"
    document.title = `${initialView === "compare" ? "Compare versions" : "Matter documents"} · doc.haus`
    const desk = mountDesk(root, connection, { floating: true, initialView, close() { connection.close() } })
    window.addEventListener("pagehide", () => { desk.dispose(); connection.dispose() }, { once: true })
  } catch {
    root.textContent = "The Word pane is unavailable. Close this window and reopen it from the doc.haus pane in Word."
  }
}
void main()
