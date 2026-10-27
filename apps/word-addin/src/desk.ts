import { isCurrentPaneVersion, WORD_UPDATE_MESSAGE } from "./client"
import { connectDeskDialog } from "./desk-host"
import { mountDesk } from "./desk-ui"

async function main() {
  const root = document.getElementById("desk-root")!
  if (!isCurrentPaneVersion(window.location.href)) { root.textContent = WORD_UPDATE_MESSAGE; return }
  try {
    const connection = await connectDeskDialog({ onDisconnect() {
      const status = root.querySelector('[data-el="status"]')
      if (status) status.textContent = "The Word pane connection is closed. Reopen Source Desk from Word."
      root.querySelectorAll("button, select, input").forEach((control) => {
        if (control instanceof HTMLButtonElement || control instanceof HTMLSelectElement || control instanceof HTMLInputElement) control.disabled = true
      })
    } })
    const desk = mountDesk(root, connection, { floating: true, close() { connection.close() } })
    window.addEventListener("pagehide", () => { desk.dispose(); connection.dispose() }, { once: true })
  } catch {
    root.textContent = "The Word pane is unavailable. Close this desk and reopen it from the doc.haus pane in Word."
  }
}
void main()
