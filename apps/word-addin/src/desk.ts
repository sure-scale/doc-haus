import { isCurrentPaneVersion, WORD_UPDATE_MESSAGE } from "./client"
import { connectDeskDialog } from "./desk-host"
import { mountDesk } from "./desk-ui"
import { isDeskView } from "./desk-protocol"

async function main() {
  const root = document.getElementById("desk-root")!
  if (!isCurrentPaneVersion(window.location.href)) { root.textContent = WORD_UPDATE_MESSAGE; return }
  let desk: ReturnType<typeof mountDesk> | undefined
  let disconnected = false
  try {
    const connection = await connectDeskDialog({ onDisconnect() {
      disconnected = true
      desk?.dispose()
      const status = document.createElement("p")
      status.setAttribute("role", "alert")
      status.textContent = "The Word pane connection is closed. Close this window and reopen it from Word."
      root.replaceChildren(status)
    } })
    if (disconnected) { connection.dispose(); return }
    const view = new URL(window.location.href).searchParams.get("view")
    const initialView = isDeskView(view) ? view : "documents"
    document.title = `${initialView === "compare" ? "Compare versions" : "Matter documents"} · doc.haus`
    desk = mountDesk(root, connection, { floating: true, initialView,
      onContextChange(context) { document.title = `${context.matterName ?? context.documentName ?? "Matter documents"} · doc.haus` },
      close() { connection.close() } })
    window.addEventListener("pagehide", () => { desk?.dispose(); connection.dispose() }, { once: true })
  } catch {
    root.textContent = "The Word pane is unavailable. Close this window and reopen it from the doc.haus pane in Word."
  }
}
void main()
