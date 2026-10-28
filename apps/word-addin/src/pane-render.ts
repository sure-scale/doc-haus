import type { PaneState } from "./sync"
import { savePresentation } from "./pane-presentation"
import { WORD_LAUNCH_URL, loopbackWordOrigin } from "./bootstrap-contract"
import { setupPaneLayout } from "./pane-layout"

function update<T, K extends keyof T>(target: T, key: K, value: T[K]): void {
  if (target[key] !== value) target[key] = value
}

/** Recovery launches the app without unloading the pane's pending save queue. */
export function configureConnectionRecovery(document: Document, origin: string) {
  const open = document.getElementById("open-dochaus") as HTMLAnchorElement
  const local = loopbackWordOrigin(origin)
  open.hidden = !local
  if (local) open.href = WORD_LAUNCH_URL
  document.getElementById("connection-help")!.textContent = local
    ? " Keep this pane open. It reconnects automatically when doc.haus is running."
    : "Check that your doc.haus server is running. Keep this pane open; it reconnects automatically."
}

/** Update only changed DOM values; importing this module never starts Office. */
export function createPaneRenderer(document: Document) {
  setupPaneLayout(document)
  const element = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
  const badge = element("status"), message = element("message")
  const save = element<HTMLButtonElement>("save")
  const version = element("save-version")
  const details = element("save-details-toggle")
  const takeover = element<HTMLButtonElement>("takeover"), release = element<HTMLButtonElement>("release")
  return (state: PaneState): void => {
    const view = savePresentation(state)
    // Polling can refresh lease metadata without changing anything visible.
    // Avoid rewriting live-region text and button attributes in that case.
    update(badge, "textContent", view.badge)
    update(badge.dataset, "state", view.tone)
    update(message, "textContent", view.message)
    update(save, "disabled", state.busy || !state.canSave || state.status === "done")
    const connection = state.connection
    const number = connection?.version.base?.number ?? state.acknowledgement?.number
    update(version, "textContent", number === undefined ? "" : `v${number}`)
    update(version, "title", number === undefined ? "" : `This Word copy: version ${number}`)
    update(version, "hidden", number === undefined)
    const detailLabel = `${view.badge}${number === undefined ? "" : `, version ${number}`}. Show save status and recovery details`
    if (details.getAttribute("aria-label") !== detailLabel) details.setAttribute("aria-label", detailLabel)
    const stale = connection?.version.state === "stale"
    update(save, "hidden", stale)
    update(element("load-latest"), "hidden", !stale)
    update(element("version-state"), "textContent", view.mismatch ?? "")
    update(element("version-state"), "hidden", !view.mismatch)
    update(element("blocked-help"), "textContent", view.help ?? "")
    update(element("blocked-help"), "hidden", !view.help)
    update(takeover, "hidden", connection?.version.state !== "current" || connection.lease.state !== "other" || !connection.lease.generation)
    update(takeover, "disabled", state.busy)
    update(release, "hidden", connection?.lease.state !== "owned" || (!state.problem && connection.version.state === "current"))
    update(release, "disabled", state.busy)
    update(element("connection-recovery"), "hidden", !state.connectionFailed || state.busy)
  }
}
