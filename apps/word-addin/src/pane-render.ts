import type { PaneState } from "./sync"
import { savePresentation } from "./pane-presentation"
import { WORD_LAUNCH_URL, loopbackWordOrigin } from "./bootstrap-contract"

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
    ? " Keep this pane open. Once doc.haus is running, choose Save or Check status."
    : "Check that your doc.haus server is running, then choose Save or Check status here. Keep this pane open."
}

/** Update only changed DOM values; importing this module never starts Office. */
export function createPaneRenderer(document: Document) {
  const element = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
  const badge = element("status"), message = element("message")
  const save = element<HTMLButtonElement>("save"), saveAgain = element<HTMLButtonElement>("save-again")
  const done = element<HTMLButtonElement>("done"), refresh = element<HTMLButtonElement>("refresh")
  const takeover = element<HTMLButtonElement>("takeover"), release = element<HTMLButtonElement>("release")
  return (state: PaneState): "save" | "close" => {
    const view = savePresentation(state)
    // Polling can refresh lease metadata without changing anything visible.
    // Avoid rewriting live-region text and button attributes in that case.
    update(badge, "textContent", view.badge)
    update(badge.dataset, "state", view.tone)
    update(message, "textContent", view.message)
    update(save, "textContent", view.primary === "close" ? "Save and close" : "Save")
    update(save, "disabled", state.busy || !state.canSave || state.status === "done")
    update(saveAgain, "hidden", view.primary !== "close")
    update(saveAgain, "disabled", state.busy || !state.canSave)
    update(done, "hidden", view.primary === "close")
    update(done, "disabled", state.busy || !state.canSave || state.status === "done")
    const connection = state.connection
    const stale = connection?.version.state === "stale"
    update(save, "hidden", stale)
    update(element("load-latest"), "hidden", !stale)
    update(element("version-state"), "textContent", view.mismatch ?? "")
    update(element("version-state"), "hidden", !view.mismatch)
    update(element("blocked-help"), "textContent", view.help ?? "")
    update(element("blocked-help"), "hidden", !view.help)
    update(refresh, "disabled", state.busy)
    update(takeover, "hidden", connection?.version.state !== "current" || connection.lease.state !== "other" || !connection.lease.generation)
    update(takeover, "disabled", state.busy)
    update(release, "hidden", connection?.lease.state !== "owned")
    update(release, "disabled", state.busy)
    update(element("connection-recovery"), "hidden", !state.connectionFailed || state.busy)
    return view.primary
  }
}

/** Outside dismissal does not move focus away from Word or another control. */
export function setupSaveMenu(menu: HTMLDetailsElement) {
  const close = (restoreFocus = false) => {
    menu.open = false
    if (restoreFocus) menu.querySelector("summary")?.focus()
  }
  const outside = (event: Event) => {
    if (menu.open && !menu.contains(event.target as Node | null)) close()
  }
  menu.ownerDocument.addEventListener("pointerdown", outside)
  menu.ownerDocument.addEventListener("click", outside)
  menu.ownerDocument.addEventListener("focusin", outside)
  menu.ownerDocument.defaultView?.addEventListener("blur", () => close())
  menu.addEventListener("keydown", event => {
    if (event.key !== "Escape") return
    event.preventDefault()
    close(true)
  })
  return () => close(true)
}
