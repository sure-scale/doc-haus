import type { PaneState } from "./sync"

/** Prepare the short-lived handoff before the click so Word can follow a real
 * user-activated link. The existing document and its saving queue stay open. */
export function setupLatestVersion(document: Document, latest: (base: string) => Promise<string>, now = Date.now) {
  const action = document.getElementById("load-latest") as HTMLAnchorElement
  const message = document.getElementById("latest-message")!
  let key = "", base = "", prepared = 0, pending = false, nextAttempt = 0
  const clear = () => {
    action.removeAttribute("href")
    action.setAttribute("aria-disabled", "true")
    prepared = 0
  }
  const prepare = async () => {
    if (!key || pending || now() < nextAttempt || (prepared && now() - prepared < 30_000)) return
    pending = true
    const requested = key
    const started = now()
    try {
      const uri = await latest(base)
      if (key !== requested) return
      action.href = uri
      action.setAttribute("aria-disabled", "false")
      prepared = started
      message.hidden = true
    } catch {
      if (key !== requested) return
      clear()
      nextAttempt = now() + 5_000
      message.textContent = "Cannot load the latest copy yet. Check that doc.haus is running, or open the matter below."
      message.hidden = false
    } finally { pending = false }
  }
  action.addEventListener("click", event => {
    if (!prepared || now() - prepared >= 45_000) {
      event.preventDefault()
      clear()
      void prepare()
      return
    }
    // Tickets are single-use. A real anchor with target=_blank keeps the pane
    // alive and avoids opening a popup after an asynchronous fetch.
    message.textContent = "Opening the latest version. This older copy stays open."
    message.hidden = false
    setTimeout(() => { clear(); void prepare() }, 0)
  })
  return (state: PaneState) => {
    const version = state.connection?.version
    const next = version?.state === "stale" && version.base && version.current ? `${version.base.id}:${version.current.id}` : ""
    if (next !== key) {
      key = next
      base = version?.base?.id ?? ""
      nextAttempt = 0
      clear()
      message.hidden = true
    }
    void prepare()
  }
}
