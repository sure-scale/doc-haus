import type { PaneState } from "./sync"

/** The primary action refreshes this Word copy. A separate file stays available
 * for hosts/packages that cannot safely replace the complete document. */
export function setupLatestVersion(document: Document, latest: (base: string) => Promise<string>, refresh?: () => Promise<boolean>, now = Date.now) {
  const action = document.getElementById("load-latest") as HTMLButtonElement
  const separate = document.getElementById("open-latest-copy") as HTMLAnchorElement
  const message = document.getElementById("latest-message")!
  let key = "", base = "", prepared = 0, pending = false, loading = false, busy = false, nextAttempt = 0
  const clear = () => { separate.removeAttribute("href"); separate.setAttribute("aria-disabled", "true"); prepared = 0 }
  const prepare = async () => {
    if (!key || pending || now() < nextAttempt || (prepared && now() - prepared < 30_000)) return
    pending = true
    const requested = key, started = now()
    try {
      const uri = await latest(base)
      if (key !== requested) return
      separate.href = uri; separate.setAttribute("aria-disabled", "false"); prepared = started
    } catch {
      if (key !== requested) return
      clear(); nextAttempt = now() + 5_000
    } finally { pending = false }
  }
  action.addEventListener("click", async () => {
    if (!refresh || loading || busy || !key) return
    loading = true; action.disabled = true; action.textContent = "Loading…"
    message.hidden = true
    try { await refresh() }
    finally { loading = false; action.disabled = busy; action.textContent = "Load latest version" }
  })
  separate.addEventListener("click", event => {
    if (!prepared || now() - prepared >= 45_000 || busy || loading) { event.preventDefault(); clear(); void prepare(); return }
    message.textContent = "Opening a separate copy. This document stays open."; message.hidden = false
    setTimeout(() => { clear(); void prepare() }, 0)
  })
  return (state: PaneState) => {
    const version = state.connection?.version
    const next = version?.state === "stale" && version.base && version.current ? `${version.base.id}:${version.current.id}` : ""
    busy = state.busy
    action.hidden = !next || !refresh
    separate.hidden = !next
    if (next && !refresh) {
      const help = document.getElementById("blocked-help")
      if (help) help.textContent = "This version of Word needs a separate copy. Use Save As first to keep local edits."
    }
    action.disabled = busy || loading
    if (next !== key) { key = next; base = version?.base?.id ?? ""; nextAttempt = 0; clear(); message.hidden = true }
    void prepare()
  }
}
