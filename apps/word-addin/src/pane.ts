import { WORD_UPDATE_MESSAGE, WordClient, isCurrentPaneVersion, readLink } from "./client"
import { connectOffice } from "./office"
import { browserCheckpoints, WordSync, type PaneState } from "./sync"
import { setupSourceDesk } from "./desk-entry"
import { configureConnectionRecovery, createPaneRenderer } from "./pane-render"
import { setupLatestVersion } from "./latest-version"

const element = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
const badge = element("status")
const message = element("message")
const save = element<HTMLButtonElement>("save")
const takeover = element<HTMLButtonElement>("takeover")
const release = element<HTMLButtonElement>("release")

const renderPane = createPaneRenderer(document)
let renderLatest: ((state: PaneState) => void) | undefined
function render(state: PaneState): void { renderPane(state); renderLatest?.(state) }

async function main(): Promise<void> {
  // An old cached manifest may still launch this page as a separate function
  // runtime. It must not become a writer merely because its JS was refreshed.
  if (!isCurrentPaneVersion(window.location.href)) {
    badge.textContent = "Update needed"
    message.textContent = WORD_UPDATE_MESSAGE
    element("actions").hidden = true
    element("background").hidden = true
    return
  }
  try {
    const office = await connectOffice()
    const link = readLink(office.link, window.location.origin)
    if (!link) {
      badge.textContent = "Not linked"
      message.textContent = "Add this file to a matter in doc.haus, then choose Open in Word and open that copy. Keep doc.haus running while you use the add-in."
      element("actions").hidden = true
      element("background").hidden = true
      return
    }
    element("matter").textContent = link.matterName
    element("matter").title = `${link.matterName} · ${link.documentName}`
    const back = element<HTMLAnchorElement>("back")
    back.href = `${link.origin}/matter/${encodeURIComponent(link.matter)}`
    back.hidden = false
    configureConnectionRecovery(document, link.apiOrigin)
    let storage: Storage | undefined
    try {
      storage = window.localStorage
    } catch {
      /* Private browsing may disable storage. */
    }
    const checkpoints = await browserCheckpoints(link, storage)
    const client = new WordClient(link, crypto.randomUUID())
    renderLatest = setupLatestVersion(document, base => client.latest(base))
    const sync = new WordSync(link.baseVersion, office.host, client, checkpoints, render)
    save.addEventListener("click", () => { void sync.save() })
    takeover.addEventListener("click", () => { void sync.takeover() })
    release.addEventListener("click", () => { void sync.release() })
    // The read-only desk has its own lifecycle; opening/closing it never owns
    // or stops the existing saving queue.
    setupSourceDesk(link)
    await sync.start()
    setInterval(() => {
      void sync.tick()
    }, 5_000)
    // Native close cannot await an export/upload. No unload handler claims to do so.
  } catch {
    badge.textContent = "Not linked"
    message.textContent =
      "Open this pane in a current desktop version of Microsoft Word, with a document opened from doc.haus."
    element("actions").hidden = true
    element("background").hidden = true
  }
}

void main()
