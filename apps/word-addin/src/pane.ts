import { WordClient, readLink } from "./client"
import { connectOffice } from "./office"
import { browserCheckpoints, WordSync, type PaneState } from "./sync"

const element = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
const badge = element("status")
const message = element("message")
const received = element("received")
const save = element<HTMLButtonElement>("save")
const done = element<HTMLButtonElement>("done")

function render(state: PaneState): void {
  badge.textContent = { "not-saved": "Not saved", saving: "Saving…", saved: "Saved", done: "Saved" }[state.status]
  badge.dataset.state = state.status
  message.textContent = state.message
  save.disabled = state.busy
  done.disabled = state.busy || state.status === "done"
  if (state.acknowledgement) {
    const time = new Date(state.acknowledgement.at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    received.textContent = `doc.haus received version ${state.acknowledgement.number} at ${time}.`
  }
}

async function main(): Promise<void> {
  try {
    const office = await connectOffice()
    const link = readLink(office.link, window.location.origin)
    if (!link) {
      badge.textContent = "Not linked"
      message.textContent = "Open a document using Open in Word in doc.haus to link it here."
      element("actions").hidden = true
      element("background").hidden = true
      return
    }
    element("matter").textContent = link.matterName
    element("document").textContent = link.documentName
    const back = element<HTMLAnchorElement>("back")
    back.href = link.origin
    back.hidden = false
    let storage: Storage | undefined
    try {
      storage = window.localStorage
    } catch {
      /* Private browsing may disable storage. */
    }
    const checkpoints = await browserCheckpoints(link, storage)
    const client = new WordClient(link, crypto.randomUUID())
    const sync = new WordSync(link.baseVersion, office.host, client, checkpoints, render)
    save.addEventListener("click", () => {
      void sync.save()
    })
    done.addEventListener("click", () => {
      void sync.done()
    })
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
