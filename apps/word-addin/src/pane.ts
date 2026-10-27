import { WORD_UPDATE_MESSAGE, WordClient, isCurrentPaneVersion, readLink } from "./client"
import { connectOffice } from "./office"
import { browserCheckpoints, WordSync, type PaneState } from "./sync"

const element = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
const badge = element("status")
const message = element("message")
const received = element("received")
const save = element<HTMLButtonElement>("save")
const done = element<HTMLButtonElement>("done")
const refresh = element<HTMLButtonElement>("refresh")
const takeover = element<HTMLButtonElement>("takeover")
const release = element<HTMLButtonElement>("release")

function render(state: PaneState): void {
  badge.textContent = { "not-saved": "Not saved", saving: "Saving…", saved: "Saved", done: "Saved" }[state.status]
  badge.dataset.state = state.status
  message.textContent = state.message
  save.disabled = state.busy || !state.canSave
  done.disabled = state.busy || !state.canSave || state.status === "done"
  const connection = state.connection
  const current = connection?.version.current?.number
  const base = connection?.version.base?.number
  element("version-state").textContent = connection
    ? `This copy: ${base ? `version ${base}` : "unavailable version"}. In doc.haus: ${current ? `version ${current}` : "unavailable"}.`
    : "Checking version and editing status…"
  element("editing-state").textContent = !connection
    ? "Checking editing status…"
    : connection.lease.state === "owned"
      ? "You are editing this document."
      : connection.lease.state === "other"
        ? "Another editor has this document."
        : connection.version.state === "current"
          ? "Editing is available. Save to doc.haus to start."
          : "Saving is paused for this copy."
  refresh.disabled = state.busy
  refresh.textContent =
    connection?.version.state === "current" && connection.lease.state === "available"
      ? "Start saving here"
      : "Check status"
  takeover.hidden =
    connection?.version.state !== "current" || connection.lease.state !== "other" || !connection.lease.generation
  takeover.disabled = state.busy
  release.hidden = connection?.lease.state !== "owned"
  release.disabled = state.busy
  if (state.acknowledgement) {
    const time = new Date(state.acknowledgement.at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    received.textContent = `doc.haus received version ${state.acknowledgement.number} at ${time}.`
  }
}

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
      message.textContent = "Open a document using Open in Word in doc.haus to link it here."
      element("actions").hidden = true
      element("background").hidden = true
      return
    }
    element("matter").textContent = link.matterName
    element("document").textContent = link.documentName
    const back = element<HTMLAnchorElement>("back")
    back.href = `${link.origin}/matter/${encodeURIComponent(link.matter)}`
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
    refresh.addEventListener("click", () => {
      void sync.refresh()
    })
    takeover.addEventListener("click", () => {
      void sync.takeover()
    })
    release.addEventListener("click", () => {
      void sync.release()
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
