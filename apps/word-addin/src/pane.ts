import { WORD_UPDATE_MESSAGE, WordClient, isCurrentPaneVersion, readLink } from "./client"
import { connectOffice } from "./office"
import { browserCheckpoints, WordSync, type PaneState } from "./sync"
import { setupSourceDesk } from "./desk-entry"
import { configureConnectionRecovery, createPaneRenderer } from "./pane-render"
import { setupLatestVersion } from "./latest-version"
import { chatClient } from "./chat-client"
import { mountWordChat } from "./chat-ui"
import { captureCurrentWordContext, captureWordChatSelection, applyReviewedProposal } from "./word-chat-host"
import { preferRoomyPane } from "./feature-tabs"
import type { WordChatAuthoredAmendment } from "../../../dochaus/lib/word-chat-amendment-contract"

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
    preferRoomyPane(window.Office)
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
    const session = crypto.randomUUID()
    const client = new WordClient(link, session)
    renderLatest = setupLatestVersion(document, base => client.latest(base))
    const sync = new WordSync(link.baseVersion, office.host, client, checkpoints, render)
    save.addEventListener("click", () => { void sync.save() })
    takeover.addEventListener("click", () => { void sync.takeover() })
    release.addEventListener("click", () => { void sync.release() })
    let localEditIdentity: { base: string; generation: string } | undefined
    const chat = chatClient(link, undefined, 60_000, () => {
      if (!localEditIdentity) throw new Error("Review the amendment before applying it in Word.")
      return { "x-dochaus-word-session": session, "x-dochaus-word-base-version": localEditIdentity.base, "x-dochaus-word-lease": localEditIdentity.generation }
    })
    const chatPanel = element("pane-chat")
    const conversation = mountWordChat(chatPanel, chat, {
      documentName: link.documentName,
      capture: captureCurrentWordContext,
      captureSelection: captureWordChatSelection,
      sourceUrl: source => source.provenance === "saved-version" ? `${link.origin}/matter/${encodeURIComponent(link.matter)}?view=documents` : undefined,
      async apply(proposal, request) {
        if (sync.snapshot().busy) throw new Error("Wait for the current Save to finish, then Apply.")
        return sync.localEdit(async identity => {
          localEditIdentity = identity
          try { return await applyReviewedProposal(proposal, ({ ooxml }) => chat.request<WordChatAuthoredAmendment>("prepare", { ...request, ooxml })) }
          finally { localEditIdentity = undefined }
        })
      },
    })
    setupSourceDesk(link, { panel: chatPanel, storage })
    window.addEventListener("pagehide", () => conversation.dispose(), { once: true })
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
