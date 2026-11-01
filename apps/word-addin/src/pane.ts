import { WORD_UPDATE_MESSAGE, WordClient, isCurrentPaneVersion, readLink } from "./client"
import { connectOffice } from "./office"
import { browserCheckpoints, WordSync, type PaneState } from "./sync"
import { setupSourceDesk } from "./desk-entry"
import { configureConnectionRecovery, createPaneRenderer } from "./pane-render"
import { setupLatestVersion } from "./latest-version"
import { ChatRequestError, chatClient } from "./chat-client"
import { mountWordChat } from "./chat-ui"
import { captureCurrentWordContext, captureWordChatSelection, wordChatTableTarget, applyReviewedProposal, applyReviewedTableRow, applyWordChatRowDeletion, WordChatHostError } from "./word-chat-host"
import { watchWordChatSelection } from "./word-chat-selection"
import { openChatAttachments } from "./chat-attachments"
import { preferRoomyPane } from "./feature-tabs"
import type { WordChatPreparedAmendment } from "../../../dochaus/lib/word-chat-amendment-contract"
import type { WordChatAuthoredTableRow, WordChatAuthoredTableRowDeletion } from "../../../dochaus/lib/word-chat-contract"
import { createReplaceConfirmation } from "./replace-confirmation"
import { applyWordChatIfEligible, wordChatApplyEligibility } from "./word-chat-apply"

const element = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
const badge = element("status")
const message = element("message")
const save = element<HTMLButtonElement>("save")
const takeover = element<HTMLButtonElement>("takeover")
const release = element<HTMLButtonElement>("release")

const renderPane = createPaneRenderer(document)
let renderLatest: ((state: PaneState) => void) | undefined
function render(state: PaneState): void { renderPane(state); renderLatest?.(state) }

// A linked file starts the shared runtime when it opens (see linkedWordDocx).
// Show the pane from here as well, because Word on Windows can give the dock to
// its own Comments pane and skip the automatic doc.haus pane.
function showPane(office: unknown): void {
  const runtime = office as {
    context?: { requirements?: { isSetSupported(name: string, version: string): boolean } }
    addin?: { showAsTaskpane(): Promise<void> }
  } | undefined
  try {
    if (runtime?.context?.requirements?.isSetSupported("SharedRuntime", "1.1"))
      void runtime.addin?.showAsTaskpane().catch(() => undefined)
  } catch { /* Older hosts open the pane from the manifest or the ribbon. */ }
}

async function main(): Promise<void> {
  // An old cached manifest may still launch this page as a separate function
  // runtime. It must not become a writer merely because its JS was refreshed.
  if (!isCurrentPaneVersion(window.location.href)) {
    badge.textContent = "Restart Word"
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
    showPane(window.Office)
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
    const sync = new WordSync(link.baseVersion, office.host, client, checkpoints, render)
    save.addEventListener("click", () => { void sync.save() })
    takeover.addEventListener("click", () => { void sync.takeover() })
    release.addEventListener("click", () => { void sync.release() })
    let localEditIdentity: { base: string; generation: string } | undefined
    const chat = chatClient(link, undefined, 60_000, () => {
      if (!localEditIdentity) throw new Error("Review the amendment before applying it in Word.")
      return { "x-dochaus-word-session": session, "x-dochaus-word-base-version": localEditIdentity.base, "x-dochaus-word-lease": localEditIdentity.generation }
    })
    // Authoring precedes every Word mutation, so its refusal reason reaches the
    // card instead of the host's generic fallback.
    const prepare = async <T>(input: object): Promise<T> => {
      try { return await chat.request<T>("prepare", input) }
      catch (error) {
        if (!(error instanceof ChatRequestError)) throw error
        throw new WordChatHostError(/unchanged\.$/.test(error.message) ? error.message : `${error.message} Your document is unchanged.`, "unchanged", { cause: error })
      }
    }
    const chatPanel = element("pane-chat")
    const conversation = mountWordChat(chatPanel, chat, {
      documentName: link.documentName,
      capture: captureCurrentWordContext,
      captureSelection: captureWordChatSelection,
      tableTarget: wordChatTableTarget,
      watchSelection: watchWordChatSelection,
      openAttachments: options => openChatAttachments(chat, options),
      sourceUrl: source => source.provenance === "saved-version" ? `${link.origin}/matter/${encodeURIComponent(link.matter)}?view=documents` : undefined,
      applyEligibility: proposal => wordChatApplyEligibility(proposal.target, proposal.after, proposal.formatting, true, proposal.inlineFormatting, proposal.paragraphAction, proposal.list),
      async applyRow(proposal, request) {
        if (sync.snapshot().busy) throw new Error("Wait for the current Save to finish, then Apply.")
        return sync.localEdit({ conversationId: request.conversationId, turnId: request.turnId }, async identity => {
          localEditIdentity = identity
          try {
            const prepareRow = async <T>(bytes: Uint8Array): Promise<T> => {
              const chunks: string[] = []
              for (let start = 0; start < bytes.length; start += 32_768)
                chunks.push(String.fromCharCode(...bytes.subarray(start, start + 32_768)))
              return prepare<T>({ ...request, docxBase64: btoa(chunks.join("")) })
            }
            return proposal.kind === "table_row_delete"
              ? await applyWordChatRowDeletion(proposal, () => office.host.export(), bytes => prepareRow<WordChatAuthoredTableRowDeletion>(bytes), () => !sync.isDisposed())
              : await applyReviewedTableRow(proposal, () => office.host.export(), bytes => prepareRow<WordChatAuthoredTableRow>(bytes), () => !sync.isDisposed())
          } finally { localEditIdentity = undefined }
        })
      },
      async apply(proposal, request) {
        return applyWordChatIfEligible(proposal, async () => {
          if (sync.snapshot().busy) throw new Error("Wait for the current Save to finish, then Apply.")
          return sync.localEdit({ conversationId: request.conversationId, turnId: request.turnId }, async identity => {
            localEditIdentity = identity
            try { return await applyReviewedProposal(proposal, ({ ooxml }) => prepare<WordChatPreparedAmendment>({ ...request, ooxml }), () => !sync.isDisposed()) }
            finally { localEditIdentity = undefined }
          })
        })
      },
    })
    const replacement = createReplaceConfirmation(document)
    renderLatest = setupLatestVersion(document, base => client.latest(base), office.host.replace ? async () => {
      try {
        const loaded = await sync.loadLatest(base => client.latestContent(base), version => replacement.confirm(version))
        if (loaded) conversation.resetDocumentContext()
        return loaded
      } finally { replacement.complete() }
    } : undefined)
    setupSourceDesk(link, { panel: chatPanel, storage })
    window.addEventListener("pagehide", () => { sync.dispose(); replacement.dispose(); conversation.dispose() }, { once: true })
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
