import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import type { WordChatMaterial, WordChatMaterials } from "../../../dochaus/lib/word-chat-contract"
import { ChatRequestError, type WordChatApi } from "./chat-client"
import { DeskRequestError, openDeskDialog } from "./desk-host"
import { isDeskRequest, type DeskRequest } from "./desk-protocol"

export type ChatAttachmentOptions = {
  conversationId: string
  attached: WordChatMaterial[]
  onAttach(material: WordChatMaterial): void
}

/** Allows the chat to offer its inline picker on hosts without secure dialogs. */
export class ChatAttachmentsUnavailable extends Error {
  constructor() {
    super("The attachment window is unavailable. Use Attach in the pane.")
    this.name = "ChatAttachmentsUnavailable"
  }
}

export type ChatAttachmentCatalog = WordChatMaterials & { attached: WordChatMaterial[] }
const identity = (material: { kind: string; id: string }) => `${material.kind}:${material.id}`
function publicMaterial(material: WordChatMaterial): WordChatMaterial {
  const { kind, id, hash, name, description, provenance } = material
  return { kind, id, hash, name, ...(description ? { description } : {}), provenance }
}

/** Captures the parent's conversation; no child request can choose another one. */
export function chatAttachmentRequestHandler(api: WordChatApi, options: ChatAttachmentOptions) {
  const conversationId = options.conversationId
  const attached = new Map(options.attached.map(material => [identity(material), publicMaterial(material)]))
  let disposed = false
  let changing = false
  const checkOpen = () => {
    if (disposed) throw new DeskRequestError("The attachment window is closed. Reopen Attach from the chat.")
  }
  const add = (material: WordChatMaterial) => {
    checkOpen()
    const safe = publicMaterial(material)
    if (attached.get(identity(safe))?.hash === safe.hash) return safe
    options.onAttach(safe)
    attached.set(identity(safe), safe)
    return safe
  }
  return {
    dispose() { disposed = true },
    async request(input: DeskRequest): Promise<unknown> {
      checkOpen()
      if (!isDeskRequest(input) || !["chat-materials", "chat-attach", "chat-upload"].includes(input.action))
        throw new DeskRequestError("Choose a file from Attach in the chat.")
      try {
        if (input.action === "chat-materials") {
          const catalog = await api.request<WordChatMaterials>("materials", { conversationId })
          checkOpen()
          return { materials: catalog.materials.map(publicMaterial), matterName: catalog.matterName,
            documentName: catalog.documentName, attached: [...attached.values()] } satisfies ChatAttachmentCatalog
        }
        // Serialize changes so rapid clicks cannot exceed the chat's eight-file
        // limit or attach the same material twice while a request is pending.
        if (changing) throw new DeskRequestError("Wait for the current attachment to finish.")
        changing = true
        try {
          if (input.action === "chat-attach") {
            const catalog = await api.request<WordChatMaterials>("materials", { conversationId })
            checkOpen()
            const material = catalog.materials.find(material => identity(material) === identity(input.material) && material.hash === input.material.hash)
            if (!material) throw new DeskRequestError("This material changed. Refresh the list and choose it again.")
            if (attached.size >= 8 && !attached.has(identity(material))) throw new DeskRequestError("Attach up to eight files. Remove one in the chat first.")
            return { material: add(material) }
          }
          if (input.action === "chat-upload") {
            if (attached.size >= 8) throw new DeskRequestError("Attach up to eight files. Remove one in the chat first.")
            const binary = atob(input.data)
            const bytes = new Uint8Array(binary.length)
            for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index)
            const result = await api.upload<{ material: WordChatMaterial }>(conversationId, new File([bytes], input.name), input.mode)
            return { material: add(result.material) }
          }
          throw new DeskRequestError("Choose a file from Attach in the chat.")
        } finally { changing = false }
      } catch (error) {
        if (error instanceof ChatRequestError) throw new DeskRequestError(error.message)
        throw error
      }
    },
  }
}

/** Resolves only when the native dialog closes, fencing conversation switches. */
export async function openChatAttachments(api: WordChatApi, options: ChatAttachmentOptions): Promise<void> {
  const handler = chatAttachmentRequestHandler(api, options)
  let finish!: () => void
  const closed = new Promise<void>(resolve => { finish = resolve })
  try {
    const url = new URL("./desk.html", window.location.href)
    url.search = new URLSearchParams({ version: WORD_ADDIN_VERSION, view: "attachments" }).toString()
    await openDeskDialog({ url: url.href, dimensions: { width: 50, height: 65 }, handleRequest: handler.request,
      onClose() { handler.dispose(); finish() } })
  } catch {
    handler.dispose()
    throw new ChatAttachmentsUnavailable()
  }
  await closed
}
