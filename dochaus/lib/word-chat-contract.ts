/** Word's scoped chat bridge. A linked-document token never proxies engine APIs. */
export const WORD_CHAT_ROUTES = ["conversations", "conversation", "send", "poll", "cancel", "materials", "upload", "suggest", "prepare", "events"] as const
export const wordChatPath = (pathname: string) => WORD_CHAT_ROUTES.some(route => pathname === `/word/chat/${route}`)
export type WordChatScope = "document" | "matter"
export type WordChatKind = "chat" | "amendment" | "formatting" | "suggestions"
export type WordChatMode = "auto" | "qa" | "redline"
export type WordChatAttachment = { kind: "source" | "playbook" | "upload"; id: string; hash: string }
export type WordChatMaterial = WordChatAttachment & { name: string; description?: string; provenance: "saved-version" | "instructions" | "request-upload" }
import type { WordChatSnapshot, WordChatProposal, WordChatTableRowInsertionProposal } from "./word-chat-amendment-contract"
export type { WordChatProposal } from "./word-chat-amendment-contract"
export type WordChatContext = WordChatSnapshot
export type WordChatSource = { id: string; contentHash?: string; name: string; quote: string; location: string; provenance: "word-buffer" | "saved-version" | "instructions" | "request-upload" }
export type WordChatConversation = { id: string; title: string; documentName: string; updatedAt: number }
export type WordChatSuggestion = { title: string; explanation: string; passage: string }
export type WordChatTurn = {
  id: string
  conversationId: string
  requestId: string
  kind: WordChatKind
  scope: WordChatScope
  text: string
  answer: string
  status: "running" | "complete" | "cancelled" | "error"
  sources: WordChatSource[]
  attachments: WordChatMaterial[]
  selection?: WordChatContext["selection"]
  proposal?: WordChatProposal
  rowProposal?: WordChatTableRowInsertionProposal
  proposalApplied?: boolean
  suggestions?: WordChatSuggestion[]
  suggestionCoverage?: { reviewed: number; total: number }
  error?: string
  createdAt: number
}
export type WordChatSend = {
  conversationId: string
  requestId: string
  text: string
  scope: WordChatScope
  attachments: WordChatAttachment[]
  context: WordChatContext
  kind?: WordChatKind
  /** A choice for this request, not permission to mutate the document. */
  mode?: WordChatMode
  /** Exact earlier reviewed target; refining never follows a new cursor position. */
  proposal?: WordChatProposal
  rowProposal?: WordChatTableRowInsertionProposal
  suggestionTarget?: "document" | "selection"
}
export type WordChatConversationResult = { conversation: WordChatConversation; turns: WordChatTurn[] }
export type WordChatPreferences = { defaultScope: WordChatScope; suggestions: boolean }
export type WordChatMaterials = { materials: WordChatMaterial[]; matterName: string; documentName: string; preferences?: WordChatPreferences }
/** Authoring output is a verification oracle, never a Word import payload. */
export type WordChatAuthoredTableRow = {
  kind: "table_row_insert"; proposalId: string; sourceSha256: string
  table: WordChatTableRowInsertionProposal["table"]; cells: string[]; author: "doc.haus"
  writerRequest: { source_sha256: string; target_handle: string; position: "after"; cells: string[]; author: string }
  authoredBase64: string; revisions: { id: string; family: string }[]
}
