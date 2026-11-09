import type { WordChatContext } from "./word-chat-contract"

/** Advisory review of a reviewed proposal. Never an edit or authority to apply. */
export type WordConsequenceFinding = {
  id: string
  title: string
  explanation: string
  kind: "effect" | "uncertain"
  basis: "inference"
  passage: { paragraphId: string; quote: string }
  related?: { paragraphId: string; quote: string }[]
}
export type WordConsequenceAssessment = {
  id: string
  conversationId: string
  requestId: string
  turnId: string
  status: "running" | "complete" | "partial" | "cancelled" | "stale" | "disabled"
  /** Model inference about before/after meaning, never assurance. */
  meaning?: string
  history?: WordConsequenceAssessment[]
  findings: WordConsequenceFinding[]
  coverage: { reviewed: number; total: number; candidates: number; checked: number }
  gaps: string[]
  createdAt: number
}
export type WordConsequenceAssess = { conversationId: string; requestId: string; turnId: string; context: WordChatContext }
export type WordConsequenceRead = { conversationId: string; assessmentId?: string; context?: WordChatContext }
