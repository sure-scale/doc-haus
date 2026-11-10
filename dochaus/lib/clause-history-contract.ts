import type { ChangeLocation } from "./version-compare-contract"

export type ClauseHistoryScope = ChangeLocation["scope"]
export type ClauseHistoryRequest = {
  anchorVersion: string
  /** Full captured paragraph text. An anchor identifies only these exact saved bytes. */
  paragraph: { scope: ClauseHistoryScope; text: string; anchorId?: string }
  expectedCurrentHash?: string
}
export type ClauseHistoryParagraph = {
  anchorId: string
  scope: ClauseHistoryScope
  paragraph: number
  text: string
}
export type ClauseHistorySource = { versionId: string; hash: string; kind: "comment" | "decision" | "revision" | "proposal"; id: string; exchangeId?: number; recordHash?: string }
export type ClauseHistoryReason = {
  kind: "comment" | "decision" | "proposal"
  /** A comment is a recorded note, not evidence of an agreed concession. */
  text: string
  author: string | null
  date: string | null
  source: ClauseHistorySource
  parent?: string | null
  resolved?: boolean | null
  decision?: "accept" | "reject" | "hold"
  outcome?: string
  /** A proposed rationale/wording is distinct from a human's local decision. */
  proposal?: { id: number; scope: string; oldText: string; newText: string }
}
export type ClauseHistoryChange = { kind: string; text: string; author: string | null; date: string | null; source: ClauseHistorySource }
export type ClauseHistoryVersion = {
  id: string; number: number; hash: string; parent: string | null; createdAt: number
  actor: string | null; origin: string; label: string | null; party?: string | null
  /** Physical version identity remains with its original document. */
  linkedReturn?: { exchangeId: number; sourceDocumentId: string; sourceDocumentName: string; targetDocumentId: string; baseVersionId: string }
}
export type ClauseHistoryRound = {
  version: ClauseHistoryVersion
  status: "verified" | "unresolved" | "unavailable"
  reason?: string
  paragraphs: ClauseHistoryParagraph[]
  evidence?: "anchor-version" | "word-paragraph-id" | "exact-wording" | "exact-split-merge"
  reasons: ClauseHistoryReason[]
  changes: ClauseHistoryChange[]
  reasonStatus: "recorded" | "not-recorded" | "unavailable"
}
export type ClauseHistoryResult = {
  document: { id: string; name: string; currentVersion: string | null }
  anchorVersion: string
  currentVersion: string | null
  anchor: { status: "verified" | "unresolved"; reason?: string; paragraphs: ClauseHistoryParagraph[] }
  rounds: ClauseHistoryRound[]
  /** Wording is the current-text projection; pending tracked edits are not approved by this view. */
  wordingView: "current-text"
  limitations: string[]
}
export type ClauseHistorySourceRequest = ClauseHistorySource
export type ClauseHistorySourceResult = {
  version: ClauseHistoryVersion
  source: ClauseHistorySource
  reason?: ClauseHistoryReason
  change?: ClauseHistoryChange
}
