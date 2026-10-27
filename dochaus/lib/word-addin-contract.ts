/** The document carries a capability for its own writes and scoped matter reads. */
export const WORD_ADDIN_ID = "4d857d21-6e86-47af-adba-cd7a4b056302"
export const WORD_LINK_SETTING = "dochaus.link"
export const WORD_ADDIN_VERSION = "1.1.2.1"

export type WordDocumentLink = {
  version: 1
  origin: string
  apiOrigin: string
  matter: string
  matterName: string
  document: string
  documentName: string
  baseVersion: string
  token: string
}

export type WordSaveResponse = {
  connection: WordConnectionState
  version: { id: string; number: number; parent: string | null; created: boolean }
  overlap: boolean
}

export type WordOpenDocument = {
  currentNumber?: number
  lease?: WordLeaseState
  name: string
  open: boolean
  currentVersion: string | null
  overlap: boolean
}


export type WordVersionRef = { id: string; number: number }
export type WordVersionState = {
  base: WordVersionRef | null
  current: WordVersionRef | null
  state: "current" | "stale" | "removed" | "unrecorded-change"
}
export type WordLeaseState = {
  state: "available" | "owned" | "other"
  generation?: string
  expiresAt: number | null
}
export type WordConnectionState = { version: WordVersionState; lease: WordLeaseState }
export type WordLeaseAction = "inspect" | "claim" | "takeover" | "release"
