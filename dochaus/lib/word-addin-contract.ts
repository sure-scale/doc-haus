/** The document carries a capability for its own writes and scoped matter reads. */
export const WORD_ADDIN_ID = "4d857d21-6e86-47af-adba-cd7a4b056302"
export const WORD_LINK_SETTING = "dochaus.link"
/** Word's manifest and every handed-out document name this version. Changing it
 * breaks documents already handed out ("This add-in is no longer available" on
 * Mac) and the hosted bootstrap, so keep it fixed. Pane changes follow
 * WORD_PROTOCOL_VERSION instead. */
export const WORD_ADDIN_VERSION = "1.1.2.1"
/** The pane and the local service always ship together. Bump this when their
 * requests change; a pane left open across an update reloads itself. */
export const WORD_PROTOCOL_VERSION = "2"

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
  documentId?: string
  nativeOperation?: { sequence: number; state: "pending" }
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
export type WordConnectionState = { version: WordVersionState; lease: WordLeaseState; nativeOperation?: { sequence: number; state: "idle" | "pending" | "confirmed" | "not-dispatched" } }
export type WordLeaseAction = "inspect" | "claim" | "takeover" | "release"
