/** The document carries a write-only capability for its own doc.haus file. */
export const WORD_ADDIN_ID = "4d857d21-6e86-47af-adba-cd7a4b056302"
export const WORD_LINK_SETTING = "dochaus.link"
export const WORD_ADDIN_VERSION = "1.0.1.0"

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
  version: { id: string; number: number; parent: string | null; created: boolean }
  overlap: boolean
}

export type WordOpenDocument = {
  name: string
  open: boolean
  currentVersion: string | null
  overlap: boolean
}
