import type { ChangeHunk, ProposalOutcome, SideRevisions } from "./version-compare-contract"

/** Read-only Word Source Desk data. Hashes identify exact content, not approval. */
export type DeskSource = {
  id: string
  name: string
  role: "source" | "deliverable"
  hash: string
  version: { id: string; number: number; createdAt?: number } | null
}

export type DeskSources = { sources: DeskSource[]; truncated: boolean; matterName: string; documentName: string }
export type DeskSourcePreview = DeskSource & {
  text: string
  /** Optional DOCX structure whose full text matches the canonical extraction. */
  blocks?: { kind: "paragraph" | "heading"; level?: number; text: string }[]
  characters: number
  truncated: boolean
  /** The preview is extracted from these exact retained bytes. */
  provenance: "saved-version"
}

export type DeskVersion = {
  id: string
  number: number
  hash: string
  parent: string | null
  createdAt: number
  bytes?: number
  origin: string
  label: string | null
  current: boolean
  removed: boolean
}
export type DeskVersions = { document: { id: string; name: string }; versions: DeskVersion[]; matterName?: string }
export type DeskComparison = {
  base: DeskVersion & { name: string; trackedChanges: SideRevisions }
  current: DeskVersion & { name: string; trackedChanges: SideRevisions }
  changes: ChangeHunk[]
  untracked: ChangeHunk[]
  proposals: ProposalOutcome[]
  attribution: "authors" | "label"
  /** Saved bytes only; current unsaved Word text is never substituted. */
  context: "saved-versions"
  truncated: boolean
}
