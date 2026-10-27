// Shared comparison types contain no runtime dependencies so browser clients
// can consume the saved-version result without importing the comparison engine.

export type CompareOptions = { ignoreFormatting: boolean; detectMoves: boolean; headersFooters: boolean }

export type ChangeLocation = {
  scope: "body" | "header" | "footer" | "footnote" | "endnote"
  /** 1-based paragraph ordinal within the scope, counted in `side`. */
  paragraph: number
  side: "left" | "right"
  /** Nearest body heading at or before the paragraph, with its Word numbering. */
  heading?: string
  /** Word auto-number of the paragraph itself. */
  number?: string
}

export type ChangeHunk = {
  id: string
  kind: "modified" | "inserted" | "deleted" | "moved"
  location: ChangeLocation
  before?: string
  after?: string
  untracked: boolean
}

export type SideRevisions = { count: number; authors: string[] }

/** What became of one of left's pending tracked insertions or deletions in right. */
export type ProposalOutcome = {
  /** P1, P2… in left's document order. */
  id: string
  kind: "inserted" | "deleted"
  /** The proposed inserted text, or the text proposed for deletion. */
  text: string
  author: string
  /**
   * pending: right still carries it as ours. countered: carried, but right also
   * strikes or restores it with a tracked change of its own. accepted: its
   * effect is in right as plain text. rejected: removed or changed in right with
   * no tracked change, which is what Word's Reject Change leaves.
   */
  outcome: "pending" | "accepted" | "rejected" | "countered"
  /**
   * U-id of an untracked edit in the proposal's own paragraph: for a rejected
   * proposal the one that removed it, otherwise the first there.
   */
  untracked?: string
  /** Its paragraph in accept-all(left), counted as the hunks count (a deleted paragraph: in left); paragraph 0 if not located. */
  location: ChangeLocation
}

export type VersionComparison = {
  blackline: Uint8Array
  changes: ChangeHunk[]
  untracked: ChangeHunk[]
  attribution: "authors" | "label"
  left: SideRevisions
  right: SideRevisions
  proposals: ProposalOutcome[]
}
