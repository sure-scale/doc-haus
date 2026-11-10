/** The admitted matter-source input set, not a claim that an output cites or relies on
 * every passage. Missing legacy versions remain explicitly incomplete. */
export type SourceVersionSnapshot = {
  versions: { documentId: string; versionId: string; sha256: string }[]
  complete: boolean
}
export type SourceDependencyStatus = {
  status: "current" | "stale" | "unknown"
  complete: boolean
  sources: {
    documentId: string; name: string; versionId: string; versionNumber: number | null; sha256: string
    currentVersionId: string | null; currentVersionNumber: number | null
    status: "current" | "newer-version" | "newer-received-branch" | "removed" | "unavailable" | "unrecorded-change"
    linkedReturn?: { exchangeId: number; versionId: string; sourceName: string; versionNumber: number }
  }[]
}
