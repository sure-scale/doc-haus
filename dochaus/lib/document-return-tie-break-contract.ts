/** A model suggestion over a closed set, followed by a lawyer's explicit
 * choice. Exact quoted evidence verifies provenance, not semantic ancestry. */
export type ReturnBasePassage = { scope: string; story: string; paragraph: number; text: string }
export type ReturnBaseContrast = {
  leftVersionId: string
  rightVersionId: string
  coverage: { changedPassages: number; processedPassages: number }
  unresolved: string[]
  parts: {
    part: number
    support: "left" | "right" | "neither" | "uncertain"
    reason: string
    limitations: string[]
    evidence?: { incoming: ReturnBasePassage; left: ReturnBasePassage; right: ReturnBasePassage }
    model: { providerID: string; id: string; variant?: string }
  }[]
}
export type DocumentReturnTieBreak = {
  choice: "candidate" | "new-document" | "uncertain"
  versionId?: string
  reason: string
  evidence: { versionId: string; incomingQuote: string; candidateQuote: string }[]
  model: { providerID: string; id: string; variant?: string }
  judgments: { versionId: string; relation: "same-document" | "different-document" | "uncertain"; reason: string; evidence: { versionId: string; incomingQuote: string; candidateQuote: string }[]; model: { providerID: string; id: string; variant?: string } }[]
  incomingHash: string
  candidates: { versionId: string; sha256: string }[]
  baseContrasts: ReturnBaseContrast[]
  baseSources: { documentId: string; documentName: string; versionId: string; versionNumber: number; sha256: string; current: boolean; sent: boolean; associationCorrected: boolean; correctedAssociations: { documentId: string; baseVersionId: string }[] }[]
  limitations: string[]
}
