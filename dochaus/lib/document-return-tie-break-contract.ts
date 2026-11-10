/** A model suggestion over a closed set, followed by a lawyer's explicit
 * choice. Exact quoted evidence verifies provenance, not semantic ancestry. */
export type DocumentReturnTieBreak = {
  choice: "candidate" | "new-document" | "uncertain"
  versionId?: string
  reason: string
  evidence: { versionId: string; incomingQuote: string; candidateQuote: string }[]
  model: { providerID: string; id: string; variant?: string }
  judgments: { versionId: string; relation: "same-document" | "different-document" | "uncertain"; reason: string; evidence: { versionId: string; incomingQuote: string; candidateQuote: string }[]; model: { providerID: string; id: string; variant?: string } }[]
  incomingHash: string
  candidates: { versionId: string; sha256: string }[]
  limitations: string[]
}
