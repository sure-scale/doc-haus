import type { ReturnIdentityClarification } from "../../../dochaus/lib/document-return-tie-break-contract"
import type { BaseContrastPart } from "./document-return-base-contrast"

export function checkIdentityClarification(text: string, context: { versionId: string; rivalVersionId: string; side: "left" | "right"; triple: BaseContrastPart; incomingQuotes: string[]; candidateQuotes: string[] }): Omit<ReturnIdentityClarification["parts"][number], "model"> {
  const value = JSON.parse(text) as Record<string, unknown>
  if (!value || Object.keys(value).sort().join(",") !== "evidence,limitations,reason,relation,versionId" || value.versionId !== context.versionId
    || !["same-document", "different-document", "neutral", "uncertain"].includes(String(value.relation)) || typeof value.reason !== "string" || !value.reason.trim() || value.reason.length > 1200
    || !Array.isArray(value.limitations) || value.limitations.length > 3 || value.limitations.some(item => typeof item !== "string" || !item.trim() || item.length > 300)
    || !Array.isArray(value.evidence) || value.evidence.length > 4 || value.relation !== "uncertain" && (!value.evidence.length || value.limitations.length))
    throw new Error("The identity clarification did not follow its bounded evidence contract.")
  const evidence = value.evidence.map(item => {
    const row = item as Record<string, unknown>
    if (!row || Object.keys(row).sort().join(",") !== "candidateQuote,incomingQuote,versionId" || row.versionId !== context.versionId
      || typeof row.incomingQuote !== "string" || typeof row.candidateQuote !== "string"
      || ![...context.incomingQuotes, context.triple.incoming.text].includes(row.incomingQuote)
      || ![...context.candidateQuotes, context.triple[context.side].text].includes(row.candidateQuote))
      throw new Error("An identity clarification quotation is not an exact supplied source quotation.")
    return { versionId: context.versionId, incomingQuote: row.incomingQuote, candidateQuote: row.candidateQuote }
  })
  if (value.relation !== "uncertain" && !evidence.some(pair => pair.incomingQuote === context.triple.incoming.text && pair.candidateQuote === context.triple[context.side].text))
    throw new Error("A resolved identity clarification must quote its complete new passage pair, not repeat identity-only evidence.")
  return { rivalVersionId: context.rivalVersionId, part: context.triple.part, relation: value.relation as ReturnIdentityClarification["parts"][number]["relation"], reason: value.reason.trim(), limitations: value.limitations as string[], evidence }
}

/** New evidence may clarify an uncertain identity, but never rewrite an
 * original different-document finding. Every required passage stays visible. */
export function clarifyIdentity(initial: { versionId: string; relation: string }, coverage: ReturnIdentityClarification["coverage"], unresolved: string[], parts: ReturnIdentityClarification["parts"]): ReturnIdentityClarification {
  const complete = coverage.requiredParts > 0 && coverage.processedParts === coverage.requiredParts && parts.length === coverage.requiredParts && !unresolved.length
  const relation = initial.relation === "uncertain" && complete && parts.some(part => part.relation === "same-document") && parts.every(part => (part.relation === "same-document" || part.relation === "neutral") && !part.limitations.length && part.evidence.length)
    ? "same-document" : initial.relation === "uncertain" && complete && parts.some(part => part.relation === "different-document") && parts.every(part => (part.relation === "different-document" || part.relation === "neutral") && !part.limitations.length && part.evidence.length)
      ? "different-document" : "uncertain"
  return { versionId: initial.versionId, relation, reason: relation === "same-document" ? "New complete passage evidence explicitly resolves the original identity concern as the same logical document; every other required passage is consistent or independently non-identifying. Base support remains a separate decision."
    : relation === "different-document" ? "New complete passage evidence identifies a different logical document; every other required passage is consistent or independently non-identifying."
      : "The original identity uncertainty is not fully resolved by complete, consistent new passage evidence.", coverage, unresolved, parts }
}
