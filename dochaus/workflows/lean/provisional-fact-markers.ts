import type { ChecklistItem, ChecklistSourceReference } from "./checklist-contract"
import { currentFactDisposition } from "./fact-answers"
import { readLeanPlan } from "./plan"
import { readLeanText, readQuestions } from "./state"

export const PROVISIONAL_FACT_MARKER_CONTRACT = "question-specific-provisional-markers-2-product-verification"

/** The producer cannot opt into a product task by naming a reserved prefix or
 * copying its wording. Reconstruct its current permission, provenance and scope. */
export function isProvisionalFactMarkerItem(directory: string, id: string, item: ChecklistItem) {
  return provisionalFactMarkerItems(directory).some(expected => expected.id === id && JSON.stringify(expected.item) === JSON.stringify(item))
}

/** Product handling requirements, not quotations falsely attributed to the
 * lawyer. The existing checklist checker decides this one question's marker
 * coverage separately for each body; the factual answer is never changed. */
export function provisionalFactMarkerItems(directory: string, questionID?: string): { id: string; item: ChecklistItem }[] {
  const plan = readLeanPlan(directory)
  if (!plan) return []
  return readQuestions(directory).flatMap(question => {
    if (questionID !== undefined && question.id !== questionID || question.stage !== "intake" || question.kind !== "fact" ||
      question.status !== "answered" || question.pid || !question.answer?.trim() ||
      question.answer_kind === "adopted" || question.answer_kind === "instruction" || currentFactDisposition(directory, question) !== "provisional") return []
    const answer: ChecklistSourceReference = { document: "request", locator: `.dochaus/lean/questions.json answer (${question.id})`,
      question: question.question, quote: question.answer }
    // A checked later permission can supersede an earlier wait. Preserve its
    // actual human words beside the original answer, without rewriting either.
    const later: ChecklistSourceReference[] = (question.fact_resolution?.basis ?? []).flatMap(ref =>
      ref.source_id === "updated-notes" && (readLeanText(directory, "updated-notes.md") ?? "").includes(ref.quote)
        ? [{ document: "request", locator: ".dochaus/lean/updated-notes.md", quote: ref.quote }] : [])
    const documents = question.document ? plan.documents.filter(document => document.id === question.document) : plan.documents
    return documents.map(document => ({
      id: `provisional-fact-marker:${question.id}:${document.id}`,
      item: {
        check: `For this one unresolved question, verify that every passage in this document which depends on its unknown fact visibly marks that fact as unresolved: ${question.question}`,
        basis: "open_question" as const, required_action: "flag" as const,
        source_refs: [answer, ...later],
        qualifications: [
          "Current provisional permission allows preparation only. The fact remains unknown; do not answer it, assume its truth, or infer sending, signature or release approval.",
          "This marker requirement is the product's handling rule, not a quotation of a human instruction. The quoted answer and any checked later permission retain their own provenance.",
          "Assess only this question and this document. Inspect all passages that depend on the fact, including conditions and references. A marker for another unknown does not satisfy this check.",
          "Return satisfied with action_status fulfilled, no open marker decisions, and exact visibly marked current body evidence only when every dependent passage is marked. The underlying factual uncertainty is not an open marker decision.",
          "Return not_applicable with action_status not_applicable and no open marker decisions only after inspecting the current document and establishing it contains no passage dependent on this fact. Missing inspection, body evidence or uncertain applicability is indeterminate, not not_applicable.",
        ],
        target: { document: document.id, role: `${document.title}: ${document.purpose}`, treatment: "open_issue" as const },
      },
    }))
  })
}
