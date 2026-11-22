import type { ChecklistItem, ChecklistSourceReference } from "./checklist-contract"
import { currentFactDisposition } from "./fact-answers"
import { readLeanPlan } from "./plan"
import { readLeanText, readQuestions } from "./state"
import { escapeTagOpeners } from "../../lib/untrusted"

export const PROVISIONAL_FACT_MARKER_CONTRACT = "question-specific-provisional-markers-4-positive-evidence"

/** The positive-evidence standard the marker checker (markerVerificationIssues)
 * and final gate (checklistFactMarked) already enforce, stated for repair
 * adjudication so a repair rejection cannot drop a duty the gate still blocks on.
 * Code owns the duty's identity (isProvisionalFactMarkerItem); the model judges
 * dependency from exact current passages and the existing checkers recheck. This
 * is a framing handoff, not a second adjudication contract or a free-text scan. */
export const PROVISIONAL_FACT_MARKER_REPAIR_STANDARD = "This is the product's verified handling duty for one unresolved provisional fact, reconstructed from the recorded permission and scope — not an adverse substantive-law or authority claim, and not an unverified assessment to discount. The quoted human answer is provenance for provisional preparation, not a per-passage marking command; the absence of a separate human marking instruction is not grounds to reject the duty. Apply the same positive-evidence standard as the marker checker and final gate: every passage that semantically depends on the unknown fact must visibly mark it as unresolved, and a marker or shared qualification elsewhere does not cure an unmarked dependent passage. Reject the duty only when current evidence actually disproves dependency — inspection of the current document establishes that no passage relies on the unknown. Uncertain dependency or missing inspection remains deferred uncertainty, never a rejection that drops the duty. Never answer the fact, assume its truth, or infer sending, signature or release approval."

/** The same duties handed to the early source-fidelity review, so an unmarked
 * dependent passage is reported as a finding while amendments still run. The
 * later independent checklist remains each duty's sole verifier. */
export const PROVISIONAL_MARKER_REVIEW_INSTRUCTION = "These are the product's verified handling duties for unresolved provisional facts, reconstructed from the recorded permission and scope — not adverse substantive-law or authority claims, and not quotations of a human marking command. For each duty, judge from the actual draft wording which passages semantically rely on its unknown fact: a passage relies on the unknown only if its truth or operative effect changes with that fact's value. Topic overlap alone is not dependency. A passage that merely states a separate future verification duty or refers to the fact as background, and whose truth and effect stand regardless of the unknown's current value, does not rely on it and stays unmarked. Report each dependent passage that does not visibly mark the fact as unresolved as a finding at its pid. Its instead keeps the passage's supported wording, qualifies the assertion or effect that relies on the unknown, and adds a visible [open item: ...] marker naming the unknown; a marker added to otherwise unchanged wording is not a drafting change. Such a finding may carry empty sources, the duty's source_refs being its provenance. Never answer the fact, assume its truth, or infer sending, signature or release approval. The verdict-shaped qualifications below (satisfied, not_applicable, action_status, open marker decisions) govern the later independent checklist reply, the sole verifier of these duties; here, report findings only."

/** The duties scoped to one planned document, rendered once for the early
 * review's prompt and key: the same authoritative items the later checklist
 * verifies, with the review-facing instruction. */
function documentMarkerDuties(directory: string, document: string) {
  return provisionalFactMarkerItems(directory).filter(({ item }) => item.target.document === document)
}

/** Empty when this document carries no duty, so a run without one keeps its
 * prompt. Human-derived quotes are escaped so they cannot close the block. */
export function provisionalMarkerReviewText(directory: string, document: string) {
  const duties = documentMarkerDuties(directory, document)
  if (!duties.length) return ""
  return escapeTagOpeners([
    PROVISIONAL_MARKER_REVIEW_INSTRUCTION,
    ...duties.map(duty => JSON.stringify(duty, null, 2)),
  ].join("\n\n"))
}

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
          "Dependency requires that the passage's truth or operative effect changes with the unknown's value — semantic reliance of an assertion or stated/implied effect on the unknown, not topic overlap alone. A separate future verification duty or background reference whose truth and effect stand regardless of the unknown's current value does not rely on it. Summaries, conditions and cross-references remain covered when they rely on it. An adverse omission must identify the dependent assertion or effect and its missing qualification.",
          "Return satisfied with action_status fulfilled, no open marker decisions, and exact visibly marked current body evidence only when every dependent passage is marked. The underlying factual uncertainty is not an open marker decision.",
          "Return not_applicable with action_status not_applicable and no open marker decisions only after inspecting the current document and establishing it contains no passage dependent on this fact. Missing inspection, body evidence or uncertain applicability is indeterminate, not not_applicable.",
        ],
        target: { document: document.id, role: `${document.title}: ${document.purpose}`, treatment: "open_issue" as const },
      },
    }))
  })
}
