import { z } from "zod"

const text = z.string().refine(value => value.trim().length > 0, "Expected nonempty text")

export const checklistSourceReferenceSchema = z.object({
  document: text,
  locator: text,
  quote: text,
}).strict()

export const checklistTargetDocumentSchema = z.object({ document: text, role: text }).strict()

/** Source-derived meaning, shared by producer and consumer. Routing and current
 * Word evidence belong outside this object; neither parser rewrites its text. */
export const checklistItemSchema = z.object({
  check: text,
  basis: z.enum(["directive", "proposal", "observation", "open_question"]),
  required_action: z.enum(["adopt", "propose", "flag", "assess"]).optional(),
  source_refs: z.array(checklistSourceReferenceSchema).min(1),
  qualifications: z.array(text),
  target: checklistTargetDocumentSchema.extend({
    treatment: z.enum(["operative", "advice", "comment", "open_issue"]),
  }).strict(),
}).strict()

export const checklistExtractionSchema = z.object({ items: z.array(checklistItemSchema) }).strict()

export const checklistProducerInputSchema = z.object({
  source: checklistTargetDocumentSchema,
  passages: z.array(checklistSourceReferenceSchema).min(1),
  instructions: z.array(checklistSourceReferenceSchema),
  targets: z.array(checklistTargetDocumentSchema).min(1),
  retrieval: text,
}).strict()

export const checklistCurrentEvidenceSchema = z.object({
  document: text,
  locator: text,
  text,
  kind: z.enum(["body", "comment", "deleted_text", "artifact"]),
}).strict()

export const checklistComparisonAssignmentSchema = z.object({
  kind: z.enum(["target_body", "process_records", "document_properties"]),
  ranges: z.array(z.object({ paragraph_index: z.number().int().nonnegative(), start_char: z.number().int().nonnegative(), end_char: z.number().int().nonnegative() }).strict()),
}).strict()
export type ChecklistComparisonAssignment = z.infer<typeof checklistComparisonAssignmentSchema>

export const checklistCheckerInputSchema = z.object({
  item_id: text,
  item: checklistItemSchema,
  assignment: checklistComparisonAssignmentSchema.optional(),
  current_evidence: z.array(checklistCurrentEvidenceSchema),
  settled_instructions: z.array(checklistSourceReferenceSchema),
  retrieval: text,
}).strict()

export const checklistFindingSchema = z.object({
  item_id: text,
  status: z.enum(["satisfied", "unsatisfied", "open", "indeterminate", "not_applicable"]),
  reason: text,
  evidence: z.array(checklistCurrentEvidenceSchema),
  source_refs: z.array(checklistSourceReferenceSchema),
  inspection_scope: text,
  limitation: z.string(),
  action_status: z.enum(["fulfilled", "not_fulfilled", "indeterminate", "not_applicable"]).optional(),
  open_decisions: z.array(text).optional(),
}).strict()

export type ChecklistTargetDocument = z.infer<typeof checklistTargetDocumentSchema>
export type ChecklistSourceReference = z.infer<typeof checklistSourceReferenceSchema>
export type ChecklistItem = z.infer<typeof checklistItemSchema>
export type ChecklistExtraction = z.infer<typeof checklistExtractionSchema>
export type ChecklistProducerInput = z.infer<typeof checklistProducerInputSchema>
export type ChecklistCurrentEvidence = z.infer<typeof checklistCurrentEvidenceSchema>
export type ChecklistCheckerInput = z.infer<typeof checklistCheckerInputSchema>
export type ChecklistFinding = z.infer<typeof checklistFindingSchema>

/** Recovery routing only: do not rewrite a finding or decide legal support.
 * A fulfilled proposal can remain open for approval. Conflicting assessments
 * require inspection rather than assuming a wording defect or a pass. */
export function checklistFindingRecovery(item: ChecklistItem, finding?: ChecklistFinding): "none" | "inspect" | "repair" {
  if (!finding) return "inspect"
  if (!item.required_action) return finding.status === "unsatisfied" ? "repair" : finding.status === "indeterminate" ? "inspect" : "none"
  if (!finding.action_status || finding.action_status === "indeterminate" || finding.status === "indeterminate") return "inspect"
  if (finding.action_status === "not_fulfilled")
    return finding.status === "open" || finding.status === "unsatisfied" ? "repair" : "inspect"
  return finding.status === "unsatisfied" ? "inspect" : "none"
}

// Parse shape only. Source support and applicability remain explicit review
// questions, rather than word matching rules hidden in a postprocessor.
export function parseChecklistItem(value: unknown): ChecklistItem {
  return checklistItemSchema.parse(value)
}

export function parseChecklistExtraction(value: unknown): ChecklistExtraction {
  return checklistExtractionSchema.parse(value)
}

export function parseChecklistCheckerInput(value: unknown): ChecklistCheckerInput {
  return checklistCheckerInputSchema.parse(value)
}

export function parseChecklistFinding(value: unknown): ChecklistFinding {
  return checklistFindingSchema.parse(value)
}

export const CHECKLIST_ITEM_SHAPE = '{"check":"...","basis":"directive|proposal|observation|open_question","source_refs":[{"document":"...","locator":"...","quote":"..."}],"qualifications":[],"target":{"document":"filename or declared document identity","role":"document purpose","treatment":"operative|advice|comment|open_issue"}}'

const sourceMeaning = "Keep directives, proposals, observations and open questions distinct. Original draft wording is an observation, not by itself an instruction to retain, revise or disclose it. Actual source directives remain instructions with their conditions. Do not infer approval, exclusivity, precedence or a memo duty from silence. An independent rationale cannot promote a proposal into a directive. Preserve material qualifications and unresolved conflicts."
const treatmentMeaning = "Treatment describes presentation, separately from basis and applicability: operative means substantive document wording; advice means explanatory or advisory prose, including memo commentary; comment means a native document annotation; open_issue means an unresolved matter without an invented resolution. Commentary in prose does not itself require native annotations. Require native comments only when source instructions support that presentation. If treatment is ambiguous, preserve the ambiguity rather than infer an annotation duty or an annotation-absence failure."
const applicabilityMeaning = "Proposal or question provenance alone creates no duty to reproduce it or keep it open in a target document. Before assigning or checking open_issue, determine whether the decision remains unsettled for the requested document under expressly controlling sources and current human instructions. If a source or instruction identifies controlling terms, inspect the relevant passage through targeted retrieval and preserve that relationship in source_refs and qualifications or the finding reason. A controlling position may settle the requested document's treatment without proving that an earlier proposal was accepted or withdrawn. Do not infer precedence from chronology or silence. Preserve genuinely unresolved proposals, material conflicts and any independent instruction to disclose or flag them; missing or conflicting authority remains uncertainty, never invented approval."
const retrieval = "The input is already inline. Use exact supplied coordinates when inspection is needed, then targeted read-record/read-document/search-document/verify-quote for a material gap. Do not discover directories, reread the packet/schema or read the full matter by default."

export function buildChecklistProducerPrompt(value: ChecklistProducerInput): string {
  const input = checklistProducerInputSchema.parse(value)
  return [
    "Extract source-grounded checklist items from this one original source and the relevant request. Do not inspect generated outputs, decide compliance, draft documents or author links.",
    sourceMeaning,
    treatmentMeaning,
    applicabilityMeaning,
    "Each item asks one focused check for one target document and states the source-supported treatment. Use the declared target identity/role, never invent a final paragraph pointer. Separate independent checks or target documents; keep necessary conditions attached. Structural artifact questions must say which property is to be inspected. Collection-wide questions must identify their scope rather than imply one passage proves completeness. If the sources do not settle treatment or applicability, preserve that uncertainty in qualifications. Do not create a duty merely to retain an observation.",
    retrieval,
    `Return only JSON {"items":[ITEM]}; ITEM uses exactly this shape: ${CHECKLIST_ITEM_SHAPE}. Choose one enum value, not a pipe-separated string. Copy exact source quotes and locators. No model-assigned IDs, extra fields or current-output evidence. An empty items array is permitted when this source yields no applicable checks.`,
    "SOURCE_INPUT:",
    JSON.stringify(input),
  ].join("\n\n")
}

// Whole-target properties have one assignment; passage and record assignments defer them.
const documentPropertiesDeferral = "A document-wide property is one established from the whole target's artifact properties, such as its exact text measures or revision counts, rather than from the wording of any passage. Only the separate document-properties assignment decides it. If the parent concerns only such a property, report not_applicable for this component rather than indeterminate; otherwise decide its other aspects here without treating that property as an inspection gap."
const documentPropertiesAssignment = "Decide ONLY document-wide properties of the named target from its supplied exact artifact properties, including accepted_body_text_measures; separate assignments compare every paragraph and the process records. Report not_applicable when the parent states no document-wide property. Do not assess passage wording here, and do not treat these properties as proof of content, layout or compliance beyond what they state; a property they cannot establish remains indeterminate."

export function buildChecklistCheckerPrompt(value: ChecklistCheckerInput): string {
  const input = parseChecklistCheckerInput(value)
  return [
    "Check this one unchanged source-derived checklist item against its target document. Do not rewrite, decompose or strengthen the item; do not draft or review the whole package.",
    `The producer and checker share this canonical item shape: ${input.item.required_action ? CHECKLIST_ITEM_SHAPE.slice(0, -1) + ",\"required_action\":\"adopt|propose|flag|assess\"}" : CHECKLIST_ITEM_SHAPE}. Current evidence and later instructions are supplied separately and do not mutate the item.`,
    ...(input.assignment ? [
      "This explicit comparison assignment is one bounded component of the preserved parent, not a new requirement. Decide only this component; other assignments cover the remaining paragraphs, process records and document-wide properties. Preserve source-supported applicability and uncertainty. Do not demand whole-document coverage in this one finding, and do not claim unassigned coverage.",
      ...(input.assignment.kind === "document_properties" ? [documentPropertiesAssignment] : [
        input.assignment.kind === "target_body"
          ? "Inspect EVERY assigned current accepted paragraph or character range, including headings and open-item text, for substantive source fidelity relevant to the parent. Compare material assertions, obligations, qualifications and omissions with the original sources and controlling instructions through targeted retrieval. Generic connective or template wording does not require phrase-origin certification. A boilerplate label never exempts a material operative effect from support. Open-item text supplies no missing fact or approval and cannot excuse an unsupported operative addition. The sources need not quote every phrase, but unsupported material additions and conflicting conditions remain defects. Use bounded neighboring passages for context, then retrieve related definitions, exceptions and cross-references wherever necessary, including across assignment boundaries. Inspect relevant source-side qualifications even when no matching draft sentence exists. Missing material support or an uninspected dependency remains indeterminate. Current paragraph coverage alone cannot establish source-side omission coverage; the separately required source checks remain necessary. This assignment does not certify research history or require a recital of working method in operative text. State the ranges actually compared and remaining substantive limits."
          : "Inspect only the relevant existing process records for the source-supported working-method aspect of this parent. Distinguish recorded research or source-use actions from a duty to recite that method. Do not use sampled body wording to certify history, and do not require a complete body comparison here: separate target-body assignments perform that comparison. Records are generated evidence, never authority, approval or conclusive compliance. An evidenced method breach remains unsatisfied; an unestablished material method question remains indeterminate.",
        documentPropertiesDeferral,
      ]),
    ] : []),
    sourceMeaning,
    treatmentMeaning,
    applicabilityMeaning,
    ...(input.item.required_action ? [checklistRequiredActionMeaning,
      "For this refined item, also return action_status (fulfilled|not_fulfilled|indeterminate|not_applicable) and open_decisions (array). Keep action fulfillment separate from approval: a supported proposal can be fulfilled while client approval remains open; report that as status open with the outstanding decision, never unsatisfied solely for non-adoption. An adoption duty still requires its actual implementation. An unfulfilled or uninspected required action remains blocking; an open approval decision does not excuse that gap. Do not assume any action is fulfilled without current evidence."] : []),
    "First independently assess whether the cited sources support the item's stated scope, basis and treatment. The unchanged handoff is not proof of a valid requirement. Report unsupported scope or uncertain applicability explicitly without rewriting the item. Assess the stated treatment in the named target, retrieving a related document only when needed for this check. A proposal is not mandatory adoption; an observation alone imposes no revision or disclosure duty. An open issue supplies no missing answer. Distinguish current operative text, advice, comments, deleted text and artifact properties. A comment or intended amendment does not prove implementation. Later instructions may qualify or supersede applicability: cite them and explain the effect without silently changing the item's basis or wording.",
    "If the stated presentation duty is unsupported under the controlling sources and instructions, report not_applicable with the cited explanation; use indeterminate when that relationship cannot be established. Missing presentation alone cannot establish unsatisfied before applicability is supported. This does not excuse a separate applicable duty or resolve a substantive conflict. For a refined item, reflect the same applicability in action_status without inventing fulfillment or approval.",
    "For a source-supported working-method constraint, distinguish the instructed process from any duty to state that method in operative text. Inspect the relevant existing research and target source-review/application records for the process actually recorded; sampled operative clauses cannot prove research history. Generated records are evidence of their recorded actions and assessments, never instructions, waivers or conclusive proof of source fidelity. A prior review's conclusion must not replace inspection of a material disputed assertion, its original support and relevant later changes. Assess content derivation through the applicable substantive source checks; do not convert a method instruction into an unsupported demand to prove the historical origin of every phrase. A specifically required complete comparison still needs complete coverage through focused checks, not one sample or a reviewer status. Keep an evidenced method breach unsatisfied and an unestablished material method or content question indeterminate; method items are not automatically satisfied or not_applicable.",
    retrieval,
    "Missing excerpts do not prove absence. Inspect the relevant target sufficiently before asserting missing wording; otherwise report indeterminate with the gap. Satisfied means the applicable stated treatment is evidenced; unsatisfied requires an evidenced applicable mismatch; open preserves a substantive unresolved fact or decision; indeterminate means insufficient evidence or inspection; not_applicable needs an evidence-based explanation. State the actual inspection scope and do not claim whole-package coverage.",
    `Return only JSON {"item_id":"supplied ID","status":"satisfied|unsatisfied|open|indeterminate|not_applicable","reason":"...","evidence":[{"document":"...","locator":"...","text":"exact current evidence","kind":"body|comment|deleted_text|artifact"}],"source_refs":[{"document":"...","locator":"...","quote":"exact source or instruction"}],"inspection_scope":"...","limitation":"..."${input.item.required_action ? ',"action_status":"fulfilled|not_fulfilled|indeterminate|not_applicable","open_decisions":[]' : ""}}. Choose one enum value. Empty evidence arrays preserve gaps; do not invent quotations or locators.`,
    "CHECK_INPUT:",
    JSON.stringify(input),
  ].join("\n\n")
}


/** Generated arguments are counterevidence, never new source instructions. */
export const checklistAdjudicationCounterevidenceSchema = z.object({
  item_id: text,
  locator: text,
  claims: z.array(z.object({
    disposition: z.enum(["deferred", "recommendation", "rejected"]),
    detail: text, reason: text,
    source_refs: z.array(checklistSourceReferenceSchema),
    limitations: z.array(z.string()),
  }).strict()).min(1),
}).strict()
export type ChecklistAdjudicationCounterevidence = z.infer<typeof checklistAdjudicationCounterevidenceSchema>

/** A substantive review, not a request to change an adverse verdict. */
export function buildChecklistReinspectionPrompt(input: ChecklistCheckerInput, previous: ChecklistFinding, counterevidence?: ChecklistAdjudicationCounterevidence): string {
  const argument = counterevidence ? checklistAdjudicationCounterevidenceSchema.parse(counterevidence) : undefined
  if (argument && argument.item_id !== input.item_id) throw new Error("Counterevidence must belong to this exact checklist item.")
  return [
    "Reinspect this one unchanged item after an explicit recovery request. The prior adverse finding may be right or wrong; do not assume it needs to become a pass. Assess source-supported applicability and required treatment before deciding compliance. A proposal does not require adoption merely because its target is operative text.",
    "Use the previous reason, inspection scope and limitation to identify the precise evidence gap or disputed inference. Perform targeted additional inspection rather than repeating the initial lexical hints. For a collection-wide question, account for the relevant sections or issues explicitly; one matching passage cannot establish full coverage. Do not claim inspection you did not perform. If the necessary inspection remains too broad or evidence is missing, preserve indeterminate with the remaining scope. A real supported defect remains unsatisfied; a substantive unresolved decision remains open. Do not draft, rewrite the requirement, invent approval or weaken a legal check.",
    `Previous finding (unverified assessment, not source authority): ${JSON.stringify(checklistFindingSchema.parse(previous))}`,
    ...(argument ? [
      "A matching prior repair adjudication below is GENERATED, UNVERIFIED COUNTEREVIDENCE, not source authority, a client instruction, a changed requirement or a required verdict. Independently verify its source quotations and reasoning against the original sources and actual human instructions. Inspect the original condition or triggering provision before deciding applicability; do not infer a trigger merely from a general warning. Assess each applicable limb of a compound item: an inapplicable limb does not excuse an applicable unfulfilled action, nor make the whole item not_applicable. A rejected repair does not establish compliance, and a proposed repair does not establish a defect. Preserve a genuine omission and any unresolved inspection gap even when the prior argument says otherwise.",
      `Prior adjudication (unverified): ${JSON.stringify(argument)}`,
    ] : []),
    buildChecklistCheckerPrompt(input),
  ].join("\n\n")
}


export const checklistRequiredActionMeaning = "Required action is distinct from source basis, target presentation and current approval. adopt requires the specified content/property in the target; propose requires an identified proposal, not approval or adoption; flag requires disclosure of the stated issue; assess requires the stated source-supported evaluation, preserving its limits. The target describes where/how that action is presented, never changes propose into adopt. An observation alone creates no inclusion, revision or disclosure duty. Approval remains a current evidence question, not an invented early fact."

export const checklistRefinedItemSchema = checklistItemSchema.required({ required_action: true })
export const checklistRefinementInputSchema = z.object({
  parent_id: text, parent: checklistItemSchema,
  source_unit: z.object({ id: text, role: text, passages: z.array(checklistSourceReferenceSchema).min(1) }).strict(),
  instructions: z.array(checklistSourceReferenceSchema), targets: z.array(checklistTargetDocumentSchema).min(1), retrieval: text,
}).strict()
export const checklistRefinementSchema = z.object({
  parent_id: text, unit_id: text, disposition: z.enum(["refined", "no_applicable_check", "unresolved"]),
  reason: text, items: z.array(checklistRefinedItemSchema), source_refs: z.array(checklistSourceReferenceSchema).min(1), limitations: z.array(text),
}).strict()
export type ChecklistRefinementInput = z.infer<typeof checklistRefinementInputSchema>
export type ChecklistRefinement = z.infer<typeof checklistRefinementSchema>
export function parseChecklistRefinement(value: unknown) { return checklistRefinementSchema.parse(value) }

export function buildChecklistRefinementPrompt(value: ChecklistRefinementInput): string {
  const input = checklistRefinementInputSchema.parse(value)
  return [
    "Refine one provisional parent requirement ONLY within the assigned original source unit. Do not inspect generated documents or prior findings, decide compliance, draft, or infer the desired result. The original parent stays preserved; its wording is not proof of source authority.",
    sourceMeaning, treatmentMeaning, applicabilityMeaning, checklistRequiredActionMeaning,
    "Return focused child checks, one source-supported action and target each, preserving necessary conditions, material exceptions and uncertainty. Do not bundle independent concerns. A collection-wide parent is not a child: extract the concrete concerns supported in THIS unit. Coverage of other units is a separate job. Original draft text alone does not impose retention or memo duties. If no applicable check is supported, explain that disposition; if scope/treatment cannot be settled, preserve unresolved and the precise gap rather than invent approval or omit the concern. Preserve every source-supported applicable aspect of the parent within this assigned unit, including its target document and treatment; action may clarify the source-supported duty but cannot silently change it. Do not replace a parent duty with a different target or presentation merely because another action is supported. If the parent target, treatment or scope is unsupported, or any applicable aspect cannot be preserved in focused children, return unresolved with a precise source-grounded reason; partial children may remain, but cannot certify the parent. No_applicable_check is only for a unit that contributes no applicable aspect, not an unsupported parent duty.",
    retrieval,
    `Return only JSON {"parent_id":"supplied ID","unit_id":"supplied source_unit.id","disposition":"refined|no_applicable_check|unresolved","reason":"source-grounded explanation","items":[ITEM],"source_refs":[{"document":"...","locator":"...","quote":"exact source"}],"limitations":[]}. ITEM uses ${CHECKLIST_ITEM_SHAPE.slice(0, -1)},"required_action":"adopt|propose|flag|assess"}. Children have no model-assigned IDs. Empty items requires an explicit no_applicable_check or unresolved disposition; unresolved may retain partial children. This job cannot certify the parent or unassigned source units complete.`,
    "REFINEMENT_INPUT:", JSON.stringify(input),
  ].join("\n\n")
}
