import { z } from "zod"
import {
  buildChecklistRefinementPrompt,
  checklistRefinementInputSchema,
  parseChecklistRefinement,
  type ChecklistRefinementInput,
  type ChecklistRefinement,
  buildChecklistCheckerPrompt,
  buildChecklistReinspectionPrompt,
  buildChecklistProducerPrompt,
  parseChecklistExtraction,
  parseChecklistFinding,
  suppliedEvidenceContains,
  type ChecklistAdjudicationCounterevidence,
  type ChecklistCheckerInput,
  type ChecklistExtraction,
  type ChecklistFinding,
  type ChecklistProducerInput,
  type ChecklistSourceReference,
  type ChecklistCurrentEvidence,
} from "./checklist-contract"
import { LEAN_MAX_RETRIES, type PermissionRule, type StageSessionPort } from "./stages"
import type { LeanModel } from "./state"
import { quoteFound, quoteMissIssue } from "./quotes"
import { normalize } from "../../lib/quote-verify"
import { buildChecklistDocumentContext, checklistWorkflowEvidence } from "./checklist-document-context"

export type ChecklistJobDeps = { sessions: StageSessionPort; model: LeanModel | null; parentID?: string }
export type ChecklistJobResult<T> =
  | { status: "succeeded"; sessionID: string; attempts: number; reply: string; value: T }
  | { status: "failed"; sessionID: string; attempts: number; reply: string; issues: string[] }

export const CHECKLIST_JOB_AGENTS = { extract: "lean-checklist-extract", check: "lean-checklist-check" } as const
export const CHECKLIST_JOB_PERMISSIONS: PermissionRule[] = [
  { action: "*", resource: "*", effect: "deny" },
  ...["read-record", "read-document", "search-document", "verify-quote"].map(action => ({ action, resource: "*", effect: "allow" as const })),
  { action: "external_directory", resource: "*", effect: "deny" },
]

export function runChecklistExtractionJob(directory: string, deps: ChecklistJobDeps, input: ChecklistProducerInput): Promise<ChecklistJobResult<ChecklistExtraction>> {
  return runChecklistJob(directory, deps, CHECKLIST_JOB_AGENTS.extract, buildChecklistProducerPrompt(input), parseChecklistExtraction, value => value.items.flatMap((item, index) => [
    ...sourceReferenceIssues(directory, item.source_refs, input.instructions, `items[${index}].source_refs`),
    ...(!input.targets.some(target => target.document === item.target.document)
      ? [`items[${index}].target.document must use a declared document identity: ${JSON.stringify(input.targets)}. Do not invent or retarget a document.`] : []),
  ]))
}

export function runChecklistCheckJob(directory: string, deps: ChecklistJobDeps, input: ChecklistCheckerInput, previous?: ChecklistFinding, counterevidence?: ChecklistAdjudicationCounterevidence): Promise<ChecklistJobResult<ChecklistFinding>> {
  return runChecklistJob(directory, deps, CHECKLIST_JOB_AGENTS.check, previous ? buildChecklistReinspectionPrompt(input, previous, counterevidence) : buildChecklistCheckerPrompt(input), parseChecklistFinding, value => [
    ...sourceReferenceIssues(directory, value.source_refs, input.settled_instructions, "source_refs"),
    ...currentEvidenceIssues(directory, value.evidence, input),
    ...(input.item.required_action && (!value.action_status || !value.open_decisions) ? ["Refined-item findings must include action_status and open_decisions, separately from overall status and source applicability."] : []),
    ...(value.item_id !== input.item_id ? [`item_id must equal the supplied ${JSON.stringify(input.item_id)}; do not replace the item identity.`] : []),
    ...(value.deferred_to && input.assignment?.kind !== "target_body" && input.assignment?.kind !== "process_records" ? ["deferred_to is only for a range or process-records assignment that leaves a document-wide property to the document-properties assignment; omit it here."] : []),
    ...(value.indeterminate_cause && value.status !== "indeterminate" && value.action_status !== "indeterminate" ? ["indeterminate_cause applies only when status or action_status is indeterminate; otherwise omit it."] : []),
    ...disclosureIssues(value, "conflicting_source_refs", value.conflicting_source_refs, refs => sourceReferenceIssues(directory, refs, input.settled_instructions, "conflicting_source_refs")),
    ...disclosureIssues(value, "departure_disclosure", value.departure_disclosure, refs => currentEvidenceIssues(directory, refs, input, "departure_disclosure")),
  ], [...CHECKLIST_JOB_PERMISSIONS, ...checklistWorkflowEvidence(directory, input.item).map(record => ({ action: "read", resource: record.artifact, effect: "allow" as const }))])
}

/** A conflict or departure explains an adverse finding only; an empty array states none. */
function disclosureIssues<T>(finding: ChecklistFinding, field: string, refs: T[] | undefined, verify: (refs: T[]) => string[]): string[] {
  if (!refs?.length) return []
  if (finding.status === "satisfied" || finding.status === "not_applicable") return [`${field} applies only to an unsatisfied, open or indeterminate status; otherwise omit it.`]
  return verify(refs)
}

/** A source-only unit refinement uses the same native read boundary and correction. */
export function runChecklistRefinementJob(directory: string, deps: ChecklistJobDeps, value: ChecklistRefinementInput): Promise<ChecklistJobResult<ChecklistRefinement>> {
  const input = checklistRefinementInputSchema.parse(value)
  return runChecklistJob(directory, deps, CHECKLIST_JOB_AGENTS.extract, buildChecklistRefinementPrompt(input), parseChecklistRefinement, result => [
    ...(result.parent_id !== input.parent_id ? ["parent_id must equal the supplied parent_id."] : []),
    ...(result.unit_id !== input.source_unit.id ? ["unit_id must equal the assigned source_unit.id."] : []),
    ...(!result.items.length && result.disposition === "refined" ? ["An empty items array needs no_applicable_check or unresolved disposition with a source-grounded reason."] : []),
    ...(result.items.length && result.disposition === "no_applicable_check" ? ["no_applicable_check cannot silently retain child duties; use refined or unresolved as supported."] : []),
    ...sourceReferenceIssues(directory, result.source_refs, input.instructions, "source_refs"),
    ...result.items.flatMap((item, index) => [
      ...sourceReferenceIssues(directory, item.source_refs, input.instructions, `items[${index}].source_refs`),
      ...(result.disposition === "refined" && (item.target.document !== input.parent.target.document || item.target.treatment !== input.parent.target.treatment)
        ? [`items[${index}] changes the preserved parent target document or treatment. If the parent duty cannot be supported and preserved, return unresolved with the source-grounded reason; otherwise return genuinely source-supported same-target children. Do not invent a duty to satisfy this constraint.`] : []),
      ...(!input.targets.some(target => target.document === item.target.document) ? [`items[${index}].target.document must use a declared identity.`] : []),
    ]),
  ])
}

/** One product job owns the native session and its bounded contract correction.
 * Diagnostics call this once; they do not continue or repair its session. */
async function runChecklistJob<T>(directory: string, deps: ChecklistJobDeps, agent: string, prompt: string, parse: (value: unknown) => T, validate: (value: T) => string[], permissions: PermissionRule[] = CHECKLIST_JOB_PERMISSIONS): Promise<ChecklistJobResult<T>> {
  const sessionID = await deps.sessions.createSession({
    directory, parentID: deps.parentID, model: deps.model, agent,
    title: agent === CHECKLIST_JOB_AGENTS.extract ? "Extract one source checklist" : "Check one checklist item",
    permissions,
  })
  const attempt = async (index: number, text: string): Promise<ChecklistJobResult<T>> => {
    // Engine/provider failures propagate; they are not malformed-output retries.
    const reply = await deps.sessions.prompt({ directory, sessionID, agent, text })
    const checked = parseReply(reply, parse)
    const issues = checked.ok ? validate(checked.value) : checked.issues
    if (checked.ok && !issues.length) return { status: "succeeded", sessionID, attempts: index + 1, reply, value: checked.value }
    if (index >= LEAN_MAX_RETRIES) return { status: "failed", sessionID, attempts: index + 1, reply, issues }
    return attempt(index + 1, `Your reply did not pass the checklist contract:\n${issues.map(issue => `- ${issue}`).join("\n")}\nReturn the complete JSON output using the original task and canonical evidence. Correct only these issues; do not strengthen requirements, answer open questions or invent evidence.`)
  }
  return attempt(0, prompt)
}

function parseReply<T>(reply: string, parse: (value: unknown) => T): { ok: true; value: T } | { ok: false; issues: string[] } {
  try {
    return { ok: true, value: parse(JSON.parse(reply)) }
  } catch (error) {
    if (!(error instanceof SyntaxError) && !(error instanceof z.ZodError)) throw error
    return { ok: false, issues: error instanceof z.ZodError
      ? error.issues.map(issue => `${issue.path.join(".") || "reply"}: ${issue.message}`)
      : ["Reply must be a complete JSON object only, without Markdown fences or surrounding prose."] }
  }
}

function sourceReferenceIssues(directory: string, refs: ChecklistSourceReference[], instructions: ChecklistSourceReference[], where: string): string[] {
  return refs.flatMap((ref, index) => {
    // Only caller-supplied canonical instructions supplement the matter. An
    // output's own references never establish their own source authority.
    const supplied = instructions.some(instruction => instruction.document === ref.document &&
      normalize(instruction.quote).includes(normalize(ref.quote)))
    if (supplied || quoteFound(directory, { source_id: ref.document, quote: ref.quote })) return []
    return [quoteMissIssue(directory, `${where}[${index}]`, { source_id: ref.document, quote: ref.quote }, "json")]
  })
}

/** Output quotations never establish their own authority. Initial caller-owned
 * hints and freshly read native Word records are independent evidence pools;
 * annotations cannot establish operative body wording. */
function currentEvidenceIssues(directory: string, refs: ChecklistCurrentEvidence[], input: ChecklistCheckerInput, where = "evidence"): string[] {
  const supplied = (ref: ChecklistCurrentEvidence) => input.current_evidence.some(evidence => suppliedEvidenceContains(evidence, ref))
  const unresolved = refs.filter(ref => !supplied(ref))
  if (!unresolved.length) return []
  const current = buildChecklistDocumentContext(directory, { item_id: input.item_id, item: input.item }).artifacts.map(artifact => artifact.value)
  const workflow = checklistWorkflowEvidence(directory, input.item)
  return refs.flatMap((ref, index) => {
    if (supplied(ref)) return []
    const needle = normalize(ref.text)
    if (ref.kind === "artifact" && needle && workflow.some(record => record.artifact === ref.document &&
      [record.text, ...(record.artifact.endsWith(".json") ? recordStrings(JSON.parse(record.text)) : [])].some(text => normalize(text).includes(needle)))) return []
    // A name shared by several planned records is not a usable identity.
    const records = current.filter(record => [...record.aliases, record.artifact, ...(record.annotations ?? []).map(annotation => annotation.artifact)].includes(ref.document))
    let texts: string[] = []
    if (records.length === 1) {
      const record = records[0]!
      const bodyIdentity = [...record.aliases, record.artifact].includes(ref.document)
      const annotations = (record.annotations ?? []).filter(annotation => record.aliases.includes(ref.document) || annotation.artifact === ref.document)
      if (ref.kind === "body" && bodyIdentity) texts = [record.paragraphs.map(paragraph => paragraph.paragraph_text).join("\n")]
      if (ref.kind === "comment") texts = annotations.flatMap(annotation => annotation.comments.map(comment => comment.text))
      if (ref.kind === "deleted_text") texts = annotations.flatMap(annotation => annotation.revisions.deleted_text)
      // Exact native property summaries and inserted revision strings establish
      // only their artifact observations, never whole-document compliance.
      if (ref.kind === "artifact") texts = [...(bodyIdentity ? [record.artifact, record.sha256] : []), ...annotations.flatMap(annotation => [annotation.artifact, annotation.sha256, annotation.properties, ...annotation.revisions.inserted_text])]
    }
    const found = needle && texts.some(text => ref.kind === "artifact" ? normalize(text) === needle : normalize(text).includes(needle))
    return found ? [] : [`${where}[${index}] text is not verified in current ${ref.kind} evidence for document ${JSON.stringify(ref.document)}: ${JSON.stringify(ref.text)}. Retrieve exact current evidence from that document and evidence kind; comments, deleted text and artifact properties do not prove operative body wording. Correct the quote/kind/document, or omit unsupported evidence and preserve the uncertainty.`]
  })
}

/** Native read-record returns decoded scalar strings, including escaped quotes. */
function recordStrings(value: unknown): string[] {
  if (typeof value === "string") return [value]
  if (Array.isArray(value)) return value.flatMap(recordStrings)
  return value && typeof value === "object" ? Object.values(value).flatMap(recordStrings) : []
}
