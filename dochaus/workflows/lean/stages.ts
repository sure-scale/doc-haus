import { draftingStyleAdvisoryNotes } from "./style-warnings"
import { checklistCheckResultSchema, readEffectiveChecklistItems, reconcileChecklistParentCoverage } from "./checklist-flow"
import { leanClarifications, checklistRefinementSelections } from "./requests"
import { z } from "zod"
import { checklistFindingRecovery, checklistItemSchema, type ChecklistFinding, type ChecklistItem } from "./checklist-contract"
import { acceptChangePlan, acceptChangeReview, preparationOutputValid, readChangePlan, readChangeReview, deferredFindingItems, reviewDispositionText, uniqueReplacementIssues } from "./change-pipeline"
import { comparisonBasis } from "./comparison"
import { checkPackageReview, readPackageReview } from "./package-review"
import { evidenceOutputValid } from "./evidence"
import { researchOutputValid } from "./research"
import { readFirmDraftingPreferences } from "../../lib/drafting-preferences"
import { existsSync, readFileSync } from "node:fs"
import { ensureParagraphPointers, paragraphSha256 } from "../../lib/drafting-amendment"
import { renderMarkdownToDocx } from "../../lib/render"
import { presentDraft, type DraftPresentation } from "./draft-presentation"
import { checkChanges } from "./change-checks"
import {
  checkDraft,
  finalBodyParagraphs,
  checkDraftingNotes,
  checkReview,
  readChangeList,
  readParagraphs,
  type ApplyResult,
  type ParagraphRow,
  type StageCheck,
} from "./checks"
import { reviewStages, FIRST_DOCUMENT, planActingFor, operativeDocument, operativePlan, plannedDocument, readLeanPlan, stepDocument, stepStage, type StepID } from "./plan"
import { completedDocumentsText } from "./prompts"
import { matterSourceBasis, matterSourceIDs } from "./quotes"
import { documentRepairs, readRepairRound } from "./repairs"
import {
  isReviewStage,
  answeredQuestionsText,
  hashStageInputs,
  leanFile,
  readLeanText,
  readLeanJson,
  readOpenItems,
  readQuestions,
  writeLeanBytes,
  writeLeanText,
  type LeanModel,
  type OpenItem,
  type Question,
  type StageID,
} from "./state"

// Step table for the lean drafting profile. Each model step is one native child
// session of a dedicated read-only lean agent (dochaus/agent/lean-*.md). The coordinator hands the agent its
// inputs, saves the agent's final reply as the step's output files, and checks
// it deterministically. It never interprets the legal content: the change list
// goes to the pointer applier as data, addressed by paragraph pointer. Steps of
// one planned document keep their files under that document's id (d1/, d2/...).

export const LEAN_PROMPT_VERSION = "lean-14"
const OUTLINE_CONTRACT = "document-playbook-selection-3"
const DELIVERY_REVIEW_CONTRACT = "grounded-delivery-selection-3"
const STRUCTURAL_AMENDMENT_CONTRACT = "explicit-paragraph-cardinality-1"
const DRAFTING_NOTES_CONTRACT = "lawyer-facing-source-comparison-notes-1"
export const LEAN_MAX_RETRIES = 3

export const LEAN_STAGE_AGENTS = {
  discrepancy: "lean-discrepancy",
  spine: "lean-spine",
  "review-independent": "lean-review-independent",
  "review-playbook": "lean-review-playbook",
  "review-wording": "lean-review-wording",
  "review-client-disadvantage": "lean-review-client-disadvantage",
  "package-review": "lean-package-review",
  intake: "lean-intake",
  research: "lean-research",
  "research-review": "lean-research-review",
  draft: "lean-draft",
  review: "lean-review",
  "change-plan": "lean-change-plan",
  "change-prepare": "lean-change-prepare",
  "change-review": "lean-change-review",
  changes: "lean-changes",
  merge: "lean-merge",
  finalize: "lean-drafting-notes",
} as const

export type PermissionRule = { action: string; resource: string; effect: "allow" | "deny" | "ask" }

/** The lean agents' own frontmatter allows exactly these tools. Native rules apply
 * last-match-wins after the agent's rules, so the session restates the same
 * boundary: stage sessions read the matter and its sources and write nothing. */
export const LEAN_READ_TOOLS = ["read", "search-document", "read-document", "read-record", "verify-quote"] as const
export const LEAN_CHILD_PERMISSIONS: PermissionRule[] = [
  { action: "*", resource: "*", effect: "deny" },
  ...LEAN_READ_TOOLS.map(action => ({ action, resource: "*", effect: "allow" as const })),
]

export type StageSessionPort = {
  createSession(input: {
    directory: string
    parentID?: string
    agent: string
    title: string
    model: LeanModel | null
    permissions?: PermissionRule[]
    metadata?: Record<string, unknown>
  }): Promise<string>
  /** Admits one prompt and resolves with the agent's final reply once the native
   * session is idle again. Throws on an engine or provider failure. */
  prompt(input: { directory: string; sessionID: string; agent: string; text: string }): Promise<string>
}

export type LeanDeps = {
  sessions: StageSessionPort
  /** Registers one finished file as a matter deliverable under `fileName`:
   * created when `priorSha256` is null, otherwise replacing a file that must
   * still hold exactly those bytes. Returns the matter-relative path, or
   * undefined when that name is held by a document the run must not overwrite. */
  registerDeliverable?: (directory: string, fileName: string, bytes: Uint8Array, priorSha256: string | null) => Promise<string | undefined>
}

export { engineSessionPort } from "./session-port"

export type ModelStageID = Exclude<StageID, "apply" | "gate" | "checklist-check">
export type StageAcceptance = StageCheck & { outputs: string[]; questions?: Question[] }

/** Native session title of a step's child session: the document's planned
 * title once the plan exists. */
export function stepSessionTitle(directory: string, step: StepID, detail: string = stepStage(step)) {
  const document = stepDocument(step)
  const title = document ? plannedDocument(directory, document)?.title : undefined
  return title ? `Lean drafting: ${title}: ${detail}` : `Lean drafting: ${detail}`
}

/** Output files of one step, relative to the lean directory. The first
 * document's draft also saves the run's plan. */
export function stepOutputs(step: StepID): string[] {
  if (step === "discrepancy" || step === "spine" || step === "package-review") return [`${step}.json`]
  if (step === "checklist-check") return ["checklist/extraction.json", "checklist/refinements.json", "checklist/checks.json", "checklist/report.md"]
  if (step === "intake") return ["intake.json"]
  if (step === "research") return ["research.json", "research-context.txt"]
  if (step === "research-review") return ["research-review.json"]
  const document = stepDocument(step)
  const at = (names: string[]) => names.map(name => `${document}/${name}`)
  const stage = stepStage(step)
  if (stage === "draft") return [...(document === FIRST_DOCUMENT ? ["plan.json"] : []), ...at(["draft.md", "body.docx", "paragraphs.json"])]
  if (isReviewStage(stage)) return at([`${stage}.md`])
  if (stage === "change-plan") return at(["change-plan.json"])
  if (stage === "change-prepare") return at(["change-preparation.json"])
  if (stage === "change-review") return at(["change-review.json"])
  if (stage === "changes") return at(["changes.json"])
  if (stage === "apply") return at(["apply.json", "body-applied.docx", "paragraphs-applied.json", "redline.docx"])
  if (stage === "finalize") return ["drafting-notes.md", "drafting-notes.docx"]
  return ["gate.json", "drafting-notes-coverage.md"]
}

// Items the drafting notes must disclose. Draft and review findings are re-derived from the
// final body at apply, so the notes disclose what the delivered body still has.
const DISCLOSED_STAGES: StageID[] = ["research", "research-review", "change-plan", "change-prepare", "changes", "apply"]

/** Named inputs of each step: everything its prompt or code step reads. Their
 * content hash is the step's resume key and nothing more. Answers are inputs of
 * every step, including the outline: an answer can change the basis of earlier
 * drafting and review, so those steps must be revisited before applying changes.
 * Every draft is keyed on its own plan entry, so a lawyer's edit of the plan
 * redrafts the documents whose entry changed, whatever their position.
 * An earlier document's final text and accepted explanations are keyed into
 * every later stage that consumes them, including drafting; changing either
 * revisits that dependent work. The changes step's own previous output and open questions
 * are shown to it but are not inputs, so its output never changes its own key.
 * Whom the lawyer acts for is an input of every model step (finalize reads it
 * through plan.json), so a lawyer's edit of it redrafts and re-reviews. The
 * planned documents' ids, titles and purposes are inputs of review and changes,
 * which are scoped by them; an edit of a document's parts re-keys neither. */
export function stageInputs(directory: string, step: StepID, options: { repairContext?: false } = {}): { name: string; bytes: string | Uint8Array }[] {
  const updates = readLeanText(directory, "updated-notes.md")
  const clarifications = leanClarifications(directory)
  const inputs = baseStageInputs(directory, step, options)
  return [...inputs,
    ...(!inputs.some(part => part.name === "drafting-preferences") ? [{ name: "drafting-preferences", bytes: readFirmDraftingPreferences() }] : []),
    ...(["change-plan", "change-prepare", "changes", "change-review", "apply", "finalize", "package-review", "gate"].includes(stepStage(step))
      ? [{ name: "structural-amendment-contract", bytes: STRUCTURAL_AMENDMENT_CONTRACT }] : []),
    ...(step === "finalize" ? [{ name: "drafting-notes-contract", bytes: DRAFTING_NOTES_CONTRACT }] : []),
    ...(["finalize", "package-review", "gate"].includes(step) ? [{ name: "review-dispositions", bytes: JSON.stringify((readLeanPlan(directory)?.documents ?? []).map(row => reviewDispositionText(directory, row.id))) }] : []),
    ...(["finalize", "package-review"].includes(step) ? [{ name: "delivery-review-contract", bytes: DELIVERY_REVIEW_CONTRACT }] : []),
    ...(!["discrepancy", "intake", "research", "research-review"].includes(step) ? ["research.json", "research-review.json"].map(name => ({ name, bytes: readLeanText(directory, name) ?? "" })) : []),
    ...(!["discrepancy", "intake", "research", "research-review", "spine"].includes(step) ? ["discrepancy.json", "spine.json"].map(name => ({ name, bytes: readLeanText(directory, name) ?? "" })) : []),
    ...(clarifications.length && ["checklist-check", "package-review", "gate"].includes(step) ? [{ name: "inspection-clarifications", bytes: JSON.stringify(clarifications) }] : []),
    { name: "source-content", bytes: matterSourceBasis(directory) },
    ...(!inputs.some(part => part.name === "answers") ? [{ name: "answers", bytes: answeredQuestionsText(directory) }] : []),
    ...(updates && !inputs.some(part => part.name === "updated-notes.md") ? [{ name: "updated-notes.md", bytes: updates }] : [])]
}

function baseStageInputs(directory: string, step: StepID, options: { repairContext?: false }): { name: string; bytes: string | Uint8Array }[] {
  const file = (name: string) => ({ name, bytes: existsSync(leanFile(directory, name)) ? readFileSync(leanFile(directory, name)) : "" })
  const answers = { name: "answers", bytes: answeredQuestionsText(directory) }
  const sources = { name: "sources", bytes: matterSourceIDs(directory).join("\n") }
  const planned = readLeanPlan(directory)?.documents.map(row => row.id) ?? []
  if (step === "checklist-check") return [{ name: "checklist-refinement-contract", bytes: "source-unit-refinement-1" }, { name: "checklist-refinement-selection", bytes: JSON.stringify(checklistRefinementSelections(directory)) }, file("checklist/refinements.json"), file("request.md"), file("plan.json"), file("checklist/extraction.json"), file("drafting-notes.docx"), file("research-context.txt"), ...planned.flatMap(id => [file(`${id}/body.docx`), file(`${id}/body-applied.docx`), file(`${id}/redline.docx`), file(`${id}/review-independent.md`), file(`${id}/changes.json`), file(`${id}/apply.json`)])]
  if (step === "discrepancy") return outlineInputs(directory)
  if (step === "intake") return [...outlineInputs(directory), file("discrepancy.json")]
  if (step === "spine") return [...outlineInputs(directory), file("discrepancy.json"), file("intake.json"), file("research.json"), file("research-review.json")]
  // Applied deletion evidence includes the original paragraph's exact wording.
  if (step === "package-review") return [file("checklist/refinements.json"), file("checklist/checks.json"), file("checklist/report.md"), file("request.md"), file("plan.json"), file("drafting-notes.md"), ...planned.flatMap(id => [file(`${id}/paragraphs.json`), file(`${id}/paragraphs-applied.json`), file(`${id}/apply.json`), file(`${id}/changes.json`), file(`${id}/redline.docx`), ...(plannedDocument(directory, id)?.playbook_review?.sources.length ? [file(`${id}/review-playbook.md`)] : [])]), file("open-items.json")]
  if (step === "research") return [...outlineInputs(directory), file("intake.json"), { name: "research-repairs", bytes: JSON.stringify(readRepairRound(directory)?.items.filter(item => item.ref.repair_target === "research") ?? []) }]
  if (step === "research-review") return [...outlineInputs(directory), file("intake.json"), file("research.json"), file("research-context.txt")]
  if (step === "finalize")
    return [file("repairs.json"),file("request.md"), file("plan.json"), ...planned.flatMap(id => [file(`${id}/paragraphs.json`), file(`${id}/paragraphs-applied.json`), file(`${id}/changes.json`), file(`${id}/apply.json`), file(`${id}/redline.docx`)]),
      { name: "open-items", bytes: JSON.stringify(disclosableItems(directory)) },
      { name: "questions", bytes: JSON.stringify(readQuestions(directory)) }]
  if (step === "gate")
    return [...planned.flatMap(id => [file(`${id}/paragraphs-applied.json`), file(`${id}/apply.json`), file(`${id}/change-review.json`)]),
      file("package-review.json"), file("checklist/refinements.json"), file("checklist/checks.json"), file("drafting-notes.md"), file("open-items.json"), file("questions.json")]
  const document = stepDocument(step) ?? FIRST_DOCUMENT
  const at = (name: string) => file(`${document}/${name}`)
  const finished = { name: "completed-documents", bytes: completedDocumentsText(directory, document) }
  const acting = { name: "acting-for", bytes: JSON.stringify(planActingFor(directory)) }
  const plan = readLeanPlan(directory)?.documents ?? []
  const entry = { name: "plan-entry", bytes: JSON.stringify(operativeDocument(plannedDocument(directory, document))) }
  const scope = { name: "planned-documents", bytes: plan.length > 1 ? JSON.stringify(plan.map(row => [row.id, row.title, row.purpose])) : "" }
  const stage = stepStage(step)
  const repairs = options.repairContext === false || !["change-plan", "change-prepare", "changes", "change-review"].includes(stage) ? undefined : documentRepairs(directory, document)
  if (stage === "draft") return [{ name: "outline-contract", bytes: OUTLINE_CONTRACT }, ...outlineInputs(directory), { name: "plan-entry", bytes: JSON.stringify(operativeDocument(plannedDocument(directory, document))) }, acting, ...(planned.indexOf(document) > 0 ? [finished] : [])]
  if (isReviewStage(stage)) return [{ name: "playbook-selection", bytes: JSON.stringify(plannedDocument(directory, document)?.playbook_review ?? null) }, file("request.md"), at("paragraphs.json"), sources, finished, acting, scope]
  if (stage === "change-prepare") return [file("request.md"), entry, at("paragraphs.json"), at("change-plan.json"), answers, sources, finished, acting, scope, { name: "repairs", bytes: JSON.stringify(repairs ?? null) }]
  if (stage === "change-review") return [...(repairs?.items.every(item => item.stage === "checklist-check") ? [{ name: "focused-review-contract", bytes: "prepared-dependency-component-review-2-proposed-body" }] : []), { name: "review-contract", bytes: "phase-aware-owned-body-review-2" }, file("request.md"), entry, at("paragraphs.json"), at("changes.json"), at("change-plan.json"), at("change-preparation.json"), ...reviewStages(directory, document).map(stage => at(`${stage}.md`)), answers, sources, finished, acting, scope, { name: "questions", bytes: JSON.stringify(readQuestions(directory)) }, { name: "repairs", bytes: JSON.stringify(repairs ?? null) }]
  if (stage === "changes" || stage === "change-plan") {
    return [...(stage === "changes" && repairs?.items.every(item => item.stage === "checklist-check") ? [{ name: "focused-reconciliation-contract", bytes: "prepared-operation-components-1" }] : []), ...(stage === "change-plan" && repairs?.items.every(item => item.stage === "checklist-check") ? [{ name: "focused-repair-contract", bytes: "one-assessment-source-adjudication-2" }] : []), file("request.md"), entry, ...(stage === "changes" ? [at("change-plan.json"), at("change-preparation.json")] : []), ...reviewStages(directory, document).map(stage => at(`${stage}.md`)), at("paragraphs.json"), { name: "repairs", bytes: repairs ? JSON.stringify(repairs) : "" }, answers, sources,
      { name: "draft-items", bytes: JSON.stringify(readOpenItems(directory).filter(item =>
        (item.stage === "draft" || isReviewStage(item.stage)) && (item.ref.document ?? FIRST_DOCUMENT) === document)) },
      finished, acting, scope]
  }
  return [at("body.docx"), at("changes.json"), at("change-review.json"), { name: "operative-plan", bytes: JSON.stringify(operativePlan(readLeanPlan(directory))) }, sources, { name: "comparison-original", bytes: comparisonBasis(directory, document) }]
}

/** Hash of what the outline is built from: the request and the matter's sources. */
export function outlineBasis(directory: string) {
  return hashStageInputs(`${FIRST_DOCUMENT}.draft`, LEAN_PROMPT_VERSION, [{ name: "outline-contract", bytes: OUTLINE_CONTRACT }, ...outlineInputs(directory), ...["research.json", "research-review.json"].map(name => ({ name, bytes: readLeanText(directory, name) ?? "" }))])
}

/** A saved plan is followed unless it records a different outline basis. A plan
 * that records none was written by hand and is followed as it is. */
export function planIsCurrent(directory: string) {
  const plan = readLeanPlan(directory)
  return Boolean(plan) && (plan?.basis === undefined || plan.basis === outlineBasis(directory))
}

function outlineInputs(directory: string) {
  const request = leanFile(directory, "request.md")
  const updates = readLeanText(directory, "updated-notes.md")
  return [{ name: "drafting-preferences", bytes: readFirmDraftingPreferences() }, { name: "source-content", bytes: matterSourceBasis(directory) }, { name: "answers", bytes: answeredQuestionsText(directory) }, ...(updates ? [{ name: "updated-notes.md", bytes: updates }] : []), { name: "request.md", bytes: existsSync(request) ? readFileSync(request) : "" }, { name: "sources", bytes: matterSourceIDs(directory).join("\n") }]
}

/** Checks a model reply and, when it passes, writes the step's output files. */
export async function acceptStageReply(directory: string, step: StepID, reply: string,
  options: { final?: boolean; presentation?: DraftPresentation } = {}): Promise<StageAcceptance> {
  const stage = stepStage(step)
  const document = stepDocument(step) ?? FIRST_DOCUMENT
  if (stage === "package-review") {
    let value: unknown
    try { value = JSON.parse(reply.trim().replace(/^```(?:json)?\s*/, "").replace(/\s*```$/, "")) } catch {
      return { issues: ["Reply with the package review JSON only."], outputs: [], openItems: [] }
    }
    const issues = checkPackageReview(directory, value)
    if (issues.length) return { issues, outputs: [], openItems: [] }
    writeLeanText(directory, "package-review.json", `${JSON.stringify(value, null, 2)}\n`)
    return { issues: [], outputs: stepOutputs(step), openItems: [] }
  }
  if (stage === "change-plan") return acceptChangePlan(directory, document, reply)
  if (stage === "change-review") return acceptChangeReview(directory, document, reply)
  if (stage === "draft") {
    const check = checkDraft(directory, reply)
    if (check.issues.length) return { ...check, outputs: [] }
    const rendered = await renderMarkdownToDocx(options.presentation?.source ?? reply)
      .then(result => options.presentation ? presentDraft(result.bytes, options.presentation.parts) : { bytes: result.bytes, attribution: undefined })
      .then(result => result, (error: unknown) => ({ error }))
    if ("error" in rendered)
      return { issues: [`The document could not be rendered to Word (${message(rendered.error)}). Use plain Markdown headings, paragraphs, lists and tables.`], openItems: [], outputs: [] }
    writeLeanText(directory, `${document}/draft.md`, ensureNewline(reply))
    writeLeanBytes(directory, `${document}/body.docx`, rendered.bytes)
    // Pointer bookmarks go into the body as soon as it is rendered, so every later
    // step and every snapshot addresses the same paragraphs.
    const rows = (await paragraphRows(leanFile(directory, `${document}/body.docx`))).map(row => ({ ...row, ...rendered.attribution?.get(row.pid) }))
    if (!rows.length) return { issues: ["The document has no body paragraphs."], openItems: [], outputs: [] }
    writeLeanText(directory, `${document}/paragraphs.json`, `${JSON.stringify(rows, null, 2)}\n`)
    return { ...check, outputs: stepOutputs(step) }
  }
  if (isReviewStage(stage)) {
    const check = checkReview(directory, reply, readParagraphs(directory, document))
    if (!check.issues.length) writeLeanText(directory, `${document}/${stage}.md`, ensureNewline(reply))
    return { ...check, openItems: check.openItems.map(item => ({ ...item, stage })), outputs: stepOutputs(step) }
  }
  if (stage === "changes") {
    const check = checkChanges(directory, reply, readParagraphs(directory, document), { withholdUnverified: options.final, document })
    const issues = [...check.issues, ...(check.changes && readLeanText(directory, `${document}/change-plan.json`) ? uniqueReplacementIssues(check.changes) : [])]
    if (issues.length || !check.changes) return { issues, openItems: [], outputs: [] }
    writeLeanText(directory, `${document}/changes.json`, `${JSON.stringify(check.changes, null, 2)}\n`)
    return { issues: [], openItems: check.openItems, outputs: stepOutputs(step), questions: check.changes.questions }
  }
  const check = checkDraftingNotes(directory, reply, disclosableItems(directory))
  if (check.issues.length) return { ...check, outputs: [] }
  const advisory = draftingStyleAdvisoryNotes([
    ...(readLeanPlan(directory)?.documents ?? []).map(row => ({ title: row.title, rows: finalBodyParagraphs(directory, row.id) })),
    { title: "Drafting notes", rows: [{ pid: "drafting-notes", text: reply }] },
  ])
  const notes = advisory ? `${reply.trimEnd()}\n\n${advisory}\n` : reply
  writeLeanText(directory, "drafting-notes.md", ensureNewline(notes))
  writeLeanBytes(directory, "drafting-notes.docx", (await renderMarkdownToDocx(notes)).bytes)
  return { ...check, outputs: stepOutputs(step) }
}

/** Gives every body paragraph a pointer (rewriting the docx when one is added)
 * and lists the paragraphs with the text hash amendments are authored against. */
export async function paragraphRows(docxPath: string): Promise<ParagraphRow[]> {
  return Object.entries(await ensureParagraphPointers(docxPath)).map(([pid, text]) => ({ pid, sha256: paragraphSha256(text), text }))
}

/** Pure re-check of a completed step's saved output, used before skipping it on
 * resume. A lawyer's edit that still passes is kept; one that does not re-runs. */
export function stageOutputValid(directory: string, step: StepID) {
  if (step === "checklist-check") return checklistOutputState(directory).valid
  if (step === "package-review") return Boolean(readPackageReview(directory))
  if (step === "discrepancy" || step === "spine") return evidenceOutputValid(directory, step)
  if (step === "intake" || step === "research" || step === "research-review") return researchOutputValid(directory, step)
  if (!stepOutputs(step).every(name => existsSync(leanFile(directory, name)))) return false
  const text = (name: string) => readLeanText(directory, name) ?? ""
  const stage = stepStage(step)
  const document = stepDocument(step) ?? FIRST_DOCUMENT
  if (stage === "change-plan") return Boolean(readChangePlan(directory, document))
  if (stage === "change-prepare") return preparationOutputValid(directory, document)
  if (stage === "change-review") return Boolean(readChangeReview(directory, document))
  if (stage === "draft")
    return Boolean(plannedDocument(directory, document)) && !checkDraft(directory, text(`${document}/draft.md`)).issues.length &&
      readParagraphs(directory, document).length > 0
  if (isReviewStage(stage)) return !checkReview(directory, text(`${document}/${stage}.md`), readParagraphs(directory, document)).issues.length
  if (stage === "changes") {
    const pids = new Set(readParagraphs(directory, document).map(row => row.pid))
    const list = readChangeList(directory, document)
    if (!list?.amendments.every(amendment => pids.has(amendment.pid))) return false
    if (!readLeanText(directory, `${document}/change-plan.json`)) return true
    return !uniqueReplacementIssues(list).length && !checkChanges(directory, text(`${document}/changes.json`), readParagraphs(directory, document), { document }).issues.length
  }
  if (stage === "finalize") return !checkDraftingNotes(directory, text("drafting-notes.md"), disclosableItems(directory)).issues.length
  return true
}

/** Open items the drafting notes must disclose. */
export function disclosableItems(directory: string): OpenItem[] {
  return [...readOpenItems(directory).filter(item => DISCLOSED_STAGES.includes(item.stage)), ...deferredFindingItems(directory)]
}

function ensureNewline(text: string) {
  return text.endsWith("\n") ? text : `${text}\n`
}

function message(error: unknown) {
  return error instanceof Error ? error.message : String(error)
}


export type ChecklistOpenFinding = { item_id: string; item: ChecklistItem; finding: ChecklistFinding }

/** Validate authoritative effective IDs and recompute coverage; archived originals
 * remain evidence, not the current delivery verdict. */
export function checklistOutputState(directory: string): { valid: boolean; open: boolean; openFindings: ChecklistOpenFinding[] } {
  try {
    const extraction = z.object({
      expected_sources: z.array(z.string()),
      sources: z.array(z.object({ key: z.string(), status: z.literal("succeeded"), items: z.array(z.object({ id: z.string(), item: checklistItemSchema })) })),
    }).parse(readLeanJson(directory, "checklist/extraction.json"))
    const checks = z.object({ extraction_complete: z.literal(true), expected_items: z.array(z.string()),
      checks: z.array(checklistCheckResultSchema),
    }).parse(readLeanJson(directory, "checklist/checks.json"))
    const sameIDs = (expected: string[], actual: string[]) => new Set(expected).size === expected.length &&
      new Set(actual).size === actual.length && expected.length === actual.length && expected.every(id => actual.includes(id))
    if (!sameIDs(extraction.expected_sources, extraction.sources.map(row => row.key))) return { valid: false, open: false, openFindings: [] }
    const originals = extraction.sources.flatMap(source => source.items)
    if (new Set(originals.map(row => row.id)).size !== originals.length) return { valid: false, open: false, openFindings: [] }
    const effective = readEffectiveChecklistItems(directory, originals)
    const expected = effective.items.map(row => row.id)
    const selected = checklistRefinementSelections(directory)
    if (selected.some(row => !effective.parents.some(parent => parent.parent_id === row.parent_id && parent.scope === row.scope))) return { valid: false, open: false, openFindings: [] }
    if (!effective.complete || !sameIDs(expected, checks.expected_items) || new Set(checks.checks.map(row => row.item_id)).size !== checks.checks.length ||
      checks.checks.some(row => row.finding && row.finding.item_id !== row.item_id)) return { valid: false, open: false, openFindings: [] }
    // Inactive original/refinement findings retain their consumed review allowance.
    // Only authoritative current IDs participate in delivery validity.
    const active = expected.map(id => checks.checks.find(row => row.item_id === id))
    if (active.some((row, index) => !row || row.status !== "succeeded" || row.reinspection?.status === "failed" || !row.finding || checklistFindingRecovery(effective.items[index]!.item, row.finding) !== "none")) return { valid: false, open: false, openFindings: [] }
    const coverage = reconcileChecklistParentCoverage(directory, checks.checks, originals)
    // Every current finding left open (or with open decisions), as recorded; a
    // parent's coverage is open only through such a child finding.
    const openFindings = active.flatMap((row, index) => row?.finding && (row.finding.status === "open" || row.finding.open_decisions?.length)
      ? [{ item_id: row.item_id, item: effective.items[index]!.item, finding: row.finding }] : [])
    return { valid: coverage.every(row => row.status === "satisfied" || row.status === "open") && existsSync(leanFile(directory, "checklist/report.md")),
      open: openFindings.length > 0 || coverage.some(row => row.status === "open"), openFindings }
  } catch { return { valid: false, open: false, openFindings: [] } }
}
