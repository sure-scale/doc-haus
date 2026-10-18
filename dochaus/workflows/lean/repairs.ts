import { existsSync, readFileSync } from "node:fs"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { checklistFindingSchema, checklistFindingRecovery, isDeliveryArtifactEvidence, type ChecklistItem } from "./checklist-contract"
import { readChecklistAdjudicationCounterevidence, readEffectiveChecklistItems } from "./checklist-flow"
import { checklistDeliveryEvidence } from "./checklist-document-context"
import { normalize } from "../../lib/quote-verify"
import { readChangeList, readParagraphs, type ApplyResult, type ParagraphRow } from "./checks"
import { FIRST_DOCUMENT, readLeanPlan, type DocID } from "./plan"
import { excerpt, openItem, withSourceSnapshot } from "./quotes"
import { hashStageInputs, isReviewStage, leanFile, readLeanJson, readLeanState, readOpenItems, readQuestions, writeLeanText, type LeanState, type OpenItem } from "./state"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"

// Repair inputs for the changes stage. The changes stage owns every fixable
// defect in the body: unsupported content found in the draft, and whatever the
// final gate still blocks on when the lawyer presses Continue. A Continue on a
// blocked gate starts one repair round (repairs.json), which is an input of the
// changes stage, so each Continue re-runs changes once and then apply, finalize
// and the gate. One round per human action keeps the loop bounded and never
// leaves Continue with nothing to do. A round's items carry their document, so
// each document's changes step sees and is keyed by its own repairs only.

export const LEAN_MAX_SUBSTANTIVE_REPAIRS = 5

type DocumentRepairContext = { round: number; items: OpenItem[]; basis: string }
export type RepairRound = { round: number; requested_at: string; items: OpenItem[]; substantive_attempts?: number; checklist_items?: string[]; document_contexts?: Partial<Record<DocID, DocumentRepairContext>> }

const REPAIRABLE_DRAFT_KINDS = new Set<OpenItem["kind"]>(["quote_miss", "citation_unsupported"])

export function readRepairRound(directory: string): RepairRound | undefined {
  const value = readLeanJson(directory, "repairs.json")
  if (typeof value !== "object" || value === null || !Array.isArray((value as RepairRound).items)) return
  const round = value as RepairRound
  // Older saved reviews called a separate memo deliverable "memo" even when
  // they supplied its document owner. Finish that admitted repair in its editor;
  // keep run-wide notes and the recorded repair allowance unchanged.
  return { ...round, items: round.items.map(item => item.ref.repair_target === "memo" && item.ref.document
    ? { ...item, ref: { ...item.ref, repair_target: "body" } }
    : item) }
}

export function startRepairRound(directory: string, blocking: OpenItem[], substantive = false): RepairRound {
  const previous = readRepairRound(directory)
  const addressed = new Set(blocking.filter(item => item.ref.document || !["memo", "research"].includes(item.ref.repair_target ?? "body"))
    .map(item => item.ref.document ?? FIRST_DOCUMENT))
  const contexts = { ...previous?.document_contexts }
  for (const document of Object.keys(contexts) as DocID[]) {
    if (addressed.has(document) || contexts[document]?.basis !== documentRepairBasis(directory, document)) delete contexts[document]
  }
  // Retiring an unrelated document's completed repair must not itself re-key
  // that document. Retain only a genuinely accepted, still-current projection.
  for (const document of new Set(previous?.items.map(item => item.ref.document ?? FIRST_DOCUMENT))) {
    if (addressed.has(document)) continue
    const items = bodyRepairItems(previous!.items, document)
    if (!items.length) continue
    const step = `${document}.change-plan` as const
    const accepted = readLeanState(directory)?.stages[step]
    if (accepted?.status !== "done" || accepted.input_sha !== hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step))) continue
    contexts[document] = { round: previous!.round, items, basis: documentRepairBasis(directory, document) }
  }
  // Like the allowance, the checklist items submitted so far outlive the round
  // that submitted them, so a later round cannot re-arm their one repair.
  const checklist = new Set([...repairedChecklistItems(previous), ...checklistItemIDs(blocking)])
  const round: RepairRound = { round: (previous?.round ?? 0) + 1, requested_at: new Date().toISOString(), items: blocking, substantive_attempts: (previous?.substantive_attempts ?? 0) + (substantive ? 1 : 0),
    ...(checklist.size ? { checklist_items: [...checklist] } : {}),
    ...(Object.keys(contexts).length ? { document_contexts: contexts } : {}) }
  writeLeanText(directory, "repairs.json", `${JSON.stringify(round, null, 2)}\n`)
  return round
}

/** Checklist item IDs any repair round of this run submitted. */
export function repairedChecklistItems(round: RepairRound | undefined) {
  const items = [...round?.items ?? [], ...Object.values(round?.document_contexts ?? {}).flatMap(context => context?.items ?? [])]
  return new Set([...round?.checklist_items ?? [], ...checklistItemIDs(items)])
}

function checklistItemIDs(items: OpenItem[]) {
  return items.flatMap(item => {
    if (item.stage !== "checklist-check") return []
    try {
      const value = JSON.parse(item.detail.slice(item.detail.indexOf("\n") + 1)) as { item_id?: unknown } | null
      return typeof value?.item_id === "string" ? [value.item_id] : []
    } catch { return [] }
  })
}

/** Last admitted repair input for this document, while its non-repair inputs
 * remain current. Retained context is not a new blocking finding or admission. */
export function documentRepairs(directory: string, document: DocID) {
  const round = readRepairRound(directory)
  const items = bodyRepairItems(round?.items ?? [], document)
  if (round && items.length) return { round: round.round, items }
  const retained = round?.document_contexts?.[document]
  if (!retained) return
  if (retained.basis === documentRepairBasis(directory, document)) return { round: retained.round, items: retained.items }
}

function bodyRepairItems(items: OpenItem[], document: DocID) {
  return items.filter(item => ofDocument(item, document) && item.ref.repair_target !== "research" && item.ref.repair_target !== "memo" && (item.stage !== "package-review" || item.ref.document !== undefined))
}

function documentRepairBasis(directory: string, document: DocID) {
  const step = `${document}.change-plan` as const
  return hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step, { repairContext: false })
    .filter(part => part.name !== "repairs"))
}

/** Prompt sections for one document's changes step beyond its findings: the
 * defects it must repair, the change list it replaces, and the questions
 * already asked. */
export function changesRepairSections(directory: string, document: DocID) {
  const paragraphs = readParagraphs(directory, document)
  const items = readOpenItems(directory).filter(item => ofDocument(item, document))
  const required = [
    ...items.filter(item => item.stage === "draft" && REPAIRABLE_DRAFT_KINDS.has(item.kind)),
    ...(documentRepairs(directory, document)?.items ?? []),
  ]
  const checklistFindings = required.filter(item => item.stage === "checklist-check")
  const defects = required.filter(item => item.stage !== "checklist-check")
  const unverifiedFindings = items.filter(item => isReviewStage(item.stage) && item.kind === "quote_miss")
  return [
    defects.length ? tagged("required_repairs", [
      "Each item is a defect the draft still has. Fix each one with an amendment to the paragraph named:",
      "quote the source exactly, remove the unsupported point, or replace it with a visible [open item: what is needed] marker.",
      "For an unanswered question, repeat it with its \"id\" and the \"pid\" of the paragraph it affects, and put a visible [open item: ...] marker in that paragraph.",
      ...defects.map(item => `- ${item.id} ${pidLabel(item, paragraphs)}: ${item.detail}${quoteLabel(item)}`),
    ].join("\n")) : "",
    checklistFindings.length ? tagged("checklist_findings", [
      "These are unverified adverse assessments for source adjudication, not established defects. Reassess source support, applicability, qualifications and actual current text before proposing a correction. Preserve proposals as proposals; do not force an unsupported or uncertain duty into the document. Carry forward every still-supported existing amendment.",
      ...checklistFindings.map(item => `- ${item.id} ${pidLabel(item, paragraphs)}: ${item.detail}${quoteLabel(item)}`),
    ].join("\n")) : "",
    unverifiedFindings.length ? tagged("unverified_findings", [
      "These quotations in the review were not found in the named source. Check the source before relying on the finding:",
      ...unverifiedFindings.map(item => `- "${excerpt(item.ref.quote ?? "")}" [source: ${item.ref.source_id ?? "unknown"}]`),
    ].join("\n")) : "",
    previousChanges(directory, document, paragraphs),
    openQuestions(directory, document),
  ].filter(Boolean)
}

/** The last change list, limited to amendments still addressed to the current
 * draft text. Not a stage input: the stage's own output must not change its key. */
function previousChanges(directory: string, document: DocID, paragraphs: ParagraphRow[]) {
  const current = new Map(paragraphs.map(row => [row.pid, row.sha256]))
  const list = readChangeList(directory, document)
  const applied = (readLeanJson(directory, `${document}/apply.json`) as ApplyResult | undefined)?.applied ?? []
  const rows = [
    ...(list?.amendments ?? []).map(amendment => ({ amendment, note: "" })),
    ...(list?.withheld ?? []).map(row => ({ amendment: row.amendment, note: ` (withheld, not applied: ${row.issues.join(" ")})` })),
  ].filter(row => current.get(row.amendment.pid) === row.amendment.existing_sha256)
  const explanations = (list?.explanations ?? []).filter(row => current.has(row.pid))
  if (!rows.length && !explanations.length) return ""
  return tagged("previous_changes", [
    "Your reply replaces this change list. Carry forward, with their sources, the amendments that are still right; correct or drop the rest.",
    ...explanations.map(row => `- Initial revision explanation: ${JSON.stringify(row)}`),
    ...rows.map(row => `- ${row.amendment.id}${row.note}: ${JSON.stringify({
      pid: row.amendment.pid, ...(row.amendment.operation ? { operation: row.amendment.operation } : {}),
      ...(applied.find(outcome => outcome.id === row.amendment.id)?.result_pids ? { final_pids: applied.find(outcome => outcome.id === row.amendment.id)!.result_pids } : {}),
      updated: row.amendment.updated, reason: row.amendment.reason, sources: row.amendment.sources,
    })}`),
  ].join("\n"))
}

/** Every open question of the run, so a later document can repeat a question
 * an earlier one asked rather than ask it again in other words. Paragraph
 * pointers are shown only for this document's own paragraphs. The final
 * gate's question is not a drafting question and is left out. */
function openQuestions(directory: string, document: DocID) {
  const open = readQuestions(directory).filter(question => question.status === "open" && question.stage !== "gate")
  if (!open.length) return ""
  return tagged("open_questions", [
    "Questions already put to the lawyer and not yet answered. To keep one, repeat it with its \"id\"; do not ask it again in other words.",
    ...open.map(question => `- ${question.id}${question.pid && (question.document ?? FIRST_DOCUMENT) === document ? ` [${question.pid}]` : ""}: ${question.question}`),
  ].join("\n"))
}

function ofDocument(item: OpenItem, document: DocID) {
  return (item.ref.document ?? FIRST_DOCUMENT) === document
}

function pidLabel(item: OpenItem, paragraphs: ParagraphRow[]) {
  if (item.ref.pid) return paragraphs.some(row => row.pid === item.ref.pid) ? `[${item.ref.pid}]` : "(earlier draft paragraph; reassess the finding in the current draft)"
  const needle = normalize(item.ref.quote ?? item.ref.citation ?? "")
  const pid = needle ? paragraphs.find(row => normalize(row.text).includes(needle))?.pid : undefined
  return pid ? `[${pid}]` : "(paragraph not located)"
}

function quoteLabel(item: OpenItem) {
  if (item.kind === "question" && item.ref.question_id) return ` (question id ${item.ref.question_id})`
  return item.kind === "quote_miss" && item.ref.quote ? ` Quoted: "${excerpt(item.ref.quote)}"` : ""
}

function tagged(name: string, body: string) {
  return `<${name}>\n${body.trim()}\n</${name}>`
}


/** Explicit Continue may submit an adverse assessment to the existing native
 * amendment adjudicator. It is evidence to reassess, never a predetermined fix.
 * `skip` withholds an item the caller will not resubmit. */
export function checklistRepairItems(directory: string, skip: (item_id: string, item: ChecklistItem) => boolean = () => false): OpenItem[] {
  const saved = readLeanJson(directory, "checklist/checks.json") as { extraction_complete?: unknown; expected_items?: unknown; checks?: unknown } | undefined
  if (!saved || saved.extraction_complete !== true || !Array.isArray(saved.expected_items) || !Array.isArray(saved.checks)) return []
  const expected = saved.expected_items, checks = saved.checks
  const effective = readEffectiveChecklistItems(directory)
  if (!effective.complete) return []
  const plan = readLeanPlan(directory)
  if (!plan) return []
  // One snapshot builds the delivered files' properties once for every item.
  return withSourceSnapshot(() => effective.items.flatMap(record => {
    if (!expected.includes(record.id)) return []
    const rows = checks.filter((value): value is { item_id: string; status: string; finding: unknown } => typeof value === "object" && value !== null &&
      (value as { item_id?: unknown }).item_id === record.id && (value as { status?: unknown }).status === "succeeded")
    if (rows.length !== 1) return []
    const parsed = checklistFindingSchema.safeParse(rows[0]!.finding)
    if (!parsed.success || parsed.data.item_id !== record.id || checklistFindingRecovery(record.item, parsed.data) !== "repair") return []
    if (!["operative", "advice"].includes(record.item.target.treatment)) return []
    const matches = plan.documents.filter(document => document.id === record.item.target.document || document.file === record.item.target.document)
    const notes = !matches.length && ["notes", plan.notes_file ?? "drafting-notes.docx"].includes(record.item.target.document)
    // Code sets a delivered file's name, Word format and composition, so no
    // text repair can change them. A finding resting only on the delivery
    // rows supplied for its target stays disclosed with its class; one that
    // also cites text keeps its repair, which receives only that text.
    const target = matches.length === 1 ? matches[0]!.id : notes ? "notes" : undefined
    const supplied = target === undefined ? [] : checklistDeliveryEvidence(directory, target).evidence
    const evidence = parsed.data.evidence.filter(row => !isDeliveryArtifactEvidence(row, supplied))
    if (parsed.data.evidence.length && !evidence.length) return []
    const finding = { ...parsed.data, evidence }
    // The run-wide drafting notes have no editing stages: their omission is a
    // required repair for the notes writer, like the package review's notes findings.
    if (notes) {
      if (!existsSync(leanFile(directory, "drafting-notes.md")) || skip(record.id, record.item)) return []
      return [openItem("checklist-check", "omission", [
        "Reassess this focused checklist assessment of the drafting notes against actual sources, qualifications and the current notes; its adverse verdict is not authority or an instruction to force a change. Address a supported omission in the notes and record what remains unsupported as uncertain. Do not invent approval, facts or advice from the verdict.",
        JSON.stringify({ item_id: record.id, item: record.item, finding, current_artifact: ".dochaus/lean/drafting-notes.md" }),
      ].join("\n"), { repair_target: "memo" })]
    }
    if (matches.length !== 1) return []
    const document = matches[0]!
    const file = [`${document.id}/body-applied.docx`, `${document.id}/body.docx`].find(file => existsSync(leanFile(directory, file)))
    if (!file || skip(record.id, record.item)) return []
    let current: ReturnType<typeof readDocxCommentParagraphs>
    try { current = readDocxCommentParagraphs(readFileSync(leanFile(directory, file))) } catch { return [] }
    const originals = readParagraphs(directory, document.id)
    const applied = (readLeanJson(directory, `${document.id}/apply.json`) as ApplyResult | undefined)?.applied ?? []
    let pid: string | undefined
    const inspected: typeof current = []
    for (const evidence of finding.evidence.filter(evidence => evidence.kind === "body" && [document.id, document.file].includes(evidence.document))) {
      const needle = normalize(evidence.text)
      const paragraphs = needle ? current.filter(paragraph => normalize(paragraph.paragraph_text).includes(needle)) : []
      if (paragraphs.length !== 1) continue
      const paragraph = paragraphs[0]!
      if (!inspected.includes(paragraph)) inspected.push(paragraph)
      const initial = new Set(paragraph.pids.flatMap(currentPID => [
        ...originals.filter(row => row.pid === currentPID).map(row => row.pid),
        ...applied.filter(row => row.result_pids?.includes(currentPID) && originals.some(original => original.pid === row.pid)).map(row => row.pid),
      ]))
      if (!pid && initial.size === 1) pid = [...initial][0]
    }
    return [openItem("checklist-check", "omission", [
      "Reassess this focused checklist assessment against actual sources, qualifications and current text; its adverse verdict is not authority or an instruction to force a change. Preserve proposals as proposals and unsupported requirements as uncertain. Do not invent approval or replacement wording from the verdict. Retain every still-supported existing amendment; correct or drop an earlier amendment if source adjudication shows it caused this defect.",
      JSON.stringify({ item_id: record.id, item: record.item, finding, current_artifact: `.dochaus/lean/${file}`, current_target_paragraphs: inspected }),
    ].join("\n"), { document: document.id, repair_target: "body", ...(pid ? { pid } : {}) })]
   }))
}

/** The document's current accepted focused adjudication of this exact item
 * declined a repair: no correction group, and a source-cited disposition. Its
 * claims are the counterevidence the checklist stage's reinspection weighs. */
export function declinedChecklistRepair(directory: string, state: LeanState, item_id: string, item: ChecklistItem) {
  const documents = readLeanPlan(directory)?.documents.filter(row => row.id === item.target.document || row.file === item.target.document)
  if (documents?.length !== 1) return
  const document = documents[0]!.id, file = `${document}/repair-adjudications.json`
  const step = state.stages[`${document}.change-plan`]
  if (step?.status !== "done" || !step.outputs.includes(file)) return
  const saved = readLeanJson(directory, file) as { assessments?: unknown } | undefined
  const entries = Array.isArray(saved?.assessments) ? saved.assessments.filter(row => (row as { item_id?: unknown } | null)?.item_id === item_id) as { value?: { groups?: unknown } }[] : []
  if (entries.length !== 1 || !Array.isArray(entries[0]!.value?.groups) || entries[0]!.value.groups.length) return
  const counterevidence = readChecklistAdjudicationCounterevidence(directory, item_id, item)
  return counterevidence?.claims.some(claim => claim.source_refs.length) ? counterevidence : undefined
}
