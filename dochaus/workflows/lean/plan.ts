import { isDateFormat, isPosture, normalizeLegend, type DateFormat, type Posture } from "../../lib/drafting-preferences"
import { OPEN_ITEM_PATTERN } from "../../lib/marker-patterns"
import { effectivePosture, matterDateFormat } from "../../lib/matter-client"
import { DOCUMENT_FORMS, type DocumentForm } from "../../lib/page-furniture"
import { TOOL_SYNTAX_ISSUE, printedToolCall } from "./checks"
import { attachmentHead } from "./consistency"
import { REVIEW_STAGES, readLeanJson, readLeanText, writeLeanText, type Question, type ReviewStageID, type StageID } from "./state"
import { templateSource } from "./template-contract"

// The run's plan: every document the request asks for, in drafting order, each
// with its top-level parts. It comes from the draft agent's outline reply (one
// "# " title per document, "## " parts under it); code assigns the document ids
// d1..dK in that order and saves it as plan.json. The walk of steps follows it:
// each document goes through draft, review, changes and apply, then one
// finalize and one gate cover the run. The plan is a drafting record, not an
// authority: it grants nothing, and nothing here knows what a document is for.
// It records the hash of the request and sources the outline was built from, so
// the outline is redone only when they change and a lawyer's edit is followed.
// It also records whom the lawyer acts for: the party the outline reads from the
// request, else the client recorded on the matter. Every model stage is told,
// and a stage that is not told takes no side. Code never infers the party from
// the sources; a lawyer can still edit plan.json. Beside it the plan freezes the
// Posture the run drafts to (runPosture), so a later settings or matter change
// cannot split the run's stages.

export const DOC_STAGES = ["draft", ...REVIEW_STAGES, "review-playbook", "change-plan", "change-prepare", "changes", "change-review", "apply"] as const
export type DocStage = (typeof DOC_STAGES)[number]
export type DocID = `d${number}`
export type StepID = `${DocID}.${DocStage}` | "instructions" | "discrepancy" | "intake" | "research" | "research-review" | "acts" | "spine" | "finalize" | "checklist-check" | "package-review" | "gate"

/** heading identifies the drafting part; absent display_heading preserves it
 * in publication (a plan saved before labels were planned unheaded, or one
 * whose outline shows its headings), while an empty display_heading makes the
 * part unheaded. `instrument` marks the first part of a separate instrument
 * the document's file also holds (foldRequestedFiles): it and the parts after
 * it, up to the next such part, are that instrument's, printed under its own
 * title on a new page. `attachment` marks a part headed as a schedule, annex,
 * appendix, exhibit or attachment ("Exhibit A", consistency.ts attachmentHead):
 * it trails the document, its heading prints and it starts on a new page. */
export type PlannedPart = { heading: string; scope: string; display_heading?: string; instrument?: PlannedInstrument; attachment?: true }
/** A separate instrument held in another document's file: its own title, what
 * it is for, the parties who execute it and its form, as its outline entry
 * gave them. */
export type PlannedInstrument = { title: string; purpose?: string; executed_by?: string[]; form?: DocumentForm }
/** A template part the outline leaves out on purpose, with its reason. */
export type LeftOutPart = { part: string; reason: string }
/** `template` names the matter document the document is drafted from, when the
 * outline gives one on a "Template:" line under the title; `left_out` lists the
 * template parts its "Left out:" lines leave out on purpose; `file` is the file
 * name the request asks the document to be delivered under, from a "File:" line.
 * `mode` distinguishes a revision foundation from a new document; `deliverable`
 * assigns the requested filename to the body or annotated redline, or, as
 * `combined`, to one file holding the requested memo and then the body.
 * `executed_by` names the parties who execute the document, from an
 * "Executed by:" line; an empty list says it takes no legal act, and an absent
 * one that the outline did not say; a document holding separate instruments is
 * executed by the parties of each of them. `form` is the kind of document it is, from
 * a "Form:" line; it sets the document's page furniture, and a letter's parts
 * are planned unheaded unless the outline shows them. `furniture` is what the
 * request asks of its page furniture beyond the firm's settings, from "Legend:"
 * and "Page numbers: none" lines; it is rendered, never drafted.
 * Legacy plans without these fields retain their existing behaviour. The title
 * is the document's own title and never carries the file name. */
export type PlaybookReview = { sources: { source_id: string; quote: string }[]; reason: string }
export type PlannedDocument = {
  id: DocID; title: string; purpose: string; parts: PlannedPart[]; playbook_review?: PlaybookReview; mode?: "new" | "revise"; deliverable?: "body" | "redline" | "combined"; template?: string; comparison?: string; left_out?: LeftOutPart[]; file?: string; executed_by?: string[]; form?: DocumentForm
  furniture?: { legend?: string; page_numbers?: false }
}
/** Whom the lawyer acts for and why: the party the lawyer's words name
 * ("stated"), the client recorded on the matter ("matter"), the party intake
 * took from what the requested document does and whom the sources show
 * instructing the work ("inferred", with its source evidence as `reason`), or
 * no side ("none"). */
export type ActingForBasis = "stated" | "matter" | "inferred" | "none"
export type ActingFor = { party: string; basis: ActingForBasis; reason?: string }
/** The dates the work turns on, as intake found them: the date the documents
 * bear, from the lawyer's words ("request"), a date the sources fix for them
 * ("sources") or the drafting day ("today", its date frozen with the plan), and
 * the operative dates the sources give, each with the words that show it. */
export type DocumentDateBasis = "request" | "sources" | "today"
export type MatterDate = { label: string; date: string; source_id: string; quote: string }
export type MatterDates = { document_date: { date: string; basis: DocumentDateBasis; source_id?: string; quote?: string }; operative_dates: MatterDate[] }
/** `memo` is the request's own words for the memo it asks for, a naming input
 * only; `memo_document` is the planned document that is that memo.
 * `record_file` is the lawyer's file name for the drafting record, set only by
 * a delivery selection's notes row (selectedDeliveryPlan); it never asks for a
 * memo. `date_format` is the date style the matter's sources used when the
 * outline was built (runDateFormat), absent when they settled on none; the
 * run writes it until the firm sets its own. */
export type LeanPlan = { schema: "dochaus.lean.plan/2"; documents: PlannedDocument[]; acting_for?: ActingFor; posture?: Posture; basis?: string; memo?: string; memo_document?: DocID; record_file?: string; dates?: MatterDates; date_format?: DateFormat }
/** `memo` is the request's own words for a memo that accompanies the documents,
 * from a "Memo:" line; with it, the last document is that memo (parsePlan). */
export type LeanOutline = { documents: Omit<PlannedDocument, "id">[]; acting_for: ActingFor; memo?: string }

export const LEAN_PLAN_SCHEMA = "dochaus.lean.plan/2"
export const FIRST_DOCUMENT: DocID = "d1"
/** Once saved, the plan fixes which documents a draft prepares: a question,
 * option or finding that adds, drops or swaps one is one nothing can carry
 * out, and asking it again cannot settle it. The lawyer changes the plan only
 * by stopping the draft (coverageScopeQuestion's context says how). */
export const FIXED_PLAN = "Which documents this draft prepares is fixed by its saved document plan: no question, option or finding may add, drop or replace a planned document."
export const NOT_STATED: ActingFor = { party: "", basis: "none" }
const ACTING_FOR_BASES: ActingForBasis[] = ["stated", "matter", "inferred", "none"]

type Checked<T> = { issues: string[]; value?: T }

/** The steps of the run in walk order: before a plan exists, only the first
 * document's steps are known. */
export function leanSteps(directory: string): StepID[] {
  const ids = readLeanPlan(directory)?.documents.map(document => document.id) ?? [FIRST_DOCUMENT]
  return ["instructions", "discrepancy", "intake", "research", "research-review", "acts", "spine", ...ids.flatMap(id => DOC_STAGES.filter(stage => stage !== "review-playbook" || reviewStages(directory, id).includes(stage)).map((stage): StepID => `${id}.${stage}`)), "finalize", "checklist-check", "package-review", "gate"]
}

/** Which answered questions `step` reads of those a document's step asked (one
 * with `document`): its own and those of the steps before it in walk order,
 * never a later step's, which that step applies and an earlier step's output
 * cannot depend on. A question with no document (a run-level one, the gate's
 * package-scope one) or whose asking step is not in the walk is read by every
 * step. The step's prompt and its key read the same, so neither changes for
 * an answer the step never reads. */
export function askedAtOrBefore(directory: string, step: StepID) {
  const order = leanSteps(directory), at = order.indexOf(step)
  return (question: Question) => {
    if (!question.document || at < 0) return true
    const asker = order.indexOf(`${question.document}.${question.stage}` as StepID)
    return asker < 0 || asker <= at
  }
}

/** Planning selects guidance per body; legacy and unassigned bodies retain the ordinary reviews. */
export function reviewStages(directory: string, document: DocID): ReviewStageID[] {
  return [...REVIEW_STAGES, ...(reviewGuidance(directory, document).length ? ["review-playbook" as const] : [])]
}

/** The template a new document is drafted from, which its guidance review
 * checks it against. A revision's original is its baseline, not guidance: the
 * instructions may remove its provisions. */
export function guidanceTemplate(entry: Pick<PlannedDocument, "mode" | "template" | "comparison"> | undefined) {
  return entry?.mode === "revise" ? undefined : templateSource(entry)
}

/** The sources the guidance review checks a document against: the playbook
 * sources planning selected for it and the template a new document is drafted from. */
export function reviewGuidance(directory: string, document: DocID): string[] {
  const planned = plannedDocument(directory, document)
  const template = guidanceTemplate(planned)
  return [...new Set([...(planned?.playbook_review?.sources.map(source => source.source_id) ?? []), ...(template ? [template] : [])])]
}

export function stepStage(step: StepID): StageID {
  return !step.includes(".") ? step as StageID : step.slice(step.indexOf(".") + 1) as DocStage
}

export function stepDocument(step: StepID): DocID | null {
  return !step.includes(".") ? null : step.slice(0, step.indexOf(".")) as DocID
}

export function isStepID(value: string): value is StepID {
  return ["instructions", "discrepancy", "intake", "research", "research-review", "acts", "spine", "finalize", "checklist-check", "package-review", "gate"].includes(value) || new RegExp(`^d[1-9]\\d*\\.(${DOC_STAGES.join("|")})$`).test(value)
}

export function readLeanPlan(directory: string): LeanPlan | undefined {
  const value = readLeanJson(directory, "plan.json")
  return planIssue(value) ? undefined : value as LeanPlan
}

/** Why the saved plan.json is not a plan: its parse error or the first rule it
 * breaks. Undefined when there is no plan.json or it is a valid plan. */
export function leanPlanIssue(directory: string): string | undefined {
  const text = readLeanText(directory, "plan.json")
  if (text === undefined) return
  try {
    JSON.parse(text)
  } catch (error) {
    return error instanceof Error ? error.message : String(error)
  }
  return planIssue(readLeanJson(directory, "plan.json"))
}

function planIssue(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null || (value as LeanPlan).schema !== LEAN_PLAN_SCHEMA) return `its schema is not "${LEAN_PLAN_SCHEMA}"`
  const documents = (value as LeanPlan).documents
  if (!Array.isArray(documents) || !documents.length) return "it lists no documents"
  const invalid = documents.findIndex(document => !isPlannedDocument(document))
  if (invalid >= 0) return `document ${invalid + 1} is not a valid planned document`
  const memo = (value as LeanPlan).memo
  if (memo !== undefined && (typeof memo !== "string" || !memo.trim())) return "its memo is not a name"
  const memoDocument = (value as LeanPlan).memo_document
  if (memoDocument !== undefined && !documents.some(document => document.id === memoDocument)) return "its memo document is not a planned document"
  const recordFile = (value as LeanPlan).record_file
  if (recordFile !== undefined && (typeof recordFile !== "string" || !recordFile.trim())) return "its record file is not a name"
}

/** A requested memo is saved as its name and the planned document that is the
 * memo; a plan without one carries neither key, nor a plan without a Posture
 * the Posture key, nor one whose sources settled on no date style the
 * date_format key. */
export function writeLeanPlan(directory: string, documents: Omit<PlannedDocument, "id">[], basis?: string, actingFor?: ActingFor, memo?: { name: string; document?: DocID }, posture?: Posture, dates?: MatterDates, dateFormat?: DateFormat): LeanPlan {
  const plan: LeanPlan = {
    schema: LEAN_PLAN_SCHEMA, documents: withDocumentIDs(documents), acting_for: actingFor, posture, basis,
    ...(memo ? { memo: memo.name, ...(memo.document ? { memo_document: memo.document } : {}) } : {}), dates, date_format: dateFormat,
  }
  writeLeanText(directory, "plan.json", `${JSON.stringify(plan, null, 2)}\n`)
  return plan
}

/** Outline documents with the ids the plan gives them: d1..dK in drafting order. */
export function withDocumentIDs(documents: Omit<PlannedDocument, "id">[]): PlannedDocument[] {
  return documents.map((document, index) => ({ id: `d${index + 1}`, ...document }))
}

export function plannedDocument(directory: string, id: DocID): PlannedDocument | undefined {
  return readLeanPlan(directory)?.documents.find(document => document.id === id)
}

/** Output naming, the requested memo's name included, does not change operative
 * drafting inputs, nor does page furniture, which is rendered and never drafted.
 * Consumers of the full delivery plan (notes, package review and registration)
 * retain them; the draft step's key keeps the furniture it renders. */
export function operativeDocument(document: PlannedDocument | undefined) {
  if (!document) return null
  const { file, deliverable, furniture, ...operative } = document
  return operative
}

export function operativePlan(plan: LeanPlan | undefined) {
  if (!plan) return null
  const { memo, record_file, ...operative } = plan
  return { ...operative, documents: plan.documents.map(operativeDocument) }
}

/** Whom the lawyer acts for. A plan without it, or with no party named, takes
 * no side. */
export function planActingFor(directory: string): ActingFor {
  return actingForValue(readLeanPlan(directory)?.acting_for)
}

/** A saved acting-for record, or no side when it names no party or basis. */
export function actingForValue(value: unknown): ActingFor {
  if (typeof value !== "object" || value === null) return NOT_STATED
  const row = value as Partial<ActingFor>
  if (typeof row.party !== "string" || !row.party.trim() || !ACTING_FOR_BASES.includes(row.basis as ActingForBasis) || row.basis === "none") return NOT_STATED
  return { party: row.party.trim(), basis: row.basis!, ...(typeof row.reason === "string" && row.reason.trim() ? { reason: row.reason.trim() } : {}) }
}

/** How far the run's drafts lean toward the party the lawyer acts for: the
 * Posture the outline froze into the plan. Before the outline, and for a plan
 * saved without one, the matter's Posture, else the firm's, read live as
 * before. Every lean stage renders its drafting preferences with it. */
export function runPosture(directory: string): Posture {
  const posture: unknown = readLeanPlan(directory)?.posture
  return isPosture(posture) ? posture : effectivePosture(directory)
}

/** The date style the matter's sources use, which the run writes until the
 * firm sets its own: the style the outline froze into the plan, none when
 * the plan records none (the sources settled on none, or a plan saved before
 * styles were frozen). Before the outline, the style the sources and the
 * request use, read live. Every lean stage renders its drafting preferences
 * with it; the firm's own style, once set, is rendered instead. */
export function runDateFormat(directory: string): DateFormat | undefined {
  const plan = readLeanPlan(directory)
  if (plan) return isDateFormat(plan.date_format) ? plan.date_format : undefined
  return matterDateFormat(directory, readLeanText(directory, "request.md") ?? "")
}

/** The documents planned before `id`, in drafting order. */
export function earlierDocuments(directory: string, id: DocID): PlannedDocument[] {
  const documents = readLeanPlan(directory)?.documents ?? []
  const at = documents.findIndex(document => document.id === id)
  return at < 0 ? [] : documents.slice(0, at)
}

/** Reads the outline reply: an "Acting for:" line before the first title, then
 * each "# " line starts a document, the plain lines under it before its first
 * part are its purpose (a "Template:" line among them names the matter document
 * it is drafted from, a "File:" line the file name the request gives it, an
 * "Executed by:" line the parties who execute it, "none" for none), and
 * each "## " line under it is a part with the lines below as its scope note. A
 * "Left out:" line anywhere under a title records a template part left out on
 * purpose and why. A "###" or deeper heading line is a scope line of the
 * current part, its heading text kept. A document's own line ("Form:", "File:"
 * and the others below) after its first part is read as nothing and named as
 * an issue: a line misplaced so must not pass silently as scope text. Only
 * mechanical defects are
 * issues; how many documents the request needs is the model's judgment. Every
 * defect is named in one pass; a shape defect (no title, a document without
 * parts, a repeated title or part heading) withholds the outline. A
 * missing "Acting for:" line is retried, and on the last attempt the outline
 * stands as not stated. With the request text, a stated party must be named in
 * the request; a party the request does not name is retry feedback and, if it
 * stands, not stated. A requested file name is held to the same standard: it
 * must be one the request gives, and a title must not carry a file name or
 * Markdown; both are retry feedback, and on the last attempt a file name the
 * request does not give is dropped and the title kept without Markdown code
 * marks; a title that is only a file name stands as that name's words
 * (plannedTitle).
 * A part's scope describes the part; a scope carrying an open-item marker is
 * retry feedback, and on the last attempt stands as the marker's words
 * (scopeWords), since the drafter sets the marker.
 * A missing "Executed by:" line is not an issue; the package-coverage check
 * holds the parties it names to the request and sources.
 * A "Form:" line under each title gives the document's form, one of
 * DOCUMENT_FORMS. A missing line or another value is retried, and on the last
 * attempt the document stands with no form, as a legacy plan does. A part's
 * label names a drafting job, not a heading of the document: every part is
 * planned unheaded, except a part with its own "Display heading:" line, unless
 * a "Headings: shown" line under its title keeps that document's labels. A part
 * headed as an attachment ("Schedule 1 …", "Exhibit A …") is marked as one and
 * prints its heading (plannedPart), except in a document drafted from a
 * template, whose own attachments are its template parts.
 * Optional "Legend:" and "Page numbers: none" lines under a title give what the
 * request asks of that document's page furniture; the first of each stands. A
 * legend must be the request's own words and is shown with the firm's, and
 * "Legend: none" asks nothing: a request never removes the firm's legend. An
 * ungrounded legend or another page numbers value is retry feedback and, on the
 * last attempt, dropped.
 * An optional "Memo:" line names, in the request's own words, a memo the
 * request asks to accompany the documents. The prompt asks for it before the
 * first title; it is read wherever it stands, so a memo named late is still
 * planned under the rules below. It must be words
 * the request uses and fit in a file name; otherwise it is retry feedback and,
 * on the last attempt, dropped, so the run plans no memo.
 * The memo is a planned document: the Memo line must name exactly one planned
 * document, the one whose title is the memo (repeatsMemo). That document is the
 * workflow's memo (memoDocument): Form: advice, executed by no one, drafted
 * last. It keeps its own "## " parts under the ordinary display rule; planned
 * without parts, it takes the cover-memo skeleton MEMO_PARTS. A Memo line
 * naming no planned document, or more than one, is retry feedback; on the last
 * attempt, of several, the one planned as advice, else the last, is the memo;
 * of none, the one document planned as advice or titled a memo, if only one
 * is, else the memo is added as its own document. With no Memo line standing,
 * the one document planned as advice, when no other is and another document is
 * planned, is the memo whatever its title: a memo accompanies something.
 * "Deliverable: combined" puts one document and that memo in one file: it needs
 * the memo, goes under a document other than the memo, and only one document
 * may have it; otherwise it is retry feedback and, on the last attempt, the
 * document is delivered as "body" and the memo keeps no deliverable. */
export function parsePlan(reply: string, request?: string, inferred?: ActingFor): Checked<LeanOutline> {
  if (!reply.trim()) return { issues: [`The reply was empty. ${OUTLINE_SHAPE}`] }
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const documents: {
    title: string; purpose: string[]; playbook_review?: PlaybookReview; mode?: "new" | "revise"; deliverable?: "body" | "redline" | "combined"; template?: string; comparison?: string; file?: string; executed_by?: string[]; form?: string; headings?: "shown"; furniture: FurnitureLines; left_out: LeftOutPart[]; parts: { heading: string; scope: string[]; display_heading?: string }[]
  }[] = []
  const acting: string[] = []
  const memos: string[] = []
  const selectionIssues: string[] = []
  // A document's own lines after its parts, which no document reads, named
  // back in one issue per document.
  const late: { title: string; line: string }[] = []
  for (const raw of reply.split(/\r?\n/)) {
    const line = raw.trim()
    const party = /^acting for\s*:\s*(.*)$/i.exec(line.replace(/^[-*]\s+/, "").replace(/[*`]/g, ""))?.[1]
    if (party !== undefined && !documents.length) {
      acting.push(party.replace(/[.\s]+$/, ""))
      continue
    }
    // The value keeps its own characters; outlineMemo cleans and checks it.
    const memoLine = /^(?:[-*]\s+)?[*`]*memo[*`]*\s*:[*`]*\s*(.*)$/i.exec(line)?.[1]
    if (memoLine !== undefined) {
      memos.push(memoLine)
      continue
    }
    const title = /^#\s+(.+)$/.exec(line)?.[1]
    const heading = /^##\s+(.+)$/.exec(line)?.[1]
    const current = documents.at(-1)
    const plain = line.replace(/\*\*/g, "")
    const template = /^[-*]?\s*template\s*:\s*(.+)$/i.exec(plain)?.[1]?.replace(/^["'`“]+|["'`”]+$/g, "").trim()
    const comparison = /^[-*]?\s*compare\s+against\s*:\s*(.+)$/i.exec(plain)?.[1]?.replace(/^["'`“]+|["'`”]+$/g, "").trim()
    const mode = /^[-*]?\s*mode\s*:\s*(new|revise)\s*$/i.exec(plain)?.[1]?.toLowerCase() as "new" | "revise" | undefined
    const deliverable = /^[-*]?\s*deliverable\s*:\s*(body|redline|combined)\s*$/i.exec(plain)?.[1]?.toLowerCase() as "body" | "redline" | "combined" | undefined
    const leftOut = /^[-*]?\s*left\s+out\s*:\s*(.+)$/i.exec(plain)?.[1]
    const file = /^[-*]?\s*file(?:\s+name)?\s*:\s*(.+)$/i.exec(plain)?.[1]?.replace(/^["'`“]+|["'`”.]+$/g, "").trim()
    const executed = /^executed\s+by\s*:\s*(.*)$/i.exec(line.replace(/^[-*]\s+/, "").replace(/[*`]/g, ""))?.[1]
    // The value as written, so an unknown form can be named back in the retry.
    const form = /^form\s*:\s*(.*)$/i.exec(line.replace(/^[-*]\s+/, "").replace(/[*`]/g, ""))?.[1]?.replace(/^["'“]+|["'”.\s]+$/g, "").trim()
    const headingsShown = /^headings\s*:\s*shown[.\s]*$/i.test(line.replace(/^[-*]\s+/, "").replace(/[*`]/g, ""))
    // Furniture values as written; documentFurniture cleans and checks them.
    const unmarked = line.replace(/^[-*]\s+/, "").replace(/[*`]/g, "")
    const legend = /^legend\s*:\s*(.*)$/i.exec(unmarked)?.[1]
    const pageNumbers = /^page\s+numbers\s*:\s*(.*)$/i.exec(unmarked)?.[1]
    const playbook = /^[-*]?\s*playbook review\s*:\s*(.+)$/i.exec(line)?.[1]
    const displayHeading = /^display heading\s*:\s*(.*)$/i.exec(line)?.[1]
    // A "###" or deeper heading writes a sub-part's name: it is scope text of
    // the current part, as a plain line under it would be.
    const sub = /^#{3,}(?!#)\s*(.*)$/.exec(line)?.[1]
    if (sub !== undefined) {
      const text = headingText(sub)
      if (text && current) (current.parts.at(-1)?.scope ?? current.purpose).push(text)
      continue
    }
    // A document's own line after its first part belongs to no document: read
    // as the part's scope, it would hide a document planned without its title.
    const field = playbook !== undefined || mode || deliverable || template !== undefined || comparison !== undefined || file !== undefined
      || executed !== undefined || form !== undefined || headingsShown || legend !== undefined || pageNumbers !== undefined
    if (field && current?.parts.length) {
      late.push({ title: current.title, line })
      continue
    }
    if (playbook && current && !current.parts.length) {
      let value: unknown
      try { value = JSON.parse(playbook) } catch {}
      if (isPlaybookReview(value) && !current.playbook_review) current.playbook_review = value
      else selectionIssues.push(`In "${current.title}", Playbook review must be one JSON object with sources [{source_id,quote}] and a nonempty reason, once per document.`)
      continue
    }
    if (title) documents.push({ title: headingText(title), purpose: [], furniture: {}, left_out: [], parts: [] })
    else if (heading && current) current.parts.push({ heading: headingText(heading), scope: [] })
    else if (displayHeading !== undefined && current?.parts.length) current.parts.at(-1)!.display_heading = /^none$/i.test(displayHeading.trim()) ? "" : headingText(displayHeading)
    else if (mode && current && !current.parts.length) current.mode = mode
    else if (deliverable && current && !current.parts.length) current.deliverable = deliverable
    else if (template && current && !current.parts.length) current.template = template
    else if (comparison && current && !current.parts.length) current.comparison = comparison
    else if (file && current && !current.parts.length) current.file = file
    else if (executed !== undefined && current && !current.parts.length) current.executed_by = executingParties(executed, current.executed_by)
    else if (form !== undefined && current && !current.parts.length) current.form = form
    else if (headingsShown && current && !current.parts.length) current.headings = "shown"
    else if (legend !== undefined && current && !current.parts.length) current.furniture.legend ??= legend
    else if (pageNumbers !== undefined && current && !current.parts.length) current.furniture.page_numbers ??= pageNumbers
    else if (leftOut && current) current.left_out.push(leftOutPart(leftOut))
    else if (line && !line.startsWith("#") && current)
      (current.parts.at(-1)?.scope ?? current.purpose).push(line.replace(/^[-*]\s+/, ""))
  }
  const titles = duplicates(documents.map(document => document.title))
  // A memo planned without parts takes the skeleton, so it may list none.
  const lined = outlineMemo(memos[0], request)
  // With no Memo line standing, the one document planned as advice, when no
  // other is and it is not the only document, is the memo whatever its title,
  // named by its title: a memo accompanies something. Only the internal advice
  // form is adopted: an advice letter is not.
  const advice = lined.name ? [] : documents.filter(document => documentForm(document.form) === "advice")
  const lone = advice.length === 1 && documents.length > 1 ? advice[0] : undefined
  const memo = lone ? { name: plainTitle(lone.title), issues: lined.issues } : lined
  const named = lone ? [lone] : memo.name ? documents.filter(document => repeatsMemo(plainTitle(document.title), memo.name!)) : []
  // With no title naming the memo, the last attempt adopts the one document
  // planned as advice or titled a memo before adding the memo as its own.
  const adoptable = memo.name && !named.length ? documents.filter(document => documentForm(document.form) === "advice" || titledMemo(plainTitle(document.title))) : []
  const memoAt = named.length ? documents.indexOf(named.find(document => documentForm(document.form) === "advice") ?? named.at(-1)!)
    : adoptable.length === 1 ? documents.indexOf(adoptable[0]!) : -1
  const own = documents.filter((_, index) => index !== memoAt)
  // Shape defects withhold the outline; every other defect is named with them,
  // so one revision can answer all of them.
  const shapeIssues = [
    ...selectionIssues,
    ...(!documents.length ? [`The outline has no "# " title line. ${OUTLINE_SHAPE}`] : []),
    ...own.filter(document => !document.parts.length)
      .map(document => `The document "${document.title}" lists no parts. Add one "## " line per top-level part of that document, in order.`),
    ...(titles.length ? [`These document titles appear more than once: ${titles.map(name => `"${name}"`).join(", ")}. Give each document a distinct title.`] : []),
    ...documents.flatMap(document => {
      const repeated = duplicates(document.parts.map(part => part.heading))
      return repeated.length
        ? [`In "${document.title}", these part headings appear more than once: ${repeated.map(name => `"${name}"`).join(", ")}. Give each part a distinct heading.`]
        : []
    }),
  ]
  const party = statedParty(acting[0] ?? "")
  // The party intake quoted, named again by the outline, keeps intake's basis.
  const asInferred = party !== undefined && inferred !== undefined && inferred.basis !== "none" && nameKey(party) === nameKey(inferred.party)
  const ungrounded = party !== undefined && request !== undefined && !namedIn(request, party) && !asInferred
  const files = documents.map(document => requestedFile(document.title, document.file, request))
  const repeatedFiles = duplicates(files.flatMap(file => file.name ?? []))
  const memoName = memo.name
  const combined = own.filter(document => document.deliverable === "combined").map(document => plainTitle(document.title))
  const combinedIssues = [
    ...(combined.length && !memo.name ? [COMBINED_NEEDS_MEMO] : []),
    ...(combined.length > 1 ? [combinedMoreThanOne(combined)] : []),
  ]
  const memoIssues = !memoName || !documents.length ? [] : !named.length ? [memoUnplanned(memoName, documents.map(document => plainTitle(document.title)))]
    : [...(named.length > 1 ? [memoNamesMore(memoName, named.map(document => plainTitle(document.title)))] : []),
      ...(documents[memoAt]!.deliverable === "combined" ? [combinedOnMemo(plainTitle(documents[memoAt]!.title))] : [])]
  const forms = documents.map((document, index) => index === memoAt ? "advice" as const : documentForm(document.form))
  const furniture = documents.map(document => documentFurniture(plainTitle(document.title), document.furniture, request))
  const issues = [
    ...shapeIssues,
    ...lateIssues(late),
    ...documents.flatMap(document => document.parts.flatMap(part => scopeIssues(plainTitle(document.title), part))),
    ...(!acting.length ? [ACTING_FOR_MISSING] : ungrounded ? [actingForUngrounded(party)] : []),
    ...documents.flatMap((document, index) => forms[index] ? [] : [formIssue(plainTitle(document.title), document.form)]),
    ...furniture.flatMap(row => row.issues),
    ...documents.flatMap(document => titleIssues(document.title)),
    ...files.flatMap(file => file.issues),
    ...(repeatedFiles.length ? [`These file names are given to more than one document: ${repeatedFiles.map(name => `"${name}"`).join(", ")}. Give each document its own file name.`] : []),
    ...memo.issues,
    ...memoIssues,
    ...combinedIssues,
  ]
  if (shapeIssues.length) return { issues }
  return {
    issues,
    value: {
      documents: [...documents.flatMap((document, index) => index === memoAt ? [] : [{
        title: plannedTitle(document.title),
        purpose: document.purpose.join(" "),
        ...(document.playbook_review ? { playbook_review: document.playbook_review } : {}),
        // Planning labels are not the document's headings.
        // A template's own attachments arrive as its parts (templateContract),
        // presented as the template presents them; they are not marked again.
        parts: document.parts.map(part => plannedPart(part, !document.headings, !document.template)),
        ...(document.mode ? { mode: document.mode } : {}),
        // A combined file that cannot be built stands as the body on the last attempt.
        ...(document.deliverable ? { deliverable: document.deliverable === "combined" && (combinedIssues.length || !memoName) ? "body" as const : document.deliverable } : {}),
        ...(document.comparison ? { comparison: document.comparison } : {}),
        ...(document.template ? { template: document.template } : {}),
        ...(document.left_out.length ? { left_out: document.left_out } : {}),
        ...(document.executed_by ? { executed_by: document.executed_by } : {}),
        ...(files[index]?.name && !repeatedFiles.includes(files[index].name) ? { file: files[index].name } : {}),
        ...(forms[index] ? { form: forms[index] } : {}),
        ...(furniture[index]?.value ? { furniture: furniture[index].value } : {}),
      }]),
      // The lawyer asked for the memo: it is planned last, as the workflow's memo.
      ...(memoName ? [memoDocument(memoAt < 0 ? { title: memoTitle(memoName), purpose: MEMO_PURPOSE, parts: memoParts() } : {
        title: plannedTitle(documents[memoAt]!.title),
        purpose: documents[memoAt]!.purpose.join(" ") || MEMO_PURPOSE,
        parts: documents[memoAt]!.parts.length ? documents[memoAt]!.parts.map(part => plannedPart(part, !documents[memoAt]!.headings)) : memoParts(),
        ...(documents[memoAt]!.playbook_review ? { playbook_review: documents[memoAt]!.playbook_review } : {}),
        ...(files[memoAt]?.name && !repeatedFiles.includes(files[memoAt].name) ? { file: files[memoAt].name } : {}),
        ...(furniture[memoAt]?.value ? { furniture: furniture[memoAt].value } : {}),
      })] : [])],
      acting_for: party === undefined || ungrounded ? NOT_STATED : request !== undefined && namedIn(request, party) || !asInferred ? { party, basis: "stated" } : inferred!,
      ...(memoName ? { memo: memoName } : {}),
    },
  }
}

/** A planned part; with `unheaded`, a part without its own "Display heading:"
 * line is unheaded. With `attachments` (a document not drafted from a
 * template), a part whose printed heading (its display heading, else its
 * label) is an attachment's heading is that attachment, and prints it. */
function plannedPart(part: { heading: string; scope: string[]; display_heading?: string }, unheaded: boolean, attachments = true): PlannedPart {
  const attachment = attachments && Boolean(attachmentHead(part.display_heading || part.heading))
  const display = attachment ? part.display_heading || undefined : part.display_heading ?? (unheaded ? "" : undefined)
  return { heading: part.heading, scope: scopeWords(part.scope.join(" ")), ...(display !== undefined ? { display_heading: display } : {}), ...(attachment ? { attachment: true as const } : {}) }
}

/** An open-item marker or blank in a scope line (OPEN_ITEM_PATTERN, without its global flag). */
const OPEN_ITEM_IN_SCOPE = new RegExp(OPEN_ITEM_PATTERN.source, "i")

/** A part whose scope carries an open-item marker: a scope describes the part,
 * and a marker in it is copied into the part beside the drafter's own. */
function scopeIssues(title: string, part: { heading: string; scope: string[] }) {
  return OPEN_ITEM_IN_SCOPE.test(part.scope.join(" "))
    ? [`In "${title}", the scope of "${part.heading}" carries an open-item marker. Scope describes the part; write the open point as scope words, the drafter sets the marker.`]
    : []
}

/** The scope the plan keeps: as written, except that an open-item marker still
 * in it on the last attempt stands as its words, and a blank as nothing. */
function scopeWords(scope: string) {
  if (!OPEN_ITEM_IN_SCOPE.test(scope)) return scope
  return scope.replace(OPEN_ITEM_PATTERN, marker => /^[[(]\s*open item:\s*(.*?)\s*[\])]$/i.exec(marker)?.[1] ?? "").replace(/\s+/g, " ").trim()
}

export function headingText(text: string) {
  return text.replace(/\s+#+\s*$/, "").trim()
}

export function normalizedHeading(text: string) {
  return headingText(text).replace(/\s+/g, " ").toLowerCase()
}

/** `"<part>" – <reason>`: a quoted part name, else the text before the first
 * dash or colon. */
function leftOutPart(text: string): LeftOutPart {
  const quoted = /^["“]([^"”]+)["”](.*)$/.exec(text.trim()) ?? /^[`'‘](.+?)[`'’](?=\s|$)(.*)$/.exec(text.trim())
  const [part = "", reason = ""] = quoted ? [quoted[1]!, quoted[2]!] : (() => {
    const at = /\s[-–—]\s|:\s/.exec(text)
    return at ? [text.slice(0, at.index), text.slice(at.index + at[0].length)] : [text, ""]
  })()
  return { part: part.trim(), reason: reason.replace(/^[\s\-–—:,;.]+/, "").trim() }
}

/** A Word file name, as an outline title may wrongly carry one. */
const FILE_IN_TITLE = /[^\s`'"“”]+\.docx?\b/i
const FILE_NAME_FORBIDDEN = /[\\/:*?"<>|\u0000-\u001f`]/

/** Whether `file` can stand as a requested file name: no folder, Markdown code
 * mark or character a file name cannot carry, and at most 120 characters. */
export function plainFileName(file: string) {
  return !FILE_NAME_FORBIDDEN.test(file) && file.length <= 120
}

/** Begins every misplaced-line issue: the order the outline's shape reads. */
const DOCUMENT_ORDER = 'Begin each document with its own "# " title line, then its lines such as "Form:" and "File:", then its "## " parts.'
const quotedLines = (lines: string[]) => lines.map(line => `"${line}"`).join(", ")

/** A document's own lines written after a document's first part, one issue per document. */
function lateIssues(rows: { title: string; line: string }[]) {
  return [...new Set(rows.map(row => row.title))].map(title => {
    const lines = rows.filter(row => row.title === title).map(row => row.line)
    return `${lines.length === 1 ? `The line ${quotedLines(lines)} comes` : `The lines ${quotedLines(lines)} come`} after the "## " parts of "${plainTitle(title)}", so no document reads ${lines.length === 1 ? "it" : "them"}. ${DOCUMENT_ORDER}`
  })
}

function titleIssues(title: string) {
  if (FILE_IN_TITLE.test(title))
    return [`The title "${title}" carries a file name. Write the document's own title on the "# " line, and give the file name the request asks for on a "File:" line under it.`]
  if (title.includes("`")) return [`The title "${title}" contains Markdown code marks. Write the title as plain text.`]
  return []
}

/** The title as plain text: Markdown code marks are not part of it, nor is a
 * leading "Subject:", "Re:", "Fw:" or "Fwd:" label a letter or email's heading
 * line carries (the label is not the document's name, and the file name drops
 * its colon). A label is only a label with its colon or a spaced dash after it:
 * "Subject Access Request", "Subject to contract" and a case-style "Re Smith"
 * stay. */
function plainTitle(title: string) {
  const plain = title.replaceAll("`", "").replace(/\s+/g, " ").trim()
  return plain.replace(/^(?:(?:subject|re|fwd?)(?:\s*:|\s+[-–—])\s*)+/i, "").trim() || plain
}

/** The title the plan keeps: the plain title, except that a title which is
 * only a file name, still so on the last attempt, stands as the words of that
 * name ("fee-letter.docx" as "Fee letter"), not as a file name. */
function plannedTitle(title: string) {
  const plain = plainTitle(title)
  const named = /^["'“‘]?([^\s`'"“”‘’]+)\.docx?["'”’]?$/i.exec(plain)?.[1]
  if (!named) return plain
  const words = named.replace(/[-_.]+/g, " ").replace(/\s+/g, " ").trim()
  if (!words) return plain
  return words === words.toLowerCase() ? words.charAt(0).toUpperCase() + words.slice(1) : words
}

/** A "File:" line's name, kept only when it is a plain file name the request gives. */
function requestedFile(title: string, file: string | undefined, request: string | undefined): { name?: string; issues: string[] } {
  if (!file) return { issues: [] }
  if (!plainFileName(file))
    return { issues: [`The "File:" line of "${title}" gives "${file}", which is not a plain file name. Give the file name only, as the request writes it, without a folder or Markdown.`] }
  if (request !== undefined && !namedIn(request, file))
    return { issues: [`The "File:" line of "${title}" gives "${file}", but the request gives no such file name. Add a "File:" line only for a file name the request gives, in its words.`] }
  return { name: file, issues: [] }
}

/** A "Memo:" line's name, kept only when it is words the request uses and can
 * be part of a file name. A file name the request gives counts as its words,
 * its hyphens and underscores as spaces ("cover-memo.docx" uses "cover memo"). List markers, bold and code marks, wrapping quotes,
 * trailing full stops and one leading article are not part of the name; only a
 * value that is wholly a non-answer is no memo. */
function outlineMemo(line: string | undefined, request: string | undefined): { name?: string; issues: string[] } {
  const name = (line ?? "").trim().replace(/^[-*]\s+/, "").replace(/\*\*|`/g, "").trim().replace(/^["'“‘]+|["'”’.]+$/g, "").trim()
    .replace(/^(?:a|an|the)\s+/i, "").replace(/\s+/g, " ")
  if (!name || MEMO_NON_ANSWER.test(name)) return { issues: [] }
  if (FILE_NAME_FORBIDDEN.test(name) || FILE_IN_TITLE.test(name))
    return { issues: [`The "Memo:" line gives "${name}", which cannot be part of a file name. Name the memo in a few of the request's own words, without a folder, file extension, Markdown or any of \\ / : * ? " < > |.`] }
  if (request !== undefined && !namedIn(request, name) && !namedIn(request.replace(/[-_]+/g, " "), name))
    return { issues: [`The "Memo:" line names "${name}", but the request does not use those words. Name the memo in a few of the request's own words, as it writes them, and keep the memo planned as its "# " document. Leave the line and that document out only when the request does not ask for a memo or note.`] }
  return { name, issues: [] }
}

/** Whether a planned title is the requested memo itself, not merely the subject
 * the memo is about. A memo name reads "<what it is> <connective> <its subject>",
 * so the words after its first connective ("on", "for", "to", "about"...) name
 * what it covers. A title whose words all sit in that subject names the memo's
 * subject, not the memo. Otherwise the title is the memo when the content words
 * of one are all among the other's, ignoring case, punctuation, function
 * words and "memorandum" for "memo", so "Memo" names "Cover Memo". */
function repeatsMemo(title: string, memo: string) {
  const words = (tokens: string[]) => new Set(tokens.filter(word => word && !FUNCTION_WORDS.has(word)))
  const tokens = (text: string) => text.toLowerCase().split(/[^\p{L}\p{N}]+/u).map(word => /^memorand(?:um|a)$|^memos$/.test(word) ? "memo" : word)
  const memoTokens = tokens(memo)
  const at = memoTokens.findIndex(word => MEMO_CONNECTIVES.has(word))
  const subject = words(at < 0 ? [] : memoTokens.slice(at + 1))
  const [planned, named] = [words(tokens(title)), words(memoTokens)]
  const within = (some: Set<string>, all: Set<string>) => some.size > 0 && [...some].every(word => all.has(word))
  if (within(planned, subject)) return false
  return within(planned, named) || within(named, planned)
}

/** Whether a planned title calls itself a memo ("memorandum" counting as "memo"). */
const titledMemo = (title: string) => title.toLowerCase().split(/[^\p{L}\p{N}]+/u).some(word => /^memo(?:s|randum|randa)?$/.test(word))

const MEMO_CONNECTIVES = new Set(["on", "for", "to", "about", "regarding", "concerning", "re", "of", "in", "with", "from", "at", "by"])
const FUNCTION_WORDS = new Set(["a", "an", "the", "to", "for", "of", "on", "and", "with", "that", "in"])
const memoUnplanned = (memo: string, titles: string[]) =>
  `The "Memo:" line names "${memo}", but no planned title is that memo; the planned titles are ${titles.map(title => `"${title}"`).join(", ")}. Plan the memo as one "# " document whose title has the memo's words, with "Form: advice", and plan its parts as for any document.`
const memoNamesMore = (memo: string, titles: string[]) =>
  `The "Memo:" line names "${memo}", which ${titles.map(title => `"${title}"`).join(", ")} all repeat. Plan the memo as one "# " document, and give the others titles that do not repeat it.`
const combinedOnMemo = (title: string) =>
  `"Deliverable: combined" is under the memo "${title}" itself. Write it under the document the request asks to deliver with the memo, and leave it out under the memo.`

/** The workflow's memo: drafted last, as advice that takes no legal act, in the
 * parts given. The open points are not a part: the gate lists them from the
 * final register. */
function memoDocument(memo: Pick<PlannedDocument, "title" | "purpose" | "playbook_review" | "file" | "furniture" | "parts">): Omit<PlannedDocument, "id"> {
  return { ...memo, mode: "new", executed_by: [], form: "advice" }
}

/** A fresh copy of the cover-memo skeleton, for a memo planned without parts. */
const memoParts = () => MEMO_PARTS.map(part => ({ ...part }))

/** A cover memo's parts from general practice, each scope line its content rule
 * and word cap. */
const MEMO_PARTS: PlannedPart[] = [
  { heading: "What this is and what you need to do", scope: "At most 150 words: the documents delivered, whom they are drafted for, and the decisions the reader needs to take." },
  { heading: "Decisions taken and why", scope: "At most 300 words: each point decided in drafting on a delegated or advised question, one short paragraph each, saying what was decided and why." },
  { heading: "Discrepancies and how handled", scope: "At most 300 words: each discrepancy between the sources or instructions, one short entry each: the point, the sources that differ, how the documents handle it, and whether it stays open." },
  { heading: "What the documents do", scope: "At most 250 words: what each delivered document does, part by part, briefly." },
  { heading: "Research relied on and its limits", scope: "At most 200 words: one paragraph on the research relied on and its limits, then the authorities relied on." },
]
const MEMO_PURPOSE = "The cover memo the request asks for: what the documents are, the decisions taken and what the reader needs to do."
/** The memo name as a title: its first letter capitalised. */
const memoTitle = (memo: string) => memo.charAt(0).toLocaleUpperCase() + memo.slice(1)
const MEMO_NON_ANSWER = /^(?:not\s+(?:requested|stated|specified|given|known|clear)|no(?:\s+memo)?|none|nothing|unknown|unspecified|unclear|n\/?a)$/i
const COMBINED_NEEDS_MEMO = "\"Deliverable: combined\" puts a document and the requested memo in one file, so it needs a \"Memo:\" line before the first \"# \" title " +
  "naming that memo in the request's own words; otherwise write \"Deliverable: body\"."
const combinedMoreThanOne = (titles: string[]) =>
  `"Deliverable: combined" is under ${titles.map(title => `"${title}"`).join(", ")}, but only one document can share the file with the memo. ` +
  "Keep it under the document the request asks to deliver with the memo, and write \"Deliverable: body\" under the others."

/** A "Form:" value as one of DOCUMENT_FORMS, ignoring case and spelling
 * "advice letter" with a hyphen or underscore; undefined when the line is
 * missing or names another form. */
function documentForm(value: string | undefined) {
  return DOCUMENT_FORMS.find(form => form === value?.toLowerCase().replace(/[-_]/g, " "))
}

const FORM_VALUES = DOCUMENT_FORMS.map(form => `"Form: ${form}"`).join(", ")
const formIssue = (title: string, value: string | undefined) => (!value
  ? `The document "${title}" has no "Form:" line.`
  : `The "Form:" line of "${title}" gives "${value}", which is not one of the forms.`) +
  ` Under its title, before its first "## " part, write one of ${FORM_VALUES}.`

type FurnitureLines = { legend?: string; page_numbers?: string }

/** A document's furniture lines as the plan keeps them, each key only when its
 * line stands: a legend the request uses as written, cleaned as a firm's legend
 * is, or "Page numbers: none". Wrapping quotes and trailing full stops are not
 * part of a value; an empty legend or "Legend: none" is no line. */
function documentFurniture(title: string, lines: FurnitureLines, request: string | undefined): { value?: PlannedDocument["furniture"]; issues: string[] } {
  const clean = (value: string) => normalizeLegend(value.trim().replace(/^["'“‘]+|["'”’.]+$/g, ""))
  const written = clean(lines.legend ?? "")
  const asked = /^none$/i.test(written) ? "" : written
  const legend = asked && (request === undefined || wordsIn(request, asked)) ? asked : undefined
  const pageNumbers = lines.page_numbers === undefined ? undefined : /^none$/i.test(clean(lines.page_numbers))
  const value = { ...(legend ? { legend } : {}), ...(pageNumbers ? { page_numbers: false as const } : {}) }
  const issues = [
    ...(asked && !legend ? [legendUngrounded(title, asked)] : []),
    ...(pageNumbers === false ? [pageNumbersNotNone(title, lines.page_numbers!.trim())] : []),
  ]
  return { ...(Object.keys(value).length ? { value } : {}), issues }
}

/** Whether `text` uses `words` as written, ignoring case, spacing and the
 * forms of quotes and dashes, not starting or ending inside a word. */
function wordsIn(text: string, words: string) {
  const fold = (value: string) => value.normalize("NFKC").toLowerCase().replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/[‐-―−]/g, "-").replace(/\s+/g, " ")
  const key = fold(words).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
  return new RegExp((WORD_START.test(key) ? `(?<!${WORD})` : "") + key + (WORD_END.test(key) ? `(?!${WORD})` : ""), "u").test(fold(text))
}

const legendUngrounded = (title: string, legend: string) =>
  `The "Legend:" line of "${title}" gives "${legend}", but the request does not use those words. Write a "Legend:" line only when the request asks for a document's header legend, ` +
  "in the request's own words; otherwise leave the line out."
const pageNumbersNotNone = (title: string, value: string) =>
  `The "Page numbers:" line of "${title}" gives "${value}", but the only value it takes is "Page numbers: none". ` +
  "Write it only when the request asks for the document without page numbers; otherwise leave the line out."

const OUTLINE_SHAPE = "Reply with the outline: for each document, one \"# \" title line, then one \"## \" line per part."
const ACTING_FOR_MISSING = "The outline has no \"Acting for:\" line. Before the first \"# \" title, write \"Acting for: <party>\" when the request says " +
  "which party the lawyer acts for, otherwise \"Acting for: not stated\"; then the complete outline."

/** A party answer without wrapping brackets or quotes and a trailing
 * explanation: a final bracket that opens with a lower-case word or a quotation
 * mark, or what follows a semicolon or a dash before a lower-case word ("Acme \u2013
 * the seller"). A dash or bracket that continues the name stays, as in
 * "Heron \u2013 Vale LLP" or "Acme Ltd (UK)", and so does a bracket the name closes. */
function partyText(text: string) {
  return text.trim().replace(/[.\s]+$/, "").replace(/\s+(?:\((?=[\p{Ll}"'\u201C\u2018])[^()]*\)|\[(?=[\p{Ll}"'\u201C\u2018])[^\]]*\])[.\s]*$/u, "")
    .replace(/\s*(?:;|\s-\s+(?=\p{Ll})|[\u2013\u2014]\s*(?=\p{Ll})).*$/u, "")
    .replace(/^[[("'\u201C\u2018<]+/, "").replace(/(?<![([][^()[\]]*)[\])"'\u201D\u2019>]+$/, "").replace(/[.:\s]+$/, "").trim()
}

/** "none" as the whole answer, not the first word of a name such as "None Corp":
 * only an explanation, a closing bracket or quote, or nothing follows it. */
const NONE = /^none\s*(?:$|[;.,:()[\]"\u201D>])/i
const NON_ANSWER = /^(?:(?:not|none)\s+(?:stated|specified|given|known|clear|applicable|named|identified|provided)|no\s+(?:party|parties|one)|nobody|unknown|unspecified|unclear|unnamed|n[/.]?a|tb[cd]|to\s+be\s+(?:confirmed|determined))\b/i

/** The party an "Acting for:" line names, or undefined when the line gives a
 * non-answer. Wrapping brackets or quotes and a trailing explanation are not
 * part of the name. */
function statedParty(text: string) {
  const party = partyText(text)
  return !party || NONE.test(party) || NON_ANSWER.test(party) ? undefined : party
}

/** An "Executed by:" line's parties, added to any the document already
 * declares under the same name: "none" declares none, and a non-answer
 * declares nothing. */
function executingParties(text: string, declared?: string[]) {
  if (NONE.test(partyText(text))) return declared ?? []
  const parties = [...declared ?? []]
  for (const party of text.split(";").flatMap(part => statedParty(part) ?? []))
    if (nameKey(party) && !parties.some(name => nameKey(name) === nameKey(party))) parties.push(party)
  return parties.length || declared ? parties : undefined
}

/** Case, whitespace, quotation marks and the punctuation . , ; : do not count:
 * "Acme, Inc." folds to "acme inc". Curly apostrophes fold to "'". */
const foldName = (text: string) =>
  text.normalize("NFC").replace(/[\u2018\u2019]/g, "'").replace(/["\u201C\u201D.,;:]/g, " ").replace(/\s+/g, " ").trim().toLowerCase()

/** Common spellings of a legal form that ends a name, as foldName folds them,
 * longest first; the first of each form's spellings stands for it. */
const LEGAL_FORMS = [["ltd", "limited"], ["inc", "incorporated"], ["co", "company"], ["corp", "corporation"], ["llc", "l l c"], ["llp", "l l p"], ["lp", "l p"], ["plc", "p l c"]]
  .flatMap(spellings => spellings.map(spelling => ({ spelling, spellings }))).toSorted((a, b) => b.spelling.length - a.spelling.length)

/** A folded name without a leading "the", split into the words before its
 * trailing legal forms and each form's spellings, in order. A form word alone
 * is a name, not a form: "the Company" is "company". */
function nameParts(party: string) {
  let head = foldName(party).replace(/^the /, "")
  const forms: string[][] = []
  for (;;) {
    const form = LEGAL_FORMS.find(row => head.endsWith(` ${row.spelling}`))
    if (!form) return { head, forms }
    head = head.slice(0, -form.spelling.length - 1)
    forms.unshift(form.spellings)
  }
}

/** A party's name as compared: case, whitespace, quotation marks, the
 * punctuation . , ; :, a leading "the" and how a trailing legal form is spelled
 * do not count. The form does: "Acme Ltd." is "Acme Limited" but not "Acme LLC". */
export function nameKey(party: string) {
  const { head, forms } = nameParts(party)
  return [head, ...forms.map(spellings => spellings[0])].join(" ")
}

const WORD = "[\\p{L}\\p{M}\\p{N}]"
const WORD_START = new RegExp(`^${WORD}`, "u")
const WORD_END = new RegExp(`${WORD}$`, "u")

/** Whether `text` names `party` as whole words, compared as nameKey does: the
 * name does not start or end inside a word, so "Ann" is not named by "Hannah".
 * An empty name is named nowhere. */
export function namedIn(text: string, party: string) {
  const { head, forms } = nameParts(party)
  if (!head) return false
  const key = [head.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), ...forms.map(spellings => `(?:${spellings.join("|")})`)].join(" ")
  const pattern = (WORD_START.test(head) ? `(?<!${WORD})` : "") + key + (forms.length || WORD_END.test(head) ? `(?!${WORD})` : "")
  return new RegExp(pattern, "u").test(foldName(text))
}

const actingForUngrounded = (party: string) =>
  `The "Acting for:" line names "${party}", but the request does not name that party. When the request says whom the lawyer acts for, ` +
  "name that party in the request's own words; otherwise write \"Acting for: not stated\". The run takes the side from the matter and its sources itself."

function duplicates(names: string[]) {
  const seen = names.map(normalizedHeading)
  return [...new Set(names.filter((_, index) => seen.indexOf(seen[index]!) !== index))]
}

function isPlannedDocument(value: unknown): value is PlannedDocument {
  if (typeof value !== "object" || value === null) return false
  const row = value as PlannedDocument
  return typeof row.id === "string" && /^d[1-9]\d*$/.test(row.id) && typeof row.title === "string" && typeof row.purpose === "string" &&
    Array.isArray(row.parts) && row.parts.length > 0 && row.parts.every(part => typeof part?.heading === "string" && typeof part.scope === "string" && (part.display_heading === undefined || typeof part.display_heading === "string") &&
      (part.instrument === undefined || isPlannedInstrument(part.instrument)) && (part.attachment === undefined || part.attachment === true)) &&
    (row.playbook_review === undefined || isPlaybookReview(row.playbook_review)) &&
    (row.mode === undefined || ["new", "revise"].includes(row.mode)) && (row.deliverable === undefined || ["body", "redline", "combined"].includes(row.deliverable)) &&
    (row.comparison === undefined || typeof row.comparison === "string") && (row.template === undefined || typeof row.template === "string") && (row.file === undefined || typeof row.file === "string") &&
    (row.executed_by === undefined || (Array.isArray(row.executed_by) && row.executed_by.every(party => typeof party === "string"))) &&
    (row.form === undefined || DOCUMENT_FORMS.includes(row.form)) && (row.furniture === undefined || isPlannedFurniture(row.furniture)) &&
    (row.left_out === undefined || (Array.isArray(row.left_out) && row.left_out.every(item => typeof item?.part === "string" && typeof item.reason === "string")))
}

function isPlannedInstrument(value: unknown) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const row = value as PlannedInstrument
  return typeof row.title === "string" && !!row.title.trim() && (row.purpose === undefined || typeof row.purpose === "string") &&
    (row.executed_by === undefined || (Array.isArray(row.executed_by) && row.executed_by.every(party => typeof party === "string"))) &&
    (row.form === undefined || DOCUMENT_FORMS.includes(row.form))
}

/** The separate instrument part `index` is in, when the plan folded one into
 * the document (PlannedPart.instrument): the nearest marked part at or before
 * it. Undefined for the document's own parts. */
export function partInstrument(parts: PlannedPart[], index: number): PlannedInstrument | undefined {
  for (let at = index; at >= 0; at--) if (parts[at]?.instrument) return parts[at]!.instrument
}

function isPlannedFurniture(value: unknown) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const row = value as NonNullable<PlannedDocument["furniture"]>
  return (row.legend === undefined || typeof row.legend === "string") && (row.page_numbers === undefined || row.page_numbers === false)
}

function isPlaybookReview(value: unknown): value is PlaybookReview {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const row = value as PlaybookReview
  return typeof row.reason === "string" && !!row.reason.trim() && Array.isArray(row.sources) && row.sources.every(source =>
    typeof source?.source_id === "string" && !!source.source_id.trim() && typeof source.quote === "string" && !!source.quote.trim())
}
