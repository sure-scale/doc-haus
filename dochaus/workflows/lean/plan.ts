import { isPosture, normalizeLegend, type Posture } from "../../lib/drafting-preferences"
import { effectivePosture } from "../../lib/matter-client"
import { DOCUMENT_FORMS, type DocumentForm } from "../../lib/page-furniture"
import { TOOL_SYNTAX_ISSUE, printedToolCall } from "./checks"
import { REVIEW_STAGES, readLeanJson, readLeanText, writeLeanText, type ReviewStageID, type StageID } from "./state"

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
export type StepID = `${DocID}.${DocStage}` | "discrepancy" | "intake" | "research" | "research-review" | "spine" | "finalize" | "checklist-check" | "package-review" | "gate"

/** heading identifies the drafting part; absent display_heading preserves it
 * in publication (a plan saved before labels were planned unheaded, or one
 * whose outline shows its headings), while an empty display_heading makes the
 * part unheaded. */
export type PlannedPart = { heading: string; scope: string; display_heading?: string }
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
 * one that the outline did not say. `form` is the kind of document it is, from
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
export type ActingFor = { party: string; stated: boolean }
/** `memo` is the name of the memo the request asks for and `notes_file` the
 * file the notes are delivered as; both are delivery metadata. */
export type LeanPlan = { schema: "dochaus.lean.plan/1"; documents: PlannedDocument[]; acting_for?: ActingFor; posture?: Posture; basis?: string; memo?: string; notes_file?: string }
/** `memo` is the request's own words for a memo that accompanies the documents,
 * from a "Memo:" line; the run writes and names the drafting notes as that memo. */
export type LeanOutline = { documents: Omit<PlannedDocument, "id">[]; acting_for: ActingFor; memo?: string }

export const LEAN_PLAN_SCHEMA = "dochaus.lean.plan/1"
export const FIRST_DOCUMENT: DocID = "d1"
export const NOT_STATED: ActingFor = { party: "", stated: false }

type Checked<T> = { issues: string[]; value?: T }

/** The steps of the run in walk order: before a plan exists, only the first
 * document's steps are known. */
export function leanSteps(directory: string): StepID[] {
  const ids = readLeanPlan(directory)?.documents.map(document => document.id) ?? [FIRST_DOCUMENT]
  return ["discrepancy", "intake", "research", "research-review", "spine", ...ids.flatMap(id => DOC_STAGES.filter(stage => stage !== "review-playbook" || reviewStages(directory, id).includes(stage)).map((stage): StepID => `${id}.${stage}`)), "finalize", "checklist-check", "package-review", "gate"]
}

/** Planning selects guidance per body; legacy and unassigned bodies retain the ordinary reviews. */
export function reviewStages(directory: string, document: DocID): ReviewStageID[] {
  return [...REVIEW_STAGES, ...(plannedDocument(directory, document)?.playbook_review?.sources.length ? ["review-playbook" as const] : [])]
}

export function stepStage(step: StepID): StageID {
  return !step.includes(".") ? step as StageID : step.slice(step.indexOf(".") + 1) as DocStage
}

export function stepDocument(step: StepID): DocID | null {
  return !step.includes(".") ? null : step.slice(0, step.indexOf(".")) as DocID
}

export function isStepID(value: string): value is StepID {
  return ["discrepancy", "intake", "research", "research-review", "spine", "finalize", "checklist-check", "package-review", "gate"].includes(value) || new RegExp(`^d[1-9]\\d*\\.(${DOC_STAGES.join("|")})$`).test(value)
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
}

/** A requested memo is saved as its name and the notes file it is delivered
 * as; a plan without one carries neither key, nor a plan without a Posture
 * the Posture key. */
export function writeLeanPlan(directory: string, documents: Omit<PlannedDocument, "id">[], basis?: string, actingFor?: ActingFor, memo?: { name: string; file: string }, posture?: Posture): LeanPlan {
  const plan: LeanPlan = {
    schema: LEAN_PLAN_SCHEMA, documents: withDocumentIDs(documents), acting_for: actingFor, posture, basis,
    ...(memo ? { memo: memo.name, notes_file: memo.file } : {}),
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
  const { notes_file, memo, ...operative } = plan
  return { ...operative, documents: plan.documents.map(operativeDocument) }
}

/** Whom the lawyer acts for. A plan without it, or with no party named, says
 * the request did not state it. */
export function planActingFor(directory: string): ActingFor {
  const value: unknown = readLeanPlan(directory)?.acting_for
  if (typeof value !== "object" || value === null) return NOT_STATED
  const row = value as Partial<ActingFor>
  return typeof row.party === "string" && row.party.trim() && row.stated !== false ? { party: row.party.trim(), stated: true } : NOT_STATED
}

/** How far the run's drafts lean toward the party the lawyer acts for: the
 * Posture the outline froze into the plan. Before the outline, and for a plan
 * saved without one, the matter's Posture, else the firm's, read live as
 * before. Every lean stage renders its drafting preferences with it. */
export function runPosture(directory: string): Posture {
  const posture: unknown = readLeanPlan(directory)?.posture
  return isPosture(posture) ? posture : effectivePosture(directory)
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
 * purpose and why. Only mechanical defects are
 * issues; how many documents the request needs is the model's judgment. A
 * missing "Acting for:" line is retried, and on the last attempt the outline
 * stands as not stated. With the request text, a stated party must be named in
 * the request; a party the request does not name is retry feedback and, if it
 * stands, not stated. A requested file name is held to the same standard: it
 * must be one the request gives, and a title must not carry a file name or
 * Markdown; both are retry feedback, and on the last attempt a file name the
 * request does not give is dropped and the title kept without Markdown code marks.
 * A missing "Executed by:" line is not an issue; the package-coverage check
 * holds the parties it names to the request and sources.
 * A "Form:" line under each title gives the document's form, one of
 * DOCUMENT_FORMS. A missing line or another value is retried, and on the last
 * attempt the document stands with no form, as a legacy plan does. A part's
 * label names a drafting job, not a heading of the document: every part is
 * planned unheaded, except a part with its own "Display heading:" line, unless
 * a "Headings: shown" line under its title keeps that document's labels.
 * Optional "Legend:" and "Page numbers: none" lines under a title give what the
 * request asks of that document's page furniture; the first of each stands. A
 * legend must be the request's own words and is shown with the firm's, and
 * "Legend: none" asks nothing: a request never removes the firm's legend. An
 * ungrounded legend or another page numbers value is retry feedback and, on the
 * last attempt, dropped.
 * An optional "Memo:" line before the first title names, in the request's own
 * words, a memo the request asks to accompany the documents. It must be words
 * the request uses and fit in a file name; otherwise it is retry feedback and,
 * on the last attempt, dropped, so the run keeps its ordinary drafting notes.
 * A document whose title repeats the memo is retry feedback too; on the last
 * attempt the memo is dropped and the document kept.
 * "Deliverable: combined" puts one document and that memo in one file: it needs
 * the memo and only one document may have it; otherwise it is retry feedback
 * and, on the last attempt, the document is delivered as "body". */
export function parsePlan(reply: string, request?: string): Checked<LeanOutline> {
  if (!reply.trim()) return { issues: [`The reply was empty. ${OUTLINE_SHAPE}`] }
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const documents: {
    title: string; purpose: string[]; playbook_review?: PlaybookReview; mode?: "new" | "revise"; deliverable?: "body" | "redline" | "combined"; template?: string; comparison?: string; file?: string; executed_by?: string[]; form?: string; headings?: "shown"; furniture: FurnitureLines; left_out: LeftOutPart[]; parts: { heading: string; scope: string[]; display_heading?: string }[]
  }[] = []
  const acting: string[] = []
  const memos: string[] = []
  const selectionIssues: string[] = []
  for (const raw of reply.split(/\r?\n/)) {
    const line = raw.trim()
    const party = /^acting for\s*:\s*(.*)$/i.exec(line.replace(/^[-*]\s+/, "").replace(/[*`]/g, ""))?.[1]
    if (party !== undefined && !documents.length) {
      acting.push(party.replace(/[.\s]+$/, ""))
      continue
    }
    // The value keeps its own characters; outlineMemo cleans and checks it.
    const memoLine = /^(?:[-*]\s+)?[*`]*memo[*`]*\s*:[*`]*\s*(.*)$/i.exec(line)?.[1]
    if (memoLine !== undefined && !documents.length) {
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
  const issues = [
    ...selectionIssues,
    ...(!documents.length ? [`The outline has no "# " title line. ${OUTLINE_SHAPE}`] : []),
    ...documents.filter(document => !document.parts.length)
      .map(document => `The document "${document.title}" lists no parts. Add one "## " line per top-level part of that document, in order.`),
    ...(titles.length ? [`These document titles appear more than once: ${titles.map(name => `"${name}"`).join(", ")}. Give each document a distinct title.`] : []),
    ...documents.flatMap(document => {
      const repeated = duplicates(document.parts.map(part => part.heading))
      return repeated.length
        ? [`In "${document.title}", these part headings appear more than once: ${repeated.map(name => `"${name}"`).join(", ")}. Give each part a distinct heading.`]
        : []
    }),
  ]
  if (issues.length) return { issues }
  const party = statedParty(acting[0] ?? "")
  const ungrounded = party !== undefined && request !== undefined && !namedIn(request, party)
  const files = documents.map(document => requestedFile(document.title, document.file, request))
  const repeatedFiles = duplicates(files.flatMap(file => file.name ?? []))
  const memo = outlineMemo(memos[0], request)
  const repeated = memo.name ? documents.map(document => plainTitle(document.title)).filter(title => repeatsMemo(title, memo.name!)) : []
  // The lawyer asked for the memo: a repeat is retry feedback only, never a reason to drop it.
  const memoName = memo.name
  const combined = documents.filter(document => document.deliverable === "combined").map(document => plainTitle(document.title))
  const combinedIssues = [
    ...(combined.length && !memo.name ? [COMBINED_NEEDS_MEMO] : []),
    ...(combined.length > 1 ? [combinedMoreThanOne(combined)] : []),
  ]
  const forms = documents.map(document => documentForm(document.form))
  const furniture = documents.map(document => documentFurniture(plainTitle(document.title), document.furniture, request))
  return {
    issues: [
      ...(!acting.length ? [ACTING_FOR_MISSING] : ungrounded ? [actingForUngrounded(party)] : []),
      ...documents.flatMap((document, index) => forms[index] ? [] : [formIssue(plainTitle(document.title), document.form)]),
      ...furniture.flatMap(row => row.issues),
      ...documents.flatMap(document => titleIssues(document.title)),
      ...files.flatMap(file => file.issues),
      ...(repeatedFiles.length ? [`These file names are given to more than one document: ${repeatedFiles.map(name => `"${name}"`).join(", ")}. Give each document its own file name.`] : []),
      ...memo.issues,
      ...repeated.map(title => memoRepeated(title, memo.name!)),
      ...combinedIssues,
    ],
    value: {
      documents: documents.map((document, index) => ({
        title: plainTitle(document.title),
        purpose: document.purpose.join(" "),
        ...(document.playbook_review ? { playbook_review: document.playbook_review } : {}),
        // Planning labels are not the document's headings.
        parts: document.parts.map(part => plannedPart(part, !document.headings)),
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
      })),
      acting_for: party !== undefined && !ungrounded ? { party, stated: true } : NOT_STATED,
      ...(memoName ? { memo: memoName } : {}),
    },
  }
}

/** A planned part; with `unheaded`, a part without its own "Display heading:"
 * line is unheaded. */
function plannedPart(part: { heading: string; scope: string[]; display_heading?: string }, unheaded: boolean): PlannedPart {
  const display = part.display_heading ?? (unheaded ? "" : undefined)
  return { heading: part.heading, scope: part.scope.join(" "), ...(display !== undefined ? { display_heading: display } : {}) }
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

function titleIssues(title: string) {
  if (FILE_IN_TITLE.test(title))
    return [`The title "${title}" carries a file name. Write the document's own title on the "# " line, and give the file name the request asks for on a "File:" line under it.`]
  if (title.includes("`")) return [`The title "${title}" contains Markdown code marks. Write the title as plain text.`]
  return []
}

/** The title as plain text: Markdown code marks are not part of it. */
function plainTitle(title: string) {
  return title.replaceAll("`", "").replace(/\s+/g, " ").trim()
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
 * be part of a file name. List markers, bold and code marks, wrapping quotes,
 * trailing full stops and one leading article are not part of the name; only a
 * value that is wholly a non-answer is no memo. */
function outlineMemo(line: string | undefined, request: string | undefined): { name?: string; issues: string[] } {
  const name = (line ?? "").trim().replace(/^[-*]\s+/, "").replace(/\*\*|`/g, "").trim().replace(/^["'“‘]+|["'”’.]+$/g, "").trim()
    .replace(/^(?:a|an|the)\s+/i, "").replace(/\s+/g, " ")
  if (!name || MEMO_NON_ANSWER.test(name)) return { issues: [] }
  if (FILE_NAME_FORBIDDEN.test(name) || FILE_IN_TITLE.test(name) || name.length > MEMO_NAME_MAX)
    return { issues: [`The "Memo:" line gives "${name}", which cannot be part of a file name. Name the memo in a few of the request's own words, without a folder, file extension, Markdown or any of \\ / : * ? " < > |.`] }
  if (request !== undefined && !namedIn(request, name))
    return { issues: [`The "Memo:" line names "${name}", but the request does not use those words. Write a "Memo:" line only when the request asks for a memo or note to accompany the documents, naming it in the request's own words; otherwise leave the line out.`] }
  return { name, issues: [] }
}

/** Whether a planned title is the requested memo itself, not merely the subject
 * the memo is about. A memo name reads "<what it is> <connective> <its subject>",
 * so the words after its first connective ("on", "for", "to", "about"...) name
 * what it covers. A title whose words all sit in that subject names the memo's
 * subject, not the memo. Otherwise the title is the memo when the content words
 * of one are all among the other's, ignoring case, punctuation and function
 * words; a memo name of one word does not make a longer title a repeat. */
function repeatsMemo(title: string, memo: string) {
  const words = (tokens: string[]) => new Set(tokens.filter(word => word && !FUNCTION_WORDS.has(word)))
  const tokens = (text: string) => text.toLowerCase().split(/[^\p{L}\p{N}]+/u)
  const memoTokens = tokens(memo)
  const at = memoTokens.findIndex(word => MEMO_CONNECTIVES.has(word))
  const subject = words(at < 0 ? [] : memoTokens.slice(at + 1))
  const [planned, named] = [words(tokens(title)), words(memoTokens)]
  const within = (some: Set<string>, all: Set<string>) => some.size > 0 && [...some].every(word => all.has(word))
  if (within(planned, subject)) return false
  return within(planned, named) || (named.size > 1 && within(named, planned))
}

const MEMO_CONNECTIVES = new Set(["on", "for", "to", "about", "regarding", "concerning", "re", "of", "in", "with", "from", "at", "by"])
const FUNCTION_WORDS = new Set(["a", "an", "the", "to", "for", "of", "on", "and", "with", "that", "in"])
const memoRepeated = (title: string, memo: string) =>
  `The document "${title}" repeats the requested memo "${memo}", which the run writes as its notes. ` +
  "Remove that document, or, if the request asks for a separate memo with its own audience or purpose, leave out the \"Memo:\" line."
const MEMO_NAME_MAX = 60
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
  "name that party in the request's own words; otherwise write \"Acting for: not stated\". Do not infer the party from the sources."

function duplicates(names: string[]) {
  const seen = names.map(normalizedHeading)
  return [...new Set(names.filter((_, index) => seen.indexOf(seen[index]!) !== index))]
}

function isPlannedDocument(value: unknown): value is PlannedDocument {
  if (typeof value !== "object" || value === null) return false
  const row = value as PlannedDocument
  return typeof row.id === "string" && /^d[1-9]\d*$/.test(row.id) && typeof row.title === "string" && typeof row.purpose === "string" &&
    Array.isArray(row.parts) && row.parts.length > 0 && row.parts.every(part => typeof part?.heading === "string" && typeof part.scope === "string" && (part.display_heading === undefined || typeof part.display_heading === "string")) &&
    (row.playbook_review === undefined || isPlaybookReview(row.playbook_review)) &&
    (row.mode === undefined || ["new", "revise"].includes(row.mode)) && (row.deliverable === undefined || ["body", "redline", "combined"].includes(row.deliverable)) &&
    (row.comparison === undefined || typeof row.comparison === "string") && (row.template === undefined || typeof row.template === "string") && (row.file === undefined || typeof row.file === "string") &&
    (row.executed_by === undefined || (Array.isArray(row.executed_by) && row.executed_by.every(party => typeof party === "string"))) &&
    (row.form === undefined || DOCUMENT_FORMS.includes(row.form)) && (row.furniture === undefined || isPlannedFurniture(row.furniture)) &&
    (row.left_out === undefined || (Array.isArray(row.left_out) && row.left_out.every(item => typeof item?.part === "string" && typeof item.reason === "string")))
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
