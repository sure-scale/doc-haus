import { TOOL_SYNTAX_ISSUE, printedToolCall } from "./checks"
import { REVIEW_STAGES, readLeanJson, writeLeanText, type ReviewStageID, type StageID } from "./state"

// The run's plan: every document the request asks for, in drafting order, each
// with its top-level parts. It comes from the draft agent's outline reply (one
// "# " title per document, "## " parts under it); code assigns the document ids
// d1..dK in that order and saves it as plan.json. The walk of steps follows it:
// each document goes through draft, review, changes and apply, then one
// finalize and one gate cover the run. The plan is a drafting record, not an
// authority: it grants nothing, and nothing here knows what a document is for.
// It records the hash of the request and sources the outline was built from, so
// the outline is redone only when they change and a lawyer's edit is followed.
// It also records whom the lawyer acts for, as the outline read it from the
// request: every model stage is told, and a stage that is not told takes no side.
// Code never infers the party; a lawyer sets it by editing plan.json.

export const DOC_STAGES = ["draft", ...REVIEW_STAGES, "review-playbook", "change-plan", "change-prepare", "changes", "change-review", "apply"] as const
export type DocStage = (typeof DOC_STAGES)[number]
export type DocID = `d${number}`
export type StepID = `${DocID}.${DocStage}` | "discrepancy" | "intake" | "research" | "research-review" | "spine" | "finalize" | "checklist-check" | "package-review" | "gate"

/** heading identifies the drafting part; absent display_heading preserves it
 * in publication, while an empty display_heading makes the part unheaded. */
export type PlannedPart = { heading: string; scope: string; display_heading?: string }
/** A template part the outline leaves out on purpose, with its reason. */
export type LeftOutPart = { part: string; reason: string }
/** `template` names the matter document the document is drafted from, when the
 * outline gives one on a "Template:" line under the title; `left_out` lists the
 * template parts its "Left out:" lines leave out on purpose; `file` is the file
 * name the request asks the document to be delivered under, from a "File:" line.
 * `mode` distinguishes a revision foundation from a new document; `deliverable`
 * assigns the requested filename to the body or annotated redline. Legacy plans
 * without these fields retain their existing behaviour. The title is the
 * document's own title and never carries the file name. */
export type PlaybookReview = { sources: { source_id: string; quote: string }[]; reason: string }
export type PlannedDocument = {
  id: DocID; title: string; purpose: string; parts: PlannedPart[]; playbook_review?: PlaybookReview; mode?: "new" | "revise"; deliverable?: "body" | "redline"; template?: string; comparison?: string; left_out?: LeftOutPart[]; file?: string
}
export type ActingFor = { party: string; stated: boolean }
export type LeanPlan = { schema: "dochaus.lean.plan/1"; documents: PlannedDocument[]; acting_for?: ActingFor; basis?: string; notes_file?: string }
export type LeanOutline = { documents: Omit<PlannedDocument, "id">[]; acting_for: ActingFor }

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
  if (typeof value !== "object" || value === null || (value as LeanPlan).schema !== LEAN_PLAN_SCHEMA) return
  const documents = (value as LeanPlan).documents
  if (!Array.isArray(documents) || !documents.length || !documents.every(isPlannedDocument)) return
  return value as LeanPlan
}

export function writeLeanPlan(directory: string, documents: Omit<PlannedDocument, "id">[], basis?: string, actingFor?: ActingFor): LeanPlan {
  const plan: LeanPlan = {
    schema: LEAN_PLAN_SCHEMA, documents: documents.map((document, index) => ({ id: `d${index + 1}`, ...document })), acting_for: actingFor, basis,
  }
  writeLeanText(directory, "plan.json", `${JSON.stringify(plan, null, 2)}\n`)
  return plan
}

export function plannedDocument(directory: string, id: DocID): PlannedDocument | undefined {
  return readLeanPlan(directory)?.documents.find(document => document.id === id)
}

/** Output naming does not change operative drafting inputs. Consumers of the
 * full delivery plan (notes, package review and registration) retain it. */
export function operativeDocument(document: PlannedDocument | undefined) {
  if (!document) return null
  const { file, deliverable, ...operative } = document
  return operative
}

export function operativePlan(plan: LeanPlan | undefined) {
  if (!plan) return null
  const { notes_file, ...operative } = plan
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

/** The documents planned before `id`, in drafting order. */
export function earlierDocuments(directory: string, id: DocID): PlannedDocument[] {
  const documents = readLeanPlan(directory)?.documents ?? []
  const at = documents.findIndex(document => document.id === id)
  return at < 0 ? [] : documents.slice(0, at)
}

/** Reads the outline reply: an "Acting for:" line before the first title, then
 * each "# " line starts a document, the plain lines under it before its first
 * part are its purpose (a "Template:" line among them names the matter document
 * it is drafted from, a "File:" line the file name the request gives it), and
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
 * request does not give is dropped and the title kept without Markdown code marks. */
export function parsePlan(reply: string, request?: string): Checked<LeanOutline> {
  if (!reply.trim()) return { issues: [`The reply was empty. ${OUTLINE_SHAPE}`] }
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  const documents: {
    title: string; purpose: string[]; playbook_review?: PlaybookReview; mode?: "new" | "revise"; deliverable?: "body" | "redline"; template?: string; comparison?: string; file?: string; left_out: LeftOutPart[]; parts: { heading: string; scope: string[]; display_heading?: string }[]
  }[] = []
  const acting: string[] = []
  const selectionIssues: string[] = []
  for (const raw of reply.split(/\r?\n/)) {
    const line = raw.trim()
    const party = /^acting for\s*:\s*(.*)$/i.exec(line.replace(/^[-*]\s+/, "").replace(/[*`]/g, ""))?.[1]
    if (party !== undefined && !documents.length) {
      acting.push(party.replace(/[.\s]+$/, ""))
      continue
    }
    const title = /^#\s+(.+)$/.exec(line)?.[1]
    const heading = /^##\s+(.+)$/.exec(line)?.[1]
    const current = documents.at(-1)
    const plain = line.replace(/\*\*/g, "")
    const template = /^[-*]?\s*template\s*:\s*(.+)$/i.exec(plain)?.[1]?.replace(/^["'`“]+|["'`”]+$/g, "").trim()
    const comparison = /^[-*]?\s*compare\s+against\s*:\s*(.+)$/i.exec(plain)?.[1]?.replace(/^["'`“]+|["'`”]+$/g, "").trim()
    const mode = /^[-*]?\s*mode\s*:\s*(new|revise)\s*$/i.exec(plain)?.[1]?.toLowerCase() as "new" | "revise" | undefined
    const deliverable = /^[-*]?\s*deliverable\s*:\s*(body|redline)\s*$/i.exec(plain)?.[1]?.toLowerCase() as "body" | "redline" | undefined
    const leftOut = /^[-*]?\s*left\s+out\s*:\s*(.+)$/i.exec(plain)?.[1]
    const file = /^[-*]?\s*file(?:\s+name)?\s*:\s*(.+)$/i.exec(plain)?.[1]?.replace(/^["'`“]+|["'`”.]+$/g, "").trim()
    const playbook = /^[-*]?\s*playbook review\s*:\s*(.+)$/i.exec(line)?.[1]
    const displayHeading = /^display heading\s*:\s*(.*)$/i.exec(line)?.[1]
    if (playbook && current && !current.parts.length) {
      let value: unknown
      try { value = JSON.parse(playbook) } catch {}
      if (isPlaybookReview(value) && !current.playbook_review) current.playbook_review = value
      else selectionIssues.push(`In "${current.title}", Playbook review must be one JSON object with sources [{source_id,quote}] and a nonempty reason, once per document.`)
      continue
    }
    if (title) documents.push({ title: headingText(title), purpose: [], left_out: [], parts: [] })
    else if (heading && current) current.parts.push({ heading: headingText(heading), scope: [] })
    else if (displayHeading !== undefined && current?.parts.length) current.parts.at(-1)!.display_heading = /^none$/i.test(displayHeading.trim()) ? "" : headingText(displayHeading)
    else if (mode && current && !current.parts.length) current.mode = mode
    else if (deliverable && current && !current.parts.length) current.deliverable = deliverable
    else if (template && current && !current.parts.length) current.template = template
    else if (comparison && current && !current.parts.length) current.comparison = comparison
    else if (file && current && !current.parts.length) current.file = file
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
  return {
    issues: [
      ...(!acting.length ? [ACTING_FOR_MISSING] : ungrounded ? [actingForUngrounded(party)] : []),
      ...documents.flatMap(document => titleIssues(document.title)),
      ...files.flatMap(file => file.issues),
      ...(repeatedFiles.length ? [`These file names are given to more than one document: ${repeatedFiles.map(name => `"${name}"`).join(", ")}. Give each document its own file name.`] : []),
    ],
    value: {
      documents: documents.map((document, index) => ({
        title: plainTitle(document.title),
        purpose: document.purpose.join(" "),
        ...(document.playbook_review ? { playbook_review: document.playbook_review } : {}),
        parts: document.parts.map(part => ({ heading: part.heading, scope: part.scope.join(" "), ...(part.display_heading !== undefined ? { display_heading: part.display_heading } : {}) })),
        ...(document.mode ? { mode: document.mode } : {}),
        ...(document.deliverable ? { deliverable: document.deliverable } : {}),
        ...(document.comparison ? { comparison: document.comparison } : {}),
        ...(document.template ? { template: document.template } : {}),
        ...(document.left_out.length ? { left_out: document.left_out } : {}),
        ...(files[index]?.name && !repeatedFiles.includes(files[index].name) ? { file: files[index].name } : {}),
      })),
      acting_for: party !== undefined && !ungrounded ? { party, stated: true } : NOT_STATED,
    },
  }
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
  if (FILE_NAME_FORBIDDEN.test(file) || file.length > 120)
    return { issues: [`The "File:" line of "${title}" gives "${file}", which is not a plain file name. Give the file name only, as the request writes it, without a folder or Markdown.`] }
  if (request !== undefined && !namedIn(request, file))
    return { issues: [`The "File:" line of "${title}" gives "${file}", but the request gives no such file name. Add a "File:" line only for a file name the request gives, in its words.`] }
  return { name: file, issues: [] }
}

const OUTLINE_SHAPE = "Reply with the outline: for each document, one \"# \" title line, then one \"## \" line per part."
const ACTING_FOR_MISSING = "The outline has no \"Acting for:\" line. Before the first \"# \" title, write \"Acting for: <party>\" when the request says " +
  "which party the lawyer acts for, otherwise \"Acting for: not stated\"; then the complete outline."

/** The party an "Acting for:" line names, or undefined when the line gives a
 * non-answer. Wrapping brackets or quotes and a trailing explanation are not
 * part of the name. */
function statedParty(text: string) {
  const party = text.trim().replace(/\s+(?:\([^()]*\)|\[[^\]]*\])[.\s]*$/, "").replace(/\s*(?:;|\s[-\u2013\u2014]\s|[\u2013\u2014]).*$/, "")
    .replace(/^[[("'\u201C\u2018<]+|[\])"'\u201D\u2019>]+$/g, "").replace(/[.:\s]+$/, "").trim()
  if (!party || /^(?:not\s+(?:stated|specified|given|known|clear)|unknown|unspecified|unclear|none|n\/?a)\b/i.test(party)) return undefined
  return party
}

function namedIn(request: string, party: string) {
  const fold = (text: string) => text.normalize("NFC").replace(/[\u2018\u2019]/g, "'").replace(/\s+/g, " ").trim().toLowerCase()
  return fold(request).includes(fold(party).replace(/^the /, ""))
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
    (row.mode === undefined || ["new", "revise"].includes(row.mode)) && (row.deliverable === undefined || ["body", "redline"].includes(row.deliverable)) &&
    (row.comparison === undefined || typeof row.comparison === "string") && (row.template === undefined || typeof row.template === "string") && (row.file === undefined || typeof row.file === "string") &&
    (row.left_out === undefined || (Array.isArray(row.left_out) && row.left_out.every(item => typeof item?.part === "string" && typeof item.reason === "string")))
}

function isPlaybookReview(value: unknown): value is PlaybookReview {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false
  const row = value as PlaybookReview
  return typeof row.reason === "string" && !!row.reason.trim() && Array.isArray(row.sources) && row.sources.every(source =>
    typeof source?.source_id === "string" && !!source.source_id.trim() && typeof source.quote === "string" && !!source.quote.trim())
}
