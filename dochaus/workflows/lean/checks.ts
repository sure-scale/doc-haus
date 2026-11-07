import { createHash } from "node:crypto"
import type { AttachedDocxComment } from "../../lib/docx-comments"
import { citationRowsForLeanText } from "../../lib/citation-extraction"
import { amendmentText, type Amendment } from "../../lib/drafting-amendment"
import type { AppliedAmendment, UnappliedAmendment } from "../../lib/drafting-amendment-apply"
import { openItemMarkers } from "../../lib/open-items"
import { draftingPlaceholders } from "../../lib/placeholder-resolution"
import { retypedListLabel } from "../../lib/list-label"
import { normalize } from "../../lib/quote-verify"
import { changeListIDs, memoMode, mintedQuestionID, runItemIDs, withoutInternalIDs, withoutRefMarkers } from "./delivered-memo"
import { notesForReader } from "./gate"
import { currentInstructions } from "./instructions"
import { excerpt, openItem, proseQuoteIssues, quoteMissItems } from "./quotes"
import type { SourceQuoteScanOptions } from "./source-quote-scan"
import { readLeanPlan, type DocID } from "./plan"
import { researchFindingIDs } from "./research"
import { raisedFindings } from "./review-findings"
import { delegatedQuestions, readLeanJson, type OpenItem, type Question, type StageID } from "./state"

// Deterministic checks for the lean drafting stages. Each check returns retry
// feedback (`issues`, shown verbatim to the model on its next attempt) and legal
// open items (never retried; settled only at the final gate). The checks route on
// structure and on verification against the matter's canonical source text; they
// never judge or rewrite the legal substance.
// The change list is admitted in change-checks.ts.

export type StageCheck = { issues: string[]; openItems: OpenItem[] }

/** One body paragraph: `text` is its own text, which `sha256` hashes and
 * amendments replace; `label` is the number Word's list numbering prints
 * before it ("1.1", "(a)"), which is not part of that text. */
export type ParagraphRow = { pid: string; sha256: string; text: string; label?: string; clause?: true; draft_part?: number | null; draft_heading?: boolean }

// A paragraph's list label is shown before its text (shownText). Amendment
// text drops a leading copy of that label (withoutListLabel), also just inside
// an opening bracket (change-checks.ts): the list supplies the number. A
// different number typed in its place is not dropped; retypedLabelIssues
// names it for a retry, as it would print beside the label.
export { shownText, withoutListLabel } from "../../lib/list-label"

/** Retry feedback for replacement text `text`, at `at`, that opens with a
 * number written like its paragraph's list label but not it ("3.2 " on a
 * paragraph Word numbers 2.2), unless `current`, the paragraph's present
 * text, opens with that number. */
export function retypedLabelIssues(at: string, text: string, label: string | undefined, current?: string) {
  const typed = retypedListLabel(text, label, current)
  return typed ? [`${at} begins with "${typed}", but its paragraph's number ${label} is Word list numbering, which prints and renumbers itself. Leave the number out of the text.`] : []
}

export type CommentExport = {
  author: string
  attached: (AttachedDocxComment & { text: string })[]
  unresolved: { id: string; reason: string; text: string }[]
}

export type ApplyResult = { applied: AppliedAmendment[]; unapplied: UnappliedAmendment[]; comments?: CommentExport }

export type ChangeExplanation = { id: string; pid: string; reason: string; sources: { source_id: string; quote: string }[] }

export type ChangeList = {
  /** Source-adjudicated withdrawal of an earlier operation; initial wording is retained without a no-op body edit. */
  retained_initial?: Omit<ChangeExplanation, "id">[]
  /** Reasons for source-to-initial-draft revisions; these never mutate the body. */
  explanations?: ChangeExplanation[]
  amendments: Amendment[]
  /** Amendments that still failed admission on the last allowed attempt: a
   * quotation not found in its source, or markers that do not show the lawyer
   * this change. They are not applied; each is an open item for the drafting notes. */
  withheld?: { amendment: Amendment; issues: string[] }[]
  questions: Question[]
}

export const TOOL_SYNTAX_ISSUE =
  "The reply contains tool-call syntax as text. Call tools through the tool interface, then reply with the requested output only."

export const PID_REFERENCE = /\[(p[0-9a-f]{12})\]/g
const PRINTED_TOOL_CALL = [
  /<\/?\s*(?:tool_call|tool_use|function_call|function_calls|invoke)\b[^>]*>/i,
  /^\s*(?:```(?:json)?\s*)?\{\s*"(?:tool|tool_name|function|name)"\s*:\s*"[^"]+"\s*,\s*"(?:args|arguments|input|parameters)"\s*:/,
]

export function printedToolCall(text: string) {
  return PRINTED_TOOL_CALL.some(pattern => pattern.test(text))
}

export function checkDraft(directory: string, content: string): StageCheck {
  const issues = documentShapeIssues(content, "document")
  if (issues.length) return { issues, openItems: [] }
  return { issues, openItems: bodyOpenItems(directory, "draft", content) }
}

/** Settles the applier's report against the change list. Every requested
 * amendment ends with exactly one outcome: a missing outcome, or a run that
 * produced no amended document, leaves the amendment unapplied (apply_failed)
 * and visible as an open item. Nothing here fails the stage. */
export function settleApply(changes: ChangeList, result: ApplyResult | undefined, outputExists: boolean, failure?: string) {
  const bodyUnchanged = !result || !outputExists
  const outcomes = new Map<string, AppliedAmendment | UnappliedAmendment>()
  if (result && !bodyUnchanged)
    [...result.applied, ...result.unapplied].forEach(row => outcomes.has(row.id) || outcomes.set(row.id, row))
  const issue = !result ? `The applier stopped: ${failure ?? "unknown error"}.`
    : !outputExists ? "The applier did not write the amended document." : "The applier reported no outcome for this amendment."
  const settled = changes.amendments.map(amendment => outcomes.get(amendment.id) ?? {
    id: amendment.id, pid: amendment.pid, reason: "apply_failed" as const,
    detail: { target: amendmentText(amendment), reason: amendment.reason, issue },
  })
  const applied = settled.filter((row): row is AppliedAmendment => "mode" in row)
  const unapplied = settled.filter((row): row is UnappliedAmendment => !("mode" in row))
  const amendments = new Map(changes.amendments.map(amendment => [amendment.id, amendment]))
  return {
    result: { applied, unapplied } satisfies ApplyResult,
    bodyUnchanged,
    openItems: unapplied.map((row): OpenItem => ({
      ...openItem("apply", "unapplied_amendment", [
        `Amendment ${row.id} to paragraph [${row.pid}] was not applied (${row.reason})${row.detail?.issue ? `: ${row.detail.issue}` : "."}`,
        `Proposed text: "${excerpt(row.detail?.target ?? (amendments.get(row.id) ? amendmentText(amendments.get(row.id)!) : ""))}"`,
        `Reason given: ${excerpt(row.detail?.reason ?? amendments.get(row.id)?.reason ?? "")} The paragraph does not include this change.`,
      ].join(" "), { amendment_id: row.id, pid: row.pid }),
      ...(row.detail ? { applier_detail: Object.fromEntries(Object.entries(row.detail).filter(([, text]) => typeof text === "string")) } : {}),
    })),
  }
}

/** The drafting notes must disclose every open item: they say what it is and
 * end that statement with its id in a "[ref: …]" marker. The notes are
 * delivered with only those markers removed, so a run id, planned document id,
 * recorded change id or research finding id written anywhere else, which would
 * reach the reader, is retried. Their own quotes and citations must be supported: they cannot
 * disclose themselves, so a defect there is retried. So is a placeholder in the
 * notes' own text, a gap no open item records. The notes may quote a body
 * placeholder exactly as its open item records it, and neither a "[ref: …]"
 * marker, an item's id, quoted text nor an editorial mark (EDITORIAL_MARK) is a
 * placeholder of their own. On the last attempt (`final`) such an id or
 * placeholder alone does not fail the notes: they are saved as the model wrote
 * them, never rewritten, and every other check still applies; delivery keeps
 * the ids from the reader (deliveredNotes). Without a requested memo an item's
 * whole id anywhere in them then discloses it, as the gate reads them; a memo
 * still needs its marker. Before the last attempt the notes must also keep a
 * heading once delivery has removed their markers and ids (notesForReader);
 * notes saved without one are delivered under a plain title (deliveredNotes),
 * so they are never delivered empty. */
export function checkDraftingNotes(directory: string, content: string, items: OpenItem[], options: { final?: boolean } = {}): StageCheck {
  const shape = documentShapeIssues(content, "drafting notes")
  if (shape.length) return { issues: shape, openItems: [] }
  if (!options.final && !/^#{1,6}[ \t]+\S/m.test(notesForReader(directory, content, items)))
    return { issues: ['The drafting notes have no heading once their "[ref: …]" markers and internal ids are removed. Begin with a "# " title in words, and say each point in words before the marker that ends it.'], openItems: [] }
  const memo = memoMode(directory)
  const markers = [...content.matchAll(REF_MARKER)].map(match => match[1]!)
  const undisclosed = items.filter(item => !markers.some(marker => standingID(item.id).test(marker)) && !(options.final && !memo && standingID(item.id).test(content)))
    .map(item => item.id)
  // A marker holding only an id's last part ("[ref: 798598f814]") names no
  // item; the retry shows the whole marker to copy.
  const shortened = undisclosed.filter(id => id.includes(":") && markers.some(marker => standingID(id.slice(id.lastIndexOf(":") + 1)).test(marker)))
  // Each delegated question is listed (delegatedDecisionsText); on the last
  // attempt notes that miss one are saved as written.
  const unlisted = options.final ? [] : delegatedQuestions(directory).map(question => question.id).filter(id => !markers.some(marker => standingID(id).test(marker)))
  const delivered = withoutRefMarkers(content)
  const ids = !options.final
  const stray = ids ? runItemIDs(directory, items).filter(id => readerVisibleID(id) && standingID(id).test(delivered)) : []
  // A planned document's short id ("d1") is internal too: the reader knows the
  // document by its title or file name.
  // Quoted text, a citation, a footnote marker and a locator's number are not
  // prose (NOT_PROSE): only a standalone id there is a stray.
  const prose = NOT_PROSE.reduce((text, mask) => text.replace(mask, " "), delivered)
  const strayDocuments = ids ? (readLeanPlan(directory)?.documents ?? []).map(row => row.id).filter(id => standingID(id).test(prose)) : []
  // So is a recorded change id ("a1"): the reader knows a change by its effect.
  const strayChanges = ids ? runChangeIDs(directory, items).filter(id => standingID(id).test(prose)) : []
  // So is a research finding id ("r1"): the reader knows an authority by its
  // title or citation.
  const strayResearch = ids ? researchFindingIDs(directory).filter(id => standingID(id).test(prose)) : []
  const recorded = new Set(items.flatMap(item => item.ref.placeholder ? [item.ref.placeholder.toLowerCase()] : []))
  const own = items.reduce((text, item) => text.replaceAll(item.id, " "), delivered).replace(QUOTED, " ")
  const gaps = options.final ? [] : textPlaceholders(own)
    .filter(placeholder => !recorded.has(placeholder.toLowerCase()) && !EDITORIAL_MARK.test(placeholder))
  const misquoted = proseQuoteIssues(directory, "The drafting notes", content)
  const unsupported = [...new Set(citationRowsForLeanText(directory, "drafting-notes.md", content)
    .filter(row => row.status === "unsupported").map(row => row.citation_text))]
  const reader = memo ? "The memo is" : "The notes are"
  // The run's own record vocabulary in the notes' prose is workflow text the
  // lawyer cannot act on; on the last attempt the notes are saved as written.
  const workflow = options.final ? [] : workflowTerms(delivered.replace(QUOTED, " "))
  return {
    issues: [
      ...(undisclosed.length
        ? [`The drafting notes do not mention these open items: ${undisclosed.join(", ")}. Say what each is and what the lawyer needs to do about it, ending that statement with its exact id as "[ref: <id>]".`]
        : []),
      ...(shortened.length
        ? [`These markers hold only the last part of an id: ${shortened.map(id => `"[ref: ${id.slice(id.lastIndexOf(":") + 1)}]" must be "[ref: ${id}]"`).join("; ")}. Copy each marker whole, exactly as open_items lists it.`]
        : []),
      ...(unlisted.length
        ? [`The drafting notes do not list these delegated questions: ${unlisted.join(", ")}. For each, say in words what was decided and why, or why the point stays open, ending with its exact id as "[ref: <id>]".`]
        : []),
      ...(stray.length
        ? [`These internal ids appear outside a "[ref: …]" marker: ${stray.join(", ")}. ${reader} for the reader: describe each point in words and keep each id only in the marker that ends its statement.`]
        : []),
      ...(strayDocuments.length
        ? [`These document ids appear outside a "[ref: …]" marker: ${strayDocuments.join(", ")}. ${reader} for the reader: name each document by its title or file name, not its id.`]
        : []),
      ...(strayChanges.length
        ? [`These change ids appear outside a "[ref: …]" marker: ${strayChanges.join(", ")}. ${reader} for the reader: describe each change by what it does, not its id.`]
        : []),
      ...(strayResearch.length
        ? [`These research finding ids appear outside a "[ref: …]" marker: ${strayResearch.join(", ")}. ${reader} for the reader: name each authority by its title or approved citation, not its research id.`]
        : []),
      ...(gaps.length
        ? [`The drafting notes leave these placeholders in their own text: ${gaps.join(", ")}. State the point in words, or quote the body's placeholder exactly as its open item records it.`]
        : []),
      ...(workflow.length
        ? [`The drafting notes describe the drafting workflow rather than the documents: ${workflow.map(term => `"${term}"`).join(", ")}. ${reader} for the lawyer: say what the delivered documents and their redlines contain, which changes were not made and what remains to do, in plain words, without the workflow's record names, counts or steps.`]
        : []),
      ...misquoted,
      ...unsupported.map(citation => `The drafting notes cite "${citation}", which no matter source, research row or verified authority supports. Remove the citation or describe the point as needing research.`),
    ],
    openItems: [],
  }
}

/** The run's own vocabulary: phrases its records and stage prompts use for
 * its own steps, and a record field name (lower-case snake_case, not part of
 * a file name). None is how a lawyer describes a document or a change to it.
 * Code spans and source-marker names are not read: they name, not describe. */
const WORKFLOW_PHRASES = /\bthis run(?:['’]s)?(?![\w-])|\b(?:initial )?generated draft\b|\breview amendments?\b|\bunapplied (?:review )?(?:amendments?|changes?)\b|\bwithheld[- ]proposals?\b|\bpackage review\b/giu
const RECORD_NAME = /(?<![\w./\\@:-])[a-z]+(?:_[a-z]+)+(?![\w./\\@:-])/gu
const NAMED = /`[^`\n]*`|[[(][ \t]*source[ \t]*:[^\])\n]*[\])]/giu

function workflowTerms(text: string) {
  const read = text.replace(NAMED, " ")
  const terms = [...read.matchAll(WORKFLOW_PHRASES), ...read.matchAll(RECORD_NAME)].toSorted((a, b) => a.index - b.index)
  return [...new Set(terms.map(match => match[0].toLowerCase()))].slice(0, 6)
}

/** Text in straight or curly double quotes. */
const QUOTED = /"[^"\n]*"|\u201c[^\u201d\n]*\u201d/g

/** An editorial mark in quoted or cited text records an edit already made, not
 * text to supply: a bracket holding only "sic" or an edit's past participle
 * ("added", "inserted", "supplied", "omitted"), optionally after "emphasis". */
const EDITORIAL_MARK = /^\[\s*(?:(?:emphasis\s+)?(?:added|inserted|supplied|omitted)|sic)\s*\]$/i

/** Text the notes check does not read as prose: quoted text (the notes' verified
 * quotes), a footnote marker "[^d1]", a citation such as "3(d1)" or
 * "3(1)(a1)", and the number a locator word gives. That number starts with a
 * digit ("clause 12.1(d)", "Section 21A", "para. 3", "reg. 4(2)(b)"), or is in
 * parentheses after a sub-division ("Sub-paragraph (a1)"). Any other token
 * after a locator word is prose: "the clause a1 adds", "the new clause (a1)"
 * and "Part d2" name an internal id. */
const NOT_PROSE = [QUOTED, /\[\^[^\]\n]*\]/g, /(?<=[A-Za-z0-9)]\()[^)\n]*\)/g,
  /(?<![\w'\u2019])(?:(?:section|regulation|paragraph|clause|schedule|article|part|rule)s?\b|(?:ss?|regs?|paras?|cl|sch|arts?)\.)[ \t]*(?:\(?\d[^\s(),;]*\)?|(?<=sub-[a-z]+[ \t]*)\([^\s(),;]+\))/gi]

/** The text NOT_PROSE masks other than quotations: delivery keeps it as
 * written, so a citation's number is never read as an id. */
export const CITED = NOT_PROSE.slice(1)

/** A "[ref: …]" or "[refs: …]" marker in any case, as withoutRefMarkers
 * (delivered-memo.ts) removes it; group 1 is its ids, one or several separated
 * by "," or ";". */
const REF_MARKER = /\[[ \t]*refs?[ \t]*:([^\]\n]*)\]/gi

/** The ids the notes check keeps out of the reader's text: each run id no
 * lawyer would write as prose (readerVisibleID), each planned document's id,
 * each recorded change id, each review finding id and each research finding
 * id. Delivery removes the same ids (notesForReader). */
export function readerIDs(directory: string, items: OpenItem[]) {
  return [...new Set([...runItemIDs(directory, items).filter(readerVisibleID), ...(readLeanPlan(directory)?.documents ?? []).map(row => row.id), ...runChangeIDs(directory, items),
    ...raisedFindings(directory).map(row => row.id), ...researchFindingIDs(directory)])]
}

/** The ids lawyer-facing text a model writes (a question, its context, advice)
 * may not name: the ids readerIDs keeps from the notes and each instruction id,
 * matched as the notes check matches them (a short id only in prose). */
function lawyerTextIDs(directory: string) {
  const ids = [...new Set([...readerIDs(directory, []), ...currentInstructions(directory).map(item => item.id)])]
  return (text: string) => {
    const prose = NOT_PROSE.reduce((masked, mask) => masked.replace(mask, " "), text)
    return ids.filter(id => standingID(id).test(readerVisibleID(id) ? text : prose))
  }
}

/** Retry feedback for each run id a lawyer-facing field names, as
 * [where, text] pairs: the id found and how to say it in words. */
export function lawyerTextIDIssues(directory: string, fields: [where: string, text: string][]) {
  const named = lawyerTextIDs(directory)
  return fields.flatMap(([where, text]) => named(text).map(id => `${where} names the internal id "${id}". The lawyer sees only words: ${
    /^p[0-9a-f]{12}$/.test(id) ? "name the paragraph by its heading or quote its words"
      : /^d\d+$/.test(id) ? "name the document by its title"
      : mintedQuestionID(id) ? "name the question or answer by what it asked"
      : "say what it refers to in plain words"}.`))
}

/** `text` as the lawyer may see it, for a field saved on a last attempt: each
 * id lawyerTextIDIssues reports goes as the notes' ids go for the reader
 * (notesForReader), an instruction id by its kind (withoutInternalIDs). Text
 * naming none is returned as written. */
export function inLawyerWords(directory: string, text: string) {
  if (!lawyerTextIDs(directory)(text).length) return text
  return withoutInternalIDs(notesForReader(directory, text, []), currentInstructions(directory).map(item => item.id))
}

/** A question's two lawyer-facing fields. */
type LawyerQuestion = { question: string; context: string }

/** Retry feedback for each internal id a question's asking or context names
 * (lawyerTextIDIssues), at the question's index. */
export function questionWordIssues(directory: string, questions: (LawyerQuestion | undefined)[]) {
  return lawyerTextIDIssues(directory, questions.flatMap((question, index) => question
    ? [[`questions[${index}].question`, question.question], [`questions[${index}].context`, question.context]] as [string, string][] : []))
}

/** A question saved on a last attempt with ids still in it: its asking and
 * context in the lawyer's words (inLawyerWords). */
export function inQuestionWords<T extends LawyerQuestion>(directory: string, question: T): T {
  return { ...question, question: inLawyerWords(directory, question.question), context: inLawyerWords(directory, question.context) }
}

/** The change ids the run records for its planned documents (each change
 * list's amendments, withheld amendments and explanations) and the amendment
 * each open item names. Only these are matched, so a list label "(a)" or other
 * text of the same shape is not. */
function runChangeIDs(directory: string, items: OpenItem[]) {
  return [...new Set([
    ...(readLeanPlan(directory)?.documents ?? []).flatMap(({ id }) => changeListIDs(directory, id)),
    ...items.map(item => item.ref.amendment_id),
  ].filter((id): id is string => typeof id === "string" && id.length > 0))]
}

/** A whole occurrence of `id`, case-sensitive and not part of a longer id, word
 * or path, with the boundaries withoutInternalIDs (delivered-memo.ts) uses. */
function standingID(id: string) {
  return new RegExp(standingIDPattern(id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")))
}

/** The pattern of a whole occurrence of one of `ids`, an alternation of
 * escaped ids, as standingID bounds it. An id joined by "/", "-" or an en dash
 * to a short id such as "d2" or "a3" ("d1/d2", "a1-a3") still stands alone. */
export function standingIDPattern(ids: string) {
  const edge = String.raw`(?<![\w:./-])`, end = String.raw`(?![\w-]|[.:/]\w)`, joint = String.raw`[/\u2013-]`
  return String.raw`(?:${edge}|(?<=${edge}(?:[a-z]\d+${joint})+))(?:${ids})(?:${end}|(?=(?:${joint}[a-z]\d+)+${end}))`
}

/** A run id no lawyer would write as prose: at least eight characters with a
 * digit and no space, joined by ":" or "-" (open item and checklist ids, which
 * carry a hash), a paragraph pointer, or a question id the run minted, whose
 * hash may hold no digit. Short labels such as a document or amendment id
 * ("d1", "a1", "w1", "i1") and other word-like ids with no digit are not
 * counted here; the notes' planned document ids, recorded change ids and
 * research finding ids are checked on their own. */
function readerVisibleID(id: string) {
  return mintedQuestionID(id) || id.length >= 8 && /\d/.test(id) && !/\s/.test(id) && (/[:-]/.test(id) || /^p[0-9a-f]{12}$/.test(id))
}

/** A planned document's final body, as its apply step left it. */
export function finalBodyParagraphs(directory: string, document: DocID): ParagraphRow[] {
  return readParagraphs(directory, document, "paragraphs-applied.json")
}

/** The final-body text a review finding is about, as a hash: its paragraph's
 * sha256, or one over its document's body when it names no paragraph. None
 * for a finding on no document or on text the final body does not hold. */
export function itemTextSha(directory: string, item: OpenItem) {
  if (!item.ref.document) return
  const rows = finalBodyParagraphs(directory, item.ref.document)
  if (item.ref.pid) return rows.find(row => row.pid === item.ref.pid)?.sha256
  return rows.length ? createHash("sha256").update(rows.map(row => row.sha256).join("\n")).digest("hex") : undefined
}

/** The item with the text it was found on (itemTextSha). */
export function withTextSha(directory: string, item: OpenItem): OpenItem {
  const text_sha = itemTextSha(directory, item)
  return text_sha ? { ...item, text_sha } : item
}

/** Whether a recorded finding's text is still the final body's: a finding a
 * later review does not repeat stands while the text it was found on does.
 * One recorded on no final-body text (on no document, or on a paragraph
 * not in the final body) stands: disclosing it again is the safe failure. */
export function findingTextUnchanged(directory: string, item: OpenItem) {
  return item.text_sha === undefined || item.text_sha === itemTextSha(directory, item)
}

export function readParagraphs(directory: string, document: DocID, name = "paragraphs.json"): ParagraphRow[] {
  const value = readLeanJson(directory, `${document}/${name}`)
  return Array.isArray(value) ? value.filter(isParagraphRow) : []
}

export function readChangeList(directory: string, document: DocID): ChangeList | undefined {
  const value = readLeanJson(directory, `${document}/changes.json`)
  if (!record(value) || !Array.isArray(value.amendments) || !Array.isArray(value.questions)) return
  return value as ChangeList
}

/** Open items for the final body: each finding is located to the paragraph
 * that carries it, so the gate can check that paragraph for a visible marker. */
export function finalBodyItems(directory: string, rows: ParagraphRow[]) {
  const holding = (ref: OpenItem["ref"]) => {
    const needle = normalize(ref.placeholder ?? ref.citation ?? ref.quote ?? ref.statement ?? "")
    return needle ? { pids: paragraphTextHolders(rows, needle), whole: rows.some(row => normalize(row.text).includes(needle)) }
      : { pids: [], whole: false }
  }
  // The same marker in several paragraphs is one open point in each, disclosed
  // and settled where it stands. Blank ___ fields and other findings are
  // located at their first paragraph.
  return bodyOpenItems(directory, "apply", rows.map(row => row.text).join("\n\n"), { format: "plain" }).flatMap(item => {
    const { pids, whole } = holding(item.ref)
    // A quotation crossing Word paragraphs is disclosed in the last affected
    // paragraph, where its closing mark/source citation appears. The gate uses
    // the same joined-text lookup to keep the complete quotation in scope.
    const located = item.kind === "placeholder" && !/^_+$/.test(item.ref.placeholder ?? "") ? pids
      : item.kind === "quote_miss" && !whole ? pids.slice(-1) : pids.slice(0, 1)
    return located.length
      ? located.map(pid => openItem("apply", item.kind, item.detail, { ...item.ref, pid }))
      : [openItem("apply", item.kind, item.detail, item.ref)]
  })
}

export function bodyOpenItems(directory: string, stage: StageID, content: string, options: SourceQuoteScanOptions = { format: "markdown" }) {
  const placeholders = textPlaceholders(content)
  const citations = [...new Map(citationRowsForLeanText(directory, `${stage}.md`, content)
    .filter(row => row.status === "unsupported").map(row => [row.normalized_key, row.citation_text])).values()]
  return [
    ...placeholders.map(placeholder => openItem(stage, "placeholder", `The body contains the placeholder ${placeholder}.`, { placeholder })),
    ...citations.map(citation => openItem(stage, "citation_unsupported",
      `No matter source, research row or verified authority supports the citation "${citation}".`, { citation })),
    ...quoteMissItems(directory, stage, content, {}, options),
  ]
}

/** Read-only locations in the actual paragraph sequence, including quotations
 * that span paragraph boundaries. These offsets are local to this check; no
 * new persisted locator or authority record is introduced. */
export function paragraphTextHolders(rows: { pid: string; text: string }[], text: string): string[] {
  const needle = normalize(text)
  if (!needle) return []
  let length = 0
  const spans = rows.flatMap(row => {
    const value = normalize(row.text)
    if (!value) return []
    const start = length
    length += value.length + 1
    return [{ pid: row.pid, text: value, start, end: length - 1 }]
  })
  const body = spans.map(row => row.text).join(" ")
  const holding = new Set<string>()
  for (let start = body.indexOf(needle); start >= 0; start = body.indexOf(needle, start + 1)) {
    const end = start + needle.length
    for (const span of spans) if (span.start < end && span.end > start) holding.add(span.pid)
  }
  return [...holding]
}

/** The drafting placeholders and open-item markers in `content`; a line of
 * underscores only (a signature or blank line) is none. */
function textPlaceholders(content: string) {
  const prose = content.split("\n").filter(line => !/^\s*_{3,}\s*$/.test(line)).join("\n")
  return [...new Set([...draftingPlaceholders(prose), ...openItemMarkers(prose)])]
}

function documentShapeIssues(content: string, noun: "document" | "drafting notes") {
  if (!content.trim()) return [`The reply was empty. Reply with the complete ${noun} in Markdown.`]
  return [
    ...(!/^#{1,6}\s+\S/m.test(content) ? [`The ${noun} ${noun === "document" ? "has" : "have"} no Markdown heading. Begin with a "# " title and use headings for its parts.`] : []),
    ...(printedToolCall(content) ? [TOOL_SYNTAX_ISSUE] : []),
  ]
}

function isParagraphRow(value: unknown): value is ParagraphRow {
  return record(value) && typeof value.pid === "string" && typeof value.sha256 === "string" && typeof value.text === "string"
    && (value.label === undefined || typeof value.label === "string") && (value.clause === undefined || value.clause === true)
}

export function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}

/** A fenced block of a reply, plain or labelled json in either case. */
const REPLY_FENCE = /```(?:json)?\s*([\s\S]*?)```/gi
const NOT_OBJECT = "The JSON value is not an object; return one JSON object."

/** The one JSON object of a stage reply, the reader every lean stage reply
 * goes through: the whole reply when it is JSON, or else its one fenced
 * block, or else the one object from the first "{" to the last "}" in
 * prose around it. A reply that could be read two ways is refused, never
 * guessed at: braces outside the fenced blocks, or an array (the whole
 * value is preferred, so an array holding an object is not taken for the
 * object). Several fenced blocks are a model's drafts of its one answer,
 * the last its final one: the last is taken when every block is an object
 * with the same keys as the last and no brace stands outside the fences;
 * any other several blocks are refused. The error is one or more sentences;
 * a syntax error says where the JSON breaks (jsonBreakNote). */
export function parseJsonObject(content: string): { ok: true; value: Record<string, unknown> } | { ok: false; error: string } {
  try {
    const value: unknown = JSON.parse(content.trim())
    return record(value) ? { ok: true, value } : { ok: false, error: NOT_OBJECT }
  } catch { /* A prose wrapper can still contain one complete JSON object. */ }
  const start = content.indexOf("{")
  const end = content.lastIndexOf("}")
  if (start < 0) return { ok: false, error: "No JSON object was found in the reply." }
  if (end < start) return { ok: false, error: jsonBreakNote(content) ?? "The reply ends before its JSON object is closed." }
  const blocks = [...content.matchAll(REPLY_FENCE)]
  const bracesOutside = /[{}]/.test(content.replace(REPLY_FENCE, ""))
  if (blocks.length > 1) {
    const values = blocks.map(block => { try { return JSON.parse(block[1]!.trim()) as unknown } catch { return undefined } })
    const last = values.at(-1)
    const shape = (value: unknown) => record(value) ? JSON.stringify(Object.keys(value).toSorted()) : undefined
    return record(last) && !bracesOutside && values.every(value => shape(value) === shape(last))
      ? { ok: true, value: last }
      : { ok: false, error: "The reply holds more than one fenced block; return one JSON object." }
  }
  const candidate = blocks[0]?.[1]?.trim() ?? content.slice(start, end + 1)
  if (!blocks.length && (content.slice(0, start).trimEnd().endsWith("[") || content.slice(end + 1).trimStart().startsWith("]")))
    return { ok: false, error: NOT_OBJECT }
  try {
    const value: unknown = JSON.parse(candidate)
    if (blocks.length && record(value) && bracesOutside)
      return { ok: false, error: "The reply has braces outside its fenced block, so its object is ambiguous; return one JSON object." }
    return record(value) ? { ok: true, value } : { ok: false, error: NOT_OBJECT }
  } catch (error) {
    return { ok: false, error: jsonBreakNote(content) ?? `The reply is not valid JSON: ${error instanceof Error ? error.message : "it could not be parsed"}.` }
  }
}

const JSON_SPACE = new Set([" ", "\t", "\n", "\r"])
const JSON_LITERAL = /true|false|null|-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/y

/** Retry feedback for a reply that parseJsonObject cannot read: where its
 * JSON first breaks, as a character offset in the reply, what was expected
 * there (such as the bracket that closes the list or object open at that
 * point, found by tracking nesting outside strings) and the words around it.
 * The JSON is read where parseJsonObject reads it: inside the one fenced
 * block, or else from the first "{" to the end of the reply, so a reply cut
 * off inside its object is placed at its end. A model told only the format
 * tends to resend the same long reply unchanged; the place lets it mend the
 * one slip. Undefined when no break is found, so the caller's format
 * restatement stands alone. */
export function jsonBreakNote(reply: string): string | undefined {
  const blocks = [...reply.matchAll(REPLY_FENCE)]
  const block = blocks.length === 1 ? blocks[0]! : undefined
  // A fenced block's JSON starts after its label and spaces: the match less the closing fence and the JSON.
  const start = block ? block.index + block[0].length - 3 - block[1]!.length : Math.max(0, reply.indexOf("{"))
  const body = block ? block[1]!.trimEnd() : reply.slice(start).trimEnd()
  const stack: { bracket: "{" | "["; at: number; name?: string }[] = []
  // What the last token was: nothing yet, an opening bracket, a comma, a colon, an object's name or a complete value.
  let last = "start" as "start" | "open" | "comma" | "colon" | "name" | "value"
  let name: string | undefined
  const where = (open: typeof stack[number]) => `the ${open.name === undefined ? "" : `"${open.name}" `}${open.bracket === "[" ? "list" : "object"} opened at character ${start + open.at}`
  const closer = (open: typeof stack[number]) => open.bracket === "[" ? "]" : "}"
  const note = (at: number, what: string) => {
    const words = (text: string) => text.replace(/\s+/g, " ")
    return `The reply is not valid JSON: at character ${start + at}, ${what}. ` +
      `Around it, with >>> marking the place: \`${words(body.slice(Math.max(0, at - 40), at))}>>>${words(body.slice(at, at + 20))}\`. Fix that and resend the complete JSON.`
  }
  for (let at = 0; at < body.length; at++) {
    const char = body[at]!
    if (JSON_SPACE.has(char)) continue
    const open = stack.at(-1)
    if (last === "value" && !open) return note(at, "the JSON value has already ended, so nothing may follow it")
    if (last === "name" && char !== ":") return note(at, `":" was expected after the name "${name}", but "${char}" was found`)
    if (char === '"') {
      if (last === "value") return note(at, `"," was expected between two values, but "${char}" was found`)
      let end = at + 1
      for (; end < body.length && body[end] !== '"'; end++) {
        if (body[end] === "\\") end++
        else if (body[end]! < " ") return note(end, "a line break or other control character was found inside a string; write it as an escape such as \\n")
      }
      if (end >= body.length) return note(at, "the string that starts here is never closed")
      name = body.slice(at + 1, end)
      last = open?.bracket === "{" && (last === "open" || last === "comma") ? "name" : "value"
      at = end
    } else if (char === ":") {
      if (last !== "name") return note(at, `":" was found where it cannot stand`)
      last = "colon"
    } else if (char === ",") {
      if (last !== "value") return note(at, `"," was found where a value was expected`)
      last = "comma"
    } else if (char === "{" || char === "[") {
      if (last === "value") return note(at, `"," was expected between two values, but "${char}" was found`)
      stack.push({ bracket: char, at, ...(last === "colon" ? { name } : {}) })
      last = "open"
    } else if (char === "}" || char === "]") {
      if (!open) return note(at, `"${char}" closes nothing`)
      if (char !== closer(open)) return note(at, `"${closer(open)}" was expected to close ${where(open)}, but "${char}" was found`)
      if (last === "comma") return note(at, `a comma stands before "${char}" with no value after it`)
      if (last === "colon") return note(at, `a value was expected after the name "${name}", but "${char}" was found`)
      stack.pop()
      last = "value"
    } else {
      JSON_LITERAL.lastIndex = at
      const literal = last === "value" ? null : JSON_LITERAL.exec(body)
      if (!literal) return note(at, last === "value" ? `"," was expected between two values, but "${char}" was found` : `"${char}" does not start a JSON value`)
      at += literal[0].length - 1
      last = "value"
    }
  }
  const open = stack.at(-1)
  if (open) return note(body.length, `the reply ends before "${closer(open)}" closes ${where(open)}`)
}
