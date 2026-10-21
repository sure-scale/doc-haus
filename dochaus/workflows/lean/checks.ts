import type { AttachedDocxComment } from "../../lib/docx-comments"
import { citationRowsForLeanText } from "../../lib/citation-extraction"
import { amendmentText, type Amendment } from "../../lib/drafting-amendment"
import type { AppliedAmendment, UnappliedAmendment } from "../../lib/drafting-amendment-apply"
import { openItemMarkers } from "../../lib/open-items"
import { draftingPlaceholders } from "../../lib/placeholder-resolution"
import { normalize } from "../../lib/quote-verify"
import { changeListIDs, memoMode, mintedQuestionID, runItemIDs, withoutRefMarkers } from "./delivered-memo"
import { notesForReader } from "./gate"
import { excerpt, openItem, proseQuoteFound, quoteMissIssue, quoteMissItems, sourceQuotes } from "./quotes"
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

export type ParagraphRow = { pid: string; sha256: string; text: string; draft_part?: number | null; draft_heading?: boolean }

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
  const misquoted = sourceQuotes(content).filter(quote => !proseQuoteFound(directory, quote))
  const unsupported = [...new Set(citationRowsForLeanText(directory, "drafting-notes.md", content)
    .filter(row => row.status === "unsupported").map(row => row.citation_text))]
  const reader = memo ? "The memo is" : "The notes are"
  return {
    issues: [
      ...(undisclosed.length
        ? [`The drafting notes do not mention these open items: ${undisclosed.join(", ")}. Say what each is and what the lawyer needs to do about it, ending that statement with its exact id as "[ref: <id>]".`]
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
      ...misquoted.map(quote => quoteMissIssue(directory, "The drafting notes", quote)),
      ...unsupported.map(citation => `The drafting notes cite "${citation}", which no matter source, research row or verified authority supports. Remove the citation or describe the point as needing research.`),
    ],
    openItems: [],
  }
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
    const needle = normalize(ref.placeholder ?? ref.citation ?? ref.quote ?? "")
    return needle ? rows.filter(row => normalize(row.text).includes(needle)).map(row => row.pid) : []
  }
  // The same marker in several paragraphs is one open point in each, disclosed
  // and settled where it stands. Blank ___ fields and other findings are
  // located at their first paragraph.
  return bodyOpenItems(directory, "apply", rows.map(row => row.text).join("\n\n")).flatMap(item => {
    const pids = holding(item.ref)
    const located = item.kind === "placeholder" && !/^_+$/.test(item.ref.placeholder ?? "") ? pids : pids.slice(0, 1)
    return located.length
      ? located.map(pid => openItem("apply", item.kind, item.detail, { ...item.ref, pid }))
      : [openItem("apply", item.kind, item.detail, item.ref)]
  })
}

export function bodyOpenItems(directory: string, stage: StageID, content: string) {
  const placeholders = textPlaceholders(content)
  const citations = [...new Map(citationRowsForLeanText(directory, `${stage}.md`, content)
    .filter(row => row.status === "unsupported").map(row => [row.normalized_key, row.citation_text])).values()]
  return [
    ...placeholders.map(placeholder => openItem(stage, "placeholder", `The body contains the placeholder ${placeholder}.`, { placeholder })),
    ...citations.map(citation => openItem(stage, "citation_unsupported",
      `No matter source, research row or verified authority supports the citation "${citation}".`, { citation })),
    ...quoteMissItems(directory, stage, content),
  ]
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
}

export function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
