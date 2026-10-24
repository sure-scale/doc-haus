import { amendmentParagraphs, type Amendment } from "../../lib/drafting-amendment"
import { normalizeLegend, normalizePreferences } from "../../lib/drafting-preferences"
import { furnitureLegend } from "../../lib/furniture-legend"
import { readFirmPreferences } from "../../lib/preferences"
import type { ParagraphRow } from "./checks"
import { plannedDocument, type DocID, type PlannedDocument } from "./plan"
import { excerpt, matterSourceIDs, openItem } from "./quotes"
import type { OpenItem } from "./state"
import { templateSource } from "./template-contract"
import { sourceText } from "./template-structure"

// Draft-status statements in an agreement's text (#286). An agreement's text
// holds its terms. Whether this version is a draft, under negotiation or
// approved for signature is the state of the text, not a term: the firm's draft
// line and legend carry it on every page from the firm's settings
// (lib/page-furniture.ts), and the points still open are open-item markers
// listed in the drafting notes. A statement of that state the model writes into
// the agreement is found here, in the draft (a required repair for the change
// cycle), in a proposed amendment (retry feedback) and in the final body (an
// open item the drafting notes disclose).
//
// Found only in a sentence about this text itself: a marking line (a short
// sentence that opens with the status wording: "Consultation draft — not for
// signature") or an assertion that names this text as a draft, document or
// version ("This draft is circulated for discussion only", "This is a working
// draft"). In such a sentence the status wording is: calling the text a draft
// for discussion or negotiation, not approved, ready or final for signature,
// not for signature, for discussion or negotiation only, or describing its own
// placeholders or bracketed wording. The same words in any other sentence are
// operative wording or name another document ("The Guarantee is not ready for
// signing until ...", "a Draft Status Report", "Clause 9 is a placeholder
// pending ...") and are never found: a missed status line costs the lawyer a
// deletion, a removed condition is a legal error. Not found either: whether the
// agreement binds ("subject to contract"), which is a term the parties may
// agree or a firm legend; quoted words; other bracketed text. An open-item
// marker says what is missing or to be confirmed, so each of its clauses (its
// sentences and the parts set off by a spaced dash) is checked like a sentence
// for the approval and readiness wording ("[open item: confirm approvals —
// draft not ready]"); its own placeholder or bracketed wording is what a marker
// is, so it is not status there. Kept: a sentence
// taken from the firm's legend for the document's form, from a legend the
// request asks for (both applied by code as page furniture) or, in a
// revision, from one paragraph of the lawyer's own original, so a marking the
// lawyer wrote is never taken out of their document.

const S = String.raw`\s+`
const HYPHEN = "[-‐‑–]"
const STATE = "(?:negotiation|discussion|working|consultation|revised|first|initial|clean|further|current)"

/** The text's approval or readiness state: a draft for discussion or
 * negotiation, not approved, ready or final for signature. */
const STATE_WORDING = [
  String.raw`\bnot${S}(?:(?:yet|been|to${S}be|being|currently)${S})*(?:approved|cleared|authori[sz]ed|released|ready|final(?:i[sz]ed)?)${S}for${S}(?:signature|signing|execution|completion|implementation|issue|release|circulation)\b`,
  String.raw`\bnot${S}for${S}(?:signature|signing|execution)\b`,
  String.raw`\bdraft${S}(?:(?:is|remains)${S})?not${S}(?:yet${S})?(?:ready|final(?:i[sz]ed)?|approved|agreed)\b`,
  String.raw`\bnegotiation${HYPHEN}only\b`,
  String.raw`\bfor${S}(?:discussion|negotiation)(?:${S}(?:and|or)${S}(?:discussion|negotiation))?${S}(?:purposes${S})?only\b`,
  String.raw`\b${STATE}(?:${HYPHEN}only)?${S}draft\b`,
  String.raw`\bdraft${S}(?:(?:issued|prepared|circulated|provided)${S})?for${S}(?:discussion|negotiation)\b`,
  String.raw`\b(?:issued|circulated|provided|prepared|sent|supplied|produced)${S}(?:(?:only|solely|purely)${S})?for${S}(?:the${S}purposes?${S}of${S})?(?:discussion|negotiation)\b`,
]

/** The text describing its own unfinished wording. */
const SELF_WORDING = [
  String.raw`\bbracketed${S}(?:wording|text|drafting|language|provisions?)\b`,
  String.raw`\bplaceholders?\b`,
]

/** The status wording, and a marking line's own opening: status wording, a
 * bare "Draft" set off by a dash, colon, bar or bracket, or a "Draft status"
 * heading. */
function wording(alternatives: string[]) {
  const status = new RegExp(alternatives.join("|"), "giu")
  const marking = new RegExp(String.raw`^(?:(?:${status.source})|draft(?=\s*$|\s+[-–—|]\s|\s*[:(])|draft${S}status$)`, "iu")
  return { status, marking }
}

/** Status wording, counted only in a sentence about this text, and in an
 * open-item marker's clause. */
const SENTENCE = wording([...STATE_WORDING, ...SELF_WORDING])
const MARKER = wording(STATE_WORDING)
const MARKING_WORDS = 14

const THIS_TEXT = String.raw`(?:this|the${S}present)${S}(?:${STATE}(?:${HYPHEN}only)?${S})*(?:draft|document|version)\b`

/** This text as the subject of a clause, not a reference such as "under this
 * document" modifying another subject. A status can also label the subject
 * itself ("this working draft"). A pronoun needs an explicit draft/version
 * classification: after sentence splitting, "it is not approved" could still
 * refer to another document mentioned in the previous clause. */
const CLASSIFIED_PRONOUN = String.raw`(?:it|this)(?=${S}(?:is|remains)${S}(?:(?:a|an|the|only)${S})*(?:[\p{L}-]+${S}){0,3}?(?:draft|version)\b)`
const OWN_SUBJECT = new RegExp(String.raw`(?:^|[,:(–—]\s*|\b(?:and|but|while|that)${S})(${THIS_TEXT}|${CLASSIFIED_PRONOUN})`, "giu")

/** Only linking words may connect the subject to a status. An intervening
 * object or operative predicate ("requires delivery of a ...") breaks that
 * relationship. Contents verbs apply only to unfinished-wording assertions. */
const LINK_WORD = "(?:is|are|was|remains|remain|will|shall|be|has|have|been|being|a|an|the|only|solely|purely|still|currently|yet|not|explicit|intended|as|issued|circulated|provided|prepared|sent|supplied|produced)"
const STATUS_LINK = new RegExp(String.raw`^\s*(?:${LINK_WORD}${S})*$`, "iu")
const CONTENT_LINK = new RegExp(String.raw`^\s*(?:(?:${LINK_WORD}|contains|includes|uses|retains)${S})*$`, "iu")
const COORDINATED_LINK = new RegExp(String.raw`(?:^|\b(?:and|or|but)${S})\s*(?:${LINK_WORD}${S})*$`, "iu")
const SELF_STATUS = new RegExp(SELF_WORDING.join("|"), "iu")
const OWN_CONTENT = new RegExp(String.raw`^(?:${S}(?:in|within|throughout|of)${S}${THIS_TEXT}|[^.;!?]*\b(?:as${S}(?:set${S}out|shown|marked)|appearing|contained)${S}in${S}${THIS_TEXT})`, "iu")

/** Markers, citations and quoted words carry no statement of the text's own. */
const NOT_STATEMENT = /\[[^\]\n]*\]|\(open item:[^)\n]*\)|"[^"\n]*"|“[^”\n]*”/giu

/** An open-item marker, in either bracket. */
const OPEN_ITEM = /\[open item:[^\]\n]*\]|\(open item:[^)\n]*\)/giu

/** Leading numbering, bullets and emphasis before a sentence's first word. */
const LEAD = /^[\s#>*_\-•–—]*(?:\(?(?:\d+(?:\.\d+)*|[a-z]|[ivxlc]+)[.)]\s*)?[\s*_]*/iu

/** Share of a sentence's words one kept paragraph must hold for the sentence
 * to be taken from it. */
const TAKEN_FROM = 0.6

export type DraftStatusStatement = { statement: string; sentence: string }

/** The statements of its own draft status or approval state in one paragraph
 * of an agreement, leaving out a sentence taken from one paragraph of `kept`
 * (a firm legend, the lawyer's original, the paragraph an amendment replaces)
 * that holds the same status wording. */
export function draftStatusStatements(text: string, kept: string[] = []): DraftStatusStatement[] {
  const keep = kept.flatMap(block => block.split(/\n+/)).map(flat).filter(words => words.trim())
  const statements = (raw: string, rules: typeof SENTENCE) => {
    const sentence = raw.replace(LEAD, "").trim()
    const marking = rules.marking.test(sentence) && wordsOf(sentence).length <= MARKING_WORDS
    if (!sentence) return []
    const spans = [...sentence.matchAll(rules.status)].map(match => ({ start: match.index, end: match.index + match[0].length, statement: match[0] }))
    if (marking && !spans.length) spans.push({ start: 0, end: sentence.length, statement: sentence })
    const subjects = [...sentence.matchAll(OWN_SUBJECT)].map(match => ({ start: match.index + match[0].length - match[1]!.length, end: match.index + match[0].length }))
    const own: typeof spans = []
    for (const span of spans) {
      const self = SELF_STATUS.test(span.statement)
      const linked = subjects.some(subject => span.start >= subject.start && (span.start < subject.end
        || (self ? CONTENT_LINK : STATUS_LINK).test(sentence.slice(subject.end, span.start))))
      const previous = own.at(-1)
      const coordinated = previous && COORDINATED_LINK.test(sentence.slice(previous.end, span.start))
      if (marking || linked || coordinated || (self && OWN_CONTENT.test(sentence.slice(span.end)))) own.push(span)
    }
    const words = wordsOf(sentence)
    const taken = (statement: string) => keep.some(paragraph => paragraph.includes(flat(statement))
      && words.filter(word => paragraph.includes(` ${word} `)).length >= TAKEN_FROM * words.length)
    return own.filter(span => !taken(span.statement)).map(span => span.statement.replace(/\s+/g, " "))
  }
  const found = [
    ...text.replace(NOT_STATEMENT, " ").split(/\n+|(?<=[.;!?])\s+/)
      .flatMap(sentence => statements(sentence, SENTENCE).map(statement => ({ statement, sentence: sentence.replace(LEAD, "").trim() }))),
    ...[...text.matchAll(OPEN_ITEM)].flatMap(([marker]) => marker.slice(1, -1).replace(/^open item:/iu, "").split(/(?<=[.;!?])\s+|\s+[-–—]\s+/)
      .flatMap(clause => statements(clause, MARKER).map(statement => ({ statement, sentence: marker })))),
  ]
  return [...new Map(found.map(row => [flat(row.statement), row])).values()]
}

function wordsOf(text: string) {
  return flat(text).trim().split(" ").filter(Boolean)
}

/** Lowercased words between single spaces, so a phrase matches whole words. */
function flat(text: string) {
  return ` ${text.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim()} `
}

/** What a document's text may hold that reads as status wording: the firm's
 * legend for its form, a legend the request asks for, and, in a revision, the
 * lawyer's original. Undefined for a document that is not an agreement. */
function keptText(directory: string, document: DocID): string[] | undefined {
  const plan = plannedDocument(directory, document)
  if (plan?.form !== "agreement") return
  return [furnitureLegend(plan.form, normalizePreferences(readFirmPreferences())), normalizeLegend(plan.furniture?.legend), revisedOriginal(directory, plan)]
    .filter(Boolean)
}

function revisedOriginal(directory: string, plan: PlannedDocument) {
  const source = plan.mode === "revise" ? templateSource(plan) : undefined
  return source && matterSourceIDs(directory).includes(source) ? sourceText(directory, source) : ""
}

const HOME = "The agreement's text holds only its terms: the firm's draft line and legend are applied to every page from the firm's settings, and each open point is a visible open-item marker in the paragraph it concerns, listed in the drafting notes. A marker says what is missing or to be confirmed, never the draft's status."

/** Open items for the paragraphs of an agreement that state its draft status
 * or approval state: in the draft, required repairs for the change cycle; in
 * the final body, items the drafting notes disclose. */
export function draftStatusItems(directory: string, document: DocID, rows: ParagraphRow[], stage: "draft" | "apply"): OpenItem[] {
  const kept = keptText(directory, document)
  if (!kept) return []
  return rows.flatMap(row => {
    const found = draftStatusStatements(row.text, kept)
    if (!found.length) return []
    return [openItem(stage, "draft_status",
      `The agreement states its own draft status or approval state: "${excerpt(found.map(row => row.sentence).join(" "))}". ${HOME} Remove the statement, keeping any open point it names as such a marker.`,
      { document, pid: row.pid, statement: found[0]!.statement })]
  })
}

/** Retry feedback for amendments to an agreement that write a statement of
 * its draft status or approval state the paragraph does not already hold. */
export function draftStatusAmendmentIssues(directory: string, document: DocID) {
  const kept = keptText(directory, document)
  return (amendment: Amendment, index: number, current: string): string[] => {
    if (!kept) return []
    const found = amendmentParagraphs(amendment).flatMap(text => draftStatusStatements(text, [...kept, current]))
    return found.length
      ? [`amendments[${index}] ([${amendment.pid}]) states the agreement's own draft status or approval state ("${found.map(row => row.statement).join("\", \"")}"). ${HOME} Leave the statement out; put an open point it names in a visible [open item: ...] marker.`]
      : []
  }
}
