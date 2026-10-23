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
// Found: a sentence that calls this text a draft, says it is not approved or
// not for signature, says it is circulated for discussion or negotiation only,
// or describes its own placeholders or bracketed wording. Some of these words
// also name other documents in operative wording ("the Consultant shall
// circulate a draft for discussion"), so those count only in a sentence that
// opens with them, as a marking line does, or that is about this text ("this
// document", "It is ..."). Not found: what a clause does to the parties' rights,
// including whether the agreement binds ("subject to contract", "not legally
// binding"), which is a term the parties may agree or a firm legend; quoted
// words; bracketed markers. Kept: the firm's legend for the document's form and
// a legend the request asks for (applied by code as page furniture), and, in a
// revision, the wording of the lawyer's own original, so a marking the lawyer
// wrote is never taken out of their document.

const S = String.raw`\s+`
const HYPHEN = "[-‐‑–]"

/** Status wording that is about the text it stands in wherever it stands. */
const ALWAYS = new RegExp([
  String.raw`\bnot${S}(?:(?:yet|been|to${S}be|being|currently)${S})*(?:approved|cleared|authori[sz]ed|released|ready|final(?:i[sz]ed)?)${S}for${S}(?:signature|signing|execution|completion|implementation|issue|release|circulation)\b`,
  String.raw`\bnot${S}for${S}(?:signature|signing|execution)\b`,
  String.raw`\bnegotiation${HYPHEN}only\b`,
  String.raw`\bfor${S}(?:discussion|negotiation)(?:${S}(?:and|or)${S}(?:discussion|negotiation))?${S}(?:purposes${S})?only\b`,
  String.raw`\bdraft${S}status\b`,
  String.raw`\b(?:this|the${S}present)${S}(?:(?:negotiation|discussion|working|consultation|revised|first|initial|clean|further|current)(?:${HYPHEN}only)?${S})*draft\b`,
  String.raw`\bbracketed${S}(?:wording|text|drafting|language|provisions?)\b`,
  String.raw`\b(?:explicit|bracketed)${S}placeholders?\b`,
  String.raw`\bplaceholders?${S}(?:pending|until|awaiting)\b`,
].join("|"), "giu")

/** Status wording that also names other documents: only in a sentence that
 * opens with it or that is about this text. */
const IN_CONTEXT = new RegExp([
  String.raw`\b(?:negotiation|discussion|working|consultation)${S}draft\b`,
  String.raw`\bdraft${S}(?:(?:issued|prepared|circulated|provided)${S})?for${S}(?:discussion|negotiation)\b`,
  String.raw`\b(?:issued|circulated|provided|prepared|sent|supplied|produced)${S}(?:(?:only|solely|purely)${S})?for${S}(?:the${S}purposes?${S}of${S})?(?:discussion|negotiation)\b`,
  String.raw`^draft\b(?=\s*(?:$|[-–—:|(]))`,
].join("|"), "giu")

/** A sentence about this text itself. */
const ABOUT_THIS_TEXT = /\b(?:this|the\s+present)\s+(?:[\p{L}-]+\s+){0,3}?(?:draft|document|version)\b|^(?:it|this)\s+(?:is|was|has|remains|does)\b/iu

/** Markers, citations and quoted words carry no statement of the text's own. */
const NOT_STATEMENT = /\[[^\]\n]*\]|\(open item:[^)\n]*\)|"[^"\n]*"|“[^”\n]*”/giu

/** Leading numbering, bullets and emphasis before a sentence's first word. */
const LEAD = /^[\s#>*_\-•–—]*(?:\(?(?:\d+(?:\.\d+)*|[a-z]|[ivxlc]+)[.)]\s*)?[\s*_]*/iu

export type DraftStatusStatement = { statement: string; sentence: string }

/** The statements of its own draft status or approval state in one paragraph
 * of an agreement, leaving out any whose words `kept` (firm legends, the
 * lawyer's original) already holds. */
export function draftStatusStatements(text: string, kept: string[] = []): DraftStatusStatement[] {
  const keep = kept.map(flat)
  const found = text.replace(NOT_STATEMENT, " ").split(/\n+|(?<=[.;!?])\s+/).flatMap(raw => {
    const sentence = raw.replace(LEAD, "").trim()
    if (!sentence) return []
    const context = ABOUT_THIS_TEXT.test(sentence)
    const spans = [
      ...[...sentence.matchAll(ALWAYS)],
      ...[...sentence.matchAll(IN_CONTEXT)].filter(match => context || match.index === 0),
    ].map(match => ({ start: match.index, end: match.index + match[0].length, statement: match[0] }))
    // One statement per stretch of words: the widest match.
    const widest = spans.filter(span => !spans.some(other => other !== span && other.start <= span.start && other.end >= span.end
      && other.end - other.start > span.end - span.start))
    return widest.filter(span => !keep.some(words => words.includes(flat(span.statement))))
      .map(span => ({ statement: span.statement.replace(/\s+/g, " "), sentence }))
  })
  return [...new Map(found.map(row => [flat(row.statement), row])).values()]
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

const HOME = "The agreement's text holds only its terms: the firm's draft line and legend are applied to every page from the firm's settings, and each open point is a visible open-item marker in the paragraph it concerns, listed in the drafting notes."

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
