import { normalize, normalizeWithMap } from "../../lib/quote-verify"
import { TOOL_SYNTAX_ISSUE, printedToolCall, record } from "./checks"
import { FIRST_DOCUMENT, nameKey, namedIn, readLeanPlan, type ActingFor, type DocID, type PlannedDocument } from "./plan"
import { tagged } from "./prompts"
import { EXCERPT, REQUEST_SOURCE, excerpt, lawyerWords, matterSourceIDs, openItem, quoteFound, quoteMissIssue } from "./quotes"
import { answeredScopeQuestion, scopeAnswers, scopeField, scopeQuestionShows, type ScopeParties } from "./package-scope"
import { answeredQuestionsText, handedBackBlock, hashStageInputs, readLeanJson, readLeanText, readQuestions, type OpenItem, type Question } from "./state"
import { sourceText } from "./template-structure"

// Package coverage: an instructed act needs an instrument executed by the
// parties who must take it. One focused job reads the planned package against
// the lawyer's instructions and lists each legal act they direct, the parties
// who must take it and the planned document that carries it. parseCoverage
// verifies that list against the matter (each quote in its source, each party
// named by the instructions or sources, each document a planned id);
// planCoverageIssues then holds the plan to it, as retry feedback for the
// outline. Nothing here judges the legal substance or supplies a party: an act
// whose parties the record does not settle lists none, and a plan-declared
// party counts only when the record names it. The outline step records the
// result in coverage.json, keyed on the plan it checked, and asks the lawyer
// one package-scope question (coverageScopeQuestion) on any act still uncovered
// that no answered package-scope question settled, never more than twice.
// An act whose document the lawyer's own words leave out of this package is
// listed with those words (excluded_by) and no document: it raises no question,
// a revision only when its parties all execute a planned document that carries
// no instructed act, and is disclosed as the check's reading for the lawyer to
// confirm, with any such document still planned, or with the planned documents
// a last attempt also listed it in.
// What the saved plan still leaves open reaches the lawyer as open items derived
// from that record (coverageItems), never stored. The gate reads the same record
// (coverageCarries) to tell an instruction the plan carries from one it may
// leave out.

/** One instructed act: what is to be done, the words that direct it and their
 * source, every party who must take it ([] when the record does not settle
 * them) and the planned document that carries it (null when none does). With
 * no document, `excluded_by` holds the lawyer's own words that leave the act's
 * document out of this package. */
export type CoverageAct = { act: string; source: string; quote: string; parties: string[]; document: DocID | null; excluded_by?: string }

export const LEAN_COVERAGE_SCHEMA = "dochaus.lean.coverage/1"
/** "covered": no issue is left; "uncovered": issues remain after the outline's
 * revisions; "unchecked": the coverage reply never passed its check, so `issues`
 * are that check's and no act is recorded. */
export type LeanCoverage = {
  schema: typeof LEAN_COVERAGE_SCHEMA; plan: string; status: "covered" | "unchecked" | "uncovered"; acts: CoverageAct[]; issues: string[]
}

type Checked<T> = { issues: string[]; value?: T }

const COVERAGE_SHAPE = 'Reply with the JSON object only: {"acts":[{"act":"...","source":"...","quote":"...","parties":["..."],"document":"planned id or null"}]}.'

/** The coverage job's one question, over the instructions and the planned documents. */
export function coveragePrompt(directory: string, documents: PlannedDocument[]) {
  const updates = readLeanText(directory, "updated-notes.md")?.trim()
  const answers = answeredQuestionsText(directory)
  return [
    "This is a lean drafting run. Check the planned package against the lawyer's instructions: which legal acts they direct, who must take each and which planned document carries it.",
    tagged("request", readLeanText(directory, "request.md") ?? ""),
    updates ? tagged("user_updates", `Later instructions from the lawyer, in arrival order. An explicit later change supersedes the earlier instruction on that point.\n\n${updates}`) : "",
    answers ? tagged("answers", `The lawyer's answers; they remain controlling:\n${answers}`) : "",
    handedBackBlock(directory) ? tagged("left_to_drafting", handedBackBlock(directory)) : "",
    tagged("sources", [
      "Matter documents (use these exact names with search-document and read-document, and as source names):",
      ...matterSourceIDs(directory).map(name => `- ${name}`),
      `To quote the request, the lawyer's updates or answers, use the source name "${REQUEST_SOURCE}".`,
    ].join("\n")),
    tagged("planned_documents", "The plan under check, by id. Its purposes and parts are the plan's own choices, not instructions. executed_by lists the parties who execute a document; [] says it takes no legal act, and null that the plan does not say.\n" +
      JSON.stringify(documents.map(coverageView), null, 2)),
    tagged("output_contract", [
      "- List every legal act that the request, the lawyer's updates or answers direct to be taken or documented in this package. A matter document directs an act only where the request adopts it as instructions; an obligation a matter document merely describes is not an instructed act. " +
      "List an act the instructions direct even when a planned purpose leaves it out or defers it, or no planned document carries it. Do not list an act the instructions themselves exclude or defer.",
      "- For each act give \"act\": what is to be done, briefly; \"source\": the exact matter document name, or \"request\"; \"quote\": the exact words that direct it, copied verbatim from that source; " +
      "\"parties\": every party who must take the act for it to be effective (for an authorisation, those who give it), each named as the sources name it, or [] when the record does not settle who must take it; \"document\": the id of the planned document that carries it, or null when none does.",
      "- When the request, the lawyer's updates or answers leave out of this package the document that the parties who must take an act would execute, " +
      "still list the act, with \"document\": null and \"excluded_by\": the whole sentence that leaves that document out, copied verbatim from the request, updates or answers, " +
      "never from a planned document or a matter document. Leave \"excluded_by\" out of every other act.",
      "- When a planned document's executed_by names a party as the sources do, write that party exactly as executed_by declares it. Never supply a party by assumption.",
      "- Read the sources you rely on with search-document or read-document, and check each quote with verify-quote.",
      `- ${COVERAGE_SHAPE} Reply {"acts":[]} when the instructions direct no legal act.`,
      "- Read-only tools are available; file-writing and publication tools are unavailable. The run saves and checks the reply.",
    ].join("\n")),
  ].filter(Boolean).join("\n\n")
}

/** Reads the coverage reply and verifies each act against the matter: the
 * quote in its source ("request" is the request, updates and answers), the
 * document a planned id or null, each party named by the instructions or
 * sources, and any excluded_by, with no document, as whole sentences of the
 * lawyer's own words (lawyerSentences) other than the words that direct the
 * act; words that are not are shown the whole sentences of the lawyer's text
 * that contain them (sentencesAround). excluded_by is the only evidence that the
 * package may leave an act's document out, so the same act for the same parties
 * listed both with and without it is retried, never kept twice. A planned
 * document's executed_by is the plan's own writing, so it grounds no party.
 * Every defect is retry feedback; the acts are returned only when there is
 * none, each entry once (by act, parties, document, source, quote and any
 * excluded_by words, the last two as the quote checks normalise them). On the last attempt (`final`) a refused excluded_by is dropped and the
 * act is then held to the plan like any other, so the outline's revisions and
 * the package-scope question see it, and the reply is not lost to that defect.
 * Of an act listed both with and without excluded_by that passed, only the
 * entries without it are kept when none of them names a planned document: all
 * readings agree that no planned document carries the act, which then goes to
 * the outline's revisions and, still uncovered, to the lawyer's package-scope
 * question. When one of them names a planned document every entry is kept, so
 * the lawyer's words are disclosed beside that document (coverageItems) and
 * never silently overridden by the plan. Every other defect still fails the
 * last attempt. */
export function parseCoverage(reply: string, directory: string, documents: PlannedDocument[], { final = false } = {}): Checked<CoverageAct[]> {
  if (printedToolCall(reply)) return { issues: [TOOL_SYNTAX_ISSUE] }
  let value: unknown
  try { value = JSON.parse(reply.trim().replace(/^```(?:json)?\s*/, "").replace(/\s*```$/, "")) } catch {}
  if (!record(value) || !Array.isArray(value.acts)) return { issues: [COVERAGE_SHAPE] }
  const ids = documents.map(document => document.id)
  const sources = [REQUEST_SOURCE, ...matterSourceIDs(directory)]
  let texts: string[] | undefined
  const issues: string[] = []
  const rows: { index: number; act: CoverageAct; same: string }[] = []
  const listed = new Map<string, { act: string; excluded: number[]; included: number[]; carried: boolean }>()
  value.acts.forEach((row: unknown, index) => {
    const text = (key: string) => record(row) && typeof row[key] === "string" ? (row[key] as string).trim() : ""
    if (!record(row) || !text("act") || !text("source") || !text("quote") || !(Array.isArray(row.parties) && row.parties.every(party => typeof party === "string" && party.trim())) ||
      !(row.document === null || typeof row.document === "string"))
      return void issues.push(`acts[${index}] needs nonempty "act", "source" and "quote" strings, "parties" as a list of names ([] when the record does not settle them), and "document" as a planned id or null. ${COVERAGE_SHAPE}`)
    const named = (row.parties as string[]).map(party => party.trim())
    const parties = named.filter((party, at) => named.findIndex(other => nameKey(other) === nameKey(party)) === at)
    const act: CoverageAct = { act: text("act"), source: text("source"), quote: text("quote"), parties, document: (text("document") || null) as DocID | null,
      ...(text("excluded_by") ? { excluded_by: text("excluded_by") } : {}) }
    const where = `acts[${index}] ("${excerpt(act.act, 80)}")`
    const refused = row.excluded_by != null && typeof row.excluded_by !== "string"
      ? `${where} gives "excluded_by" that is not a string. Give the whole sentence of the request, one of the lawyer's updates or one answer that leaves this act's document out of this package, verbatim, or leave "excluded_by" out.`
      : !act.excluded_by ? undefined
      : act.document !== null ? `${where} gives "excluded_by" with the document "${act.document}". An act whose document the instructions leave out of this package is carried by no planned document: ` +
        'give "document": null with "excluded_by", or the planned document without it.'
      : !lawyerSentences(directory, act.excluded_by) ? `${where} gives "excluded_by": "${excerpt(act.excluded_by)}", but those words are not one or more whole sentences of the request, one of the lawyer's updates or one answer, exactly as written. ` +
        sentencesAround(directory, act.excluded_by) +
        "Copy verbatim, with its capitals and from its first word to its last, the whole sentence of the lawyer's own words that leaves this act's document out of this package; " +
        "planned documents and matter documents do not count. Otherwise leave \"excluded_by\" out."
      : overlaps(act.excluded_by, act.quote) ? `${where} gives "excluded_by": "${excerpt(act.excluded_by)}", which overlaps the words that direct this act ("${excerpt(act.quote)}"). ` +
        "\"excluded_by\" cannot be the words that direct the act or a sentence containing them: copy verbatim another whole sentence of the request, one update or one answer " +
        "that leaves this act's document out of this package, or leave \"excluded_by\" out."
      : undefined
    if (refused && final) delete act.excluded_by
    else if (refused) issues.push(refused)
    if (!sources.includes(act.source))
      issues.push(`${where} gives the source "${act.source}", which is not a matter document. Use one of ${sources.map(name => `"${name}"`).join(", ")}.`)
    else if (!quoteFound(directory, { quote: act.quote, source_id: act.source }))
      issues.push(quoteMissIssue(directory, where, { quote: act.quote, source_id: act.source }, "json"))
    if (act.document !== null && !ids.includes(act.document))
      issues.push(`${where} names the document "${act.document}", which is not planned. Use one of the planned ids ${ids.join(", ")}, or null when no planned document carries it.`)
    for (const party of parties.filter(party => !(texts ??= namingTexts(directory)).some(text => namedIn(text, party))))
      issues.push(`${where} names the party "${party}", but neither the request, the lawyer's updates and answers nor any matter document names that party. ` +
        "Name each party as the sources name it, and list only parties the record names ([] when it does not settle who must take the act).")
    const same = sameAct(act)
    const group = listed.get(same) ?? listed.set(same, { act: act.act, excluded: [], included: [], carried: false }).get(same)!
    ;(act.excluded_by ? group.excluded : group.included).push(index)
    group.carried ||= !act.excluded_by && act.document !== null
    rows.push({ index, act, same })
  })
  const entries = (indexes: number[]) => {
    const names = indexes.toSorted((a, b) => a - b).map(index => `acts[${index}]`)
    return names.length > 1 ? `${names.slice(0, -1).join(", ")} and ${names.at(-1)}` : names.join("")
  }
  const dropped = new Set<number>()
  for (const { act, excluded, included, carried } of listed.values()) if (excluded.length && included.length) {
    if (final && !carried) excluded.forEach(index => dropped.add(index))
    else if (!final) issues.push(`${entries([...excluded, ...included])} describe an act the same way ("${excerpt(act, 80)}") for the same parties: with "excluded_by" in ${entries(excluded)}, without it in ${entries(included)}. ` +
      "If these are different acts, describe each so that the descriptions differ. If they are the same act, keep one entry for it: " +
      "with \"excluded_by\" and \"document\": null only when the lawyer's own words leave its document out of this package, otherwise without \"excluded_by\".")
  }
  if (issues.length) return { issues }
  const seen = new Set<string>()
  const acts = rows.flatMap(({ index, act, same }) => {
    // A last attempt's kept words and an entry in no document are two readings, both kept.
    // Other words that direct the act or leave its document out are another instruction, also kept.
    const key = JSON.stringify([same, act.document, act.source, normalize(act.quote), act.excluded_by ? normalize(act.excluded_by) : null])
    if (dropped.has(index) || seen.has(key)) return []
    seen.add(key)
    return [act]
  })
  return { issues, value: acts }
}

// A sentence of the lawyer's text ends after a run of ".", "!", "?" or "…" (and
// any closing quotation mark, bracket or emphasis) before whitespace or the end,
// or after a ";". A "!" or "?" run always ends one. A "." or "…" run whose next
// word starts with a lowercase letter or a digit (CONTINUES) ends none: an
// abbreviation, a number after one (cl. 3 of, Facility No. 2. The) or a pause.
// Only a digit that opens an inline list label (inlineLabels) starts a new
// sentence there (Two points: 1. … 2. …).
const SENTENCE_END = /([.!?…]+)["'”’)\]*_]*(?=\s|$)|;/g
const CONTINUES = /[\s>*_"'“‘(]*([\p{Ll}\p{Nd}])/uy
// A number closed by "." or ")" before a word that does not start lowercase,
// not the end of a longer number (4.25, 1.1.): an inline list label, or the
// number that closes a sentence after an abbreviation (Facility No. 2. The).
const LABEL = /(?<![\p{L}\p{N}.,])(\d+)([.)])(?=\s+[^\p{Ll}\s])/gu
// A list, heading or blockquote marker at the start of a line: a bullet (- * •
// + · – — ◦ ▪), numbering such as 1. 1) 1.1 2.3.1 a. (a) a) (A) (iv), a
// heading's # or a blockquote's >. Only a lowercase letter with "." numbers a
// line, and only before a word that is not a number or lowercase ("A. Moss" is
// an initial, "s. 146" and "p. 5" carry on the line above), so the expression
// is case-sensitive. A number such as 1.1 without a closing "." numbers a line
// only before a capitalised word: "3.5 million", "4.25 %", "4.2 (Conditions)"
// and "2.5 EUR" carry on the line above. sentenceBounds holds a dash to its
// list (DASH) and a lettered marker to its sequence (NEXT_LABEL).
const LINE_MARKER = /^[ \t]*(?:([-*•+·–—◦▪]|#{1,6}|\d+(?:\.\d+)*[.)]|\d+(?:\.\d+)+(?=[ \t]+\p{Lu}(?!\p{Lu}))|[a-z]\.(?=[ \t]+[^\p{Ll}\p{Nd}\s])|\(?(?:\d+|[A-Za-z]|[ivxIVX]+)\))[ \t]+|>[ \t]*)/u
// A dash bullet. A hard wrap can put a spaced dash at the start of a line
// ("… out of this package\n– unless …"), so after a line that carries on, a
// dash opens a list item only when an item of the same paragraph opened with
// the same dash, or before a capitalised word.
const DASH = /^[-–—]$/
// A lettered or lowercase roman label closed by "." at the start of a line (a.
// b. ii. iv.), before a word that is not a number. sentenceBounds reads it as a
// list item, whatever the case of that word, when it comes next after a label
// of a list still open (follows), or as a list's first label (a. i.) on a line
// that starts a sentence or follows one ending as a sentence does: "a. Leave
// X.\nb. the Y" and "Notes:\na. the X\nb. the Y" are two items each, while
// "Moss\nv. the Crown" and "under\ns. 146" carry on the line above.
const NEXT_LABEL = /^[ \t]*((?:[a-z]|[ivx]+)\.)[ \t]+(?=[^\p{Nd}\s])/u
// The lowercase roman numerals from 1 to 39, each at its value.
const ROMAN = Array.from({ length: 40 }, (_, n) => "x".repeat(Math.floor(n / 10)) + ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix"][n % 10]!)
// Emphasis, quotation marks, brackets and an ellipsis around a sentence.
const AROUND = /[*_"'()[\]…]/

/** Offsets of the inline list labels of one lawyer text (where each label's
 * number starts): each LABEL in a run of two or more in one paragraph, each n
 * after the nearest n − 1 with the same closer (1. … 2. …, 1) … 2) …). The "."
 * before a label could equally close an abbreviation whose number ends the
 * sentence (Leave out Facility No. 1. Keep Facility No. 2. The rest stays.),
 * so a run closed by "." counts only when its first label opens the
 * paragraph, a line, or follows ":", ";", "!" or "?" (and any opening quotation
 * mark, bracket or emphasis). */
function inlineLabels(text: string) {
  const labels = new Set<number>()
  let offset = 0
  for (const paragraph of text.split(/(?<=\n[ \t]*\n)/)) {
    const runs = new Map<string, { opened: boolean; starts: number[] }>()
    for (const label of paragraph.matchAll(LABEL)) {
      const [, number, closer] = label
      let at = label.index
      while (at > 0 && /[ \t>*_"'“‘([]/.test(paragraph[at - 1]!)) at--
      const run = runs.get(`${closer}${Number(number) - 1}`) ?? { opened: closer === ")" || at === 0 || /[\n:;!?]/.test(paragraph[at - 1]!), starts: [] }
      run.starts.push(offset + label.index)
      runs.set(`${closer}${Number(number)}`, run)
      if (run.opened && run.starts.length > 1) for (const start of run.starts) labels.add(start)
    }
    offset += paragraph.length
  }
  return labels
}

/** Whether the list label `next` comes right after `last` (each a lowercase
 * letter or roman numeral with its "."): the next letter (b. after a.) or the
 * next roman numeral (ii. after i., v. after iv.). */
function follows(last: string, next: string) {
  const [from, to] = [last.slice(0, -1), next.slice(0, -1)]
  return (from.length === 1 && to.length === 1 && to.charCodeAt(0) === from.charCodeAt(0) + 1) || (ROMAN.indexOf(from) > 0 && ROMAN.indexOf(to) === ROMAN.indexOf(from) + 1)
}

/** Offsets of one lawyer text where a sentence may start or stop: the text's
 * start and end, each SENTENCE_END, and the start of a line (and the end of
 * each of its markers) that follows a blank line, a heading, a line ending with
 * ":" (and any closing quotation mark, bracket or emphasis) or a line whose
 * end is a SENTENCE_END, or that opens with a list or heading marker or changes
 * the blockquote depth. A dash opens a list item after a line that carries on
 * only as DASH allows. A lettered or roman label opens an item before any word
 * but a number (NEXT_LABEL), so the bound falls before it, when it comes next
 * after a label of an open list (b. after a., ii. after i.), or as a list's
 * first label (a. i.) on a line that starts a sentence or follows a line ending
 * with ".", "!", "?", "…" or ";" (and any closing quotation mark, bracket or
 * emphasis). Any other lettered marker opens one only as a list's first letter
 * before a word that is not a number or lowercase ("v. Wren" carries on
 * "Moss"). A label next after one of an open list closes the lists opened
 * since, and a paragraph that opens with no list marker closes every list. A
 * line's trailing spaces (a Markdown hard line break) are not part of its end.
 * Any other line break is a hard wrap inside a sentence. */
function sentenceBounds(text: string) {
  const bounds = new Set([0, text.length])
  const labels = inlineLabels(text)
  for (const end of text.matchAll(SENTENCE_END)) {
    const after = end.index + end[0].length
    CONTINUES.lastIndex = after
    const next = CONTINUES.exec(text)
    if (!/[.…]$/.test(end[1] ?? "") || !next || labels.has(CONTINUES.lastIndex - next[1]!.length)) bounds.add(after)
  }
  let offset = 0
  let previous: { end: number; depth: number; blank: boolean; heading: boolean; colon: boolean; stop: boolean } | undefined
  // The last label of each open lettered or roman list, outermost first.
  let bullet: string | undefined, lists: string[] = []
  for (const line of text.split("\n")) {
    const starts = [offset]
    let at = 0, depth = 0, listed = false, heading = false
    const opens = () => !previous || previous.blank || depth !== previous.depth || previous.heading || previous.colon || bounds.has(previous.end)
    const markerAt = (rest: string) => {
      const next = NEXT_LABEL.exec(rest)
      return next && (lists.some(last => follows(last, next[1]!)) || (/^[ai]\.$/.test(next[1]!) && (opens() || previous!.stop))) ? next : LINE_MARKER.exec(rest)
    }
    for (let marker = markerAt(line); marker; marker = markerAt(line.slice(at))) {
      const sign = marker[1]
      if (sign && DASH.test(sign) && !(opens() || bullet === sign || /^[*_"'“‘([]*\p{Lu}/u.test(line.slice(at + marker[0].length)))) break
      if (sign && /^[a-z]\.$/.test(sign) && !(sign === "a." || sign === "i." || lists.some(last => follows(last, sign)))) break
      at += marker[0].length
      starts.push(offset + at)
      if (!sign) depth++
      else {
        listed = true
        heading ||= sign.startsWith("#")
        if (DASH.test(sign)) bullet = sign
        if (/^(?:[a-z]|[ivx]+)\.$/.test(sign)) {
          // Next after an open list's label, it replaces that label and closes the lists opened since; any other opens a list.
          const open = lists.findLastIndex(last => follows(last, sign))
          lists = [...lists.slice(0, open < 0 ? lists.length : open), sign]
        }
      }
    }
    const blank = !line.slice(at).trim()
    if (listed || opens()) for (const start of starts) bounds.add(start)
    if (blank) bullet = undefined
    // A paragraph that opens with no list marker closes every list.
    else if (previous?.blank && !listed) lists = []
    const trimmed = line.trimEnd()
    previous = { end: offset + trimmed.length, depth, blank, heading, colon: /:["'”’)\]*_]*$/.test(trimmed), stop: /[.!?…;]["'”’)\]*_]*$/.test(trimmed) }
    offset += line.length + 1
  }
  return [...bounds].toSorted((a, b) => a - b)
}

/** The lawyer's texts (lawyerWords) as the sentence checks read them: without
 * NUL, which lawyerSentences uses to mark a bound, and with every line break
 * (a "\n", "\r\n" or bare "\r") as "\n". */
function lawyerTexts(directory: string) {
  return lawyerWords(directory).map(raw => raw.replaceAll("\0", "").replace(/\r\n?/g, "\n"))
}

/** Whether `words` are one or more whole sentences of one lawyer text (the
 * request, one included update or one answer, as lawyerTexts gives them),
 * verbatim and in the same case apart from the quote checks' whitespace and
 * quotation-mark normalisation: they start and stop at sentenceBounds, and may
 * leave off a line's markers, the emphasis, quotation marks, brackets or
 * ellipsis around them, and the closing ".", "!", "?" or ";". Markup inside
 * them (MARKUP) is kept as written: the words are quoted again as the lawyer's
 * (coverageItems) and the quote checks read them verbatim. Words running
 * across two texts never count. The bounds come from punctuation and line
 * layout, not grammar, so words that start after an abbreviation followed by a
 * capital still pass. */
function lawyerSentences(directory: string, words: string) {
  const quote = normalize(words)
  return quote.length > 0 && lawyerTexts(directory).some(text => {
    // Mark each bound, normalise, then note where the marks fall in the unmarked text.
    let marked = "", from = 0
    for (const at of sentenceBounds(text)) {
      marked += `${text.slice(from, at)}\0`
      from = at
    }
    let plain = ""
    const ends = new Set<number>()
    for (const char of normalize(marked)) {
      if (char === "\0") ends.add(plain.length)
      else if (char !== " " || (plain && !plain.endsWith(" "))) plain += char
    }
    const starts = (at: number) => {
      while (!ends.has(at) && (plain[at - 1] === " " || AROUND.test(plain[at - 1] ?? ""))) at--
      return ends.has(at)
    }
    const stops = (at: number) => {
      while (!ends.has(at) && /[ .!?;*_"'()[\]…]/.test(plain[at] ?? "")) at++
      return ends.has(at)
    }
    for (let at = plain.indexOf(quote); at !== -1; at = plain.indexOf(quote, at + 1))
      if (starts(at) && stops(at + quote.length)) return true
    return false
  })
}

// Markup that can stand inside a sentence: emphasis, and the blockquote markers
// opening a line that carries a sentence on.
const MARKUP = /[*_]+|^[ \t]*(?:>[ \t]*)+/gm

/** `text` without MARKUP, normalised as normalizeWithMap does, with each
 * character's offset in `text`. */
function withoutMarkup(text: string) {
  const markup = new Set([...text.matchAll(MARKUP)].flatMap(match => Array.from({ length: match[0].length }, (_, at) => match.index + at)))
  const kept = Array.from({ length: text.length }, (_, at) => at).filter(at => !markup.has(at))
  const { text: plain, map } = normalizeWithMap(kept.map(at => text[at]).join(""))
  return { text: plain, map: map.map(at => kept[at]!) }
}

/** The length of the line markers (LINE_MARKER) that open `line`. */
function markersLength(line: string) {
  let at = 0
  for (let marker = LINE_MARKER.exec(line); marker; marker = LINE_MARKER.exec(line.slice(at))) at += marker[0].length
  return at
}

// The longest whole sentences a refusal hint quotes. A hint never cuts what it
// offers to be copied, since a cut sentence given back is refused again; longer
// ones are pointed to where they stand in the lawyer's text instead.
const QUOTED_WHOLE = 2000
const FIND = (them: string) => `too long to quote here: find ${them} where those words stand in the request, the lawyer's updates or answers. `

/** For nonempty `words` that are not whole sentences of the lawyer's
 * (lawyerSentences), the refusal's pointer to every place in the lawyer's texts
 * that best holds them, each as its whole sentences: from the sentenceBounds at
 * or before the words' first character to the one after their last, normalised
 * to one line (given back whole, they pass lawyerSentences). The places that
 * hold the words as written (verbatim, same case) tie, so each is offered; only
 * when none does, the places that hold them once MARKUP is set aside on both
 * sides, quoted with the lawyer's markup as written, since only that passes. A
 * place is only ever quoted whole, and only up to QUOTED_WHOLE characters; past
 * that it is pointed to, never cut. One place is quoted; of several, each that
 * fits whole within the excerpt length is quoted, in the texts' order, and the
 * rest are counted (the first short enough quoted when none fits). Otherwise no
 * one lawyer text holds them. */
function sentencesAround(directory: string, words: string) {
  const texts = lawyerTexts(directory).map(text => ({ text, bounds: sentenceBounds(text) }))
  for (const read of [normalizeWithMap, withoutMarkup]) {
    const quote = read(words).text
    // Each place's whole sentences, and whether they are several.
    const places = new Map<string, boolean>()
    for (const { text, bounds } of quote ? texts : []) {
      const { text: plain, map } = read(text)
      for (let at = plain.indexOf(quote); at !== -1; at = plain.indexOf(quote, at + 1)) {
        const from = bounds.findLast(bound => bound <= map[at]!)!
        const to = bounds.find(bound => bound > map[at + quote.length - 1]!)!
        // Several sentences when a bound past the line's markers has words on both sides; a marker alone is none.
        const start = from + markersLength(text.slice(from, to))
        places.set(normalize(text.slice(from, to)), bounds.some(bound => bound > start && bound < to &&
          [text.slice(start, bound), text.slice(bound, to)].every(part => /[\p{L}\p{N}]/u.test(part))))
      }
    }
    if (!places.size) continue
    const markup = read === withoutMarkup
    const [[span, several]] = [...places]
    if (places.size === 1) return `In the lawyer's text, the whole sentence${several ? "s" : ""} containing them${markup ? `, with ${several ? "their" : "its"} markup as written,` : ""} ` +
      `${several ? "are" : "is"} ${span!.length <= QUOTED_WHOLE ? `"${span}". ` : FIND(several ? "them" : "it")}`
    let room = EXCERPT
    const quoted: string[] = []
    for (const whole of places.keys()) if (whole.length <= room) {
      quoted.push(`"${whole}"`)
      room -= whole.length
    }
    const first = [...places.keys()].find(whole => whole.length <= QUOTED_WHOLE)
    if (!quoted.length && first !== undefined) quoted.push(`"${first}"`)
    const there = `In the lawyer's text, ${places.size} places contain them; the whole sentences there${markup ? ", with their markup as written," : ""} are `
    if (!quoted.length) return `${there}each ${FIND("them")}`
    const listed = [...quoted, ...(places.size > quoted.length ? [`${places.size - quoted.length} more not quoted here`] : [])]
    return `${there}${listed.slice(0, -1).join(", ")} and ${listed.at(-1)}. `
  }
  return "No one text of the request, the lawyer's updates or answers contains those words exactly as written, capitals included. "
}

/** Whether one text contains the other, after the quote checks' normalisation
 * and ignoring case, as "request" quotes are verified. A sentence that both
 * directs the act and leaves its document out is refused on purpose: the words
 * that direct an act cannot also be the evidence that it is left out, so
 * without other words the act is listed without excluded_by and the
 * package-scope question asks the lawyer. */
function overlaps(a: string, b: string) {
  const [left, right] = [normalize(a).toLowerCase(), normalize(b).toLowerCase()]
  return !!left && !!right && (left.includes(right) || right.includes(left))
}

/** The plan held to the verified acts, as retry feedback for the outline: a
 * declared executing party the instructions and sources do not name, an act no
 * planned document carries, an act with a party who does not execute the
 * document that carries it, and an act carried by a document executed by no
 * party. An act whose parties the record does not settle raises nothing once a
 * document that does not declare "none" carries it: no party is supplied. An
 * act whose document the lawyer's own words leave out raises only the planned
 * documents executed by all of its parties that carry no instructed act
 * (leftOutDocuments), likely including the document those words leave out, in
 * one issue naming every one; a document that carries another instructed act
 * is never sent back for it. Acts whose issues read the same (another lawyer
 * sentence leaving the same document out, say) give that line once, in
 * first-occurrence order. */
export function planCoverageIssues(acts: CoverageAct[], documents: PlannedDocument[], directory: string): string[] {
  const declared = documents.flatMap(document => (document.executed_by ?? []).map(party => ({ document, party })))
  const texts = declared.length ? namingTexts(directory) : []
  return [...new Set([
    ...declared.filter(row => !texts.some(text => namedIn(text, row.party))).map(({ document, party }) =>
      `${label(document)} declares "Executed by: ${party}", but neither the request, the lawyer's updates and answers nor any matter document names "${party}". ${DECLARED_PARTY_FIX}`),
    ...acts.flatMap(act => actIssues(act, documents, acts)),
  ])]
}

/** The outline's instruction after a declared party the record names nowhere. */
const DECLARED_PARTY_FIX = 'Name each executing party as the request or sources name it, separating several parties with ";". If the record does not name who executes the document, leave the "Executed by:" line out; never supply a party the record does not name.'
/** The outline's instruction after a left-out act's leftOutDocuments: one, or several. */
const LEFT_OUT_DOCUMENT_FIX = "If the instructions leave that document out, do not plan it."
const LEFT_OUT_DOCUMENTS_FIX = "Do not plan any of those documents that the instructions leave out."
/** A recorded issue as the lawyer reads it, without the outline's instruction.
 * The record's issues were computed from the very plan it checked, so every
 * act's own issue is tied to that act's item (coverageItems) and only a
 * declared party's issue is read here. */
function withoutOutlineFix(issue: string) {
  return issue.endsWith(` ${DECLARED_PARTY_FIX}`) ? issue.slice(0, -DECLARED_PARTY_FIX.length - 1) : issue
}

/** For an act whose document the lawyer's words leave out, each planned
 * document executed by all of its parties that carries no act of `acts` (the
 * record carries no left-out act): likely including the document those words
 * leave out. */
function leftOutDocuments(act: CoverageAct, documents: PlannedDocument[], acts: CoverageAct[]) {
  if (!act.excluded_by || !act.parties.length) return []
  return documents.filter(document => act.parties.every(party => document.executed_by?.some(name => nameKey(name) === nameKey(party))) &&
    acts.every(other => other.document !== document.id))
}

/** "A", "A, and B" or "A, B, and C": entries that carry their own commas. */
function serialList(entries: string[]) {
  return entries.length > 1 ? `${entries.slice(0, -1).join(", ")}, and ${entries.at(-1)}` : entries.join("")
}

/** `act`'s issues against the planned documents; `acts` is the whole record. */
function actIssues(act: CoverageAct, documents: PlannedDocument[], acts: CoverageAct[]): string[] {
  const directed = `"${act.act}" ("${excerpt(act.quote)}" [source: ${act.source}])`
  const who = quotedList(act.parties)
  if (act.excluded_by) {
    const leftOut = leftOutDocuments(act, documents, acts)
    if (!leftOut.length) return []
    const planned = serialList(leftOut.map(document => `${label(document)}, executed by ${quotedList(document.executed_by!)}`))
    return [`The check read the instructions as leaving out of this package the document for ${directed}, which must be taken by ${who}, but ${planned}, ` +
      (leftOut.length > 1 ? `are planned and carry no instructed act. ${LEFT_OUT_DOCUMENTS_FIX}` : `is planned and carries no instructed act. ${LEFT_OUT_DOCUMENT_FIX}`)]
  }
  const move = (verb: string) => `${verb} the act to a document executed by ${who}, or add a companion document executed by ${who}.`
  const carrier = documents.find(document => document.id === act.document)
  if (!carrier) return [`The instructions direct ${directed}, but no planned document carries it. ` + (act.parties.length
    ? `Plan it in a document executed by ${who}, adding one when none is planned.`
    : "Plan it in the document that should carry it, adding a document when none is planned.")]
  const executing = carrier.executed_by
  if (!act.parties.length) return executing?.length === 0 ? [`${directed} is carried by ${label(carrier)}, which is executed by no party (Executed by: none). ` +
    'Plan the act in a document a party executes, naming that party as the request or sources name it; only when the record does not name who executes the document, leave its "Executed by:" line out instead of writing "none". ' +
    "Never supply a party the record does not name."] : []
  if (!executing) return [`${directed} must be taken by ${who}, and ${label(carrier)} carries it but has no "Executed by:" line. ` +
    `Add that line naming the parties who execute it, separated by ";". If they do not include ${who}, ${move("move")}`]
  const missing = act.parties.filter(party => !executing.some(name => nameKey(name) === nameKey(party)))
  if (!missing.length) return []
  const joined = executing.find(name => missing.some(party => namedIn(name, party)))
  const executed = !executing.length ? "no party (Executed by: none)"
    : quotedList(executing) + (missing.length < act.parties.length ? `, not by ${quotedList(missing)}` : "")
  return [`${directed} must be taken by ${who}, but ${label(carrier)} is executed by ${executed}. ` + (joined
    ? `It declares "${joined}" as one executing party: if that names several parties who execute the document, list each separately on the "Executed by:" line, separated by ";", as the request or sources name them; otherwise ${move("move")}`
    : move("Move"))]
}

/** The plan step's package-scope question: the instructed acts the plan,
 * after the outline's revisions, still leaves without a document executed by
 * the parties who must take them. It lists each act (what is to be done, the
 * words that direct it and their source, who must take it and where the plan
 * has it) and asks whether the package should add or change a document so
 * those parties take it, or whether it is outside this request. Its id keys on
 * each act's verbatim quote with its parties (actKeys), so another party's act
 * in the same words is another question. An act one of the `answered`
 * package-scope questions already settled (actScopeAnswers: it showed the act
 * (showsAct), or has the id this question would have) is left out of it: that
 * answer settles it. None is asked when no act is left uncovered and unsettled, or once
 * SCOPE_QUESTIONS_PER_ASKER of the plan step's own are answered; coverageItems
 * discloses the acts left. An act whose document the lawyer's words leave out
 * is never asked about: those words already speak to the package's scope, so its
 * leftOutDocuments go only to the outline's revisions and, any still planned,
 * to its disclosure. The step never answers it. Asked about a `saved` plan
 * (the draft step's recheck of a plan it will not outline again), it offers no
 * new document: it asks only which planned document carries each act, or
 * whether it is outside this request, since a document is added only by a new
 * draft, which plans the package from the start. */
export function coverageScopeQuestion(acts: CoverageAct[], documents: PlannedDocument[], answered: Question[] = [], options: { saved?: boolean } = {}): Question | undefined {
  const { ask: open, id } = actScopeAnswers(acts.filter(act => !act.excluded_by && actIssues(act, documents, acts).length), answered, acts)
  if (!open.length) return undefined
  const several = open.length > 1
  return {
    id,
    stage: "draft",
    question: [`The document plan leaves ${several ? "these instructed acts" : "this instructed act"} without a document executed by the parties who must take ${several ? "them" : "it"}:`,
      ...open.map(act => actLine(act, documents, excerpt(act.quote))),
      options.saved
        ? `${several ? "For each, which" : "Which"} planned document, executed by those parties, should carry it, or is it outside this request?`
        : `${several ? "For each, should" : "Should"} the package add a document, or change a planned one, so that those parties take it, or is it outside this request? If the package should carry it, say in which document and who executes it.`].join(" "),
    context: options.saved
      ? "Drafting waits for your answer. The saved document plan, including any edits to it, is kept as it stands: to add a document, start a new draft, which plans the package from the start."
      : "Drafting waits for your answer. The document plan is drawn up with your answer before any part is drafted.",
    status: "open",
  }
}

/** The open items for the plan coverage the saved plan leaves open, derived
 * from coverage.json and never stored, so the drafting notes must name them and
 * the gate delivers them. Only a record of the saved plan counts. A plan
 * written by hand with no record was never checked and
 * gives none; any other plan whose record is missing, of another plan or
 * unreadable gives one item saying its coverage was not checked, never none.
 * A covered or uncovered record gives one item per act whose document the
 * check read the lawyer's words as leaving out, quoting those words as the
 * notes' quote check resolves them, for the lawyer to confirm, and asking the
 * lawyer also to confirm each of the act's leftOutDocuments the revisions kept
 * planned (stillPlanned); their issue gives no item of its own. When the record
 * also has the same act (sameAct) in planned documents, which only a last
 * attempt keeps (parseCoverage), the item names where the plan has it instead
 * of listing it as outstanding, and asks the lawyer to confirm whether it is to
 * be dealt with there or outside this package. "uncovered"
 * also gives one item per other act still without a document executed by the
 * parties who must take it, noting a package-scope question the lawyer
 * answered (never its answer) when that question settled the act
 * (actScopeAnswers: it showed the act (showsAct), or has the id the plan step
 * would ask with), and
 * one per other recorded issue; "unchecked" gives one item saying the check
 * gave no usable result. Either single item asks the lawyer to confirm
 * coverage. They belong to the drafting notes, so a repair goes to the
 * notes writer and a re-run of the draft step, which replaces draft-stage
 * items, keeps them.
 * Each id keys on the item's content. */
export function coverageItems(directory: string): OpenItem[] {
  const plan = readLeanPlan(directory)
  if (!plan) return []
  const saved = readLeanCoverage(directory)
  const coverage = saved?.plan === coveragePlanHash(plan.documents, plan.acting_for) ? saved : undefined
  const item = (detail: string) => openItem("finalize", "omission", detail, { repair_target: "memo" })
  const unknown = "so it is not known whether each act the instructions direct has a document executed by the parties who must take it. " +
    "The lawyer should confirm that each instructed act is carried by a document executed by the parties who must take it."
  // No record of the saved plan: the check never ran on it.
  if (!coverage) return plan.basis === undefined && readLeanText(directory, "coverage.json") === undefined ? []
    : [item(`Plan coverage was not checked against the saved document plan: no current record of that check exists, ${unknown}`)]
  if (coverage.status === "unchecked") return [item(`Plan coverage could not be checked: the check of the document plan against the instructions gave no usable result, ${unknown}`)]
  const issuesOf = (act: CoverageAct) => actIssues(act, plan.documents, coverage.acts)
  const open = coverage.acts.filter(act => !act.excluded_by && issuesOf(act).length)
  // Every act's own issue reaches the lawyer in that act's item.
  const tied = new Set(coverage.acts.flatMap(issuesOf))
  // "request" names a matter document of that name in the notes' quote check;
  // the left-out words were verified against the lawyer's own words only.
  const lawyerQuote = (words: string) => `"${words.replace(/\s+/g, " ").trim()}"` + (matterSourceIDs(directory).includes(REQUEST_SOURCE) ? "" : ` [source: ${REQUEST_SOURCE}]`)
  const settled = actScopeAnswers(open, readQuestions(directory).filter(answeredScopeQuestion), coverage.acts)
  const after = (act: CoverageAct) => {
    const asked = settled.by(act).map(question => question.id)
    return asked.length ? ` The lawyer answered the package-scope question ${asked.join(", ")}; the plan still leaves this act so.` : ""
  }
  const stillPlanned = (act: CoverageAct) => {
    const planned = leftOutDocuments(act, plan.documents, coverage.acts).map(document => `${label(document)}, which is executed by ${quotedList(document.executed_by!)}`)
    return planned.length > 1 ? ` The lawyer should also confirm whether each of ${serialList(planned)}, should be in this package; none of them carries an instructed act.`
      : planned.length ? ` The lawyer should also confirm whether ${planned[0]} and carries no instructed act, should be in this package.` : ""
  }
  const excludedReading = (act: CoverageAct, quote: string) => {
    // The same act in planned documents: a last attempt's other reading, kept beside the lawyer's words.
    const carried = coverage.acts.filter(other => !other.excluded_by && other.document !== null && sameAct(other) === sameAct(act))
    return carried.length
      ? "as leaving out the document for this instructed act, yet also read the plan as carrying it: " +
        `${actLine(act, plan.documents, quote, serialList(carried.map(other => plannedIn(other, plan.documents))))} ` +
        "The lawyer should confirm whether this act is to be dealt with in this package as planned or outside it."
      : `as leaving out the document for this instructed act, so it is listed as outstanding: ${actLine(act, plan.documents, quote)} ` +
        "The lawyer should confirm that this act is to be dealt with outside this package."
  }
  const items = [
    ...coverage.acts.filter(act => act.excluded_by).map(act => item(`Plan coverage: the check of the document plan against the instructions read the lawyer's words ${lawyerQuote(act.excluded_by!)} ` +
      excludedReading(act, act.quote.replace(/\s+/g, " ").trim()) + stillPlanned(act))),
    ...open.map(act => item("Plan coverage: the document plan leaves this instructed act without a document executed by the parties who must take it: " +
      actLine(act, plan.documents, act.quote.replace(/\s+/g, " ").trim()) + after(act))),
    ...coverage.issues.filter(issue => !tied.has(issue)).map(issue => item("Plan coverage: the check of the document plan against the instructions left this open: " +
      withoutOutlineFix(issue))),
  ]
  return [...new Map(items.map(row => [row.id, row])).values()]
}

/** Whether the saved plan's coverage record has the planned `document` carrying
 * every one of `refs`: each directs an act that document carries for the
 * parties who must take it, which the record names, and no act any of them
 * directs is left without a document executed by its parties. An act is
 * directed by a ref when its source is the ref's document and either
 * normalized quote contains the other. Only a record of the saved plan counts,
 * as in coverageItems; a missing, unreadable or earlier plan's record carries
 * nothing. An act whose document the lawyer's words leave out (excluded_by)
 * carries nothing, so an instruction whose only acts are left out, when the
 * checklist leaves it open, keeps the gate's non-blocking package-scope
 * question as the backstop. */
export function coverageCarries(directory: string, refs: { document: string; quote: string }[], document: DocID) {
  const plan = readLeanPlan(directory)
  const saved = readLeanCoverage(directory)
  if (!plan || !saved || saved.plan !== coveragePlanHash(plan.documents, plan.acting_for)) return false
  const contains = (a: string, b: string) => {
    const [left, right] = [normalize(a), normalize(b)]
    return !!left && !!right && (left.includes(right) || right.includes(left))
  }
  const directed = refs.map(ref => saved.acts.filter(act => ref.document === act.source && contains(ref.quote, act.quote)))
  return directed.length > 0 && !directed.flat().some(act => actIssues(act, plan.documents, saved.acts).length) &&
    directed.every(acts => acts.some(act => act.document === document && act.parties.length > 0))
}

/** Whether a package-scope question showed `act`, one of the record's `acts`:
 * one of its instruction quotes shows the act's quote (scopeQuestionShows) for
 * the act's parties, compared as nameKey does, so another party's act in the
 * same words is still asked. An act line of the plan step's names its parties;
 * a gate row, read whole, stands for every one of `acts` its quotes direct
 * (directedActs), so an answer on an instruction that also directs another
 * party's act settles neither party's act. */
function showsAct(question: Question, act: CoverageAct, acts: CoverageAct[]) {
  return scopeQuestionShows(question, [act.quote], { acts: [act], key: partyKeys, directed: quotes => directedActs(acts, quotes), act: true })
}

/** How the gate reads who must take its rows (scopeQuestionShows): a row
 * stands for the acts the saved plan's coverage record has its quotes
 * directing (directedActs) that the plan step asks about, those with an issue
 * (actIssues) or left out by the lawyer's words (excluded_by), so the plan
 * step's act lines show it only when they showed each of those acts for its
 * own parties. An act the plan carries for its parties is never listed, so it
 * needs no showing; a row whose quotes direct only such acts stands for every
 * act they direct. Only a record of the saved plan counts, as in
 * coverageItems; with none, or when the row's quotes direct no act it records,
 * no act line shows a row. */
export function gateScopeParties(directory: string): (quotes: string[]) => ScopeParties {
  const plan = readLeanPlan(directory)
  const saved = readLeanCoverage(directory)
  const acts = plan && saved?.plan === coveragePlanHash(plan.documents, plan.acting_for) ? saved.acts : []
  const asked = new Set(acts.filter(act => act.excluded_by || actIssues(act, plan!.documents, acts).length))
  const directed = (quotes: string[]) => directedActs(acts, quotes)
  return quotes => {
    const every = directed(quotes)
    const shown = every.filter(act => asked.has(act))
    return { acts: shown.length ? shown : every, key: partyKeys, directed, act: false }
  }
}

/** The `acts` any of `quotes` directs: whose quote holds it or is held by it (overlaps). */
const directedActs = (acts: CoverageAct[], quotes: string[]) => acts.filter(act => quotes.some(quote => overlaps(quote, act.quote)))

const partyKeys = (names: string[]) => JSON.stringify([...new Set(names.map(nameKey))].toSorted())

/** An act's key in the plan step's question id: its normalized quote with its
 * parties, compared as showsAct does, so the same act asks the same question
 * however its quote is spaced or its parties written or ordered, and another
 * party's act in the same words asks another. */
const actKeys = (act: CoverageAct) => [JSON.stringify([normalize(act.quote), partyKeys(act.parties)])]

/** The plan step's scopeAnswers on `acts`, of the record's `all`: settled by
 * the questions that showed each act (showsAct) or asked under the id actKeys
 * give them. The asker and the disclosure both read it, so the step asks under
 * the id it checks. */
const actScopeAnswers = (acts: CoverageAct[], answered: Question[], all: CoverageAct[]) =>
  scopeAnswers(acts, act => [act.quote], answered, "draft", (question, act) => showsAct(question, act, all), actKeys)

/** One instructed act as the lawyer reads it: what is to be done, the words
 * that direct it and their source, who must take it and where the plan has it
 * (by default, plannedIn). The description, parties and where the plan has it
 * are written from a model's words, so each goes through scopeField and
 * questionQuotes reads only the act lines this writes; the quote is verbatim. */
function actLine(act: CoverageAct, documents: PlannedDocument[], quote: string, where = plannedIn(act, documents)) {
  return `"${scopeField(act.act)}", directed by "${quote}" [source: ${act.source}]; to be taken by ${act.parties.length
    ? scopeField(quotedList(act.parties)) : "parties the instructions and sources do not name"}; ${scopeField(where)}.`
}

/** Where the plan has an act: the document that carries it and who executes that document. */
function plannedIn(act: CoverageAct, documents: PlannedDocument[]) {
  const carrier = documents.find(document => document.id === act.document)
  if (!carrier) return "no planned document carries it"
  const executing = carrier.executed_by
  return `planned in ${label(carrier)}, executed by ${!executing ? "parties the plan does not name" : executing.length ? quotedList(executing) : "no party"}`
}

/** What the coverage job reads of a planned document: no delivery or review fields. */
function coverageView(document: PlannedDocument) {
  return {
    id: document.id, title: document.title, purpose: document.purpose, executed_by: document.executed_by ?? null,
    parts: document.parts.map(part => ({ heading: part.heading, scope: part.scope })), ...(document.left_out?.length ? { left_out: document.left_out } : {}),
  }
}

/** The texts that can name a party: the lawyer's own words and each matter document. */
function namingTexts(directory: string) {
  return [...lawyerWords(directory), ...matterSourceIDs(directory).map(id => sourceText(directory, id))]
}

/** Content key of a checked plan: what the coverage job read of its documents
 * (coverageView) and whom the lawyer acts for. A document's delivery fields,
 * which the package review may write into the plan, do not change it. */
export function coveragePlanHash(documents: PlannedDocument[], actingFor: ActingFor | undefined) {
  return hashStageInputs(`${FIRST_DOCUMENT}.draft`, LEAN_COVERAGE_SCHEMA, [
    { name: "documents", bytes: JSON.stringify(documents.map(coverageView)) }, { name: "acting_for", bytes: JSON.stringify(actingFor ?? null) },
  ])
}

/** The saved plan's coverage is settled: it was written by hand, so it is
 * followed unchecked, or its record checked this very plan. Otherwise (a
 * missing, unreadable or earlier plan's record) the first document's draft step
 * is not done and checks the saved plan again. */
export function planCoverageCurrent(directory: string, plan = readLeanPlan(directory)) {
  return Boolean(plan) && (plan!.basis === undefined || readLeanCoverage(directory)?.plan === coveragePlanHash(plan!.documents, plan!.acting_for))
}

/** The saved coverage record, or undefined when it is missing or malformed,
 * including any act or issue not shaped as parseCoverage records it. */
export function readLeanCoverage(directory: string): LeanCoverage | undefined {
  const value = readLeanJson(directory, "coverage.json")
  if (!record(value) || value.schema !== LEAN_COVERAGE_SCHEMA || typeof value.plan !== "string" || !["covered", "unchecked", "uncovered"].includes(value.status as string) ||
    !Array.isArray(value.acts) || !value.acts.every(savedAct) || !strings(value.issues)) return
  return value as LeanCoverage
}

const strings = (value: unknown) => Array.isArray(value) && value.every(row => typeof row === "string")
const savedAct = (row: unknown) => record(row) && ["act", "source", "quote"].every(key => typeof row[key] === "string") && strings(row.parties) &&
  (row.document === null || typeof row.document === "string") && (row.excluded_by === undefined || (typeof row.excluded_by === "string" && row.document === null))

const label = (document: PlannedDocument) => `${document.id} "${document.title}"`

/** The one act an entry describes: what is to be done, ignoring spacing and
 * case, and its parties by nameKey. */
const sameAct = (act: Pick<CoverageAct, "act" | "parties">) => JSON.stringify([act.act.replace(/\s+/g, " ").toLowerCase(), act.parties.map(nameKey).toSorted()])

/** "A", "A" and "B", or "A", "B" and "C". */
function quotedList(names: string[]) {
  const quoted = names.map(name => `"${name}"`)
  return quoted.length > 1 ? `${quoted.slice(0, -1).join(", ")} and ${quoted.at(-1)}` : quoted.join("")
}
