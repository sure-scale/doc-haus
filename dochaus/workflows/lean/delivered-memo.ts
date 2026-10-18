import { deferredFindingItems } from "./change-pipeline"
import { finalBodyParagraphs, readChangeList, readParagraphs, record, standingIDPattern } from "./checks"
import { requestedMemo } from "./deliverables"
import { readLeanPlan, type DocID } from "./plan"
import { readRepairRound } from "./repairs"
import { SHORT_HASH_LENGTH, readLeanJson, readOpenItems, readQuestions, type OpenItem } from "./state"

// The drafting notes, and the memo they are written as when the request asks
// for one, refer to each open item in plain words with a trailing
// "[ref: <id>]". The working drafting-notes.md and .docx keep those markers:
// the notes check, the gate and the checklist check read them. Only the notes
// delivered at the gate drop them, and the lines the gate composes there name
// no internal id.

/** Whether the run's notes are delivered as a requested memo. */
export function memoMode(directory: string) {
  return Boolean(requestedMemo(directory))
}

/** A "[ref: …]" or "[refs: …]" marker, in any case. */
const REF = String.raw`\[[ \t]*[Rr][Ee][Ff][Ss]?[ \t]*:[^\]\n]*\]`

/** The notes as delivered: every "[ref: …]" marker removed whole. Their prose
 * is otherwise kept; the notes check keeps ids out of it. */
export function withoutRefMarkers(text: string) {
  return withoutInternalIDs(text, [])
}

/** Removes every "[ref: …]" marker and, outside quotations, every whole
 * occurrence of a known id, with a wrapping code mark, bold or quote: ids that
 * label the line or sit alone in brackets go whole; any other run of them is
 * named in plain words, or goes where a word beside it names its kind
 * (idPhrase). Text inside straight or curly double quotes is kept byte for
 * byte, though a quote holding only an id is that id. Only the lines that
 * changed are tidied. A line left with no content is dropped; every other line
 * keeps its exact text. */
export function withoutInternalIDs(text: string, ids: Iterable<string>) {
  const known = [...new Set([...ids].filter(id => id.length > 1))].toSorted((a, b) => b.length - a.length).map(escape).join("|")
  // A known id stands alone as the memo check reads it (standingIDPattern,
  // which bounds the whole wrapped token): not part of a longer id, word or
  // path, though joined to another short id by "/" or a dash. Known ids joined
  // that way ("a1-a3", "**a2/w1**") go as one, with their joints.
  const run = `(?:${known})(?:[/\\u2013-](?:${known}))*`
  const id = known ? standingIDPattern([...["\\*\\*", "\\*", "`", "\""].map(wrap => `${wrap}${run}${wrap}`), `\\u201c${run}\\u201d`, run].join("|")) : undefined
  // Ids listed with commas, "and", "or" or "&" form one run.
  const joint = String.raw`(?:[ \t]*,[ \t]*(?:(?:and|or)[ \t]+)?|[ \t]+(?:and|or|&)[ \t]+)`
  // A line that opens with references loses the separator that followed them.
  const leading = new RegExp(`^(?:[ \\t]*${REF}[ \\t]*[,;]?)+[ \\t]*(?:[:\\u2013\\u2014]|-(?=[ \\t]|$))?[ \\t]*`)
  const inline = new RegExp(`[ \\t]*${REF}`, "g")
  const quoted = new RegExp(`${known ? `("${run}"|\\u201c${run}\\u201d)|` : ""}"[^"\\n]*"|\\u201c[^\\u201d\\n]*\\u201d`, "g")
  const label = id && new RegExp(`^(?:[ \\t]*${id}[ \\t]*[,;]?)+[ \\t]*(?:[:\\u2013\\u2014]|-(?=[ \\t]|$)|$)[ \\t]*`)
  const bracketed = id && new RegExp(`[ \\t]*[([][ \\t]*(?:(?:[Ss]ee(?:[ \\t]+also)?|[Cc]f\\.|[Ii]ds?|[Rr][Ee][Ff][Ss]?)[ \\t]*:?[ \\t]*)?${id}(?:(?:${joint}|[ \\t]*;[ \\t]*)${id})*[ \\t]*[)\\]]`, "g")
  const listed = id && new RegExp(`${id}(?:${joint}${id})*`, "g")
  const each = new RegExp(standingIDPattern(known), "g")
  return text.split("\n").flatMap(line => {
    const [, prefix, content, trail] = /^([ \t]*(?:>[ \t]?)*(?:(?:[-*+]|\d+[.)])[ \t]+|#{1,6}[ \t]+)?)(.*?)([ \t\r]*)$/.exec(line)!
    const unmarked = content!.replace(leading, "").replace(inline, "")
    // Each quotation stands aside as one QUOTE mark until the line is tidied;
    // one holding only a known id (group 1) is that id. With no known ids the
    // pattern has no group, and the argument after the match is its offset.
    const quotes: string[] = []
    const shielded = unmarked.replace(quoted, (quote, id: unknown) => typeof id === "string" ? id : (quotes.push(quote), QUOTE))
    const named = label && bracketed && listed ? shielded.replace(label, "").replace(bracketed, "")
      .replace(listed, (match, offset: number, whole: string) => idPhrase([...match.matchAll(each)].map(row => row[0]), /[ \t]or[ \t]/.test(match), whole.slice(0, offset), whole.slice(offset + match.length))) : shielded
    if (unmarked === content && named === shielded) return [line]
    const tidied = tidy(named).replace(new RegExp(QUOTE, "g"), () => quotes.shift()!)
    return tidied ? [`${prefix}${tidied}${trail}`] : []
  }).join("\n")
}

/** A Unicode noncharacter, never in document text, holding a quotation's place. */
const QUOTE = "\uFDD0"

/** The id the run gives a question it asks: a fixed lowercase prefix and a
 * hash of shortHash's default length, as intake and the change steps mint "q-…",
 * and the package-scope, requested-files and checklist-conflict questions
 * theirs. The model never chooses one. A question saved under an id of any
 * other form, such as a word, is not matched, so no ordinary word is. */
const MINTED_QUESTION_ID = new RegExp(`^[a-z]+(?:-[a-z]+)*-[0-9a-f]{${SHORT_HASH_LENGTH}}$`)

/** Whether `id` has the form the run gives its own question ids. */
export function mintedQuestionID(id: string) {
  return MINTED_QUESTION_ID.test(id)
}

/** What a run id names, told apart by the shape its producer gives it: a
 * planned document "d1", a proposed change "a1" or a withheld one "w1", an
 * explanation "i1", a paragraph pointer, a question the run asked
 * (MINTED_QUESTION_ID), a checklist requirement (its refined and comparison
 * checks included) and an open item "stage:kind:hash". Each
 * kind is said with its article, after a determiner, and for several; `names`
 * are the words that already name it, beside which its id is only a label. A
 * memo finding is itself a requirement, so another one it cites is "another". */
const ID_KINDS = [
  { shape: /^d\d+$/, one: "a document", bare: "document", many: "documents", names: ["document"] },
  { shape: /^[aw]\d+$/, one: "a proposed change", bare: "proposed change", many: "proposed changes", names: ["change", "amendment"] },
  { shape: /^i\d+$/, one: "an explanation", bare: "explanation", many: "explanations", names: ["explanation"] },
  { shape: /^p[0-9a-f]{12}$/, one: "a paragraph", bare: "paragraph", many: "paragraphs", names: ["paragraph"] },
  { shape: MINTED_QUESTION_ID, one: "a question", bare: "question", many: "questions", names: ["question"] },
  { shape: /^source-[0-9a-f]{12}-\d+/, one: "another requirement", bare: "other requirement", many: "other requirements", names: ["requirement", "item"] },
  { shape: /^[^\s:]+:[^\s:]+:[^\s:]+$/, one: "an open item", bare: "open item", many: "open items", names: ["item"] },
]
const ANY_ID = { one: "an item", bare: "item", many: "items", names: ["item"] }

/** A run of ids, as found in prose between `before` and `after`, in plain
 * words: each kind once, in order, singular with its article (none after a
 * determiner or a name) or plural when the run names more than one of it; a
 * run joined by "or" names one. Capitalized where a sentence or a field
 * ("Reason:") starts. Where a word beside the run names its kind the ids only
 * label it and go: the word before ("Paragraph p0…"), or the word after in the
 * number the run takes ("the a1 change", "a1/a2 changes"; in the other number
 * it is a verb, "a1 changes the price", as is a singular one after an
 * auxiliary, "does a1 change it"), the ids then giving way to "the" where no
 * determiner or name leads them. */
function idPhrase(ids: string[], or: boolean, before: string, after: string) {
  const kinds = ids.map(id => ID_KINDS.find(kind => kind.shape.test(id)) ?? ANY_ID)
  const distinct = [...new Set(kinds)]
  const names = distinct.flatMap(kind => kind.names)
  const word = /([A-Za-z]+)[ \t]+$/.exec(before)?.[1]?.toLowerCase()
  if (word && names.some(name => word === name || word === `${name}s`)) return ""
  const led = DETERMINED.test(before) || nameLeads(before)
  const next = /^[ \t]+([A-Za-z]+)(?![\w-])/.exec(after)?.[1]?.toLowerCase()
  const plural = !or && ids.length > 1
  if (next && (plural || led || !AUXILIARY.test(before)) && names.some(name => next === (plural ? `${name}s` : name))) return led ? "" : OPENING.test(before) ? "The" : "the"
  const phrases = distinct.map((kind, index) => !or && kinds.filter(row => row === kind).length > 1 ? kind.many : led && !index ? kind.bare : kind.one)
  const text = phrases.length > 1 ? `${phrases.slice(0, -1).join(", ")} ${or ? "or" : "and"} ${phrases.at(-1)}` : phrases[0]!
  return OPENING.test(before) ? text[0]!.toUpperCase() + text.slice(1) : text
}

/** An auxiliary, a modal or "let" before an id no name leads: the word after it is a verb. */
const AUXILIARY = /\b(?:[Dd]o|[Dd]oes|[Dd]id|[Ww]ill|[Ww]ould|[Cc]an|[Cc]ould|[Mm]ay|[Mm]ight|[Mm]ust|[Ss]hall|[Ss]hould|[Ll]et)[ \t]+$/

/** A determiner or possessive just before an id: its name takes no article. */
const DETERMINED = /(?:\b(?:[Tt]he|[Aa]n?|[Tt]his|[Tt]hese|[Tt]hose|[Ee]ach|[Ee]very|[Aa]ny|[Ii]ts|[Tt]heir|[Oo]ur|[Yy]our|[Ww]hose)|'s)[ \t]+$/

/** Whether a name ends the text before an id, which the id is then part of,
 * so its name takes no article ("the Deed of Variation p0…", or "the Supply
 * note p0…" where the gate put a document's title for its id): a capitalized
 * word that opens no sentence, or one a determiner leads with one lowercase
 * word after it that is not a phrase of its own (PHRASE_WORD; "the Lease
 * omits a1"). Text a caller holds aside behind U+FDEF marks, as the memo's
 * findings hold the titles put in for ids, leads as a name. */
function nameLeads(before: string) {
  if (/\uFDEF[ \t]+$/.test(before)) return true
  const match = /(?<![^ \t])[A-Z][\w'-]*(?:[ \t]+([a-z][a-z'-]*))?[ \t]+$/.exec(before)
  if (!match) return false
  const head = before.slice(0, match.index)
  return !OPENING.test(head) && (!match[1] || DETERMINED.test(head) && !PHRASE_WORD.test(match[1]))
}

/** A word after a name that is not part of it: a closed-class word (an
 * auxiliary or modal, a preposition, a conjunction or a relative), or one
 * ending "-s" (not "-ss" or "-us") or "-ed" (not "-eed"), read as a verb or as
 * a plural heading its own phrase. */
const PHRASE_WORD = new RegExp(`^(?:${[
  "am|is|are|was|were|be|been|being|has|have|had|do|does|did|will|would|shall|should|may|might|must|can|could|cannot",
  "to|of|in|on|for|with|by|from|as|at|than|about|above|across|after|against|along|among|around|before|behind|below|beneath|beside|between|beyond|despite|during|except|including|into|onto|over|per|since|through|throughout|toward|under|until|upon|via|within|without",
  "and|or|nor|but|yet|so|if|unless|because|although|though|while|whereas|whether|when|where|once|not|that|which|who|whom",
].join("|")}|[a-z'-]*(?:[a-rtv-z]s|[^e]ed))$`)

/** The start of a line, a sentence or a capitalized field ("Reason:", "Open
 * decisions:"). A period after a dotted abbreviation ("e.g.") or a vowel-less
 * word in lower case ("cl.", "Sch.") ends no sentence; one after an acronym
 * ("PLC.") does. */
const OPENING = /^$|(?:(?<!\.[A-Za-z]|\b(?=[^\W\d_]*[a-z])[^\W\dAEIOUYaeiouy_]+)\.|[!?]|(?:^|[.!?][ \t]+)[A-Z][a-z]*(?:[ \t]+[a-z]+)?:)[ \t]+$/

/** Every internal id the run records that its notes may name: open, deferred
 * and repair items (and `items`, such as the gate's) with their amendment,
 * checklist and paragraph references, each planned document's id, paragraph
 * pointers and change-list ids, the checklist's item ids, and the question ids
 * the run minted (mintedQuestionID), saved or named by an item, which the
 * finalize prompt lists. */
export function runItemIDs(directory: string, items: OpenItem[] = []) {
  const repair = readRepairRound(directory)
  const checks = readLeanJson(directory, "checklist/checks.json")
  const extraction = readLeanJson(directory, "checklist/extraction.json")
  const rows = (value: unknown, key: string) => record(value) && Array.isArray(value[key]) ? value[key] as unknown[] : []
  const recorded = [...readOpenItems(directory), ...deferredFindingItems(directory), ...(repair?.items ?? []), ...items]
  return [...new Set([
    ...recorded.flatMap(item => [item.id, item.ref.amendment_id, item.ref.checklist_item, item.ref.pid]),
    ...(readLeanPlan(directory)?.documents ?? []).flatMap(({ id: document }) =>
      [document, ...[...readParagraphs(directory, document), ...finalBodyParagraphs(directory, document)].map(row => row.pid), ...changeListIDs(directory, document)]),
    ...(repair?.checklist_items ?? []),
    ...rows(checks, "expected_items"),
    ...rows(checks, "checks").map(row => record(row) ? row.item_id : undefined),
    ...rows(extraction, "sources").flatMap(source => rows(source, "items").map(row => record(row) ? row.id : undefined)),
    ...[...readQuestions(directory).map(question => question.id), ...recorded.map(item => item.ref.question_id)].filter(id => id !== undefined && mintedQuestionID(id)),
  ].filter((value): value is string => typeof value === "string" && value.length > 0))]
}

/** The change ids one planned document's change list records: its amendments,
 * withheld amendments and explanations, in that order. */
export function changeListIDs(directory: string, document: DocID) {
  const changes = readChangeList(directory, document)
  return [...(changes?.amendments ?? []).map(row => row.id), ...(changes?.withheld ?? []).map(row => row.amendment.id), ...(changes?.explanations ?? []).map(row => row.id)]
}

/** Tidies a line whose references were removed: emptied brackets, doubled or
 * stranded separators and the spacing left behind. */
function tidy(text: string) {
  return text
    .replace(/[ \t]*[([][ \t]*(?:(?:[Ss]ee(?:[ \t]+also)?|[Cc]f\.|[Ii]ds?|[Rr][Ee][Ff][Ss]?)[ \t]*:?[ \t]*)?(?:[,;][ \t]*)*[)\]]/g, "")
    .replace(/([,;])(?:[ \t]*[,;])+/g, "$1")
    .replace(/([([])[ \t]*[,;][ \t]*/g, "$1")
    .replace(/[ \t]*[,;:]+[ \t]*(?=[.!?)\]])/g, "")
    .replace(/[ \t]*[,;]+[ \t]*$/, "")
    .replace(/[ \t]+(?=[,.;:!?)\]])/g, "")
    .replace(/(?<!\.)\.\.(?!\.)/g, ".")
    .replace(/(\S)[ \t]{2,}(?=\S)/g, "$1 ")
    .trim()
}

function escape(text: string) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}
