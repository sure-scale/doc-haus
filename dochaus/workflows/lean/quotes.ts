import { boundPlaybook } from "../../lib/playbook-source"
import { approvedResearch, researchSource } from "./research"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { SHORT_QUOTE, elidedParts, elisionGap, locateElidedQuote, nearestGap, normalize, quotePlaces, verifyQuote } from "../../lib/quote-verify"
import type { DocID } from "./plan"
import { INSTRUCTIONS_FILE } from "./instructions"
import { leanUserUpdates } from "./requests"
import { readLeanText, readQuestions, shortHash, type OpenItem, type StageID } from "./state"
import { scanSourceQuotes, type SourceQuote, type SourceQuoteScanOptions } from "./source-quote-scan"

// Source-quote helpers shared by the lean stage checks and the final gate. A
// quote passes when lib/quote-verify finds it in the named matter document
// (verbatim, after folding curly quotes, whitespace and Markdown escapes in the
// source). A miss names the quote so the authoring model can correct it.

// Inside a quotation in double marks, quoted words are written in single ones:
// "the 'Price' is payable" quotes a source's the "Price" is payable. They end
// at the next single mark, past an apostrophe within a word ('Seller's Group').
const NESTED_QUOTE = /(?<![\p{L}\p{N}])['‘](?=\S)((?:[^'‘’"“”\n]|(?<=[\p{L}\p{N}])['’](?=[\p{L}\p{N}]))+?)(?<=\S)['’](?![\p{L}\p{N}])/gu
export const EXCERPT = 240

export type { SourceQuote } from "./source-quote-scan"

/** The source name under which the lawyer's own words are quoted: the request
 * and the lawyer's answers. A matter document of the same name takes precedence. */
export const REQUEST_SOURCE = "request"

/** The source name of one lawyer's answer: "answer:" and its question's id. */
export function answerSource(questionID: string) { return `answer:${questionID}` }
const ANSWER_SOURCE = /^answer:\S+$/
export function isAnswerSource(source_id: string) { return ANSWER_SOURCE.test(source_id) }

/** The source name of the recommendation adopted for a question the lawyer
 * delegated: "adopted:" and its question's id. Its words are the drafter's
 * judgment, never the lawyer's, so "request" and answerSource never hold them. */
export function adoptedSource(questionID: string) { return `adopted:${questionID}` }
const ADOPTED_SOURCE = /^adopted:\S+$/

/** The source names of the run's adopted recommendations, ordered by question id. */
export function adoptedSources(directory: string) {
  return adoptedQuestions(directory).map(question => adoptedSource(question.id))
}

function adoptedQuestions(directory: string) {
  return readQuestions(directory).filter(question => question.status === "answered" && question.answer_kind === "adopted" && question.answer)
    .toSorted((a, b) => a.id.localeCompare(b.id))
}

export function openItem(stage: StageID, kind: OpenItem["kind"], detail: string, ref: OpenItem["ref"]): OpenItem {
  return { id: `${stage}:${kind}:${shortHash(`${detail}\0${JSON.stringify(ref)}`)}`, stage, kind, detail, ref, status: "open" }
}

/** The item as found in one planned document's body. The document is part of
 * the reference and so of the id: the same finding in two documents is two
 * items, each disclosed and settled against its own body. */
export function inDocument(item: OpenItem, document: DocID): OpenItem {
  const ref = { ...item.ref, document }
  return { ...item, ...openItem(item.stage, item.kind, item.detail, ref), status: item.status }
}

// One synchronous operation's view of the matter's sources (basis, ids and
// checklist source inputs) and processed requests. Keying every step rereads
// them; inside a scope each is read once, and the scope ends before anything can
// await, so no write, lock release or later turn sees it. Values are shared
// within the scope, so they are deeply frozen: one caller cannot alter what
// another reads. Outside a scope every read is fresh (and frozen alike).
const snapshot: { reads?: Map<string, unknown> } = {}

/** Runs `operation` with the matter reads above memoized for its duration only.
 * A nested call joins the outer scope. An async operation is refused: the
 * snapshot must not outlive an await. */
export function withSourceSnapshot<T>(operation: () => T): T {
  const outer = snapshot.reads
  snapshot.reads ??= new Map()
  try {
    const value = operation()
    if (typeof (value as { then?: unknown } | null | undefined)?.then === "function") {
      void Promise.resolve(value).catch(() => undefined)
      throw new Error("withSourceSnapshot needs a synchronous operation; a source snapshot cannot span an await.")
    }
    return value
  } finally {
    snapshot.reads = outer
  }
}

/** The scope's value for `key`, read once; a fresh read outside any scope.
 * Either way the value is deeply frozen. */
export function snapshotRead<T>(key: string, read: () => T): T {
  const reads = snapshot.reads
  if (!reads) return deepFreeze(read())
  if (!reads.has(key)) reads.set(key, deepFreeze(read()))
  return reads.get(key) as T
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== "object" || value === null || ArrayBuffer.isView(value) || Object.isFrozen(value)) return value
  Object.freeze(value)
  for (const child of Object.values(value)) deepFreeze(child)
  return value
}

/** Canonical source ids include indexed document names and the selected library
 * playbook. read-document retrieves either; search-document addresses the index. */
export function matterSourceIDs(directory: string): string[] {
  return [...snapshotRead(`source-ids\0${directory}`, () => readMatterSourceIDs(directory))]
}

function readMatterSourceIDs(directory: string): string[] {
  const file = path.join(directory, ".dochaus", "legal.db")
  const playbook = boundPlaybook(directory)
  const policy = playbook?.text ? [playbook.source_id] : []
  if (!existsSync(file)) return policy
  const db = new Database(file, { readonly: true })
  try {
    return (db.query("SELECT DISTINCT doc_name FROM chunks ORDER BY doc_name").all() as { doc_name: string }[])
      .map(row => row.doc_name).concat(policy)
  } catch {
    return policy
  } finally {
    db.close()
  }
}

/** Content-keyed resume notices revised canonical text even when its source name stays the same. */
export function matterSourceBasis(directory: string): string {
  return snapshotRead(`source-basis\0${directory}`, () => readMatterSourceBasis(directory))
}

function readMatterSourceBasis(directory: string) {
  const file = path.join(directory, ".dochaus", "legal.db")
  const playbook = boundPlaybook(directory)
  if (!existsSync(file)) return playbook ? shortHash(JSON.stringify(playbook), 40) : ""
  const db = new Database(file, { readonly: true })
  try {
    const sources = db.query("SELECT doc_name, text FROM chunks ORDER BY doc_name, rowid").all()
    // Creating an empty index during delivery introduces no new evidence.
    return sources.length || playbook ? shortHash(JSON.stringify([sources, playbook]), 40) : ""
  } finally { db.close() }
}

/** Quotations written as "exact words" [source: document name]. */
export function sourceQuotes(text: string, options: SourceQuoteScanOptions = { format: "markdown" }): SourceQuote[] {
  return scanSourceQuotes(text, options).quotes
}

/** Both unreadable attempted citations and recognized misquotations receive
 * feedback; a syntax miss must never make the source check disappear. */
export function proseQuoteIssues(directory: string, where: string, text: string, options: SourceQuoteScanOptions = { format: "markdown" }) {
  const scanned = scanSourceQuotes(text, options)
  return [
    ...scanned.issues.map(issue => `${where} could not read a cited quotation: ${issue.message}. Write the exact quoted words followed by [source: exact source name].`),
    ...scanned.quotes.filter(quote => !proseQuoteFound(directory, quote)).map(quote => quoteMissIssue(directory, where, quote)),
  ]
}

export function quoteFound(directory: string, quote: SourceQuote) {
  if (/^https?:\/\//.test(quote.source_id)) return approvedResearch(directory).some(row => row.url === quote.source_id &&
    quotePlaces(normalize(researchSource(directory, row.snapshot)?.text ?? ""), normalize(quote.quote), { first: true }).places.length > 0)
  if (quote.source_id === REQUEST_SOURCE && !matterSourceIDs(directory).includes(REQUEST_SOURCE)) {
    return instructionFound(directory, quote.quote)
  }
  const answer = answerTexts(directory, quote.source_id) ?? adoptedTexts(directory, quote.source_id)
  if (answer) return lawyerWordsHold(answer, quote.quote)
  return verifyQuote({ directory, quote: quote.quote, document: quote.source_id }).found
}

/** User instructions have their own provenance: a matter source named
 * "request" must not substitute for what the lawyer actually asked. Their
 * words are matched in lower case, with a quotation's bracketed change of case
 * (quotePlaces). */
export function instructionFound(directory: string, quote: string) {
  return lawyerWordsHold(lawyerWords(directory), quote)
}

function lawyerWordsHold(texts: string[], quote: string) {
  const words = normalize(quote).toLowerCase()
  return words.length > 0 && texts.some(text => quotePlaces(normalize(text).toLowerCase(), words, { first: true }).places.length > 0)
}

/** Where the lawyer wrote `quote`: "request" when the request or an included
 * update holds it, otherwise the first answer that does (answerSource), or
 * undefined when none does. */
export function lawyerQuoteSource(directory: string, quote: string) {
  const passages = lawyerPassages(directory)
  const holds = (source_id: string) => lawyerWordsHold(passages.filter(row => row.source_id === source_id).map(row => row.text), quote)
  return [REQUEST_SOURCE, ...new Set(passages.map(row => row.source_id))].find(holds)
}

/** The lawyer's words of the answer `source_id` names, or undefined when it
 * names no answer (or a matter source has that name). */
function answerTexts(directory: string, source_id: string) {
  if (!isAnswerSource(source_id) || matterSourceIDs(directory).includes(source_id)) return
  return lawyerPassages(directory).filter(row => row.source_id === source_id).map(row => row.text)
}

/** The adopted option's words of the question `source_id` names ([] when it
 * names no adopted question), or undefined when it is not an adopted source
 * name (or a matter source has that name). */
function adoptedTexts(directory: string, source_id: string) {
  if (!ADOPTED_SOURCE.test(source_id) || matterSourceIDs(directory).includes(source_id)) return
  return adoptedQuestions(directory).filter(question => adoptedSource(question.id) === source_id).map(question => question.answer!)
}

/** A quotation in prose ('"words" [source: id]') passes when quoteFound does,
 * or when an ellipsis at its start or end leaves one passage that is in the
 * named source (elisionGap). An ellipsis in the middle is never accepted:
 * the words it leaves out can change who must do what or drop a condition,
 * amount or limit. quoteMissIssue then says what it leaves out. Quoted words
 * in single marks inside it may stand for the source's double marks
 * (NESTED_QUOTE), tried when the words as written are not found. A JSON quote
 * field keeps the verbatim contract of quoteFound, since its consumers locate
 * those words. The citation may add a pinpoint after the source name
 * (citedSource). */
export function proseQuoteFound(directory: string, quote: SourceQuote) {
  const cited = { ...quote, source_id: citedSource(directory, quote.source_id) }
  const nested = cited.quote.replace(NESTED_QUOTE, '"$1"')
  return [cited, ...(nested === cited.quote ? [] : [{ ...cited, quote: nested }])].some(reading => {
    if (quoteFound(directory, reading)) return true
    const elided = elision(directory, reading)
    return elided?.parts.length === 1 && elided.gap?.kind === "near"
  })
}

/** The source a prose citation names. A name the check knows is taken as
 * written. Otherwise the longest known name the citation starts with is taken
 * when the rest opens as a pinpoint ("Name (2).docx, cl. 4.2"); a name that
 * continues another way ("Name.docx.pdf") is a different source and kept, so
 * it fails as unknown. */
function citedSource(directory: string, source_id: string) {
  if (!/[,;]/.test(source_id)) return source_id
  const known = [...matterSourceIDs(directory), REQUEST_SOURCE, ...approvedResearch(directory).map(row => row.url)]
  if (known.includes(source_id)) return source_id
  return known.filter(name => source_id.startsWith(name) && /^\s*[,;]\s*\S/.test(source_id.slice(name.length)))
    .sort((a, b) => b.length - a.length)[0] ?? source_id
}

/** Where the passages of a quotation with an ellipsis sit in its named source
 * (elisionGap), or undefined for a quotation without one. */
function elision(directory: string, quote: SourceQuote) {
  const parts = elidedParts(quote.quote)
  if (!parts) return
  if (/^https?:\/\//.test(quote.source_id)) return { parts, gap: nearestGap(approvedResearch(directory).filter(row => row.url === quote.source_id)
    .map(row => elisionGap(researchSource(directory, row.snapshot)?.text ?? "", parts))) }
  if (quote.source_id === REQUEST_SOURCE && !matterSourceIDs(directory).includes(REQUEST_SOURCE))
    return { parts, gap: nearestGap(lawyerWords(directory).map(text => elisionGap(text, parts, { ignoreCase: true }))) }
  const answer = answerTexts(directory, quote.source_id) ?? adoptedTexts(directory, quote.source_id)
  if (answer) return { parts, gap: nearestGap(answer.map(text => elisionGap(text, parts, { ignoreCase: true }))) }
  return { parts, gap: locateElidedQuote({ directory, parts, document: quote.source_id }) }
}

/** The request, each included update and each answer, quoted one at a time so a
 * quotation never runs across two of them, with the lawyer's own words of a
 * delegating answer recorded as an adoption. The questions are the model's
 * words and are not included. */
export function lawyerWords(directory: string) {
  return lawyerPassages(directory).map(row => row.text)
}

/** lawyerWords, each with its source: "request" for the request and the
 * updates, answerSource for an answer. */
function lawyerPassages(directory: string): { source_id: string; text: string }[] {
  return [readLeanText(directory, "request.md") ?? "", ...leanUserUpdates(directory).filter(update => update.status === "included").map(update => update.text)]
    .map(text => ({ source_id: REQUEST_SOURCE, text })).concat(readQuestions(directory).flatMap(question =>
      question.status === "answered" && question.answer && question.answer_kind !== "instruction"
        // An adopted recommendation's words are the run's, not the lawyer's.
        ? [...question.answer_kind === "adopted" ? [] : [question.answer], ...(question.lawyer_words ? [question.lawyer_words] : [])].map(text => ({ source_id: answerSource(question.id), text }))
        : []))
}

/** When words quoted as the lawyer's are the run's own record (the instructions
 * list, a question an instruction settled, or an adopted recommendation),
 * where the lawyer's words are. An adopted recommendation's words are pointed
 * to their own source name (adoptedSource) where the check accepts it
 * (`adopted`), otherwise to the request or the matter sources. */
export function derivedRecordNote(directory: string, quote: string, { adopted: quotable = false } = {}) {
  const settled = readQuestions(directory).find(question => question.answer_kind === "instruction" && question.decided_by &&
    lawyerWordsHold([question.question, question.answer ?? "", question.decided_by.text], quote))
  if (settled?.decided_by)
    return `Those words are the run's own record of question ${settled.id}, which an instruction settled; they are not the lawyer's words. The lawyer's words behind it are "${excerpt(settled.decided_by.quote)}" in "${settled.decided_by.source_id}".`
  const adopted = readQuestions(directory).find(question => question.status === "answered" && question.answer_kind === "adopted" && lawyerWordsHold([question.answer ?? ""], quote))
  if (adopted)
    return `Those words are the recommendation adopted for question ${adopted.id}, which the lawyer delegated; they are not the lawyer's words. ` +
      (quotable ? `Quote them under the source name "${adoptedSource(adopted.id)}".` : "Quote the lawyer's request or the matter sources that support the point instead.")
  return lawyerWordsHold([readLeanText(directory, INSTRUCTIONS_FILE) ?? ""], quote)
    ? "Those words are the run's list of the lawyer's instructions, not the lawyer's words; quote the words the list shows from its named source, under that source's name."
    : ""
}

/** Actionable retry feedback for one unverified quote. When the words are in a
 * different matter document, that document is named. When an ellipsis in the
 * middle joins passages of the named source, the feedback says so and shows
 * the words it leaves out (or the source's own ellipsis, written differently),
 * or that the passages are apart. A passage beside an ellipsis that is too
 * short to locate (SHORT_QUOTE) is reported as such, not as missing. Words
 * quoted as the lawyer's that are an adopted recommendation are pointed to its
 * source name when the caller accepts it (`adopted`; derivedRecordNote). */
export function quoteMissIssue(directory: string, where: string, quote: SourceQuote, format: "prose" | "json" = "prose", options: { adopted?: boolean } = {}) {
  const request = quote.source_id === REQUEST_SOURCE && !matterSourceIDs(directory).includes(REQUEST_SOURCE)
  const answer = answerTexts(directory, quote.source_id) !== undefined
  const adopted = adoptedTexts(directory, quote.source_id) !== undefined
  const elided = elision(directory, quote)
  const gap = elided && elided.parts.length > 1 ? elided.gap : null
  if (gap) {
    const quoted = `${where} quotes "${excerpt(quote.quote)}" as from "${quote.source_id}".`
    const there = request ? "the request, a user update or one of the lawyer's answers" : answer ? "that answer" : adopted ? "that adopted recommendation" : "that document"
    if (gap.kind === "apart")
      return `${quoted} The quoted parts are in ${there}, but in separate passages. Quote each passage as its own quotation, without an ellipsis.`
    if (!gap.omitted)
      return `${quoted} The quoted words are one continuous passage of ${there}; the "…" leaves nothing out. Quote the passage without an ellipsis.`
    // What the "…" stands for is the source's own ellipsis ("..." for "…").
    if (elidedParts(gap.omitted)?.length === 0)
      return `${quoted} The quoted words are in ${there}, but it writes the ellipsis as "${gap.omitted}". Copy the exact characters, including that ellipsis.`
    return `${quoted} The quoted words are in ${there}, but the "…" leaves out: "${gap.omitted}". Leaving words out can change who must do what or drop a condition, amount or limit, so quotations with words left out in the middle are not accepted. Quote one continuous passage that includes those words, without an ellipsis.`
  }
  const elsewhere = [...new Set(verifyQuote({ directory, quote: quote.quote }).hits.map(hit => hit.doc))]
    .filter(doc => doc !== quote.source_id)
  const short = elided?.parts.some(part => part.length < SHORT_QUOTE)
  return [
    short
      ? `${where} quotes "${excerpt(quote.quote)}" as from "${quote.source_id}". A passage around the "…" is too short to locate on its own.`
      : request
        ? [`${where} quotes "${excerpt(quote.quote)}" as from "${REQUEST_SOURCE}", but those words are not in the request, user updates or any one of the lawyer's answers.`, derivedRecordNote(directory, quote.quote, options)].filter(Boolean).join(" ")
        : answer
        ? [`${where} quotes "${excerpt(quote.quote)}" as from "${quote.source_id}", but those words are not in the lawyer's answer to that question.`, derivedRecordNote(directory, quote.quote, options)].filter(Boolean).join(" ")
        : adopted
        ? `${where} quotes "${excerpt(quote.quote)}" as from "${quote.source_id}", but those words are not in the recommendation adopted for that question.`
        : `${where} quotes "${excerpt(quote.quote)}" as from "${quote.source_id}", but those words are not in that document.`,
    elsewhere.length
      ? `They do appear in ${elsewhere.map(doc => `"${doc}"`).join(", ")}; cite that document if it is the source you meant.`
      : format === "json"
        ? "The JSON quote field requires verbatim source words. Copy them with read-document or check them with verify-quote. Record paraphrases outside source-quote entries; removing JSON string delimiters does not make a paraphrase a valid quote."
        : short
          ? "Quote one continuous passage without an ellipsis, or drop the quotation marks."
          : "Copy the exact words with read-document (or check them with verify-quote), or drop the quotation marks and paraphrase.",
  ].join(" ")
}

/** Open items for quotations in `text` that the named source does not contain. */
export function quoteMissItems(directory: string, stage: StageID, text: string, ref: OpenItem["ref"] = {}, options: SourceQuoteScanOptions = { format: "markdown" }) {
  const scanned = scanSourceQuotes(text, options)
  return [
    ...scanned.issues.map(issue => openItem(stage, "quote_miss", `The text could not read a cited quotation: ${issue.message}.`, { ...ref, statement: issue.statement })),
    ...scanned.quotes.filter(quote => !proseQuoteFound(directory, quote))
      .map(quote => openItem(stage, "quote_miss", `Quoted text was not found verbatim in source "${quote.source_id}".`,
        { ...ref, source_id: quote.source_id, quote: quote.quote })),
  ]
}

export function excerpt(text: string, length = EXCERPT) {
  const flat = text.replace(/\s+/g, " ").trim()
  return flat.length > length ? `${flat.slice(0, length - 1)}…` : flat
}
