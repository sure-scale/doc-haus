import { boundPlaybook } from "../../lib/playbook-source"
import { approvedResearch, researchSource } from "./research"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { SHORT_QUOTE, elidedParts, elisionGap, locateElidedQuote, nearestGap, normalize, verifyQuote } from "../../lib/quote-verify"
import type { DocID } from "./plan"
import { leanUserUpdates } from "./requests"
import { readLeanText, readQuestions, shortHash, type OpenItem, type StageID } from "./state"

// Source-quote helpers shared by the lean stage checks and the final gate. A
// quote passes when lib/quote-verify finds it in the named matter document
// (verbatim, after folding curly quotes, whitespace and Markdown escapes in the
// source). A miss names the quote so the authoring model can correct it.

const SOURCE_QUOTE = /["“]([^"“”\n]{1,2000})["”]\s*[[(]\s*source\s*:\s*([^\])\n]+?)\s*[\])]/gi
export const EXCERPT = 240

export type SourceQuote = { quote: string; source_id: string }

/** The source name under which the lawyer's own words are quoted: the request
 * and the lawyer's answers. A matter document of the same name takes precedence. */
export const REQUEST_SOURCE = "request"

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
export function sourceQuotes(text: string): SourceQuote[] {
  return [...text.matchAll(SOURCE_QUOTE)].map(match => ({ quote: match[1]!.trim(), source_id: match[2]!.trim() }))
}

export function quoteFound(directory: string, quote: SourceQuote) {
  if (/^https?:\/\//.test(quote.source_id)) return approvedResearch(directory).some(row => row.url === quote.source_id &&
    normalize(researchSource(directory, row.snapshot)?.text ?? "").includes(normalize(quote.quote)))
  if (quote.source_id === REQUEST_SOURCE && !matterSourceIDs(directory).includes(REQUEST_SOURCE)) {
    return instructionFound(directory, quote.quote)
  }
  return verifyQuote({ directory, quote: quote.quote, document: quote.source_id }).found
}

/** User instructions have their own provenance: a matter source named
 * "request" must not substitute for what the lawyer actually asked. */
export function instructionFound(directory: string, quote: string) {
  const words = normalize(quote).toLowerCase()
  return words.length > 0 && lawyerWords(directory).some(text => normalize(text).toLowerCase().includes(words))
}

/** A quotation in prose ('"words" [source: id]') passes when quoteFound does,
 * or when an ellipsis at its start or end leaves one passage that is in the
 * named source (elisionGap). An ellipsis in the middle is never accepted:
 * the words it leaves out can change who must do what or drop a condition,
 * amount or limit. quoteMissIssue then says what it leaves out. A JSON quote
 * field keeps the verbatim contract of quoteFound, since its consumers locate
 * those words. */
export function proseQuoteFound(directory: string, quote: SourceQuote) {
  if (quoteFound(directory, quote)) return true
  const elided = elision(directory, quote)
  return elided?.parts.length === 1 && elided.gap?.kind === "near"
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
  return { parts, gap: locateElidedQuote({ directory, parts, document: quote.source_id }) }
}

/** The request, each included update and each answer, quoted one at a time so a
 * quotation never runs across two of them, with the lawyer's own words of a
 * delegating answer recorded as an adoption. The questions are the model's
 * words and are not included. */
export function lawyerWords(directory: string) {
  return [readLeanText(directory, "request.md") ?? "", ...leanUserUpdates(directory).filter(update => update.status === "included").map(update => update.text), ...readQuestions(directory).flatMap(question =>
    question.status === "answered" && question.answer ? [question.answer, ...(question.lawyer_words ? [question.lawyer_words] : [])] : [])]
}

/** Actionable retry feedback for one unverified quote. When the words are in a
 * different matter document, that document is named. When an ellipsis in the
 * middle joins passages of the named source, the feedback says so and shows
 * the words it leaves out (or the source's own ellipsis, written differently),
 * or that the passages are apart. A passage beside an ellipsis that is too
 * short to locate (SHORT_QUOTE) is reported as such, not as missing. */
export function quoteMissIssue(directory: string, where: string, quote: SourceQuote, format: "prose" | "json" = "prose") {
  const request = quote.source_id === REQUEST_SOURCE && !matterSourceIDs(directory).includes(REQUEST_SOURCE)
  const elided = elision(directory, quote)
  const gap = elided && elided.parts.length > 1 ? elided.gap : null
  if (gap) {
    const quoted = `${where} quotes "${excerpt(quote.quote)}" as from "${quote.source_id}".`
    const there = request ? "the request, a user update or one of the lawyer's answers" : "that document"
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
        ? `${where} quotes "${excerpt(quote.quote)}" as from "${REQUEST_SOURCE}", but those words are not in the request, user updates or any one of the lawyer's answers.`
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
export function quoteMissItems(directory: string, stage: StageID, text: string, ref: OpenItem["ref"] = {}) {
  return sourceQuotes(text)
    .filter(quote => !proseQuoteFound(directory, quote))
    .map(quote => openItem(stage, "quote_miss", `Quoted text was not found verbatim in source "${quote.source_id}".`,
      { ...ref, source_id: quote.source_id, quote: quote.quote }))
}

export function excerpt(text: string, length = EXCERPT) {
  const flat = text.replace(/\s+/g, " ").trim()
  return flat.length > length ? `${flat.slice(0, length - 1)}…` : flat
}
