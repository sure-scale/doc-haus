// Shared helpers for quote verification. Used by tool/verify-quote.ts and
// lib/verification-loop.ts (citation-anchors gate). The `normalize`/map helpers
// are pure; `verifyQuote` reads the per-matter legal.db and is the single
// source of truth for "does this quote appear verbatim" — both the tool and
// the verification-loop gate call it so the matching logic (single-chunk and
// cross-boundary) lives in one place.

import { boundPlaybook } from "./playbook-source"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"

export const SHORT_QUOTE = 20
export const MAX_RENDERED_HITS = 5

export function normalize(text: string) {
  return text
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/\s+/g, " ")
    .trim()
}

// ASCII punctuation a Markdown writer may escape with a backslash. Ingested
// sources converted through Markdown can carry `\.` or `\(` where the document
// shows `.` or `(`; a faithful quote of the document must still match.
const MARKDOWN_ESCAPABLE = /^[!-\/:-@\[-`{-~]$/

/** Normalized text plus, for each output character, its raw offset. With
 * `unescapeMarkdown`, a backslash before ASCII punctuation is dropped (source
 * side only: a quote is never unescaped, so a quote's own backslash still has to
 * be in the source). */
export function normalizeWithMap(raw: string, options: { unescapeMarkdown?: boolean } = {}) {
  let out = ""
  const map: number[] = []
  let spaceAt = -1
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i]!
    if (options.unescapeMarkdown && ch === "\\" && MARKDOWN_ESCAPABLE.test(raw[i + 1] ?? "")) continue
    if (/\s/.test(ch)) {
      if (out.length > 0 && spaceAt === -1) spaceAt = i
      continue
    }
    if (spaceAt !== -1) {
      out += " "
      map.push(spaceAt)
      spaceAt = -1
    }
    out += ch === "\u201C" || ch === "\u201D" ? '"' : ch === "\u2018" || ch === "\u2019" ? "'" : ch
    map.push(i)
  }
  return { text: out, map }
}

export type QuoteChunk = {
  section: string
  text: string
  char_start: number
  char_end: number
}

export type QuoteHit = {
  doc: string
  section: string
  charStart: number
  charEnd: number
  verbatim: string
}

export type VerifyQuoteResult = {
  found: boolean
  hits: QuoteHit[]
  // Why verification produced no hit, so callers can distinguish structural
  // absence (no DB / no documents) from a genuine miss. The verify-quote tool
  // surfaces these as distinct user-facing strings; the citation-anchors gate
  // treats no-db as not-applicable (the matter has not been indexed).
  reason?: "no-db" | "empty-quote" | "no-docs" | "not-found"
}

// Verify `quote` appears verbatim in the matter's indexed documents. Opens the
// per-matter legal.db read-only, scans each document for a single-chunk or
// cross-boundary match, and returns every document-level hit. Single source of
// truth for both tool/verify-quote (renders hits) and lib/verification-loop
// (citation-anchors gate treats `found` as pass/fail).
export function verifyQuote(args: {
  directory: string
  quote: string
  document?: string
}): VerifyQuoteResult {
  const needle = normalize(args.quote)
  const playbook = boundPlaybook(args.directory)
  const policyHits = playbook?.text && needle && (!args.document || args.document === playbook.source_id)
    ? findHits(playbook.source_id, [{ section: "Playbook", text: playbook.text, char_start: 0, char_end: playbook.text.length }], needle, {}) ?? []
    : []
  if (args.document?.startsWith("playbook:")) return policyHits.length
    ? { found: true, hits: policyHits }
    : { found: false, hits: [], reason: !needle ? "empty-quote" : !playbook?.text || args.document !== playbook.source_id ? "no-docs" : "not-found" }
  const indexed = verifyIndexedQuote(args)
  const hits = [...indexed.hits, ...policyHits]
  return hits.length ? { found: true, hits } : indexed
}

function verifyIndexedQuote(args: { directory: string; quote: string; document?: string }): VerifyQuoteResult {
  const dbPath = path.join(args.directory, ".dochaus", "legal.db")
  if (!existsSync(dbPath)) return { found: false, hits: [], reason: "no-db" }
  const needle = normalize(args.quote)
  if (!needle) return { found: false, hits: [], reason: "empty-quote" }

  const db = new Database(dbPath, { readonly: true })
  try {
    const docNames = (
      args.document
        ? db.query("SELECT DISTINCT doc_name FROM chunks WHERE doc_name = ?").all(args.document)
        : db.query("SELECT DISTINCT doc_name FROM chunks ORDER BY doc_name").all()
    ) as { doc_name: string }[]
    if (docNames.length === 0) return { found: false, hits: [], reason: "no-docs" }

    const hits = docNames.flatMap(({ doc_name }) => findHitsInDoc(db, doc_name, needle))
    return hits.length
      ? { found: true, hits }
      : { found: false, hits: [], reason: "not-found" }
  } finally {
    db.close()
  }
}

function findHitsInDoc(db: Database, docName: string, needle: string): QuoteHit[] {
  const chunks = db
    .query("SELECT section, text, char_start, char_end FROM chunks WHERE doc_name = ? ORDER BY char_start")
    .all(docName) as QuoteChunk[]
  return findHits(docName, chunks, needle, {}) ?? findHits(docName, chunks, needle, { unescapeMarkdown: true }) ?? []
}

function findHits(docName: string, chunks: QuoteChunk[], needle: string, options: { unescapeMarkdown?: boolean }): QuoteHit[] | undefined {
  const flat = (text: string) => (options.unescapeMarkdown ? normalizeWithMap(text, options).text : normalize(text))
  const single = chunks.find((c) => flat(c.text).includes(needle))
  if (single) {
    return [
      {
        doc: docName,
        section: single.section,
        charStart: single.char_start,
        charEnd: single.char_end,
        verbatim: single.text,
      },
    ]
  }
  // The quote may straddle a chunk boundary: rebuild the document text
  // (skipped gaps were whitespace-only at ingest, so pad with spaces to keep
  // stored offsets valid — see reconstructText in services/ingest/src/structure.ts),
  // append only the uncovered suffix of overlapping chunks, then search the
  // normalized whole and map the match back to raw offsets to find
  // the covering chunks.
  const raw = documentText(chunks)
  const normalized = normalizeWithMap(raw, options)
  const at = normalized.text.indexOf(needle)
  if (at === -1) return
  const rawStart = normalized.map[at]
  const rawLast = normalized.map[at + needle.length - 1]
  if (rawStart === undefined || rawLast === undefined) return
  const rawEnd = rawLast + 1
  const covering = chunks.filter((c) => c.char_start < rawEnd && c.char_end > rawStart)
  const first = covering[0]
  if (!first) return
  // A later-starting chunk may be wholly contained in an earlier one.
  // Its end is not necessarily the end of the matching evidence span.
  const end = Math.max(...covering.map(chunk => chunk.char_end))
  return [
    {
      doc: docName,
      section: first.section,
      charStart: first.char_start,
      charEnd: end,
      verbatim: raw.slice(first.char_start, end),
    },
  ]
}

/** A document's raw text rebuilt from its chunks in char_start order. */
function documentText(chunks: QuoteChunk[]) {
  return chunks.reduce(
    (acc, c) => c.char_start > acc.length
      ? acc + " ".repeat(c.char_start - acc.length) + c.text
      : acc + c.text.slice(Math.max(0, acc.length - c.char_start)),
    "",
  )
}

/** An ellipsis: "...", "…" or spaced ". . .", optionally in square brackets;
 * a fourth dot (the period ending the sentence before it) is part of it. */
const ELLIPSIS = /\s*(?:\[\s*(?:(?:\. ?){2,3}\.|\u2026)\s*\]|(?:\. ?){2,3}\.|\u2026)\s*/

/** The normalized passages of a quotation that marks omitted words with an
 * ellipsis (ELLIPSIS), or undefined when it marks none. An ellipsis at either
 * end omits words outside the quotation and leaves no passage there. */
export function elidedParts(quote: string): string[] | undefined {
  const parts = normalize(quote).split(ELLIPSIS)
  if (parts.length < 2) return
  return parts.filter((part, index) => part || (index > 0 && index < parts.length - 1))
}

/** The most source text one ellipsis may stand for: a citation or an aside,
 * not the distance between two parts of a document. */
export const MAX_OMITTED = 300

/** Omitted words that would negate, limit or condition the words around them:
 * negation, exclusion, restriction and condition. "No." before a number
 * abbreviates "number" and negates nothing. */
const QUALIFIER = /\b(?:not|no(?!\.\s?\d)|never|nor|neither|none|nothing|cannot|without|except|unless|save|only|subject\s+to|provided|notwithstanding|other\s+than|otherwise\s+than|rather\s+than|excluding|apart\s+from|solely|exclusively|if|to\s+the\s+extent|in\s+the\s+event)\b|n't\b/i

/** An omitted amount: a currency sign or a percentage. */
const AMOUNT = /[£$€¥]\s?\d|\d\s?%|\bper\s?cent\b/i

/** An omitted modal verb. Where the passage after the omission continues the
 * sentence, the omission may hand that verb to another subject or modal. */
const MODAL = /\b(?:shall|may|must|will|should)\b/i

/** An omission that ends on a bare "non" or "un" prefix of the next word
 * ("non exclusive", "un- limited"), the hyphen also written as an en dash. */
const DROPPED_PREFIX = /(?:^|[^\p{L}\p{N}])(?:non|un)[-\u2013]?\s*$/iu

/** A passage that resumes at a list item label: "(c) ", "4. ", "iv) ". */
const LIST_LABEL = /^\(?[a-z0-9ivx]{1,4}[.)]\s/

/** Word's hyphen and non-breaking hyphen (U+2010, U+2011), which a page shows
 * as "-". */
const HYPHENS = /[\u2010\u2011]/g

/** Whether every passage appears in `raw`, normalized like a quotation (and
 * case-insensitively with `ignoreCase`; hyphens folded to "-"), in the quoted
 * order, each starting after the previous one ends, and the ellipsis only
 * shortens the source:
 * - each passage is long enough (SHORT_QUOTE) to identify its place on its own
 *   and begins and ends on whole words of the source (insideWord);
 * - each omission between passages is short (MAX_OMITTED), stays within one
 *   paragraph (no blank line in the raw text, unless the next passage resumes
 *   at a list item label) and drops no negation, qualification, amount or
 *   word prefix that changes the meaning (QUALIFIER, AMOUNT, DROPPED_PREFIX),
 *   nor a modal verb when the next passage continues the sentence (MODAL).
 * Every place a passage appears is tried, not only the first. */
export function partsInOrder(raw: string, parts: string[], options: { unescapeMarkdown?: boolean; ignoreCase?: boolean } = {}) {
  if (!parts.length || parts.some(part => part.length < SHORT_QUOTE)) return false
  const normalized = normalizeWithMap(raw, options)
  const folded = { text: normalized.text.replace(HYPHENS, "-"), map: normalized.map }
  const { text, map } = options.ignoreCase ? lowerWithMap(folded) : folded
  const wanted = parts.map(part => part.replace(HYPHENS, "-")).map(part => options.ignoreCase ? lowerChars(part) : part)
  const omissionAllowed = (from: number, to: number, next: string) => {
    const omitted = text.slice(from, to)
    if (QUALIFIER.test(omitted) || AMOUNT.test(omitted) || DROPPED_PREFIX.test(omitted)) return false
    if (MODAL.test(omitted) && /^\p{Ll}/u.test(next)) return false
    return LIST_LABEL.test(next) || !/\n[^\S\n]*\n/.test(raw.slice(map[from - 1]! + 1, map[to]))
  }
  // A failed (passage, start) pair fails the same way from any earlier placement.
  const failed = new Set<string>()
  const place = (index: number, from: number): boolean => {
    if (failed.has(`${index}:${from}`)) return false
    const part = wanted[index]!
    for (let at = text.indexOf(part, from); at !== -1; at = text.indexOf(part, at + 1)) {
      if (index > 0 && at - from > MAX_OMITTED) break
      if (insideWord(text, at) || insideWord(text, at + part.length)) continue
      if (index > 0 && !omissionAllowed(from, at, parts[index]!)) continue
      if (index === wanted.length - 1 || place(index + 1, at + part.length)) return true
    }
    failed.add(`${index}:${from}`)
    return false
  }
  return place(0, 0)
}

/** Whether one word of `text` runs across offset `at`: letters or digits on
 * both sides, or joined there by a hyphen, apostrophe or underscore
 * ("non-exclusive", "Customer's"). */
function insideWord(text: string, at: number) {
  const word = (i: number) => /[\p{L}\p{N}]/u.test(text[i] ?? "")
  const joiner = (i: number) => /[-_'\u2019]/.test(text[i] ?? "")
  return word(at - 1) && (word(at) || joiner(at) && word(at + 1)) || joiner(at - 1) && word(at - 2) && word(at)
}

/** Lower case one character at a time, so offsets keep their raw positions. */
function lowerChars(text: string) {
  return Array.from(text, ch => ch.toLowerCase()).join("")
}

function lowerWithMap(normalized: { text: string; map: number[] }) {
  let text = ""
  const map: number[] = []
  let at = 0
  for (const ch of normalized.text) {
    const lower = ch.toLowerCase()
    text += lower
    for (let i = 0; i < lower.length; i++) map.push(normalized.map[at]!)
    at += ch.length
  }
  return { text, map }
}

/** Whether a quotation's elided passages (elidedParts) all appear, in order, in
 * the one named document: an indexed matter document or the bound playbook. */
export function verifyElidedQuote(args: { directory: string; parts: string[]; document: string }) {
  const playbook = boundPlaybook(args.directory)
  const texts = args.document.startsWith("playbook:")
    ? playbook?.text && args.document === playbook.source_id ? [playbook.text] : []
    : indexedText(args.directory, args.document)
  return texts.some(raw => [{}, { unescapeMarkdown: true }].some(options => partsInOrder(raw, args.parts, options)))
}

function indexedText(directory: string, document: string): string[] {
  const dbPath = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(dbPath)) return []
  const db = new Database(dbPath, { readonly: true })
  try {
    const chunks = db
      .query("SELECT section, text, char_start, char_end FROM chunks WHERE doc_name = ? ORDER BY char_start")
      .all(document) as QuoteChunk[]
    return chunks.length ? [documentText(chunks)] : []
  } finally {
    db.close()
  }
}
