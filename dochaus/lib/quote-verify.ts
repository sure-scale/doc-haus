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
import { chunkText } from "./chunk-text"

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
    .query("SELECT section, text, char_start, char_end FROM chunks WHERE doc_name = ? ORDER BY char_start, rowid")
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
  // (chunkText), then search the normalized whole and map the match back to
  // raw offsets to find the covering chunks.
  const raw = chunkText(chunks).text
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


/** An ellipsis: "...", "…" or spaced ". . .", optionally in square brackets;
 * a fourth dot (the period ending the sentence before it) is part of it. */
const ELLIPSIS = /\s*(?:\[\s*(?:(?:\. ?){2,3}\.|\u2026)\s*\]|(?:\. ?){2,3}\.|\u2026)\s*/

/** The normalized passages of a quotation that marks omitted words with an
 * ellipsis (ELLIPSIS), or undefined when it marks none. An ellipsis at either
 * end omits words outside the quotation and leaves no passage there, nor does
 * sentence punctuation beyond it ("A …." ends the writer's sentence). */
export function elidedParts(quote: string): string[] | undefined {
  const parts = normalize(quote).split(ELLIPSIS)
  if (parts.length < 2) return
  return parts.filter((part, index) => (index > 0 && index < parts.length - 1) || !/^[\s.,;:!?]*$/.test(part))
}

/** The most source text one ellipsis may stand for when its passages are
 * read as one passage: a citation or an aside, not the distance between two
 * parts of a document. */
export const MAX_OMITTED = 300

/** Word's hyphen and non-breaking hyphen (U+2010, U+2011), the figure dash, en
 * dash and minus sign (U+2012, U+2013, U+2212), which a page shows as or in
 * place of "-". Not the em dash, which sets off an aside. */
const HYPHENS = /[\u2010-\u2013\u2212]/g

/** Where an elided quotation's passages (elidedParts) sit in `raw`. It only
 * locates: words left out between passages can change who must do what or
 * drop a condition, amount or limit, and no list of words can tell which, so
 * only a single passage (an ellipsis at the start or end) verifies a quotation.
 *
 * Passages are matched as a quotation is (normalized, case-insensitively with
 * `ignoreCase`, HYPHENS folded to "-"), each long enough (SHORT_QUOTE) to find
 * its place on its own, beginning and ending on whole words of the source
 * (insideWord: a hyphen or dash joins words) and beginning after the previous
 * one ends.
 * - "near": every passage in the quoted order, each omission short
 *   (MAX_OMITTED) and within one paragraph (no blank line in the raw text). Of
 *   such placements, the one leaving out least, the earliest on a tie.
 *   `omitted` is the source words each ellipsis leaves out (whitespace
 *   normalized), joined with " … "; a single passage leaves out nothing.
 * - "apart": every passage in the quoted order, but only with a longer
 *   omission or one across a paragraph break.
 * - null: some passage is not in the source in the quoted order. */
export function elisionGap(raw: string, parts: string[], options: { unescapeMarkdown?: boolean; ignoreCase?: boolean } = {}): ElisionGap {
  if (!parts.length || parts.some(part => part.length < SHORT_QUOTE)) return null
  const normalized = normalizeWithMap(raw, options)
  const folded = { text: normalized.text.replace(HYPHENS, "-"), map: normalized.map }
  const { text, map } = options.ignoreCase ? lowerWithMap(folded) : folded
  const wanted = parts.map(part => part.replace(HYPHENS, "-")).map(part => options.ignoreCase ? lowerChars(part) : part)
  const places = wanted.map(part => {
    const found: number[] = []
    for (let at = text.indexOf(part); at !== -1; at = text.indexOf(part, at + 1))
      if (!insideWord(text, at) && !insideWord(text, at + part.length)) found.push(at)
    return found
  })
  // The source as written between the end of one passage and the start of the next.
  const between = (from: number, to: number) => raw.slice(map[from - 1]! + 1, map[to])
  const near = (from: number, to: number) => to - from <= MAX_OMITTED && !/\n[^\S\n]*\n/.test(between(from, to))
  // Each place of the latest passage placed, with the least it can leave out
  // after near placements of the passages before it.
  type Step = { at: number; end: number; left: number; before?: Step }
  let steps: Step[] = places[0]!.map(at => ({ at, end: at + wanted[0]!.length, left: 0 }))
  for (let index = 1; index < wanted.length; index++) {
    steps = places[index]!.flatMap(at => {
      const left = (step: Step) => step.left + at - step.end
      const before = steps.filter(step => step.end <= at && near(step.end, at))
        .reduce<Step | undefined>((best, step) => !best || left(step) < left(best) ? step : best, undefined)
      return before ? [{ at, end: at + wanted[index]!.length, left: left(before), before }] : []
    })
  }
  const last = steps.reduce<Step | undefined>((best, step) => !best || step.left < best.left ? step : best, undefined)
  if (last) {
    const omitted: string[] = []
    for (let step = last; step.before; step = step.before) omitted.unshift(normalizeWithMap(between(step.before.end, step.at), options).text)
    return { kind: "near", omitted: omitted.filter(Boolean).join(" \u2026 ") }
  }
  let from = 0
  for (const [index, part] of wanted.entries()) {
    const at = places[index]!.find(at => at >= from)
    if (at === undefined) return null
    from = at + part.length
  }
  return { kind: "apart" }
}

export type ElisionGap = { kind: "near"; omitted: string } | { kind: "apart" } | null

/** The nearest of several placements of one quotation: "near" over "apart"
 * over null. */
export function nearestGap(gaps: ElisionGap[]): ElisionGap {
  return gaps.find(gap => gap?.kind === "near") ?? gaps.find(gap => gap?.kind === "apart") ?? null
}

/** Whether one word of `text` runs across offset `at`: letters or digits on
 * both sides, or joined there by a hyphen (or a dash HYPHENS folds to one),
 * apostrophe or underscore ("non-exclusive", "Customer's"), or by a point or
 * comma between two digits ("2.5", "1,000,000"). */
function insideWord(text: string, at: number) {
  const word = (i: number) => /[\p{L}\p{N}]/u.test(text[i] ?? "")
  const digit = (i: number) => /\p{N}/u.test(text[i] ?? "")
  const joiner = (i: number) => /[-_'\u2019]/.test(text[i] ?? "") || /[.,]/.test(text[i] ?? "") && digit(i - 1) && digit(i + 1)
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

/** Where a quotation's elided passages (elidedParts) sit in the one named
 * document, an indexed matter document or the bound playbook (elisionGap). */
export function locateElidedQuote(args: { directory: string; parts: string[]; document: string }) {
  const playbook = boundPlaybook(args.directory)
  const texts = args.document.startsWith("playbook:")
    ? playbook?.text && args.document === playbook.source_id ? [playbook.text] : []
    : indexedText(args.directory, args.document)
  return nearestGap(texts.flatMap(raw => [{}, { unescapeMarkdown: true }].map(options => elisionGap(raw, args.parts, options))))
}

function indexedText(directory: string, document: string): string[] {
  const dbPath = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(dbPath)) return []
  const db = new Database(dbPath, { readonly: true })
  try {
    const chunks = db
      .query("SELECT section, text, char_start, char_end FROM chunks WHERE doc_name = ? ORDER BY char_start, rowid")
      .all(document) as QuoteChunk[]
    return chunks.length ? [chunkText(chunks).text] : []
  } finally {
    db.close()
  }
}
