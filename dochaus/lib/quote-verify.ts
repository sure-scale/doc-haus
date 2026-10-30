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

/** Word's hyphen and non-breaking hyphen (U+2010, U+2011), the figure dash, en
 * dash and minus sign (U+2012, U+2013, U+2212), which a page shows as or in
 * place of "-", read as "-" on both sides of every quote check. Not the em
 * dash, which sets off an aside. */
const HYPHENS = /[\u2010-\u2013\u2212]/g

export function normalize(text: string) {
  return text
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(HYPHENS, "-")
    .replace(/\s+/g, " ")
    .trim()
}

// ASCII punctuation a Markdown writer may escape with a backslash. Ingested
// sources converted through Markdown can carry `\.` or `\(` where the document
// shows `.` or `(`; a faithful quote of the document must still match.
const MARKDOWN_ESCAPABLE = /^[!-\/:-@\[-`{-~]$/

/** One reading of a source: the format markup it reads past, so a quote of the
 * words matches (source side only: a quote is never unescaped, so a quote's
 * own marks still have to be in the source). Each reading reads the source as
 * written (normalize) and, with
 * - `escapes`: without a backslash before ASCII punctuation;
 * - `markdown` (or its older name `unescapeMarkdown`): also without paired
 *   bold and emphasis marks (strongMarks, EMPHASIS) and a link's brackets and
 *   target (LINK), with ENTITIES and numeric entities decoded;
 * - `replies`: without an email's reply marks (REPLY_MARKS), and never across
 *   a change of reply depth (replyDepthChanges);
 * - `rules`: with a PDF's page-break rule (RULE_LINE, pageBreak) read as a space;
 * - `csv`: without a spreadsheet's CSV quoting (csvQuotes).
 * sourceReadings says which a document gets. */
export type SourceReading = { escapes?: boolean; markdown?: boolean; unescapeMarkdown?: boolean; replies?: boolean; rules?: boolean; csv?: boolean }

/** The readings a quotation of `document` (a matter document's name or a
 * playbook's source id) is matched against, in order, as ingest writes each
 * format's text: every source as written and without backslash escapes; Markdown
 * (a .md file, a playbook), an email body and a PDF's Markdown-like text also
 * without Markdown formatting; an email also without its reply marks; a PDF
 * also across its page breaks; a spreadsheet without its CSV quoting. A Word
 * document's text is mammoth's plain text, where "*", "[", "&gt;" and "---"
 * are what the page shows, so it and any other format are read strictly. */
export function sourceReadings(document?: string): SourceReading[] {
  const name = document?.toLowerCase() ?? ""
  const email = /\.(?:eml|msg)$/.test(name), pdf = name.endsWith(".pdf")
  return [
    {},
    { escapes: true },
    ...(email || pdf || /\.(?:md|markdown)$/.test(name) || name.startsWith("playbook:") ? [{ markdown: true }] : []),
    ...(email ? [{ markdown: true, replies: true }] : []),
    ...(pdf ? [{ markdown: true, rules: true }] : []),
    ...(name.endsWith(".xlsx") ? [{ csv: true }] : []),
  ]
}

/** Stands for a change of reply depth in an email's reading: never part of a
 * quote, which normalize reads as a space, so no quote matches across it. */
const DEPTH_CHANGE = "\u2029"

// A rule line ("-----"), read as a space where a blank line or the text's edge
// is either side of it, as LiteParse writes the break between PDF pages: a
// page break is no paragraph break. A table's "---" cell or a form's blank line
// has words beside it and stays.
const RULE_LINE = /(?<![^\n])[^\S\n]*-{3,}[^\S\n]*(?![^\n])/g
// Reply marks opening a line of an email: "> ", or "&gt; " (or "&amp;gt; ") in
// an HTML body.
const REPLY_MARKS = /(?<![^\n])[^\S\n]*(?:(?:>|&gt;|&amp;gt;)[^\S\n]*)+/g
// A single "*" either side of emphasised words on one line: opening after a
// space, a line's start or an opening bracket or quote mark, closing after
// neither, and before neither a word nor another "*". A "*" after a figure or
// a closing bracket ("5%*", "(per claim)*") marks a footnote and stays.
const EMPHASIS = /(?<![^\s([{"'\u201C\u2018])\*(?=[^\s*)\]}])[^*\n]*?[^\s*\\([{\u201C\u2018]\*(?![*\p{L}\p{N}])/gu
// A link, "[Schedule 2](#s2)", whose target is an anchor, a URL or a path, not
// joined to a word either side. Brackets without one ("[insert date]", "[●]",
// "clause [4](a)", "[£5m](TBC:board)") stay, as does a footnote link inside a
// figure ("£5[1](#fn1),000").
const LINK = /(?<![\\!\p{L}\p{N}])\[[^[\]\n]*\]\((?:#|(?:https?|ftp|file):\/\/|mailto:|tel:|\.{0,2}\/)[^()\s]*\)(?![\p{L}\p{N}])/giu
// Entities an HTML email body keeps, read as the characters they stand for, as
// are numeric ones ("&#8217;", "&#x2019;").
const ENTITIES = new Map(Object.entries({
  lt: "<", gt: ">", amp: "&", quot: '"', apos: "'", nbsp: "\u00A0", lsquo: "\u2018", rsquo: "\u2019",
  ldquo: "\u201C", rdquo: "\u201D", ndash: "\u2013", mdash: "\u2014", hellip: "\u2026",
}))
const ENTITY = /^&(?:#(\d{1,7})|#[xX]([\da-fA-F]{1,6})|([a-z]+));/

/** The characters an entity at `at` stands for and its raw length, if any. */
function entityAt(raw: string, at: number) {
  const match = ENTITY.exec(raw.slice(at, at + 10))
  if (!match) return
  const code = match[1] ? Number(match[1]) : match[2] ? parseInt(match[2], 16) : undefined
  const text = code === undefined ? ENTITIES.get(match[3]!) : code > 0 && code <= 0x10ffff && (code < 0xd800 || code > 0xdfff) ? String.fromCodePoint(code) : undefined
  return text === undefined ? undefined : { text, length: match[0].length }
}

/** The raw offsets of each `**` or `***` run that opens or closes bold words
 * on one line: an opening run not after a letter or digit and before a
 * non-space, then a closing run of the same length after a non-space and not
 * before a letter or digit. A run with no such partner ("£10,000,000**",
 * "10**6", "Cap**: £1m; Basket**: £100k", "pay *** within") stays, as does an
 * escaped one. */
function strongMarks(raw: string, skip: Set<number>) {
  const word = (ch?: string) => !!ch && /[\p{L}\p{N}]/u.test(ch)
  let open: { at: number; length: number } | undefined
  for (const run of raw.matchAll(/\*+|\n/g)) {
    if (run[0] === "\n") open = undefined
    if (run[0].length < 2 || run[0].length > 3 || raw[run.index - 1] === "\\") continue
    const at = run.index, length = run[0].length, before = raw[at - 1], after = raw[at + length]
    if (open?.length === length && /\S/.test(before ?? "") && !word(after)) {
      for (const from of [open.at, at]) for (let mark = from; mark < from + length; mark++) skip.add(mark)
      open = undefined
    } else if (!word(before) && /\S/.test(after ?? "")) open = { at, length }
  }
}

/** The line starts of an email where the reply depth (the number of reply
 * marks opening a line) differs from the last line with any text. */
function replyDepthChanges(raw: string, depths: Map<number, number>) {
  const changes = new Set<number>()
  let previous: number | undefined
  for (let start = 0, next = 0; start <= raw.length; start = next + 1) {
    next = raw.indexOf("\n", start)
    if (next === -1) next = raw.length
    if (!/\S/.test(raw.slice(start, next))) continue
    const depth = depths.get(start) ?? 0
    if (previous !== undefined && depth !== previous) changes.add(start)
    previous = depth
  }
  return changes
}

/** Whether the rule line `raw.slice(from, to)` has a blank line or the text's
 * edge either side. */
function pageBreak(raw: string, from: number, to: number) {
  const before = from === 0 ? "" : raw.slice(raw.lastIndexOf("\n", from - 2) + 1, from - 1)
  const next = raw.indexOf("\n", to + 1)
  const after = to >= raw.length ? "" : raw.slice(to + 1, next === -1 ? raw.length : next)
  return !/\S/.test(before) && !/\S/.test(after)
}

/** The raw offsets a reading reads as a space (`space`: page-break rules) or
 * leaves out (`skip`: bold and emphasis marks, a link's brackets and target,
 * reply marks), and the line starts where the reply depth changes (`depth`). */
function readingMarks(raw: string, reading: { markdown: boolean; replies?: boolean; rules?: boolean }) {
  const space = new Set<number>(), skip = new Set<number>()
  const add = (set: Set<number>, from: number, to: number) => { for (let at = from; at < to; at++) set.add(at) }
  if (reading.rules) for (const match of raw.matchAll(RULE_LINE))
    if (pageBreak(raw, match.index, match.index + match[0].length)) add(space, match.index, match.index + match[0].length)
  if (reading.markdown) {
    strongMarks(raw, skip)
    for (const match of raw.matchAll(EMPHASIS)) skip.add(match.index).add(match.index + match[0].length - 1)
    for (const match of raw.matchAll(LINK)) {
      skip.add(match.index)
      add(skip, match.index + match[0].indexOf("]("), match.index + match[0].length)
    }
  }
  const depths = new Map<number, number>()
  if (reading.replies) for (const match of raw.matchAll(REPLY_MARKS)) {
    add(skip, match.index, match.index + match[0].length)
    depths.set(match.index, match[0].match(/&amp;gt;|&gt;|>/g)!.length)
  }
  return { space, skip, depth: reading.replies ? replyDepthChanges(raw, depths) : undefined }
}

/** The raw offsets of the quote marks a spreadsheet's CSV text (ingest writes
 * each sheet with sheet_to_csv) puts around a cell holding a comma, quote or
 * line break, and of the first of each doubled quote inside one, so the cells
 * read as the sheet shows them. Read from the start of the text, as a quoted
 * cell can hold a line break. A cell keeps its marks when its digits would run
 * into a neighbouring cell's across the comma between them: cells "5,000" and
 * "250" never read as 5,000,250. */
function csvQuotes(raw: string) {
  const skip = new Set<number>()
  type Cell = { marks: number[]; value: string }
  const runOn = (a?: Cell, b?: Cell) => !!a && !!b && /\p{N}$/u.test(a.value) && /^\p{N}/u.test(b.value)
  let row: Cell[] = []
  for (let at = 0; at <= raw.length; at++) {
    const cell: Cell = { marks: [], value: "" }
    let end = at
    if (raw[at] === '"') {
      cell.marks.push(at)
      for (end = at + 1; end < raw.length && (raw[end] !== '"' || raw[end + 1] === '"'); end++) {
        if (raw[end] === '"') cell.marks.push(end++)
        cell.value += raw[end]
      }
      // Its closing mark ends the cell; a cell with other text there is read as written.
      if (end < raw.length && /^[,\n]?$/.test(raw[end + 1] ?? "")) cell.marks.push(end++)
      else {
        cell.marks = []
        end = at
      }
    }
    if (!cell.marks.length) {
      while (end < raw.length && raw[end] !== "," && raw[end] !== "\n") end++
      cell.value = raw.slice(at, end)
    }
    row.push(cell)
    if (raw[end] !== ",") {
      row.forEach((cell, index) => {
        if (!runOn(row[index - 1], cell) && !runOn(cell, row[index + 1])) cell.marks.forEach(mark => skip.add(mark))
      })
      row = []
    }
    at = end
  }
  return skip
}

/** Normalized text in one `reading` (SourceReading) plus, for each output
 * character, the raw offset it starts at (`map`) and the one just past its last
 * raw character (`end`: past the whole of a decoded entity), so `raw.slice(
 * map[first], end[last])` is the source of a match as written. A space
 * stands for a whitespace run, read past marks included; `breaks` are the
 * output offsets of those that hold a paragraph break (a blank line, but not
 * one around a page-break rule the reading reads as a space) and of each
 * DEPTH_CHANGE. An escaped character is literal ("\*"); underscores stay, as
 * they mark blanks. */
export function normalizeWithMap(raw: string, options: SourceReading = {}) {
  const markdown = !!(options.markdown || options.unescapeMarkdown)
  const unescape = markdown || !!options.escapes
  const marks = markdown || options.replies || options.rules ? readingMarks(raw, { markdown, replies: options.replies, rules: options.rules }) : undefined
  const quotes = options.csv ? csvQuotes(raw) : undefined
  let out = ""
  const map: number[] = [], end: number[] = [], breaks: number[] = []
  // The pending whitespace run: where it starts, whether it holds a blank
  // line, a page-break rule or a reply-depth change, and whether its current
  // line is blank so far.
  let spaceAt = -1, paragraph = false, rule = false, depth = false, blank = false
  const emit = (text: string, from: number, to: number) => {
    if (spaceAt !== -1) {
      if (depth || paragraph && !rule) breaks.push(out.length)
      out += depth ? DEPTH_CHANGE : " "
      map.push(spaceAt)
      end.push(spaceAt + 1)
      spaceAt = -1
    }
    paragraph = rule = depth = blank = false
    const folded = text.replace(/[\u201C\u201D]/g, '"').replace(/[\u2018\u2019]/g, "'").replace(HYPHENS, "-")
    for (let at = 0; at < folded.length; at++) {
      map.push(from)
      end.push(to)
    }
    out += folded
  }
  for (let i = 0; i < raw.length; i++) {
    if (marks?.depth?.has(i)) depth = true
    const ch = raw[i]!
    if (quotes?.has(i) || marks?.skip.has(i)) continue
    if (unescape && ch === "\\" && MARKDOWN_ESCAPABLE.test(raw[i + 1] ?? "")) {
      emit(raw[i + 1]!, i + 1, i + 2)
      i++
      continue
    }
    const entity = markdown && ch === "&" ? entityAt(raw, i) : undefined
    const isRule = !!marks?.space.has(i)
    if (/\s/.test(entity?.text ?? ch) || isRule) {
      if (out.length > 0 && spaceAt === -1) spaceAt = i
      if (isRule) {
        rule = true
        blank = false
      } else if (ch === "\n") {
        paragraph ||= blank
        blank = true
      }
      if (entity) i += entity.length - 1
      continue
    }
    emit(entity?.text ?? ch, i, i + (entity?.length ?? 1))
    if (entity) i += entity.length - 1
  }
  return { text: out, map, end, breaks }
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
  const policyChunks = playbook?.text ? [{ section: "Playbook", text: playbook.text, char_start: 0, char_end: playbook.text.length }] : []
  const policyHits = playbook?.text && needle && (!args.document || args.document === playbook.source_id)
    ? firstHits(playbook.source_id, policyChunks, needle)
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
  return firstHits(docName, chunks, needle)
}

/** The hits in the first of a document's readings (sourceReadings) that holds `needle`. */
function firstHits(docName: string, chunks: QuoteChunk[], needle: string): QuoteHit[] {
  for (const reading of sourceReadings(docName)) {
    const hits = findHits(docName, chunks, needle, reading)
    if (hits) return hits
  }
  return []
}

function findHits(docName: string, chunks: QuoteChunk[], needle: string, reading: SourceReading): QuoteHit[] | undefined {
  // Only the reading as written reads one chunk at a time: the others depend
  // on what is around a character (a line's start, an escaping backslash, a
  // closing mark, a quoted cell), which a chunk's edges can cut off, so they
  // read the whole document, below.
  const asWritten = !Object.values(reading).some(Boolean)
  const single = asWritten ? chunks.find((c) => quotePlaces(normalize(c.text), needle, { first: true }).places.length) : undefined
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
  const normalized = normalizeWithMap(raw, reading)
  const { places: [at], length } = quotePlaces(normalized.text, needle, { first: true })
  if (at === undefined) return
  const rawStart = normalized.map[at]
  const rawEnd = normalized.end[at + length - 1]
  if (rawStart === undefined || rawEnd === undefined) return
  const covering = chunks.filter((c) => c.char_start < rawEnd && c.char_end > rawStart)
  // A match one chunk holds is that chunk's hit, as in a single-chunk match.
  const whole = asWritten ? undefined : covering.find((c) => c.char_start <= rawStart && c.char_end >= rawEnd)
  if (whole) return [{ doc: docName, section: whole.section, charStart: whole.char_start, charEnd: whole.char_end, verbatim: whole.text }]
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

/** Letters a quotation puts in square brackets ("[d]raft", "[T]he"): the
 * quoter's mark of a change of case. */
const CASE_CHANGE = /\[(\p{L}+)\]/gu

/** The offsets in `text` (a reading of a source) of `needle` (a normalized
 * quotation), each `length` characters of `text` long, in either case with
 * `ignoreCase`, and only the first with `first`. Where `needle` is not in
 * `text` as written, it is read without the brackets of each change of case
 * (CASE_CHANGE): their letters match the source's letters there in either
 * case, so "[d]raft" quotes "Draft". Brackets around anything else, and
 * letters other than the source's ("[the Company]", "[x]raft"), stay. */
export function quotePlaces(text: string, needle: string, options: { ignoreCase?: boolean; first?: boolean } = {}) {
  const read = (value: string) => options.ignoreCase ? lowerChars(value) : value
  const haystack = read(text)
  const places = (within: string, wanted: string, kept: (at: number) => boolean = () => true) => {
    const found: number[] = []
    for (let at = within.indexOf(wanted); at !== -1 && !(options.first && found.length); at = within.indexOf(wanted, at + 1))
      if (kept(at)) found.push(at)
    return found
  }
  const exact = places(haystack, read(needle))
  // The needle without its case-change brackets, and the offsets in it of
  // the letters they held.
  let plain = "", from = 0
  const changed = new Set<number>()
  for (const match of needle.matchAll(CASE_CHANGE)) {
    plain += needle.slice(from, match.index)
    for (let at = 0; at < match[1]!.length; at++) changed.add(plain.length + at)
    plain += match[1]
    from = match.index + match[0].length
  }
  if (exact.length || !changed.size) return { places: exact, length: needle.length }
  plain = read(plain + needle.slice(from))
  const asWritten = (at: number) => {
    for (let index = 0; index < plain.length; index++) if (!changed.has(index) && haystack[at + index] !== plain[index]) return false
    return true
  }
  return { places: places(options.ignoreCase ? haystack : lowerChars(text), lowerChars(plain), asWritten), length: plain.length }
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

/** Where an elided quotation's passages (elidedParts) sit in `raw`. It only
 * locates: words left out between passages can change who must do what or
 * drop a condition, amount or limit, and no list of words can tell which, so
 * only a single passage (an ellipsis at the start or end) verifies a quotation.
 *
 * Passages are matched as a quotation is (normalized, with any bracketed
 * change of case: quotePlaces, case-insensitively with `ignoreCase`), each
 * long enough (SHORT_QUOTE) to find its place on its own,
 * beginning and ending on whole words of the source (insideWord: a hyphen or
 * dash joins words) and beginning after the previous one ends.
 * - "near": every passage in the quoted order, each omission short
 *   (MAX_OMITTED) and within one paragraph of the reading (no break of
 *   normalizeWithMap: a blank line, but not one around a page-break rule the
 *   reading reads as a space, nor a change of reply depth). Of such
 *   placements, the one leaving out least, the earliest on a tie.
 *   `omitted` is the source words each ellipsis leaves out (whitespace
 *   normalized), joined with " … "; a single passage leaves out nothing.
 * - "apart": every passage in the quoted order, but only with a longer
 *   omission or one across a paragraph break.
 * - null: some passage is not in the source in the quoted order. */
export function elisionGap(raw: string, parts: string[], options: SourceReading & { ignoreCase?: boolean } = {}): ElisionGap {
  if (!parts.length || parts.some(part => part.length < SHORT_QUOTE)) return null
  const normalized = normalizeWithMap(raw, options)
  const text = normalized.text
  const breaks = new Set(normalized.breaks)
  const found = parts.map(part => quotePlaces(text, part, { ignoreCase: options.ignoreCase }))
  // The source length of each passage (quotePlaces), past which the next begins.
  const lengths = found.map(({ length }) => length)
  const places = found.map(({ places, length }) => places.filter(at => !insideWord(text, at) && !insideWord(text, at + length)))
  const parted = (from: number, to: number) => {
    for (let at = from; at < to; at++) if (breaks.has(at)) return true
    return false
  }
  const near = (from: number, to: number) => to - from <= MAX_OMITTED && !parted(from, to)
  // Each place of the latest passage placed, with the least it can leave out
  // after near placements of the passages before it.
  type Step = { at: number; end: number; left: number; before?: Step }
  let steps: Step[] = places[0]!.map(at => ({ at, end: at + lengths[0]!, left: 0 }))
  for (let index = 1; index < parts.length; index++) {
    steps = places[index]!.flatMap(at => {
      const left = (step: Step) => step.left + at - step.end
      const before = steps.filter(step => step.end <= at && near(step.end, at))
        .reduce<Step | undefined>((best, step) => !best || left(step) < left(best) ? step : best, undefined)
      return before ? [{ at, end: at + lengths[index]!, left: left(before), before }] : []
    })
  }
  const last = steps.reduce<Step | undefined>((best, step) => !best || step.left < best.left ? step : best, undefined)
  if (last) {
    const omitted: string[] = []
    for (let step = last; step.before; step = step.before) omitted.unshift(normalized.text.slice(step.before.end, step.at).trim())
    return { kind: "near", omitted: omitted.filter(Boolean).join(" \u2026 ") }
  }
  let from = 0
  for (const [index, length] of lengths.entries()) {
    const at = places[index]!.find(at => at >= from)
    if (at === undefined) return null
    from = at + length
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

/** Lower case one character at a time, keeping any whose lower case is longer
 * ("\u0130"), so offsets into `text` hold. */
function lowerChars(text: string) {
  return Array.from(text, ch => {
    const lower = ch.toLowerCase()
    return lower.length === ch.length ? lower : ch
  }).join("")
}

/** Where a quotation's elided passages (elidedParts) sit in the one named
 * document, an indexed matter document or the bound playbook (elisionGap). */
export function locateElidedQuote(args: { directory: string; parts: string[]; document: string }) {
  const playbook = boundPlaybook(args.directory)
  const texts = args.document.startsWith("playbook:")
    ? playbook?.text && args.document === playbook.source_id ? [playbook.text] : []
    : indexedText(args.directory, args.document)
  return nearestGap(texts.flatMap(raw => sourceReadings(args.document).map(reading => elisionGap(raw, args.parts, reading))))
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
