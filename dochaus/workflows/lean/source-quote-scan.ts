import { parse, postprocess, preprocess } from "micromark"

export type SourceQuote = { quote: string; source_id: string }
export type SourceQuoteIssue = { message: string; statement: string }
export type SourceQuoteScanOptions = { format: "markdown" | "plain" }

const QUOTE_CAP = 2000
const SOURCE_CAP = 4096
const INNER_CAP = 500
// This finds the explicit marker vocabulary only. Balanced names and quotation
// boundaries are read below rather than inferred by a semantic regex.
const SOURCE_MARKER = /[[(]\s*source\s*:/giu
const opening = (character: string) => character === '"' || character === "“"
const closing = (character: string) => character === '"' || character === "”"
const space = (character: string | undefined) => character !== undefined && /\s/u.test(character)
const word = (character: string | undefined) => character !== undefined && /[\p{L}\p{N}]/u.test(character)

/** Source-backed prose quotations, with unreadable attempted citations kept as
 * explicit issues. Literal code and comments are classified by CommonMark.
 * Every marker reads at most one bounded quote and source name; a long line of
 * unrelated quotation marks never causes a new unbounded backwards search. */
export function scanSourceQuotes(text: string, options: SourceQuoteScanOptions = { format: "plain" }): { quotes: SourceQuote[]; issues: SourceQuoteIssue[] } {
  const markers = [...text.matchAll(SOURCE_MARKER)]
  if (!markers.length) return { quotes: [], issues: [] }
  const literals = options.format === "markdown" ? literalRanges(text) : []
  let literal = 0
  let consumed = 0
  let previous: { quotation: { start: number; quote: string }; source: string; end: number } | undefined
  const quotes: (SourceQuote & { marker: number })[] = []
  const issues: (SourceQuoteIssue & { marker: number })[] = []
  const quotationRanges: { start: number; end: number }[] = []
  for (const marker of markers) {
    const start = marker.index
    if (start < consumed) continue
    while (literals[literal]?.end <= start) literal++
    if (literals[literal]?.start <= start && start < literals[literal].end) continue
    const source = sourceName(text, start, start + marker[0].length)
    consumed = source.end
    let end = start - 1
    while (end >= 0 && space(text[end])) end--
    const prior = previous && !text.slice(previous.end, start).trim() ? previous : undefined
    const quotation = prior?.quotation ?? (closing(text[end]) ? quotationBefore(text, end) : undefined)
    const message = source.error ?? (!prior && !closing(text[end]) ? "the source marker does not follow a closing quotation mark"
      : !quotation ? `no opening quotation mark was found within the ${QUOTE_CAP}-character quotation limit`
        : !quotation.quote.trim() ? "the quotation is empty" : undefined)
    if (message) {
      // Keep the fallback statement search bounded too: repeated malformed
      // markers on a long line must not each re-read its entire prefix.
      const floor = Math.max(0, end - QUOTE_CAP)
      const preceding = text.slice(floor, start).lastIndexOf("\n\n")
      const paragraph = preceding < 0 ? floor : floor + preceding + 2
      issues.push({ marker: start, message, statement: text.slice(quotation?.start ?? paragraph, source.end) })
      previous = undefined
      continue
    }
    quotationRanges.push({ start: quotation!.start + 1, end: quotation!.start + 1 + quotation!.quote.length })
    // Several immediately adjacent source markers can support one quotation.
    // Every different source is verified; a repeated identical marker adds no
    // duplicate finding for the same words.
    if (!prior || prior.source !== source.name)
      quotes.push({ marker: start, quote: quotation!.quote.trim(), source_id: source.name })
    previous = { quotation: quotation!, source: source.name, end: source.end }
  }
  // A marker in the actual words of an enclosing quotation is literal source
  // content. Check the complete outer quotation against its named source;
  // never interpret its bracketed words as a second application instruction.
  const ranges = quotationRanges.toSorted((a, b) => a.start - b.start || b.end - a.end)
  const outside = <T extends { marker: number }>(rows: T[]) => {
    let index = 0, end = -1
    return rows.filter((row) => {
      while (ranges[index] && ranges[index].start <= row.marker) {
        end = Math.max(end, ranges[index].end)
        index++
      }
      return row.marker >= end
    })
  }
  return {
    quotes: outside(quotes).map(({ marker: _marker, ...quote }) => quote),
    issues: outside(issues).map(({ marker: _marker, ...issue }) => issue),
  }
}

function sourceName(text: string, marker: number, start: number) {
  const open = text[marker], close = open === "[" ? "]" : ")"
  let depth = 0
  let firstClose: number | undefined
  let cursor = start
  for (; cursor < text.length && text[cursor] !== "\n" && text[cursor] !== "\r"; cursor++) {
    if (cursor - start > SOURCE_CAP)
      return { name: "", end: cursor, error: `the source name exceeds the ${SOURCE_CAP}-character limit` }
    const character = text[cursor]
    if (firstClose === undefined && (character === "]" || character === ")")) firstClose = cursor
    if (character === open) depth++
    if (character !== close) continue
    if (depth) { depth--; continue }
    const name = text.slice(start, cursor).trim()
    return { name, end: cursor + 1, error: name ? undefined : "the source name is empty" }
  }
  if (cursor - start > SOURCE_CAP)
    return { name: "", end: cursor, error: `the source name exceeds the ${SOURCE_CAP}-character limit` }
  // Compatibility: an unmatched/mismatched citation bracket still checks the
  // recoverable quotation against the name ending at the first closer or EOL.
  const stop = firstClose ?? cursor
  const name = text.slice(start, stop).trim()
  return { name, end: firstClose === undefined ? cursor : stop + 1, error: name ? undefined : "the source name is empty" }
}

function quotationBefore(text: string, end: number) {
  const floor = Math.max(0, end - QUOTE_CAP - 1)
  for (let cursor = end - 1; cursor >= floor; cursor--) {
    const character = text[cursor]
    if (!opening(character) && !closing(character)) continue
    // A nested pair such as the "Deposit" means ... is one unit. Separate
    // quotations ("A" and "B" [source: X]) stop at B's own opening mark.
    if (closing(character) && !space(text[cursor - 1]) && !word(characterAt(text, cursor + 1))) {
      const inner = innerOpening(text, cursor, floor)
      if (inner !== undefined) { cursor = inner; continue }
    }
    if (opening(character) && end - cursor - 1 <= QUOTE_CAP)
      return { start: cursor, quote: text.slice(cursor + 1, end) }
  }
}

function innerOpening(text: string, end: number, floor: number) {
  for (let cursor = end - 1; cursor >= Math.max(floor, end - INNER_CAP - 1); cursor--) {
    if (!opening(text[cursor]) && !closing(text[cursor])) continue
    return opening(text[cursor]) && !word(characterBefore(text, cursor)) && !space(text[cursor + 1]) ? cursor : undefined
  }
}

function characterAt(text: string, index: number) {
  const point = text.codePointAt(index)
  return point === undefined ? undefined : String.fromCodePoint(point)
}

function characterBefore(text: string, index: number) {
  const previous = text.charCodeAt(index - 1)
  return characterAt(text, previous >= 0xdc00 && previous <= 0xdfff ? index - 2 : index - 1)
}

function literalRanges(text: string) {
  // Plain prose needs no Markdown parse. This also keeps the existing long-line
  // quote control independent of a document-tokenizer allocation per character.
  if (!/[`<~]|^(?: {4}|\t)|\n(?: {4}|\t)/.test(text)) return []
  return postprocess(parse().document().write(preprocess()(text, undefined, true))).flatMap(([event, token]) => {
    if (event !== "enter") return []
    const literal = ["codeText", "codeFenced", "codeIndented"].includes(token.type)
      || (["htmlText", "htmlFlow"].includes(token.type) && text.startsWith("<!--", token.start.offset))
    return literal ? [{ start: token.start.offset!, end: token.end.offset! }] : []
  })
}
