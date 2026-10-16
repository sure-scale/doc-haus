import { parse, postprocess, preprocess } from "micromark"
import { decodeString } from "micromark-util-decode-string"
import { scanMarkdownStructure } from "./markdown-structure"

// One visible character run and the raw source span [start, end) it came from.
export type VisibleTextSegment = { value: string; start: number; end: number }

// How the structure scanner classified a raw line. Only "text" lines are read
// as inline Markdown; every other kind is literal (visible === raw).
export type VisibleLineKind = "text" | "fence_delimiter" | "fenced" | "indented_code" | "html_block"

export type VisibleTextLine = {
  line: number
  source: string
  visible: string
  segments: VisibleTextSegment[]
  kind: VisibleLineKind
}

// A visible fragment and its raw offset map. map[i] is the raw offset of the
// token that produced visible character i and ends[i] that token's raw end, so
// an escape or character reference is indivisible. map[text.length] is the raw
// length of the fragment.
export type VisibleInlineSpans = { text: string; map: number[]; ends: number[] }

export type VisibleLine = { line: number; raw: string; visible: string; map: number[]; ends: number[]; kind: VisibleLineKind }

// One detector match: found on visible text, bound to raw bytes. start/end are
// offsets in the whole input, line/column are 1-based raw coordinates and raw
// is the exact source slice.
export type VisibleTextMatch = {
  pattern: number
  start: number
  end: number
  line: number
  column: number
  raw: string
  visible: string
}

// Markup delimiters the CommonMark tokenizer proved to be syntax; they have no
// visible form. An unmatched `*`, an intraword `_` or a lone backtick stays a
// `data` token instead, so it is never removed as formatting.
const DELIMITERS = new Set([
  "emphasisSequence",
  "strongSequence",
  "codeTextSequence",
  "codeTextPadding",
  "labelMarker",
  "labelImageMarker",
  "autolinkMarker",
])
// A link destination or title is not displayed text.
const HIDDEN = new Set(["resource", "reference"])
// Escapes and character references show the character they encode. Code span
// contents, autolinks and raw HTML are literal (CommonMark section 2.4), so
// they are never routed through this decoder.
const DECODED = new Set(["characterEscape", "characterReference"])
const BREAKS = new Set(["hardBreakEscape", "hardBreakTrailing"])
// An HTML comment is not displayed. Other raw HTML stays literal text.
const HTML = new Set(["htmlText", "htmlFlow"])

type VisibleEdit = { start: number; end: number; value: string }

// The CommonMark visible text of one inline fragment (a heading's text, a
// scalar, or one paragraph line), tokenized as inline text only: a block
// reading would turn "1. Title" into a list item and drop its number. Every
// byte the tokenizer does not prove to be a delimiter is kept, so no visible
// character can be lost to an unrecognized token.
export function visibleInlineSegments(source: string) {
  return visibleSegments(source, 0, source.length, visibleEdits(source, postprocess(parse().text().write(preprocess()(source, undefined, true)))))
}

export function visibleInlineText(source: string) {
  return visibleInlineSegments(source).map((segment) => segment.value).join("")
}

// Line-preserving visible text of a Markdown section. Block structure is
// classified on the raw lines by scanMarkdownStructure before anything is
// decoded, so an escaped `\#` paragraph stays a paragraph. Only its visible
// lines are projected; fenced and indented code, fence delimiters and raw HTML
// blocks are kept byte-literal. The projection reads the original source as
// one CommonMark document, so a code span or comment that continues onto the
// next line keeps its context, and only container markers the parser proves
// change what a line shows. Segment spans are offsets in each line's `source`.
export function visibleSectionLines(content: string): VisibleTextLine[] {
  const structure = scanMarkdownStructure(content, [])
  const visible = new Set(structure.visibleLines.map((line) => line.line))
  const kinds = new Map<number, VisibleLineKind>([
    ...structure.fenceLines.map((line) => [line.line, line.delimiter ? "fence_delimiter" : "fenced"] as const),
    ...structure.indentedCodeLines.map((line) => [line.line, "indented_code"] as const),
  ])
  const edits = visibleEdits(content, postprocess(parse().document().write(preprocess()(content, undefined, true))))
  const starts = [0, ...[...content.matchAll(/\r?\n/g)].map((match) => match.index + match[0].length)]
  return content.split(/\r?\n/).map((source, index) => {
    const text = visible.has(index + 1)
    const segments = text
      ? visibleSegments(content, starts[index]!, starts[index]! + source.length, edits)
      : [{ value: source, start: 0, end: source.length }]
    return {
      line: index + 1,
      source,
      visible: segments.map((segment) => segment.value).join(""),
      segments,
      kind: text ? "text" : kinds.get(index + 1) ?? "html_block",
    }
  })
}

export function visibleSectionText(content: string) {
  return visibleSectionLines(content).map((line) => line.visible).join("\n")
}

export function visibleInlineSpans(source: string): VisibleInlineSpans {
  return inlineSpans(visibleInlineSegments(source), source.length)
}

// Line-preserving visible text with a raw offset map per line. Structure is
// classified on raw lines first (see visibleSectionLines).
export function visibleLines(content: string): VisibleLine[] {
  return visibleSectionLines(content).map((line) => ({
    line: line.line,
    raw: line.source,
    visible: line.visible,
    kind: line.kind,
    ...inlineSpans(line.segments, line.source.length),
  }))
}

// The minimal raw range [start, end) that produced visible [a, b). Delimiters
// before the first or after the last visible character stay outside the range;
// delimiters strictly inside the match are inside it.
export function visibleRawSpan(spans: { map: number[]; ends: number[] }, a: number, b: number) {
  if (b <= a) return { start: spans.map[a]!, end: spans.map[a]! }
  return { start: spans.map[a]!, end: spans.ends[b - 1]! }
}

// Detect on visible text, bind on raw bytes. Every global pattern runs on each
// line's visible text and each match is bound to the raw span that produced
// it. A match of a pattern on the raw text is also kept unless a visible match
// already covers any of its bytes: a CommonMark reading can hide text a raw
// reader sees (paired underscore runs read as emphasis, HTML comment bodies, a
// match across a line break), and a detector must never see less than the raw
// text shows. Matches are ordered by raw offset and deduplicated by raw span.
export function visibleTextMatches(text: string, patterns: readonly RegExp[]): VisibleTextMatch[] {
  const starts = [0, ...[...text.matchAll(/\n/g)].map((match) => match.index + 1)]
  const visible = visibleLines(text).flatMap((line, index) =>
    // The structure scanner accepts any line shaped like `<name ... >` as an
    // HTML block, which is looser than CommonMark, and a Word projection writes
    // a literal `>` as `\>`. An escaped paragraph can therefore be classified
    // as raw HTML, so detection also reads those lines as inline text.
    [line, ...(line.kind === "html_block" ? [{ ...visibleInlineSpans(line.raw), visible: visibleInlineText(line.raw) }] : [])]
      .flatMap((reading) =>
        patterns.flatMap((pattern, number) =>
          [...reading.visible.matchAll(pattern)].map((match) => {
            const span = visibleRawSpan(reading, match.index, match.index + match[0].length)
            const start = starts[index]! + span.start
            const end = starts[index]! + span.end
            return { pattern: number, start, end, line: line.line, column: span.start + 1, raw: text.slice(start, end), visible: match[0] }
          })
        )
      )
  )
  const raw = patterns.flatMap((pattern, number) =>
    [...text.matchAll(pattern)].flatMap((match) => {
      const start = match.index
      const end = start + match[0].length
      if (visible.some((found) => found.start < end && found.end > start)) return []
      const index = Math.max(0, starts.findLastIndex((offset) => offset <= start))
      return [{ pattern: number, start, end, line: index + 1, column: start - starts[index]! + 1, raw: match[0], visible: match[0] }]
    })
  )
  const seen = new Set<string>()
  return [...visible, ...raw]
    .toSorted((left, right) => left.start - right.start || left.pattern - right.pattern || left.end - right.end)
    .filter((match) => {
      const key = `${match.start}:${match.end}`
      if (seen.has(key)) return false
      seen.add(key)
      return true
    })
}

function inlineSpans(segments: VisibleTextSegment[], length: number): VisibleInlineSpans {
  const chars = segments.flatMap((segment) => {
    // A decoded token (escape, reference, hard break) shows text that differs
    // from its raw bytes and cannot be split; literal runs map one to one.
    const atomic = segment.end - segment.start !== segment.value.length
    return Array.from({ length: segment.value.length }, (_, offset) =>
      atomic ? [segment.start, segment.end] : [segment.start + offset, segment.start + offset + 1]
    )
  })
  return {
    text: segments.map((segment) => segment.value).join(""),
    map: [...chars.map((char) => char[0]!), length],
    ends: chars.map((char) => char[1]!),
  }
}

// The raw spans whose shown form differs from their bytes, in source order.
function visibleEdits(source: string, events: ReturnType<typeof postprocess>) {
  const edits: VisibleEdit[] = []
  let skipped: (typeof events)[number][1] | undefined
  events.forEach(([kind, token]) => {
    if (skipped) {
      if (kind === "exit" && token === skipped) skipped = undefined
      return
    }
    if (kind === "exit") return
    const start = token.start.offset
    const end = token.end.offset
    const raw = source.slice(start, end)
    if (HTML.has(token.type)) {
      skipped = token
      const comments = token.type === "htmlText"
        ? raw.startsWith("<!--") ? [{ index: 0, length: raw.length }] : []
        : [...raw.matchAll(/<!--[\s\S]*?(?:-->|$)/g)].map((match) => ({ index: match.index, length: match[0].length }))
      comments.forEach((comment) => edits.push({ start: start + comment.index, end: start + comment.index + comment.length, value: "" }))
      return
    }
    if (DELIMITERS.has(token.type) || HIDDEN.has(token.type) || DECODED.has(token.type) || BREAKS.has(token.type)) {
      skipped = token
      edits.push({ start, end, value: DECODED.has(token.type) ? decodeString(raw) : BREAKS.has(token.type) ? " " : "" })
    }
  })
  return edits
}

// Visible runs of source[from, to): raw bytes between edits, and each edit's
// shown value. Offsets are relative to `from`.
function visibleSegments(source: string, from: number, to: number, edits: VisibleEdit[]) {
  const segments: VisibleTextSegment[] = []
  const raw = (start: number, end: number) => {
    if (end > start) segments.push({ value: source.slice(start, end), start: start - from, end: end - from })
  }
  const cursor = edits.filter((edit) => edit.end > from && edit.start < to).reduce((position, edit) => {
    const start = Math.max(edit.start, from)
    const end = Math.min(edit.end, to)
    raw(position, start)
    if (edit.value) segments.push({ value: edit.value, start: start - from, end: end - from })
    return Math.max(position, end)
  }, from)
  raw(cursor, to)
  return segments
}
