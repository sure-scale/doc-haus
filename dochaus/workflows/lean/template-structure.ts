import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { chunkText } from "../../lib/chunk-text"

// The headed parts of a matter document read as a template. Parts come from the
// shape of the canonical source lines only: a numbered line with a short
// capitalised title, an all-capitals line, a Markdown heading, or a label line
// (a capitalised word followed by a letter, number or roman numeral). Nothing
// here knows what a part is called or what a document should contain. A source
// with no such lines has no parts, and every check built on it is a no-op.
//
// The numbered clauses are the document's backbone. An unnumbered heading among
// the clauses, followed by a clause that continues their numbering, is part of
// the clause before it; one after the clauses whose numbering does not continue
// opens a part of its own, and numbered lines under it are its own. Unnumbered
// headings before the first clause hold no clauses. So a heading that only
// divides the clauses holds no text of its own and is not a part.

export type TemplateUnit = {
  /** Position in template order. */
  index: number
  /** Heading text as written in the template (one line, or a label line and its title line). */
  label: string
  /** Heading wordings any of which names the part in a draft. */
  alternatives: string[][]
  /** 0 for an unnumbered heading, otherwise the depth of its number (1, 1.1, ...). */
  level: number
  /** Index of the enclosing part, if any. */
  parent?: number
  /** Words in the part, its heading excluded. */
  words: number
  /** The template's opening heading: its own title, which the draft's title replaces. */
  title?: boolean
}

/** Parts shorter than this are not measured: they carry too little text to tell
 * a heading from a stray capitalised line. */
const MIN_PART_WORDS = 8
const MAX_HEADING_WORDS = 12
const MAX_TITLE_WORDS = 10
const MAX_EXTRA_HEADING_WORDS = 4

/** The source text of one matter document, rebuilt from its indexed chunks in
 * offset order. A chunk that opens a new section starts a new line, as the
 * section's heading line did in the source. */
export function sourceText(directory: string, sourceID: string) {
  const file = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(file)) return ""
  const db = new Database(file, { readonly: true })
  try {
    const chunks = db.query("SELECT section, text, char_start FROM chunks WHERE doc_name = ? ORDER BY char_start, rowid")
      .all(sourceID) as { section: string | null; text: string; char_start: number }[]
    // A gap is padded to keep offsets, ending in a line break where a section opens.
    return chunkText(chunks, (gap, at) =>
      at > 0 && chunks[at]!.section !== chunks[at - 1]!.section ? `${" ".repeat(gap - 1)}\n` : " ".repeat(gap)).text
  } catch {
    return ""
  } finally {
    db.close()
  }
}

export function parseTemplateUnits(text: string): TemplateUnit[] {
  const lines = text.split(/\r?\n/).map(line => line.trim())
  const headings = mergedHeadings(lines)
  const titled = headings[0]?.level === 0 && headings[0].start === lines.findIndex(Boolean)
  const shape = structure(headings, titled)
  const inside = (at: number, holder: number): boolean => {
    const parent = shape[at]!.parent
    return parent !== undefined && (parent === holder || inside(parent, holder))
  }
  const units = headings.map((heading, at) => {
    const next = headings.findIndex((_, later) => later > at && !inside(later, at))
    const span = lines.slice(heading.start, next < 0 ? lines.length : headings[next]!.start).join(" ")
    return {
      label: heading.lines.join(" – "),
      wordings: headingWordings(heading.titles),
      level: heading.number?.length ?? 0,
      words: Math.max(0, countWords(span) - heading.titles.reduce((sum, title) => sum + countWords(title), 0) - (heading.number?.length ?? 0)),
      title: titled && at === 0,
    }
  })
  const kept = units.flatMap((unit, at) =>
    unit.title || (unit.wordings.whole.length + unit.wordings.sides.length > 0 && unit.words >= MIN_PART_WORDS) ? [at] : [])
  // A part whose enclosing heading is not kept belongs to the nearest kept one.
  const keptParent = (at: number): number | undefined => {
    const parent = shape[at]!.parent
    if (parent === undefined) return
    return kept.includes(parent) ? kept.indexOf(parent) : keptParent(parent)
  }
  // One side of a heading (the "Services" of "Schedule 1 – Services") that is
  // also how another part is headed names neither part on its own.
  const keys = kept.map(at => new Set([...units[at]!.wordings.whole, ...units[at]!.wordings.sides].map(words => words.join(" "))))
  return kept.map((at, index): TemplateUnit => {
    const unit = units[at]!
    const parent = keptParent(at)
    const sides = unit.wordings.sides.filter(words => !keys.some((other, key) => key !== index && other.has(words.join(" "))))
    const alternatives = [...new Map([...unit.wordings.whole, ...sides].map(words => [words.join(" "), words])).values()]
    return {
      index, label: unit.label, alternatives, level: unit.level, words: unit.words,
      ...(parent !== undefined ? { parent } : {}), ...(unit.title ? { title: true } : {}),
    }
  })
}

/** The template parts a heading names: those with a wording all of whose words
 * the heading holds, few other words around it; of several, those with the
 * longest such wording ("Schedule 1 – Services" names the schedule, not the
 * "Services" clause). The template's own title competes with no part. */
export function headingUnits(text: string, units: TemplateUnit[]) {
  const scores = units.map(unit => unit.title ? 0 : headingScore(text, unit))
  const best = Math.max(0, ...scores)
  return new Set(best ? units.filter(unit => scores[unit.index] === best).map(unit => unit.index) : [])
}

/** How closely a heading names one template part: the length of the longest
 * wording of the part that it holds with few other words around it, else 0. */
export function headingScore(text: string, unit: TemplateUnit) {
  return Math.max(0, ...headingSegments(text).flatMap(segment => unit.alternatives
    .filter(words => words.every(word => segment.includes(word)) && segment.length <= words.length + MAX_EXTRA_HEADING_WORDS)
    .map(words => words.length)))
}

/** Some wording of the part's heading appears in the text as a phrase. */
export function phraseNames(text: string, unit: TemplateUnit) {
  const words = tokens(text)
  return unit.alternatives.some(phrase => words.some((_, at) => phrase.every((word, offset) => words[at + offset] === word)))
}

/** Lowercased word tokens without accents, a plural "s", or leading numbering. */
export function tokens(text: string) {
  const words = text.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim().split(" ")
    .filter(Boolean).map(word => word.length > 3 && word.endsWith("s") && !word.endsWith("ss") ? word.slice(0, -1) : word)
  const start = words.findIndex(word => !/^(?:\d+|[ivx]{1,4})$/.test(word))
  return start < 0 ? [] : words.slice(start)
}

export function countWords(text: string) {
  return text.match(/[\p{L}\p{N}]+/gu)?.length ?? 0
}

type Heading = { start: number; end: number; level: number; number?: number[]; lines: string[]; titles: string[] }

/** Heading lines in order. A label line with no title of its own ("Schedule 1")
 * and the heading line directly under it are one heading. */
function mergedHeadings(lines: string[]) {
  return lines.reduce<{ headings: Heading[]; body: boolean }>((acc, line, at) => {
    if (!line) return acc
    const heading = headingLine(line)
    if (!heading) return { ...acc, body: true }
    const last = acc.headings.at(-1)
    if (last && !acc.body && last.level === 0 && heading.level === 0 && last.lines.length === 1 && labelOnly(last.lines[0]!))
      return { headings: [...acc.headings.slice(0, -1), { ...last, end: at + 1, lines: [...last.lines, heading.display], titles: [...last.titles, heading.title] }], body: false }
    return { headings: [...acc.headings, { start: at, end: at + 1, level: heading.level, number: heading.number, lines: [heading.display], titles: [heading.title] }], body: false }
  }, { headings: [], body: true }).headings
}

/** Each heading's enclosing heading, by the backbone rules above. */
function structure(headings: Heading[], titled: boolean) {
  const state: { container?: number; last?: number[] } = {}
  return headings.map((heading, at): { parent?: number } => {
    if (titled && at === 0) return {}
    if (heading.number) {
      const from = state.container === undefined ? 0 : state.container + 1
      const parent = headings.slice(from, at).findLastIndex(other => Boolean(other.number) && other.number!.length < heading.number!.length)
      if (state.container === undefined) state.last = heading.number
      return { parent: parent >= 0 ? from + parent : state.container }
    }
    const last = state.last
    if (!last) return {}
    const next = headings.slice(at + 1).find(other => other.number)?.number
    const continues = state.container === undefined && next !== undefined && (next[0]! > last[0]! || (next[0] === last[0] && next.length > 1))
    if (continues) return { parent: headings.slice(0, at).findLastIndex(other => Boolean(other.number)) }
    state.container = at
    return {}
  })
}

function headingLine(raw: string): { level: number; number?: number[]; title: string; display: string } | undefined {
  if (/[_[\]|]/.test(raw)) return
  const markdown = /^#{1,6}\s+/.test(raw)
  const line = raw.replace(/^#{1,6}\s+/, "").replace(/\s+#+$/, "").trim()
  const numbered = /^(\d+(?:\.\d+)*)[.)]?\s+(.+)$/.exec(line)
  if (numbered) {
    // A defined term with its meaning run in ("1.1 Term: a ...") or a quoted
    // term opens a definition, not a part.
    if (/^["“'‘]/.test(numbered[2]!) || /^[^:]{1,80}:\s+\p{Ll}/u.test(numbered[2]!)) return
    const title = numbered[2]!.split(/[.:]\s/)[0]!.replace(/[.:;,]$/, "").trim()
    const words = countWords(title)
    if (!words || words > MAX_TITLE_WORDS || !(capitals(title) || titleCase(title))) return
    const number = numbered[1]!.split(".").map(Number)
    return { level: number.length, number, title, display: `${numbered[1]} ${title}` }
  }
  const words = countWords(line)
  if (!words || words > MAX_HEADING_WORDS || /[.;:,]$/.test(line)) return
  if (markdown || capitals(line) || labelLine(line)) return { level: 0, title: line, display: line }
}

/** Every word with a letter is in capitals. */
function capitals(text: string) {
  return /\p{L}/u.test(text) && text === text.toUpperCase()
}

/** Every word of four or more letters starts with a capital. */
function titleCase(text: string) {
  const words = text.match(/\p{L}+/gu) ?? []
  return words.some(word => /^\p{Lu}/u.test(word)) && words.every(word => word.length < 4 || /^\p{Lu}/u.test(word))
}

/** A capitalised word followed by a letter, number or roman numeral, then
 * nothing or a title ("Annex A", "Part 2 – Charges"): the shape of a part label. */
function labelLine(line: string) {
  return /^\p{Lu}\p{Ll}+\s+(?:[A-Z]|\d+|[IVXLC]+)(?:$|\s*[-–—:.)]\s*\S|\s+\S)/u.test(line) && titleCase(line)
}

/** A label with no title of its own: one word and a letter, number or roman numeral. */
function labelOnly(line: string) {
  return /^\p{Lu}\p{L}*\s+(?:[A-Z]|\d+|[IVXLC]+)$/u.test(line)
}

/** The opening words of a paragraph that can be a heading: its first line,
 * whole or up to its first ". " or ": ", numbering removed. */
function headingSegments(text: string) {
  const line = (text.split("\n")[0] ?? "").replace(/^\s*\(?(?:\d+(?:\.\d+)*\.?|[ivxlc]+[.)]|[a-z][.)])\)?\s+/i, "")
  return [line, line.split(/[.:]\s/)[0] ?? ""].filter(segment => countWords(segment) <= MAX_HEADING_WORDS).map(tokens)
}

/** A heading's wordings: the whole heading and, when it joins a label and a
 * title (on two lines, or with a colon or dash), each side. Wordings too short
 * to name a part are dropped. */
function headingWordings(titles: string[]) {
  const named = (texts: string[]) => texts.map(tokens).filter(words => words.length >= 2 || words.some(word => word.length >= 4))
  const split = (title: string) => {
    const sides = title.split(/\s*[:–—]\s*|\s+-\s+/).filter(Boolean)
    return sides.length > 1 ? sides : []
  }
  return { whole: named([titles.join(" ")]), sides: named([...(titles.length > 1 ? titles : []), ...titles.flatMap(split)]) }
}
