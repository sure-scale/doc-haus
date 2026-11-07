import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { matterSourceIDs, REQUEST_SOURCE } from "./quotes"
import { approvedResearch } from "./research"

// Stages cite a quotation as "exact words" [source: name]: the marker is how
// every check finds and verifies it, so the working files keep it. The reader
// gets the same citation in plain words: "exact words" (name), adjacent
// markers joined as (A; B). The lawyer's own request needs no citation, so its
// marker goes; a firm playbook is named as the lawyer knows it. Only a whole
// marker, closed on its own line, naming a source the checks know (a matter
// source, the request or approved research), is rewritten; any other
// parenthetical such as "(Source: company records)" is the drafter's own words
// and stays. The quoted words never change.

const MARKER = /[[(][ \t]*source[ \t]*:/giu
const NAME_CAP = 4096

/** A source as the lawyer knows it. */
export function sourceName(document: string) {
  if (!document.startsWith("playbook:")) return document
  const label = document.slice("playbook:".length).replace(/^playbook-/, "").toUpperCase()
  return label ? `${label} playbook` : "firm playbook"
}

/** Whether a marker's name is a source the checks know. */
export type CitedSource = (name: string) => boolean

/** The matter's citable sources: its documents and playbook, the request and
 * approved research. A name may continue with a pinpoint ("a.docx, cl. 4"). */
export function citationSources(directory: string): CitedSource {
  const known = new Set([...matterSourceIDs(directory), REQUEST_SOURCE, ...approvedResearch(directory).map(row => row.url)])
  return name => known.has(name) ||
    [...known].some(source => name.startsWith(source) && /^\s*[,;]\s*\S/.test(name.slice(source.length)))
}

type Edit = { start: number; end: number; text: string }

/** `text` with each marker naming a known source in reader form. */
export function readerCitations(text: string, cited: CitedSource) {
  const edits = citationEdits(text, cited)
  if (!edits.length) return text
  let output = "", cursor = 0
  for (const edit of edits) {
    output += text.slice(cursor, edit.start) + edit.text
    cursor = edit.end
  }
  return output + text.slice(cursor)
}

function citationEdits(text: string, cited: CitedSource): Edit[] {
  const markers: { start: number; end: number; name: string }[] = []
  for (const match of text.matchAll(MARKER)) {
    if (markers.length && match.index < markers.at(-1)!.end) continue
    const marker = closedMarker(text, match.index, match.index + match[0].length)
    if (marker && cited(marker.name)) markers.push(marker)
  }
  const edits: Edit[] = []
  for (let index = 0; index < markers.length;) {
    // Markers separated only by spaces cite one quotation: one parenthesis.
    let last = index
    while (markers[last + 1] && /^[ \t]*$/.test(text.slice(markers[last]!.end, markers[last + 1]!.start))) last++
    const names = [...new Set(markers.slice(index, last + 1).map(row => row.name).filter(name => name !== REQUEST_SOURCE).map(sourceName))]
    let start = markers[index]!.start
    // A dropped citation takes the space before it, so no gap is left before punctuation.
    if (!names.length) while (start > 0 && (text[start - 1] === " " || text[start - 1] === "\t")) start--
    edits.push({ start, end: markers[last]!.end, text: names.length ? `(${names.join("; ")})` : "" })
    index = last + 1
  }
  return edits
}

function closedMarker(text: string, start: number, from: number) {
  const open = text[start], close = open === "[" ? "]" : ")"
  let depth = 0
  for (let cursor = from; cursor < text.length && cursor - from <= NAME_CAP; cursor++) {
    const character = text[cursor]
    if (character === "\n" || character === "\r") return undefined
    if (character === open) depth++
    else if (character === close && depth) depth--
    else if (character === close) {
      const name = text.slice(from, cursor).trim()
      return name ? { start, end: cursor + 1, name } : undefined
    }
  }
  return undefined
}

// A self-closing empty paragraph is matched on its own, so no match spans two.
const PARAGRAPH = /<w:p(?:\s[^>]*?)?\/>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g
// A paragraph's visible text and a redline's struck text are read apart.
const VISIBLE = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g
const DELETED = /(<w:delText(?:\s[^>]*)?>)([\s\S]*?)(<\/w:delText>)/g

/** A Word file with each paragraph's source markers in reader form. The
 * paragraph's visible text, and apart from it its deleted text, is read across
 * its runs, so a marker split between runs is rewritten as one; each run keeps
 * its own formatting. Returns `bytes` itself when no paragraph holds a marker. */
export function readerCitationsInDocx(bytes: Uint8Array, cited: CitedSource) {
  const parts = unzipSync(bytes)
  const document = parts["word/document.xml"]
  if (!document) return bytes
  const xml = strFromU8(document)
  const rewritten = xml.replace(PARAGRAPH, paragraph => runCitations(runCitations(paragraph, VISIBLE, cited), DELETED, cited))
  if (rewritten === xml) return bytes
  parts["word/document.xml"] = strToU8(rewritten)
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

function runCitations(paragraph: string, pattern: RegExp, cited: CitedSource) {
  const runs = [...paragraph.matchAll(pattern)].map(match => ({ index: match.index, xml: match[0], open: match[1]!, text: decode(match[2]!), close: match[3]! }))
  const text = runs.map(run => run.text).join("")
  const edits = citationEdits(text, cited)
  if (!edits.length) return paragraph
  // Each edit's replacement goes in the run where it starts; the rest of its
  // span is removed from whichever runs hold it.
  let offset = 0
  const changed = runs.map(run => {
    const from = offset, to = offset + run.text.length
    offset = to
    let out = ""
    for (let at = from; at < to; at++) {
      const edit = edits.find(row => row.start <= at && at < row.end)
      if (!edit) out += text[at]
      else if (at === edit.start) out += edit.text
    }
    return out
  })
  let output = "", cursor = 0
  runs.forEach((run, index) => {
    output += paragraph.slice(cursor, run.index)
    const value = changed[index]!
    output += value === run.text ? run.xml : `${preserved(run.open)}${encode(value)}${run.close}`
    cursor = run.index + run.xml.length
  })
  return output + paragraph.slice(cursor)
}

function preserved(open: string) {
  return /xml:space=/.test(open) ? open : open.replace(/^<(w:t|w:delText)/, `<$1 xml:space="preserve"`)
}

function decode(text: string) {
  return text.replace(/&(lt|gt|quot|apos|amp|#(\d+)|#x([0-9a-f]+));/gi, (entity, name: string, decimal?: string, hex?: string) =>
    decimal ? String.fromCodePoint(Number(decimal)) : hex ? String.fromCodePoint(parseInt(hex, 16))
      : ({ lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" } as Record<string, string>)[name.toLowerCase()] ?? entity)
}

function encode(text: string) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}
