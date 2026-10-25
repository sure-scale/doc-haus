import { INSTRUCTION_COMMENT_PATTERN, OPEN_ITEM_PATTERN, PLACEHOLDER_PATTERN } from "./marker-patterns"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

export const OPEN_ITEM_HIGHLIGHT_FILL = "FFF2CC"
export const FIELD_HIGHLIGHT_FILL = "DDEBF7"

const OPEN_ITEM_RE = OPEN_ITEM_PATTERN
const OPEN_ITEM_FILLS = new Set([OPEN_ITEM_HIGHLIGHT_FILL, FIELD_HIGHLIGHT_FILL])
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XML = "http://www.w3.org/XML/1998/namespace"
const RUN_RE = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g
/** A run's start tag and its properties, which every piece of a split run repeats. */
const RUN_START_RE = /^<w:r\b[^>]*>(?:\s*<w:rPr>[\s\S]*?<\/w:rPr>)?/
const TEXT_RE = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g

/** The open-item markers render highlights, as they appear in plain text. */
export function openItemMarkers(text: string) {
  return [...new Set([...text.matchAll(OPEN_ITEM_RE)].map((match) => match[0].replace(/\s+/g, " ").trim()))]
}

export function highlightOpenItemsInDocx(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  const document = parts["word/document.xml"]
  if (!document) return bytes
  const highlighted = highlightOpenItemsInDocumentXml(strFromU8(document))
  if (highlighted === strFromU8(document)) return bytes
  parts["word/document.xml"] = strToU8(highlighted)
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

function highlightOpenItemsInDocumentXml(xml: string) {
  return xml.replace(RUN_RE, (run) => highlightRunOpenItems(run))
}

/** A run with each marker of its texts (shadedMarkerSpans) in a run of its
 * own, shaded. Only a text that holds a marker is split; each piece keeps the
 * run's properties and its own share of the run's content (Word puts several
 * texts, breaks and tabs in one run), and a split text keeps its spaces. A run
 * with no marker, or shaded in a fill of its own (ownShading), is returned as
 * it is. */
function highlightRunOpenItems(run: string) {
  const start = RUN_START_RE.exec(run)?.[0]
  if (!start || !run.endsWith("</w:r>") || ownShading(/<w:shd\b[^>]*\bw:fill="([^"]*)"/.exec(start)?.[1] ?? "")) return run
  const content = run.slice(start.length, -"</w:r>".length)
  const pieces: string[] = []
  const piece = (prefix: string, inner: string) => {
    if (inner.trim()) pieces.push(`${prefix}${inner}</w:r>`)
  }
  let cursor = 0
  for (const text of content.matchAll(TEXT_RE)) {
    const value = text[2] ?? ""
    const markers = shadedMarkerSpans(value)
    if (!markers.length) continue
    piece(start, content.slice(cursor, text.index))
    const tag = /\sxml:space=/.test(text[1]!) ? text[1]! : text[1]!.replace("<w:t", '<w:t xml:space="preserve"')
    let offset = 0
    for (const marker of markers) {
      if (marker.start > offset) piece(start, `${tag}${value.slice(offset, marker.start)}</w:t>`)
      piece(highlightRun(start, marker.fill), `${tag}${value.slice(marker.start, marker.end)}</w:t>`)
      offset = marker.end
    }
    if (offset < value.length) piece(start, `${tag}${value.slice(offset)}</w:t>`)
    cursor = text.index + text[0].length
  }
  if (!pieces.length) return run
  piece(start, content.slice(cursor))
  return pieces.join("")
}

/** [start, end) and fill of each marker the DOCX shades in a paragraph's
 * text: an open item or a blank (OPEN_ITEM_PATTERN) and a drafting placeholder
 * (PLACEHOLDER_PATTERN, INSTRUCTION_COMMENT_PATTERN). A bracketed or
 * parenthetical open item ends at the bracket that closes its own, the
 * brackets inside it counted, so a heading it cites in brackets ("[open item:
 * … addressed under [Disputes]]") stays in it; one the line leaves unclosed
 * ends at its first closing bracket, as OPEN_ITEM_PATTERN reads it. A
 * placeholder is its pattern's match. Markers that overlap are one marker,
 * from the first to start (the longest of those starting together) to the
 * last to end, so an open item naming a placeholder ("[open item: … [TBC] …]")
 * stays one open item. A blank, bare or bracketed ("___", "[____]"), takes the
 * field fill, every other marker the open-item fill. */
export function shadedMarkerSpans(text: string) {
  const found: { start: number; end: number }[] = []
  const pattern = new RegExp(OPEN_ITEM_RE.source, "gi")
  for (let match = pattern.exec(text); match; match = pattern.exec(text)) {
    const open = match[0][0]!
    const close = open === "[" ? "]" : open === "(" ? ")" : undefined
    let end = match.index + match[0].length
    if (close) for (let at = match.index, depth = 0; at < text.length && text[at] !== "\n" && text[at] !== "\r"; at++) {
      if (text[at] === open) depth++
      else if (text[at] === close && !--depth) {
        end = at + 1
        break
      }
    }
    found.push({ start: match.index, end })
    pattern.lastIndex = end
  }
  for (const placeholder of [PLACEHOLDER_PATTERN, INSTRUCTION_COMMENT_PATTERN])
    for (const match of text.matchAll(placeholder)) found.push({ start: match.index ?? 0, end: (match.index ?? 0) + match[0].length })
  const spans: { start: number; end: number; fill: string }[] = []
  for (const { start, end } of found.sort((a, b) => a.start - b.start || b.end - a.end)) {
    const last = spans.at(-1)
    if (last && start < last.end) last.end = Math.max(last.end, end)
    else spans.push({ start, end, fill: "" })
  }
  for (const span of spans) span.fill = markerFill(text.slice(span.start, span.end))
  return spans
}

/** A body's open items shaded once its text is final: each marker of a
 * paragraph (shadedMarkerSpans over the text it shows, deleted text aside) shaded
 * whole in its fill, however many runs it spans, a run partly in a marker
 * split at the marker's edge. In a paragraph `clear` accepts, open-item
 * shading outside the markers is removed too, so words a change wrote into a
 * marker's run, or a marker it made clause text, are not shaded as open.
 * Every other formatting is kept, and a run shaded in a fill of its own
 * (ownShading) is left as it is. The bytes are returned themselves when
 * nothing changes. */
export function shadeOpenItemsInDocx(bytes: Uint8Array, clear: (paragraph: Element) => boolean = () => false) {
  const parts = unzipSync(bytes)
  const document = parts["word/document.xml"]
  if (!document) return bytes
  const doc = new DOMParser().parseFromString(strFromU8(document), "application/xml")
  let changed = false
  for (const paragraph of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
    const runs = paragraphRuns(paragraph)
    const spans = shadedMarkerSpans(runs.flatMap(row => row.content.map(part => part.value)).join(""))
    const clearing = clear(paragraph)
    if (spans.length || clearing) for (const row of runs) changed = shadeRun(row, spans, clearing) || changed
  }
  if (!changed) return bytes
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

/** Removes the open-item shading of each run holding part of a marker
 * (shadedMarkerSpans), as the redline compares two bodies; whether any was
 * removed. */
export function unshadeOpenItems(doc: Document) {
  let changed = false
  for (const paragraph of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
    const runs = paragraphRuns(paragraph)
    const spans = shadedMarkerSpans(runs.flatMap(row => row.content.map(part => part.value)).join(""))
    for (const { run, content } of runs) {
      const shading = runShading(run)
      if (!shading || !OPEN_ITEM_FILLS.has((shading.getAttributeNS(W, "fill") ?? "").toUpperCase())) continue
      if (!content.some(part => part.start < part.end && spans.some(span => part.start < span.end && span.start < part.end))) continue
      shading.parentNode!.removeChild(shading)
      changed = true
    }
  }
  return changed
}

type RunContent = { node: Element; value: string; start: number; end: number }

// The runs of a paragraph's own text (not a nested text box's), each with its
// content and where that sits in the text the paragraph shows: a text, tab or
// break is its characters, anything else (a field character or code, a
// drawing) and all of a tracked deletion is empty.
function paragraphRuns(paragraph: Element) {
  let at = 0
  return Array.from(paragraph.getElementsByTagNameNS(W, "r")).filter(run => owningParagraph(run) === paragraph).map(run => {
    const deleted = withinDeletion(run, paragraph)
    return {
      run,
      content: Array.from(run.childNodes).filter((node): node is Element => node.nodeType === 1 && (node as Element).localName !== "rPr").map((node): RunContent => {
        const value = deleted || node.namespaceURI !== W ? "" : node.localName === "t" ? node.textContent ?? "" : node.localName === "tab" ? "\t"
          : node.localName === "br" || node.localName === "cr" ? "\n" : ""
        return { node, value, start: at, end: (at += value.length) }
      }),
    }
  })
}

function withinDeletion(run: Element, paragraph: Element) {
  for (let parent = run.parentNode; parent && parent !== paragraph; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && (parent as Element).localName === "del") return true
  return false
}

// The run with its content in a marker shaded in the marker's fill and, when
// `clearing`, its other content's open-item shading removed: split where that
// changes, each piece keeping the run's properties and its own content.
// Content showing no text goes with the content before it, or after it when
// it comes first. A run shaded in a fill of its own is left as it is.
function shadeRun({ run, content }: { run: Element; content: RunContent[] }, spans: ReturnType<typeof shadedMarkerSpans>, clearing: boolean) {
  const shading = runShading(run)
  const fill = (shading?.getAttributeNS(W, "fill") ?? "").toUpperCase()
  if (ownShading(fill)) return false
  const current = shading?.getAttributeNS(W, "val") === "clear" ? fill : undefined
  // A piece's shading: undefined keeps the run's, "" removes it, a fill sets it.
  const outside = clearing && OPEN_ITEM_FILLS.has(fill) ? "" : undefined
  const pieces: { node: Element; value?: string; fill?: string; shown: boolean }[] = []
  for (const part of content) {
    if (part.start === part.end) {
      pieces.push({ node: part.node, shown: false })
      continue
    }
    const cuts = [...new Set([part.start, ...spans.flatMap(span => [span.start, span.end]).filter(at => part.start < at && at < part.end), part.end])].sort((a, b) => a - b)
    for (const [index, from] of cuts.slice(0, -1).entries()) {
      const wanted = spans.find(span => span.start <= from && from < span.end)?.fill ?? outside
      pieces.push({ node: part.node, value: cuts.length > 2 ? part.value.slice(from - part.start, cuts[index + 1]! - part.start) : undefined,
        fill: wanted === current ? undefined : wanted, shown: true })
    }
  }
  const shown = pieces.filter(piece => piece.shown)
  if (!shown.length) return false
  let previous = shown[0]!
  for (const piece of pieces) if (piece.shown) previous = piece; else piece.fill = previous.fill
  const groups = pieces.reduce<(typeof pieces)[]>((all, piece) => {
    const last = all.at(-1)
    if (last && last[0]!.fill === piece.fill) last.push(piece)
    else all.push([piece])
    return all
  }, [])
  if (groups.length === 1) return groups[0]![0]!.fill === undefined ? false : (restyle(run, groups[0]![0]!.fill), true)
  const properties = Array.from(run.childNodes).find(node => node.nodeType === 1 && (node as Element).localName === "rPr")
  for (const group of groups) {
    const copy = run.cloneNode(false) as Element
    if (properties) copy.appendChild(properties.cloneNode(true))
    for (const piece of group) {
      if (piece.value === undefined) {
        copy.appendChild(piece.node)
        continue
      }
      const text = copy.appendChild(piece.node.cloneNode(false)) as Element
      text.setAttributeNS(XML, "xml:space", "preserve")
      text.appendChild(run.ownerDocument!.createTextNode(piece.value))
    }
    restyle(copy, group[0]!.fill)
    run.parentNode!.insertBefore(copy, run)
  }
  run.parentNode!.removeChild(run)
  return true
}

// A run shaded in `fill`, its texts keeping their spaces as a marker's do, or
// its shading removed (""); undefined leaves it as it is.
function restyle(run: Element, fill: string | undefined) {
  if (fill === undefined) return
  if (!fill) {
    const shading = runShading(run)
    shading?.parentNode!.removeChild(shading)
    return
  }
  shadeRunElement(run, fill)
  for (const text of Array.from(run.childNodes)) if (text.nodeType === 1 && (text as Element).localName === "t") (text as Element).setAttributeNS(XML, "xml:space", "preserve")
}

function runShading(run: Element) {
  const properties = Array.from(run.childNodes).find(node => node.nodeType === 1 && (node as Element).localName === "rPr") as Element | undefined
  return properties && Array.from(properties.childNodes).find(node => node.nodeType === 1 && (node as Element).localName === "shd") as Element | undefined
}

function owningParagraph(node: Node) {
  for (let parent = node.parentNode; parent; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && (parent as Element).localName === "p") return parent
  return null
}

function shadeRunElement(run: Element, fill: string) {
  const doc = run.ownerDocument!
  const child = (parent: Element, name: string) => Array.from(parent.childNodes).find(node => node.nodeType === 1 && (node as Element).localName === name) as Element | undefined
  let properties = child(run, "rPr")
  if (!properties) properties = run.insertBefore(doc.createElementNS(W, "w:rPr"), run.firstChild)
  const current = child(properties, "shd")
  if (current?.getAttributeNS(W, "fill") === fill && current.getAttributeNS(W, "val") === "clear") return false
  if (current) properties.removeChild(current)
  const shading = doc.createElementNS(W, "w:shd")
  shading.setAttributeNS(W, "w:val", "clear")
  shading.setAttributeNS(W, "w:color", "auto")
  shading.setAttributeNS(W, "w:fill", fill)
  properties.insertBefore(shading, child(properties, "rPrChange") ?? null)
  return true
}

// Whether a run's shading fill is one of its own, such as a lawyer's
// highlight in a supplied original: not a marker fill and not "auto" (none).
// The marker in such a run shows shaded already, so the run keeps it.
function ownShading(fill: string) {
  const upper = fill.toUpperCase()
  return !!upper && upper !== "AUTO" && !OPEN_ITEM_FILLS.has(upper)
}

function markerFill(marker: string) {
  return /^(?:_+|\[\s*_+\s*\])$/.test(marker) ? FIELD_HIGHLIGHT_FILL : OPEN_ITEM_HIGHLIGHT_FILL
}

function highlightRun(runBeforeText: string, fill: string) {
  const shading = `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`
  if (/<w:shd\b[^>]*w:fill="[^"]*"[^>]*\/>/.test(runBeforeText)) return runBeforeText.replace(/<w:shd\b[^>]*w:fill="[^"]*"[^>]*\/>/, shading)
  if (runBeforeText.includes("<w:rPr>")) return runBeforeText.replace("</w:rPr>", `${shading}</w:rPr>`)
  return runBeforeText.replace(/(<w:r\b[^>]*>)/, `$1<w:rPr>${shading}</w:rPr>`)
}

export function hasOpenItemMarker(text: string) {
  return new RegExp(OPEN_ITEM_RE.source, "i").test(text)
}
