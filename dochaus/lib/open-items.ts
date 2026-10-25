import { OPEN_ITEM_PATTERN } from "./marker-patterns"
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

/** A run with each open-item marker of its texts in a run of its own,
 * shaded. Only a text that holds a marker is split; each piece keeps the run's
 * properties and its own share of the run's content (Word puts several texts,
 * breaks and tabs in one run), and a split text keeps its spaces. A run with no
 * marker is returned as it is. */
function highlightRunOpenItems(run: string) {
  const start = RUN_START_RE.exec(run)?.[0]
  if (!start || !run.endsWith("</w:r>")) return run
  const content = run.slice(start.length, -"</w:r>".length)
  const pieces: string[] = []
  const piece = (prefix: string, inner: string) => {
    if (inner.trim()) pieces.push(`${prefix}${inner}</w:r>`)
  }
  let cursor = 0
  for (const text of content.matchAll(TEXT_RE)) {
    const value = text[2] ?? ""
    const markers = [...value.matchAll(OPEN_ITEM_RE)]
    if (!markers.length) continue
    piece(start, content.slice(cursor, text.index))
    const tag = /\sxml:space=/.test(text[1]!) ? text[1]! : text[1]!.replace("<w:t", '<w:t xml:space="preserve"')
    let offset = 0
    for (const marker of markers) {
      const at = marker.index ?? 0
      if (at > offset) piece(start, `${tag}${value.slice(offset, at)}</w:t>`)
      piece(highlightRun(start, markerFill(marker[0])), `${tag}${marker[0]}</w:t>`)
      offset = at + marker[0].length
    }
    if (offset < value.length) piece(start, `${tag}${value.slice(offset)}</w:t>`)
    cursor = text.index + text[0].length
  }
  if (!pieces.length) return run
  piece(start, content.slice(cursor))
  return pieces.join("")
}

/** [start, end) and fill of each open-item marker in a paragraph's text. A
 * bracketed or parenthetical marker ends at the bracket that closes its own,
 * the brackets inside it counted, so a heading it cites in brackets ("[open
 * item: … addressed under [Disputes]]") stays in it. One the line leaves
 * unclosed ends at its first closing bracket, as OPEN_ITEM_PATTERN reads it. */
export function openItemSpans(text: string) {
  const spans: { start: number; end: number; fill: string }[] = []
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
    spans.push({ start: match.index, end, fill: markerFill(match[0]) })
    pattern.lastIndex = end
  }
  return spans
}

/** A body's open items shaded once its text is final: each marker of a
 * paragraph (openItemSpans over the text it shows, deleted text aside) shaded
 * whole in its fill, however many runs it spans, a run partly in a marker
 * split at the marker's edge. In a paragraph `clear` accepts, open-item
 * shading outside the markers is removed too, so words a change wrote into a
 * marker's run, or a marker it made clause text, are not shaded as open.
 * Every other formatting, and shading of another fill, is kept. The bytes are
 * returned themselves when nothing changes. */
export function shadeOpenItemsInDocx(bytes: Uint8Array, clear: (paragraph: Element) => boolean = () => false) {
  const parts = unzipSync(bytes)
  const document = parts["word/document.xml"]
  if (!document) return bytes
  const doc = new DOMParser().parseFromString(strFromU8(document), "application/xml")
  let changed = false
  for (const paragraph of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
    const runs = paragraphRuns(paragraph)
    const spans = openItemSpans(runs.flatMap(row => row.content.map(part => part.value)).join(""))
    const clearing = clear(paragraph)
    if (spans.length || clearing) for (const row of runs) changed = shadeRun(row, spans, clearing) || changed
  }
  if (!changed) return bytes
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

/** Removes the open-item shading of each run holding part of an open-item
 * marker (openItemSpans), as the redline compares two bodies; whether any was
 * removed. */
export function unshadeOpenItems(doc: Document) {
  let changed = false
  for (const paragraph of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
    const runs = paragraphRuns(paragraph)
    const spans = openItemSpans(runs.flatMap(row => row.content.map(part => part.value)).join(""))
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
// break is its characters, anything else (a field character or code, deleted
// text, a drawing) is empty.
function paragraphRuns(paragraph: Element) {
  let at = 0
  return Array.from(paragraph.getElementsByTagNameNS(W, "r")).filter(run => owningParagraph(run) === paragraph).map(run => ({
    run,
    content: Array.from(run.childNodes).filter((node): node is Element => node.nodeType === 1 && (node as Element).localName !== "rPr").map((node): RunContent => {
      const value = node.namespaceURI !== W ? "" : node.localName === "t" ? node.textContent ?? "" : node.localName === "tab" ? "\t"
        : node.localName === "br" || node.localName === "cr" ? "\n" : ""
      return { node, value, start: at, end: (at += value.length) }
    }),
  }))
}

// The run with its content in a marker shaded in the marker's fill and, when
// `clearing`, its other content's open-item shading removed: split where that
// changes, each piece keeping the run's properties and its own content.
// Content showing no text goes with the content before it, or after it when
// it comes first.
function shadeRun({ run, content }: { run: Element; content: RunContent[] }, spans: ReturnType<typeof openItemSpans>, clearing: boolean) {
  const shading = runShading(run)
  const fill = (shading?.getAttributeNS(W, "fill") ?? "").toUpperCase()
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

function markerFill(marker: string) {
  return /^_+$/.test(marker) ? FIELD_HIGHLIGHT_FILL : OPEN_ITEM_HIGHLIGHT_FILL
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
