import { OPEN_ITEM_PATTERN } from "./marker-patterns"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

export const OPEN_ITEM_HIGHLIGHT_FILL = "FFF2CC"
export const FIELD_HIGHLIGHT_FILL = "DDEBF7"

const OPEN_ITEM_RE = OPEN_ITEM_PATTERN
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

/** Each open-item marker whose text spans several runs of a paragraph, as a
 * redline splits one a change rewrote word by word, shaded whole: a run partly
 * in the marker is split at the marker's edge and only its marker text is
 * shaded. highlightOpenItemsInDocx shades a marker inside one text. The words
 * around a marker, deleted text and a run holding more than one text keep
 * their formatting. */
export function highlightSplitOpenItemsInDocx(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  const document = parts["word/document.xml"]
  if (!document) return bytes
  const doc = new DOMParser().parseFromString(strFromU8(document), "application/xml")
  const spans = new Map<Element, [number, number, string][]>()
  for (const paragraph of Array.from(doc.getElementsByTagNameNS(W, "p"))) {
    let at = 0
    const texts = Array.from(paragraph.getElementsByTagNameNS(W, "t")).filter(text => owningParagraph(text) === paragraph)
      .map(text => ({ text, start: at, end: (at += (text.textContent ?? "").length) }))
    for (const marker of texts.map(row => row.text.textContent ?? "").join("").matchAll(OPEN_ITEM_RE)) {
      const start = marker.index ?? 0
      const end = start + marker[0].length
      const covered = texts.filter(row => row.start < end && start < row.end)
      if (covered.length > 1) for (const row of covered)
        spans.set(row.text, [...spans.get(row.text) ?? [], [Math.max(start, row.start) - row.start, Math.min(end, row.end) - row.start, markerFill(marker[0])]])
    }
  }
  let changed = false
  for (const [text, inside] of spans) changed = shadeTextSpans(text, inside) || changed
  if (!changed) return bytes
  parts["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

function owningParagraph(node: Node) {
  for (let parent = node.parentNode; parent; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && (parent as Element).localName === "p") return parent
  return null
}

// The run of `text` split so each [from, to) of it is a run of its own, shaded
// with its fill; the rest keeps the run's formatting.
function shadeTextSpans(text: Element, inside: [number, number, string][]) {
  const run = text.parentNode as Element | null
  const content = run ? Array.from(run.childNodes).filter(node => node.nodeType === 1 && (node as Element).localName !== "rPr") : []
  if (run?.localName !== "r" || content.length !== 1) return false
  const value = text.textContent ?? ""
  const pieces: { value: string; fill?: string }[] = []
  let cursor = 0
  for (const [from, to, fill] of inside.toSorted((a, b) => a[0] - b[0])) {
    if (from > cursor) pieces.push({ value: value.slice(cursor, from) })
    pieces.push({ value: value.slice(from, to), fill })
    cursor = to
  }
  if (cursor < value.length) pieces.push({ value: value.slice(cursor) })
  if (pieces.length === 1) return shadeRunElement(run, pieces[0]!.fill!)
  for (const piece of pieces) {
    const copy = run.cloneNode(true) as Element
    const node = copy.getElementsByTagNameNS(W, "t")[0]!
    while (node.firstChild) node.removeChild(node.firstChild)
    node.appendChild(copy.ownerDocument!.createTextNode(piece.value))
    node.setAttributeNS(XML, "xml:space", "preserve")
    if (piece.fill) shadeRunElement(copy, piece.fill)
    run.parentNode!.insertBefore(copy, run)
  }
  run.parentNode!.removeChild(run)
  return true
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
