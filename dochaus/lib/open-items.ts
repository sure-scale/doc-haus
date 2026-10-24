import { OPEN_ITEM_PATTERN } from "./marker-patterns"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

export const OPEN_ITEM_HIGHLIGHT_FILL = "FFF2CC"
export const FIELD_HIGHLIGHT_FILL = "DDEBF7"

const OPEN_ITEM_RE = OPEN_ITEM_PATTERN
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
