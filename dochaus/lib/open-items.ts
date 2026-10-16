import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

export const OPEN_ITEM_HIGHLIGHT_FILL = "FFF2CC"
export const FIELD_HIGHLIGHT_FILL = "DDEBF7"

const OPEN_ITEM_RE = /\[open item:[^\]\r\n]*\]|\(open item:[^)\r\n]*\)|_{3,}/gi
const RUN_RE = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g
const TEXT_RE = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/

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
  return zipSync(parts)
}

function highlightOpenItemsInDocumentXml(xml: string) {
  return xml.replace(RUN_RE, (run) => highlightRunOpenItems(run))
}

function highlightRunOpenItems(run: string) {
  const text = TEXT_RE.exec(run)
  if (!text) return run
  const content = text[2] ?? ""
  const matches = [...content.matchAll(OPEN_ITEM_RE)]
  if (matches.length === 0) return run

  const textStart = text.index + text[1]!.length
  const textEnd = textStart + content.length
  const beforeText = run.slice(0, textStart)
  const afterText = run.slice(textEnd)
  let offset = 0
  return matches
    .flatMap((match) => {
      const start = match.index ?? 0
      const before = content.slice(offset, start)
      const marker = match[0]
      offset = start + marker.length
      return [
        before ? `${beforeText}${before}${afterText}` : "",
        `${highlightRun(beforeText, markerFill(marker))}${marker}${afterText}`,
      ].filter(Boolean)
    })
    .concat(content.slice(offset) ? `${beforeText}${content.slice(offset)}${afterText}` : [])
    .join("")
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
