import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// Plain-text extraction (mammoth) reads only text, tabs and paragraph ends, so
// lines a Word file splits with line breaks inside one paragraph (an address
// block, a letterhead) run together. Every reader of a Word file's text passes
// it through writeLineBreaks after markTrackedChanges: each line break becomes a
// newline in the text. Page and column breaks are not lines, and a break in a
// pending deletion is not current wording, so both stay unread. A file without
// line breaks is returned as the same buffer, so its text and offsets do not change.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XML = "http://www.w3.org/XML/1998/namespace"
const BODY = "word/document.xml"
const BREAK_MARKUP = /<(?:[^\s<>/:]+:)?br[\s/>]/
const REMOVALS = ["del", "moveFrom"]

export function writeLineBreaks(buffer: Buffer): Buffer {
  const body = unzipSync(buffer, { filter: file => file.name === BODY })[BODY]
  if (!body || !BREAK_MARKUP.test(strFromU8(body))) return buffer
  const fail = (message: string) => { throw new Error(`invalid_docx_xml:${message}`) }
  const doc = new DOMParser({ errorHandler: { warning: () => {}, error: () => {}, fatalError: fail } }).parseFromString(strFromU8(body), "application/xml")
  const breaks = Array.from(doc.getElementsByTagNameNS(W, "br")).filter(element =>
    ["", "textWrapping"].includes(element.getAttributeNS(W, "type") ?? "") && !removed(element))
  if (!breaks.length) return buffer
  for (const element of breaks) {
    const prefix = element.prefix ? `${element.prefix}:` : ""
    const text = doc.createElementNS(W, `${prefix}t`)
    text.setAttributeNS(XML, "xml:space", "preserve")
    text.appendChild(doc.createTextNode("\n"))
    element.parentNode!.replaceChild(text, element)
  }
  const parts = unzipSync(buffer)
  parts[BODY] = strToU8(new XMLSerializer().serializeToString(doc))
  // mammoth reads the copy once and discards it, so it is stored uncompressed.
  return Buffer.from(zipSync(parts, { level: 0 }))
}

function removed(element: Element) {
  for (let node = element.parentNode; node; node = node.parentNode)
    if (node.nodeType === 1 && (node as Element).namespaceURI === W && REMOVALS.includes((node as Element).localName)) return true
  return false
}
