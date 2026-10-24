import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { MarkdownAnchorTarget } from "docxodus"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"

// Word draws a numbered paragraph's label ("1.", "1.1", "(a)") from the list
// definition, not from the paragraph's runs, so plain-text extraction (mammoth)
// drops it: a natively numbered clause reads without its number in search,
// sections, citations and quote checks. Both lockstep readers (ingest's
// extractDocumentText and dochaus/lib/extract.ts) pass a Word file through
// writeListLabels before markTrackedChanges and mammoth: each numbered body
// paragraph (tables included) gets its label as literal text, followed by one
// space. Labels come from Docxodus, the engine that numbers paragraph rows and
// views elsewhere, so the text shows the numbers the other readers show. They
// are read from the file as saved: markTrackedChanges keeps a deleted clause as
// a paragraph and adds a summary paragraph, which Docxodus would number too.
//
// Only document.xml is labelled: mammoth's plain text reads the body and its
// tables, not headers, footers, footnotes or comments. Bullets take no label;
// nor do empty paragraphs (they still count in the numbering), text-box
// paragraphs (Docxodus numbers them apart from the body) and clauses a pending
// change deletes or moves away (they do not count). A file without a
// numbered list, or one Docxodus cannot read, is returned as the same buffer,
// so its text and offsets do not change.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XML = "http://www.w3.org/XML/1998/namespace"
const BODY = "word/document.xml"
const NUMBERING = "word/numbering.xml"
const STYLES = "word/styles.xml"
const NUM_PR = /<(?:[\w.-]+:)?numPr[\s/>]/
// Paragraph anchors in the projection (plain, list item, heading); the others
// are containers (tables, rows, cells, columns, content controls, sections).
const PARAGRAPHS = new Set(["p", "li", "h"])

export async function writeListLabels(buffer: Buffer): Promise<Buffer> {
  try {
    const head = unzipSync(buffer, { filter: file => [BODY, NUMBERING, STYLES].includes(file.name) })
    const body = head[BODY] && strFromU8(head[BODY])
    // Opening Docxodus costs more than mammoth itself, so a file whose body
    // cannot be numbered skips it.
    if (!body || !head[NUMBERING] || !usesNumbering(body, head[STYLES])) return buffer
    const levels = listLevels(head[NUMBERING])
    if (!levels.numbered) return buffer
    const anchors = await bodyParagraphs(buffer, levels.bullets)
    if (!anchors.some(anchor => anchor.label)) return buffer
    const doc = parse(body)
    // Docxodus lists every body paragraph in document order, as the DOM does,
    // so each label goes to the paragraph at its own position: two clauses of
    // the same text keep their own numbers. A paragraph the two read with
    // different text takes no label, and a file they count differently none.
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W, "p"))
    if (paragraphs.length !== anchors.length) return buffer
    let written = 0
    anchors.forEach(({ preview, label }, at) => {
      const node = paragraphs[at]!
      const text = flatText(node)
      if (!label || !text || !sameText(preview, text) || within(node, "txbxContent") || movedAway(node)) return
      const pPr = Array.from(node.childNodes).find(child => (child as Element).namespaceURI === W && (child as Element).localName === "pPr")
      node.insertBefore(labelRun(doc, node, label), pPr ? pPr.nextSibling : node.firstChild)
      written++
    })
    if (!written) return buffer
    const parts = unzipSync(buffer)
    parts[BODY] = strToU8(new XMLSerializer().serializeToString(doc))
    // mammoth reads the copy once and discards it, so it is stored uncompressed.
    return Buffer.from(zipSync(parts, { level: 0 }))
  } catch {
    return buffer
  }
}

/** The body's paragraph anchors in document order, each with its text preview
 * and, when numbered (not bulleted), its label. */
async function bodyParagraphs(buffer: Buffer, bullets: Set<string>) {
  const dx = await docxodus()
  const session = dx.openDocxSession(new Uint8Array(buffer), { smartQuotes: false, persistAnchorIds: false })
  try {
    // The session's projection type omits the prefix its anchors carry.
    return Object.entries(session.project().anchorIndex as Record<string, MarkdownAnchorTarget>)
      .filter(([, anchor]) => anchor.partUri === `/${BODY}` && PARAGRAPHS.has(anchor.kind))
      .map(([id, anchor]) => {
        const label = anchor.autoNumberPrefix?.trim()
        // Docxodus labels a bullet whose glyph is a letter (Word's second-level
        // bullet is "o"). Only a label that matches a bullet glyph is checked
        // against its list level: the lookup is slow on a long document.
        const numbered = label && /[\p{L}\p{N}]/u.test(label)
          && !(bullets.has(label) && session.getListMembership(id)?.format === "bullet")
        return { preview: anchor.textPreview, label: numbered ? label : undefined }
      })
  } finally {
    session.close()
  }
}

/** A body paragraph carries list numbering directly or through its paragraph
 * style (the default one included) or a style that style is based on. Most
 * Word files ship list styles nobody uses, and the house seed's Subtitle style
 * names a list level without a list. styles.xml runs to hundreds of kilobytes
 * of latent styles, so it is scanned, not parsed: a parse would cost about as
 * much as mammoth's whole read. */
function usesNumbering(body: string, stylesPart?: Uint8Array) {
  if (NUM_PR.test(body)) return true
  const styles = stylesPart && strFromU8(stylesPart)
  if (!styles || !NUM_PR.test(styles)) return false
  const paragraphStyles = new Map<string, { list?: string; basedOn?: string }>()
  let fallback: string | undefined
  for (const [, , tag = "", content = ""] of styles.matchAll(/<([\w.-]+:)?style\b([^>]*)>([\s\S]*?)<\/\1style>/g)) {
    const id = attribute(tag, "styleId")
    if (!id || attribute(tag, "type") !== "paragraph") continue
    paragraphStyles.set(id, { list: childValue(content, "numId"), basedOn: childValue(content, "basedOn") })
    if (["1", "true", "on"].includes(attribute(tag, "default") ?? "")) fallback = id
  }
  const numbered = (id: string | undefined, seen = new Set<string>()): boolean => {
    const style = id === undefined || seen.has(id) ? undefined : paragraphStyles.get(id)
    if (!style) return false
    seen.add(id!)
    return style.list ? style.list !== "0" : numbered(style.basedOn, seen)
  }
  const used = [...body.matchAll(/<(?:[\w.-]+:)?pStyle\b([^>]*)/g)].map(([, tag = ""]) => attribute(tag, "val"))
  return [...new Set([...used, fallback])].some(id => numbered(id))
}

function attribute(tag: string, name: string) {
  return new RegExp(`(?:^|\\s)(?:[\\w.-]+:)?${name}="([^"]*)"`).exec(tag)?.[1]
}

/** The w:val of the first `name` element in `xml`. */
function childValue(xml: string, name: string) {
  const tag = new RegExp(`<(?:[\\w.-]+:)?${name}\\b([^>]*)`).exec(xml)?.[1]
  return tag === undefined ? undefined : attribute(tag, "val")
}

/** Whether some list level numbers rather than bullets (a level without a
 * format is decimal), and the glyphs the bullet levels draw. */
function listLevels(part: Uint8Array) {
  let numbered = false
  const bullets = new Set<string>()
  for (const level of Array.from(parse(strFromU8(part)).getElementsByTagNameNS(W, "lvl"))) {
    const formats = Array.from(level.getElementsByTagNameNS(W, "numFmt")).map(format => format.getAttributeNS(W, "val"))
    if (!formats.length || formats.some(format => format !== "bullet")) numbered = true
    else bullets.add(Array.from(level.getElementsByTagNameNS(W, "lvlText"))[0]?.getAttributeNS(W, "val")?.trim() ?? "")
  }
  return { numbered, bullets }
}

/** A paragraph's run text as the projection previews it: every w:t inside it. */
function flatText(node: Element) {
  return Array.from(node.getElementsByTagNameNS(W, "t")).map(text => text.textContent ?? "").join("")
}

/** The projection previews about 80 characters and marks a cut with "…". */
function sameText(preview: string, text: string) {
  if (preview === text) return true
  const cut = preview.endsWith("…") ? preview.slice(0, -1) : undefined
  return cut !== undefined && text.length > cut.length && text.startsWith(cut)
}

function within(node: Node, localName: string) {
  for (let parent = node.parentNode; parent && parent.nodeType === 1; parent = parent.parentNode)
    if ((parent as Element).namespaceURI === W && (parent as Element).localName === localName) return true
  return false
}

/** A pending move of a whole clause: all its text is marked as moved away.
 * Docxodus numbers it like the clause after it, as it does a deleted clause,
 * whose text is gone from w:t; the number belongs where the clause moves to. */
function movedAway(node: Element) {
  return Array.from(node.getElementsByTagNameNS(W, "t")).every(text => !text.textContent || within(text, "moveFrom"))
}

function labelRun(doc: Document, like: Element, label: string) {
  const qualified = (name: string) => (like.prefix ? `${like.prefix}:${name}` : name)
  const run = doc.createElementNS(W, qualified("r"))
  const text = doc.createElementNS(W, qualified("t"))
  text.setAttributeNS(XML, "xml:space", "preserve")
  text.appendChild(doc.createTextNode(`${label} `))
  run.appendChild(text)
  return run
}

function parse(xml: string) {
  const fail = (message: string) => { throw new Error(`invalid_docx_xml:${message}`) }
  return new DOMParser({ errorHandler: { warning: () => {}, error: () => {}, fatalError: fail } }).parseFromString(xml, "application/xml")
}
