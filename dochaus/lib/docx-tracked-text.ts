import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import mammoth from "mammoth"
import { writeLineBreaks } from "./docx-line-breaks"
import { normalizeUntrustedText } from "./untrusted-text"

// Plain-text extraction (mammoth) keeps tracked insertions as ordinary text and
// drops tracked deletions and moves, so a counterparty markup reads as if every
// proposal were agreed. Both lockstep readers (ingest's extractDocumentText and
// dochaus/lib/extract.ts) pass a Word file through markTrackedChanges before
// mammoth: each pending revision becomes labelled text in the body, and a first
// paragraph says the file carries tracked changes. A file without revisions is
// returned as the same buffer, so its text and offsets do not change.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XML = "http://www.w3.org/XML/1998/namespace"
const BODY = "word/document.xml"
const COMMENTS = "word/comments.xml"
const REVISION_MARKUP = /<(?:[^\s<>/:]+:)?(?:ins|del|moveFrom|moveTo)[\s/>]/
const KINDS = ["ins", "del", "moveFrom", "moveTo"]
const REMOVALS = ["del", "moveFrom"]
const LABEL: Record<string, string> = {
  ins: "Tracked insertion",
  del: "Tracked deletion",
  moveFrom: "Tracked move from here",
  moveTo: "Tracked move to here",
}
const MAX_NAMED_AUTHORS = 5

export function markTrackedChanges(buffer: Buffer, options: { structural?: true } = {}): Buffer {
  const body = unzipSync(buffer, { filter: file => file.name === BODY })[BODY]
  if (!body || !REVISION_MARKUP.test(strFromU8(body))) return buffer
  const parts = unzipSync(buffer)
  const doc = parse(strFromU8(body))
  // A deleted paragraph mark makes mammoth join the paragraph to the next, and a
  // deleted table row makes it drop the row; keep both so their labelled text
  // stays where it was. A row is kept only when all its text is itself tracked
  // as deleted, or the row would read as agreed text.
  for (const marker of revisionElements(doc).filter(isPropertyMarker)) {
    const holder = marker.parentNode as Element
    if (!REMOVALS.includes(marker.localName) || !["rPr", "trPr"].includes(holder.localName)) continue
    if (holder.localName === "trPr" && untrackedText(holder.parentNode as Element)) continue
    holder.removeChild(marker)
  }
  const groups = revisionGroups(revisionElements(doc).filter(element => !isPropertyMarker(element) && readable(element)))
  const labelled = new Set(groups.filter(group => group.some(hasText)))
  // Innermost and latest first, so unwrapping a group never moves one still to come.
  for (const group of [...groups].reverse()) {
    if (labelled.has(group)) {
      const label = `${LABEL[group[0]!.localName]}${byline(group[0]!)}: `
      group[0]!.parentNode!.insertBefore(textRun(doc, group[0]!, options.structural ? blankText(`[${label}`) : `[${label}`), group[0]!)
      const last = group[group.length - 1]!
      last.parentNode!.insertBefore(textRun(doc, last, options.structural ? " " : "]"), last.nextSibling)
    }
    // A removal with no wording (a break or symbol) stays unread, as mammoth reads it.
    for (const element of group)
      if (labelled.has(group) || !REMOVALS.includes(element.localName)) {
        if (options.structural && REMOVALS.includes(element.localName)) blankRevision(element)
        unwrap(element)
      }
      else element.parentNode!.removeChild(element)
  }
  if (!labelled.size) return buffer
  const container = doc.getElementsByTagNameNS(W, "body")[0]!
  const notice = summary([...labelled], commentCount(parts[COMMENTS]))
  container.insertBefore(paragraph(doc, container, options.structural ? blankText(notice) : notice), container.firstChild)
  parts[BODY] = strToU8(new XMLSerializer().serializeToString(doc))
  // mammoth reads the copy once and discards it, so it is stored uncompressed.
  return Buffer.from(zipSync(parts, { level: 0 }))
}

// Structural interpretation is supplied by extraction, never inferred from source
// punctuation. Legacy documents without this projection keep literal text until
// re-ingested; canonical source text and all citation offsets remain unchanged.
export type StructuralMask = [number, number][]
export function proposedView(text: string, mask: StructuralMask = []) {
  if (!mask.length) return text
  let previous = 0
  const view = text.split("")
  for (const [from, to] of mask) {
    if (!Number.isSafeInteger(from) || !Number.isSafeInteger(to) || from < previous || to <= from || to > text.length)
      throw new Error("invalid_structural_mask")
    for (let at = from; at < to; at++) if (view[at] !== "\n" && view[at] !== "\r") view[at] = " "
    previous = to
  }
  return view.join("")
}

export function structuralMask(text: string, view: string): StructuralMask {
  if (text.length !== view.length) throw new Error("structural_text_offset_mismatch")
  const ranges: StructuralMask = []
  for (let at = 0; at < text.length; at++) {
    if (text[at] === view[at]) continue
    if (view[at] !== " " || /[\r\n]/.test(text[at]!)) throw new Error("invalid_structural_text")
    const last = ranges.at(-1)
    if (last?.[1] === at) last[1]++
    else ranges.push([at, at + 1])
  }
  return ranges
}

/** Two views of the same XML extraction. Only actual revision DOM nodes and
 * labels we insert are masked; identical marker-like source text stays literal. */
export async function docxTextViews(buffer: Buffer, options: { structural?: boolean } = { structural: true }) {
  const labelled = markTrackedChanges(buffer)
  const text = (await mammoth.extractRawText({ buffer: writeLineBreaks(labelled) })).value
  if (!options.structural || labelled === buffer) return { text, mask: [] as StructuralMask }
  const view = normalizeUntrustedText((await mammoth.extractRawText({ buffer: writeLineBreaks(markTrackedChanges(buffer, { structural: true })) })).value).text
  return { text, mask: structuralMask(normalizeUntrustedText(text).text, view) }
}

// Preserve invisible scalars until the shared normalizer removes them from both
// views. Mask UTF-16 units so astral text keeps the canonical offsets too.
function blankText(text: string) {
  return text.replace(/[^\r\n]/gu, char => normalizeUntrustedText(char).text ? " ".repeat(char.length) : char)
}
function blankRevision(element: Element) {
  for (const node of Array.from(element.getElementsByTagNameNS(W, "*"))) {
    if (["t", "delText"].includes(node.localName)) node.textContent = blankText(node.textContent ?? "")
    if (node.localName === "tab") node.parentNode!.replaceChild(textRun(element.ownerDocument!, node, " ").firstChild!, node)
  }
}

function parse(xml: string) {
  const fail = (message: string) => { throw new Error(`invalid_docx_xml:${message}`) }
  return new DOMParser({ errorHandler: { warning: () => {}, error: () => {}, fatalError: fail } }).parseFromString(xml, "application/xml")
}
/** Revision elements in document order. */
function revisionElements(doc: Document) {
  return Array.from(doc.getElementsByTagNameNS(W, "*")).filter(element => KINDS.includes(element.localName))
}
/** A revision recorded on a paragraph mark, run, row or numbering property, not around content. */
function isPropertyMarker(element: Element) {
  return /Pr$/.test((element.parentNode as Element | null)?.localName ?? "")
}
/** mammoth skips simple fields, custom XML, the preferred branch of alternate
 * content (it reads the fallback copy) and a table row still marked deleted. A
 * revision there stays unread, so it is neither labelled nor counted. */
function readable(element: Element) {
  for (let parent = element.parentNode; parent && parent.nodeType === 1; parent = parent.parentNode) {
    const { namespaceURI, localName } = parent as Element
    if (namespaceURI === W && ["fldSimple", "customXml"].includes(localName)) return false
    if (localName === "Choice") return false
    if (namespaceURI === W && localName === "tr" && Array.from(parent.childNodes).some(child => (child as Element).localName === "trPr"
      && Array.from(child.childNodes).some(marker => REMOVALS.includes((marker as Element).localName))))
      return false
  }
  return true
}
function untrackedText(container: Element) {
  return Array.from(container.getElementsByTagNameNS(W, "t")).some(node => (node.textContent ?? "").length > 0
    && !ancestors(node).some(parent => parent.namespaceURI === W && REMOVALS.includes(parent.localName)))
}
function ancestors(node: Node) {
  const found: Element[] = []
  for (let parent = node.parentNode; parent && parent.nodeType === 1; parent = parent.parentNode) found.push(parent as Element)
  return found
}
/** Word splits one edit into adjacent revisions (new rsid, bookmark, proofing
 * mark); siblings of the same kind and author with only such markup between
 * them read as one change. */
function revisionGroups(elements: Element[]) {
  const groups: Element[][] = []
  const ending = new Map<Element, Element[]>()
  for (const element of elements) {
    const previous = previousContent(element)
    const joined = previous && previous.localName === element.localName && author(previous) === author(element)
      ? ending.get(previous) : undefined
    const group = joined ?? []
    if (!joined) groups.push(group)
    group.push(element)
    ending.set(element, group)
  }
  return groups
}
function previousContent(node: Node) {
  let sibling = node.previousSibling
  while (sibling && (sibling.nodeType !== 1 || ((sibling as Element).namespaceURI === W
    && /^(bookmark(Start|End)|\w+Range(Start|End)|proofErr|perm(Start|End))$/.test((sibling as Element).localName))))
    sibling = sibling.previousSibling
  return sibling as Element | null
}
function hasText(element: Element) {
  return ["t", "delText", "tab"].some(name => Array.from(element.getElementsByTagNameNS(W, name))
    .some(node => name === "tab" || (node.textContent ?? "").length > 0))
}
/** Move a revision's content into its parent, turning deleted text into text mammoth reads. */
function unwrap(element: Element) {
  if (REMOVALS.includes(element.localName))
    for (const [from, to] of [["delText", "t"], ["delInstrText", "instrText"]])
      for (const node of Array.from(element.getElementsByTagNameNS(W, from!))) {
        const renamed = element.ownerDocument!.createElementNS(W, qualified(node, to!))
        for (const attribute of Array.from(node.attributes)) renamed.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value)
        while (node.firstChild) renamed.appendChild(node.firstChild)
        node.parentNode!.replaceChild(renamed, node)
      }
  const parent = element.parentNode!
  while (element.firstChild) parent.insertBefore(element.firstChild, element)
  parent.removeChild(element)
}
function author(element: Element) {
  return (element.getAttributeNS(W, "author") ?? "").replace(/[[\]:]/g, " ").replace(/\s+/g, " ").trim().slice(0, 80)
}
function byline(element: Element) {
  const name = author(element)
  return name ? ` by ${name}` : ""
}
function summary(groups: Element[][], comments: number) {
  const authors = [...new Set(groups.map(group => author(group[0]!)).filter(Boolean))]
  const named = authors.length > MAX_NAMED_AUTHORS
    ? `${authors.slice(0, MAX_NAMED_AUTHORS).join(", ")} and ${authors.length - MAX_NAMED_AUTHORS} others`
    : authors.join(", ")
  const count = `${groups.length} pending tracked ${groups.length === 1 ? "change" : "changes"}`
  // Comments are not in this text: the viewer's panel and the comments view list
  // them, as threads, so a count here (each reply counted) would not match.
  const note = comments ? " Its comments are listed separately." : ""
  return `[This Word file contains ${count}${named ? ` by ${named}` : ""}, marked inline below. Pending tracked changes are not accepted text.${note}]`
}
function commentCount(part: Uint8Array | undefined) {
  return part ? parse(strFromU8(part)).getElementsByTagNameNS(W, "comment").length : 0
}
function qualified(like: Element, name: string) {
  return like.prefix ? `${like.prefix}:${name}` : name
}
function textRun(doc: Document, like: Element, text: string) {
  const run = doc.createElementNS(W, qualified(like, "r"))
  const node = doc.createElementNS(W, qualified(like, "t"))
  node.setAttributeNS(XML, "xml:space", "preserve")
  node.appendChild(doc.createTextNode(text))
  run.appendChild(node)
  return run
}
function paragraph(doc: Document, like: Element, text: string) {
  const node = doc.createElementNS(W, qualified(like, "p"))
  node.appendChild(textRun(doc, like, text))
  return node
}
