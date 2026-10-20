import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

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
const REVISION_MARKUP = /<(?:[\w.-]+:)?(?:ins|del|moveFrom|moveTo)[\s/>]/
const KINDS = ["ins", "del", "moveFrom", "moveTo"]
const LABEL: Record<string, string> = {
  ins: "Tracked insertion",
  del: "Tracked deletion",
  moveFrom: "Tracked move from here",
  moveTo: "Tracked move to here",
}
const MAX_NAMED_AUTHORS = 5

export function markTrackedChanges(buffer: Buffer): Buffer {
  const body = unzipSync(buffer, { filter: file => file.name === BODY })[BODY]
  if (!body || !REVISION_MARKUP.test(strFromU8(body))) return buffer
  const parts = unzipSync(buffer)
  const doc = parse(strFromU8(body))
  // A deleted paragraph mark makes mammoth join the paragraph to the next, and a
  // deleted table row makes it drop the row; keep both so their labelled text
  // stays where it was.
  for (const marker of revisionElements(doc).filter(isPropertyMarker))
    if (["del", "moveFrom"].includes(marker.localName) && ["rPr", "trPr"].includes((marker.parentNode as Element).localName))
      marker.parentNode!.removeChild(marker)
  const groups = revisionGroups(revisionElements(doc).filter(element => !isPropertyMarker(element)))
  const labelled = groups.filter(group => group.some(hasText))
  // Innermost and latest first, so unwrapping a group never moves one still to come.
  for (const group of [...groups].reverse()) {
    if (labelled.includes(group)) {
      const label = `${LABEL[group[0]!.localName]}${byline(group[0]!)}: `
      group[0]!.parentNode!.insertBefore(textRun(doc, group[0]!, `[${label}`), group[0]!)
      const last = group[group.length - 1]!
      last.parentNode!.insertBefore(textRun(doc, last, "]"), last.nextSibling)
    }
    for (const element of group) unwrap(element)
  }
  if (!labelled.length) return buffer
  const container = doc.getElementsByTagNameNS(W, "body")[0]!
  container.insertBefore(paragraph(doc, container, summary(labelled, commentCount(parts[COMMENTS]))), container.firstChild)
  parts[BODY] = strToU8(new XMLSerializer().serializeToString(doc))
  return Buffer.from(zipSync(parts))
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
  if (["del", "moveFrom"].includes(element.localName))
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
  const note = comments ? ` It also contains ${comments} ${comments === 1 ? "comment" : "comments"}, which ${comments === 1 ? "is" : "are"} not shown.` : ""
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
