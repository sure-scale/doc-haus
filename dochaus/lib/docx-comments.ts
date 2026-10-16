import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const R = "http://schemas.openxmlformats.org/package/2006/relationships"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const COMMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
const COMMENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
const BODY = "word/document.xml"
const RELS = "word/_rels/document.xml.rels"

export type DocxComment = {
  id: string
  text: string
  author: string
  anchor: {
    paragraph_text?: string
    quote?: string
    pid?: string
    /** Zero-based ordinal among nonempty accepted body paragraphs, including tables. */
    paragraph_index?: number
    /** An exact tracked-deleted paragraph, not an accepted-text anchor. */
    deleted_paragraph_text?: string
    /** A contiguous range of exact accepted paragraphs for one explanation. */
    paragraphs?: { paragraph_index: number; paragraph_text: string }[]
  }
}
export type AttachedDocxComment = { id: string; comment_id: string; paragraph_index: number; paragraph_text: string; anchor_kind?: "deleted"; paragraphs?: { paragraph_index: number; paragraph_text: string }[] }
export type UnresolvedDocxComment = { id: string; reason: string }

/** Native accepted-text view shared with attachment. Tabs and line breaks are
 * retained verbatim; deleted-only and empty paragraphs have no ordinal. */
export function readDocxCommentParagraphs(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  if (!parts[BODY]) throw new Error("invalid_docx_package")
  return acceptedParagraphs(parse(strFromU8(parts[BODY]))).map(row => ({
    paragraph_index: row.index, paragraph_text: row.text, pids: row.pids,
  }))
}

/** Reads actual tracked text from the exported OOXML package. The accepted
 * clean body is a separate artifact and cannot establish redline presence. */
export function readDocxRevisionText(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  if (!parts[BODY]) throw new Error("invalid_docx_package")
  const body = parse(strFromU8(parts[BODY]))
  const text = (names: string[]) => names.flatMap(name => elements(body, W, name).map(row => elements(row, W, "*").map(wordText).join(""))).filter(Boolean)
  return { inserted_text: text(["ins", "moveTo"]), deleted_text: text(["del", "moveFrom"]) }
}

/** Counts actual revision elements in the main Word body only. These are
 * package observations, not a judgment that a comparison or legal review is complete. */
export function readDocxRevisionProperties(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  if (!parts[BODY]) throw new Error("invalid_docx_package")
  const body = parse(strFromU8(parts[BODY]))
  const names = ["ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange", "numberingChange"]
  return { part: BODY, revision_elements: Object.fromEntries(names.map(name => [name, elements(body, W, name).length])) }
}

/** Read native comment text separately from current body text. Anchor records
 * identify paragraphs containing range markers, not the exact commented span.
 * paragraph_index is an accepted-text ordinal; xml_paragraph_index includes
 * empty/deleted paragraphs. PIDs and ordinals belong to this artifact only. */
export function readDocxComments(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  if (!parts[BODY]) throw new Error("invalid_docx_package")
  const body = parse(strFromU8(parts[BODY]))
  const relationships = parts[RELS] ? parse(strFromU8(parts[RELS])) : undefined
  const links = relationships ? elements(relationships, R, "Relationship").filter(row => row.getAttribute("Type") === COMMENT_REL) : []
  if (!links.length) return []
  if (links.length !== 1 || links[0]!.getAttribute("TargetMode") === "External") throw new Error("invalid_docx_comments_relationship")
  const target = links[0]!.getAttribute("Target")
  if (!target) throw new Error("invalid_docx_comments_relationship")
  const commentPart = path.posix.resolve("/word", target).slice(1)
  if (!parts[commentPart]) throw new Error("missing_docx_comments_part")
  const doc = parse(strFromU8(parts[commentPart]))
  const accepted = acceptedParagraphs(body)
  const allParagraphs = elements(body, W, "p")
  return elements(doc, W, "comment").map(comment => {
    const id = comment.getAttributeNS(W, "id") ?? ""
    const markers = ["commentRangeStart", "commentRangeEnd", "commentReference"].flatMap(kind =>
      elements(body, W, kind).filter(marker => marker.getAttributeNS(W, "id") === id).map(marker => {
        const paragraph = owningParagraph(marker) as Element | null
        const row = accepted.find(candidate => candidate.node === paragraph)
        return { kind, paragraph_index: row?.index ?? null, xml_paragraph_index: paragraph ? allParagraphs.indexOf(paragraph) : null, pids: paragraph ? paragraphPIDs(paragraph) : [],
          accepted_paragraph_text: paragraph ? acceptedText(paragraph) : "",
          deleted_paragraph_text: paragraph ? deletedText(paragraph) : "" }
      }))
    return { comment_id: id, text: elements(comment, W, "p").map(acceptedText).join("\n"), markers }
  })
}

/** Comparison can discard native comments. Restore their original paragraphs
 * and package records only where every paragraph in each range is unchanged.
 * Edited, empty or non-body anchors require explicit review, never relocation. */
export function preserveDocxComments(original: Uint8Array, compared: Uint8Array) {
  const comments = readDocxComments(original)
  if (!comments.length) return compared
  if (readDocxComments(compared).length) throw new Error("comparison_already_has_native_comments")
  const source = unzipSync(original)
  const output = unzipSync(compared)
  const before = parse(strFromU8(source[BODY]))
  const after = parse(strFromU8(output[BODY]))
  const from = acceptedParagraphs(before)
  const to = acceptedParagraphs(after)
  const relationships = output[RELS] ? parse(strFromU8(output[RELS])) : parse(`<Relationships xmlns="${R}"/>`)
  const sourceRelationships = parse(strFromU8(source[RELS]))
  if (from.length !== to.length) throw new Error("original_comment_paragraph_structure_changed")
  const restored = new Set<number>()
  for (const comment of comments) {
    const indices = ["commentRangeStart", "commentRangeEnd", "commentReference"].map(kind => {
      const markers = comment.markers.filter(marker => marker.kind === kind)
      if (markers.length !== 1 || markers[0].paragraph_index === null) throw new Error("unsupported_original_comment_anchor")
      return markers[0].paragraph_index
    })
    if (indices[0] > indices[1] || indices[1] > indices[2]) throw new Error("unsupported_original_comment_range")
    for (let index = indices[0]; index <= indices[2]; index++) {
      if (from[index].text !== to[index].text || elements(to[index].node, W, "*").some(node =>
        ["ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange"].includes(node.localName)))
        throw new Error("original_comment_range_was_edited")
      restored.add(index)
    }
  }
  // A paragraph's relationship IDs and style/numbering IDs belong to its
  // package. Equal text alone does not authorize transplanting those bindings.
  for (const index of restored) {
    const nodes = [from[index].node, ...Array.from(from[index].node.getElementsByTagName("*"))]
    if (nodes.some(node => ["drawing", "pict", "object", "fldSimple", "fldChar", "instrText", "footnoteReference", "endnoteReference"].includes(node.localName)
      || (node.localName === "hyperlink" && node.hasAttributeNS(W, "anchor"))))
      throw new Error("unsupported_original_comment_paragraph_dependency")
    for (const [names, part] of [[["pStyle", "rStyle", "numPr"], "word/styles.xml"], [["numPr"], "word/numbering.xml"]] as const) {
      if (nodes.some(node => (names as readonly string[]).includes(node.localName)) && (!source[part] || !output[part]
        || source[part].length !== output[part].length || source[part].some((byte, index) => byte !== output[part][index])))
        throw new Error("original_comment_paragraph_dependency_changed")
    }
    for (const node of nodes) for (const attribute of Array.from(node.attributes).filter(attribute => attribute.namespaceURI === O)) {
      const original = elements(sourceRelationships, R, "Relationship").filter(row => row.getAttribute("Id") === attribute.value)
      const compared = elements(relationships, R, "Relationship").filter(row => row.getAttribute("Id") === attribute.value)
      if (original.length !== 1 || compared.length !== 1 || ["Type", "Target", "TargetMode"].some(name => original[0].getAttribute(name) !== compared[0].getAttribute(name)))
        throw new Error("original_comment_paragraph_dependency_changed")
      if (node.namespaceURI !== W || node.localName !== "hyperlink" || attribute.localName !== "id"
        || original[0].getAttribute("Type") !== `${O}/hyperlink` || original[0].getAttribute("TargetMode") !== "External")
        throw new Error("unsupported_original_comment_paragraph_dependency")
    }
  }
  for (const index of restored) to[index].node.parentNode!.replaceChild(after.importNode(from[index].node, true), to[index].node)
  // Threaded comments and comment attachments need their own relationship
  // preservation contract; do not export a partially restored annotation.
  if (elements(sourceRelationships, R, "Relationship").some(row => row.getAttribute("Type") !== COMMENT_REL
    && /\/(comments[^/]*|people)$/.test(row.getAttribute("Type") ?? "")))
    throw new Error("unsupported_original_comment_relationship")
  const relationship = elements(sourceRelationships, R, "Relationship").find(row => row.getAttribute("Type") === COMMENT_REL)!
  const part = path.posix.resolve("/word", relationship.getAttribute("Target")!).slice(1)
  if (source[path.posix.join(path.posix.dirname(part), "_rels", `${path.posix.basename(part)}.rels`)] || output[part])
    throw new Error("unsupported_original_comment_part_dependencies_or_collision")
  const added = relationships.importNode(relationship, true) as Element
  const used = new Set(elements(relationships, R, "Relationship").map(row => row.getAttribute("Id")))
  const allocation = { next: 1 }
  while (used.has(`rId${allocation.next}`)) allocation.next++
  added.setAttribute("Id", `rId${allocation.next}`)
  relationships.documentElement.appendChild(added)
  const contentTypes = parse(strFromU8(output["[Content_Types].xml"]))
  if (elements(contentTypes, CT, "Override").some(row => row.getAttribute("PartName") === `/${part}`))
    throw new Error("original_comment_content_type_collision")
  const override = elements(parse(strFromU8(source["[Content_Types].xml"])), CT, "Override")
    .find(row => row.getAttribute("PartName") === `/${part}` && row.getAttribute("ContentType") === COMMENT_TYPE)
  if (!override) throw new Error("invalid_original_comment_content_type")
  contentTypes.documentElement.appendChild(contentTypes.importNode(override, true))
  const serializer = new XMLSerializer()
  for (const [name, doc] of [[BODY, after], [RELS, relationships], ["[Content_Types].xml", contentTypes]] as const)
    output[name] = strToU8(serializer.serializeToString(doc))
  output[part] = source[part]
  return zipSync(output, { mtime: new Date(1980, 0, 1) })
}

/** Adds native Word comments without accepting or rewriting existing revisions.
 * A quote selects an exact containing paragraph; the comment range covers that
 * entire paragraph. Missing or conflicting pointers never select nearby text.
 * Caller-supplied attribution is kept verbatim; no author or date is invented. */
export function exportDocxComments(bytes: Uint8Array, comments: DocxComment[]) {
  if (new Set(comments.map(comment => comment.id)).size !== comments.length) throw new Error("duplicate_docx_comment_id")
  for (const comment of comments) {
    if (![comment.id, comment.text, comment.author].every(text => typeof text === "string" && text.trim()))
      throw new Error("invalid_docx_comment")
    if ([comment.text, comment.author].some(text => /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text)))
      throw new Error("invalid_docx_comment_xml_character")
  }
  const parts = unzipSync(bytes)
  if (!parts[BODY] || !parts["[Content_Types].xml"]) throw new Error("invalid_docx_package")
  const doc = parse(strFromU8(parts[BODY]))
  const paragraphs = acceptedParagraphs(doc)
  const unresolved: UnresolvedDocxComment[] = []
  const targets = comments.flatMap(comment => {
    const range = comment.anchor.paragraphs
    const anchor = { ...comment.anchor, ...(range?.length ? range[0] : {}) }
    if ((!anchor.paragraph_text && !anchor.quote && !anchor.deleted_paragraph_text) || (anchor.paragraph_index !== undefined
      && (!Number.isSafeInteger(anchor.paragraph_index) || anchor.paragraph_index < 0))) {
      unresolved.push({ id: comment.id, reason: "An exact final paragraph or quote and a valid paragraph position are required." })
      return []
    }
    const deleted = anchor.deleted_paragraph_text !== undefined
    const inventory = deleted ? elements(doc, W, "p").map((node, index) => ({ node, index, text: deletedText(node), pids: [] as string[] }))
      .filter(row => row.text && !acceptedText(row.node).trim()) : paragraphs
    const candidates = inventory.filter(row =>
      (anchor.deleted_paragraph_text === undefined || row.text === anchor.deleted_paragraph_text) &&
      (anchor.paragraph_text === undefined || row.text === anchor.paragraph_text)
      && (anchor.quote === undefined || row.text.includes(anchor.quote))
      && (anchor.paragraph_index === undefined || row.index === anchor.paragraph_index)
      && (anchor.pid === undefined || row.pids.includes(anchor.pid)))
    if (candidates.length !== 1) {
      unresolved.push({ id: comment.id, reason: candidates.length
        ? "More than one accepted paragraph matches the exact anchor. Supply a paragraph position or unique pointer."
        : "No accepted paragraph matches the exact anchor. Deleted text and stale or missing pointers are not attachment targets." })
      return []
    }
    const target = candidates[0]
    const spanning = range ? range.map(row => paragraphs.find(paragraph => paragraph.index === row.paragraph_index && paragraph.text === row.paragraph_text)) : [target]
    if (!spanning.length || spanning.some(row => !row) || spanning.some((row, index) =>
      row!.node.parentNode !== target.node.parentNode || row!.index !== target.index + index)) {
      unresolved.push({ id: comment.id, reason: "The exact accepted paragraph range is not contiguous in one document parent." })
      return []
    }
    const end = spanning.at(-1)!
    // A deleted paragraph mark merges paragraphs on acceptance. Attaching to
    // the old boundary would misrepresent the final paragraph's extent.
    if (!deleted && spanning.some(paragraph => elements(paragraph!.node, W, "del").some(node => node.parentNode?.nodeType === 1
      && (node.parentNode as Element).localName === "rPr" && (node.parentNode?.parentNode as Element)?.localName === "pPr"))) {
      unresolved.push({ id: comment.id, reason: "The paragraph boundary is deleted by a tracked revision; its accepted range is unresolved." })
      return []
    }
    return [{ comment, ...target, end: end!, deleted, range }]
  })
  const attached: AttachedDocxComment[] = []
  if (!targets.length) return { bytes: new Uint8Array(bytes), attached, unresolved }

  const relationships = parts[RELS] ? parse(strFromU8(parts[RELS])) : parse(`<Relationships xmlns="${R}"/>`)
  const existing = elements(relationships, R, "Relationship").filter(row => row.getAttribute("Type") === COMMENT_REL)
  if (existing.length > 1 || existing[0]?.getAttribute("TargetMode") === "External") throw new Error("invalid_docx_comments_relationship")
  const relationshipTarget = existing[0]?.getAttribute("Target")
  if (existing.length && !relationshipTarget) throw new Error("invalid_docx_comments_relationship")
  const commentPart = relationshipTarget ? path.posix.resolve("/word", relationshipTarget).slice(1) : "word/comments.xml"
  if (existing.length && !parts[commentPart]) throw new Error("missing_docx_comments_part")
  const commentDoc = parts[commentPart] ? parse(strFromU8(parts[commentPart])) : parse(`<w:comments xmlns:w="${W}"/>`)
  if (commentDoc.documentElement.namespaceURI !== W || commentDoc.documentElement.localName !== "comments")
    throw new Error("invalid_docx_comments_part")
  const ids = Object.entries(parts).filter(([name]) => name.endsWith(".xml")).flatMap(([, content]) => {
    const part = parse(strFromU8(content))
    return ["comment", "commentRangeStart", "commentRangeEnd", "commentReference"].flatMap(name =>
      elements(part, W, name).map(node => Number(node.getAttributeNS(W, "id"))).filter(Number.isSafeInteger))
  })
  const allocation = { next: Math.max(-1, ...ids) + 1 }
  for (const target of targets) {
    const id = String(allocation.next++)
    const start = doc.createElementNS(W, "w:commentRangeStart")
    start.setAttributeNS(W, "w:id", id)
    const properties = Array.from(target.node.childNodes).find(node => node.nodeType === 1
      && (node as Element).namespaceURI === W && (node as Element).localName === "pPr")
    target.node.insertBefore(start, properties ? properties.nextSibling : target.node.firstChild)
    const end = doc.createElementNS(W, "w:commentRangeEnd")
    end.setAttributeNS(W, "w:id", id)
    target.end.node.appendChild(end)
    const run = doc.createElementNS(W, "w:r")
    const reference = doc.createElementNS(W, "w:commentReference")
    reference.setAttributeNS(W, "w:id", id)
    run.appendChild(reference)
    target.end.node.appendChild(run)
    const native = commentDoc.createElementNS(W, "w:comment")
    native.setAttributeNS(W, "w:id", id)
    native.setAttributeNS(W, "w:author", target.comment.author)
    for (const line of target.comment.text.split(/\r\n|\r|\n/)) {
      const paragraph = commentDoc.createElementNS(W, "w:p")
      const textRun = commentDoc.createElementNS(W, "w:r")
      const text = commentDoc.createElementNS(W, "w:t")
      text.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
      text.appendChild(commentDoc.createTextNode(line))
      textRun.appendChild(text)
      paragraph.appendChild(textRun)
      native.appendChild(paragraph)
    }
    commentDoc.documentElement.appendChild(native)
    attached.push({ id: target.comment.id, comment_id: id, paragraph_index: target.index, paragraph_text: target.text, ...(target.deleted ? { anchor_kind: "deleted" as const } : {}), ...(target.range && target.range.length > 1 ? { paragraphs: target.range } : {}) })
  }
  if (!existing.length) {
    const used = new Set(elements(relationships, R, "Relationship").map(row => row.getAttribute("Id")))
    const numbering = { next: 1 }
    while (used.has(`rId${numbering.next}`)) numbering.next++
    const relationship = relationships.createElementNS(R, "Relationship")
    relationship.setAttribute("Id", `rId${numbering.next}`)
    relationship.setAttribute("Type", COMMENT_REL)
    relationship.setAttribute("Target", path.posix.relative("word", commentPart))
    relationships.documentElement.appendChild(relationship)
  }
  const contentTypes = parse(strFromU8(parts["[Content_Types].xml"]))
  const override = elements(contentTypes, CT, "Override").find(row => row.getAttribute("PartName") === `/${commentPart}`)
  if (override && override.getAttribute("ContentType") !== COMMENT_TYPE) throw new Error("invalid_docx_comments_content_type")
  if (!override) {
    const added = contentTypes.createElementNS(CT, "Override")
    added.setAttribute("PartName", `/${commentPart}`)
    added.setAttribute("ContentType", COMMENT_TYPE)
    contentTypes.documentElement.appendChild(added)
  }
  const serializer = new XMLSerializer()
  for (const [name, xml] of [[BODY, doc], [commentPart, commentDoc], [RELS, relationships], ["[Content_Types].xml", contentTypes]] as const)
    parts[name] = strToU8(serializer.serializeToString(xml))
  return { bytes: zipSync(parts, { mtime: new Date(1980, 0, 1) }), attached, unresolved }
}

function parse(xml: string) {
  const fail = (message: string) => { throw new Error(`invalid_docx_xml:${message}`) }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
}
function elements(node: Document | Element, namespace: string, name: string) {
  return Array.from(node.getElementsByTagNameNS(namespace, name))
}
function owningParagraph(node: Node) {
  for (let parent = node.parentNode; parent; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && (parent as Element).localName === "p") return parent
  return null
}
function inDeletedContent(node: Node) {
  for (let parent: Node | null = node; parent; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && ["del", "moveFrom"].includes((parent as Element).localName)) return true
  return false
}
function acceptedText(paragraph: Element) {
  return Array.from(paragraph.getElementsByTagNameNS(W, "*")).filter(node => owningParagraph(node) === paragraph && !inDeletedContent(node))
    .map(node => node.localName === "t" ? node.textContent ?? "" : node.localName === "tab" ? "\t" : ["br", "cr"].includes(node.localName) ? "\n" : "").join("")
}
function deletedText(paragraph: Element) {
  return elements(paragraph, W, "*").filter(node => owningParagraph(node) === paragraph && inDeletedContent(node)).map(wordText).join("")
}
function wordText(node: Element) {
  return ["t", "delText"].includes(node.localName) ? node.textContent ?? "" : node.localName === "tab" ? "\t" : ["br", "cr"].includes(node.localName) ? "\n" : ""
}
function paragraphPIDs(paragraph: Element) {
  return elements(paragraph, W, "bookmarkStart")
    .filter(mark => owningParagraph(mark) === paragraph && !inDeletedContent(mark))
    .flatMap(mark => {
      const name = mark.getAttributeNS(W, "name") ?? ""
      return name.startsWith("_dh_") ? [name.slice(4)] : []
    })
}
function acceptedParagraphs(doc: Document) {
  return elements(doc, W, "p").map(node => ({ node, text: acceptedText(node) })).filter(row => row.text.length > 0)
    .map((row, index) => ({ ...row, index, pids: paragraphPIDs(row.node) }))
}
