import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { openRevision, type DocxRevisionOptions } from "./docx-amendments"
import { assertLayeredRoundTrip } from "./redline-round-trip"
import { validWordChatFormatting, type WordChatParagraphFormat } from "./word-chat-amendment-contract"
import { wordChatParagraphFormatChanged } from "./word-chat-formatting-ooxml"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const MAIN = "word/document.xml"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const serializer = new XMLSerializer()
const children = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const paragraphs = (doc: Document) => Array.from(doc.getElementsByTagNameNS(W, "p"))
const properties = (node: Element) => children(node).find(n => n.namespaceURI === W && n.localName === "pPr")
function parse(bytes: Uint8Array) {
  const xml = strFromU8(bytes)
  if (/<!DOCTYPE/i.test(xml)) throw new Error("Unsupported Word document")
  const fail = () => { throw new Error("Invalid Word document XML") }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
}
function canonical(node: Node): unknown {
  if (node.nodeType === 1) {
    const element = node as Element
    return [element.namespaceURI, element.localName, Array.from(element.attributes).filter(a => a.namespaceURI !== XMLNS)
      .map(a => [a.namespaceURI ?? "", a.localName, a.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    Array.from(node.childNodes).map(canonical).filter(x => x !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) {
    const parent = node.parentNode as Element | null
    if (!node.nodeValue?.trim() && !(parent?.namespaceURI === W && ["t", "delText", "instrText"].includes(parent.localName))) return undefined
    return ["text", node.nodeValue]
  }
  if (node.nodeType === 7 || node.nodeType === 8) return [node.nodeType, node.nodeName, node.nodeValue]
  return undefined
}

/** Format one uniquely identified complete body paragraph, preserving the full package.
 * The exact resulting text is checked again after wording has been authored. */
export async function applyDocxParagraphFormatting(bytes: Uint8Array, input: {
  paragraphText: string
  format: Omit<WordChatParagraphFormat, "paragraph">
  author: string
  tracked: boolean
  allowNoop?: boolean
  options?: DocxRevisionOptions
}): Promise<Uint8Array> {
  const { paragraphText, format, author, tracked } = input
  const op = { ...format, paragraph: 0 }
  const target = { id: "paragraph", text: paragraphText, paragraphs: [{ id: "paragraph", text: paragraphText }], startOffset: 0,
    endOffset: paragraphText.length, capturedAt: "", saved: true }
  if (!author.trim() || !paragraphText || !validWordChatFormatting(target, [op]) || /[\r\n]/.test(paragraphText))
    throw new Error("Invalid paragraph formatting proposal")
  const source = unzipSync(bytes), original = source[MAIN]
  if (!original) throw new Error("Word document has no main part")
  const beforeDoc = parse(original), before = paragraphs(beforeDoc)
  const opened = await openRevision(bytes, input.options ?? {})
  let ordinal: number, anchor: string
  try {
    const candidates = opened.paragraphs.filter(p => p.part === MAIN && p.match.text === paragraphText)
    if (candidates.length !== 1) throw new Error("Formatting target is missing or ambiguous")
    const selected = candidates[0]!
    ordinal = before.indexOf(selected.node)
    // openRevision owns a separately parsed tree; map its main-part position.
    if (ordinal < 0) ordinal = Array.from(opened.documents[MAIN]!.getElementsByTagNameNS(W, "p")).indexOf(selected.node)
    if (ordinal < 0 || !before[ordinal]) throw new Error("Formatting target is not a complete body paragraph")
    anchor = selected.match.enclosingAnchor.id
    if (properties(before[ordinal]!)?.getElementsByTagNameNS(W, "pPrChange").length)
      throw new Error("Decide the paragraph's existing formatting change first")
    const session = opened.session
    const membership = (id: string) => { const value = session.getListMembership(id); return value ? { ...value, anchorId: undefined } : null }
    const numbered = JSON.stringify(membership(anchor))
    if (format.style !== undefined) {
      if (!session.listStyles().some(style => style.id === format.style && style.type === "paragraph")) throw new Error("Paragraph style is unavailable in this document")
      const changed = session.setParagraphStyle(anchor, format.style)
      if (!changed.success) throw new Error("Could not apply paragraph style")
      anchor = changed.modified[0]?.id ?? anchor
    }
    if (format.spacingBefore !== undefined || format.spacingAfter !== undefined || format.lineSpacing !== undefined) {
      const result = session.setParagraphFormat(anchor, {
        ...(format.spacingBefore !== undefined ? { spacingBefore: format.spacingBefore * 20 } : {}),
        ...(format.spacingAfter !== undefined ? { spacingAfter: format.spacingAfter * 20 } : {}),
        ...(format.lineSpacing !== undefined ? { lineSpacing: format.lineSpacing * 240, lineSpacingRule: "auto" as const } : {}),
      })
      if (!result.success) throw new Error("Could not apply paragraph spacing")
    }
    if (JSON.stringify(membership(anchor)) !== numbered) throw new Error("Formatting would change paragraph numbering")
    const main = unzipSync(session.save())[MAIN]
    if (!main) throw new Error("Formatting lost the document body")
    const afterDoc = parse(main), after = paragraphs(afterDoc)
    if (before.length !== after.length) throw new Error("Formatting changed paragraph boundaries")
    for (let i = 0; i < before.length; i++) wordChatParagraphFormatChanged(before[i]!, after[i]!, i === ordinal ? op : undefined)
    if (!wordChatParagraphFormatChanged(before[ordinal]!, after[ordinal]!, op)) {
      if (input.allowNoop) return bytes
      throw new Error("Paragraph already has the proposed formatting")
    }
    // Replacing the one approved paragraph must reproduce the entire main XML.
    const changedParagraph = after[ordinal]!, parent = changedParagraph.parentNode!
    parent.replaceChild(afterDoc.importNode(before[ordinal]!, true), changedParagraph)
    if (JSON.stringify(canonical(beforeDoc.documentElement)) !== JSON.stringify(canonical(afterDoc.documentElement)))
      throw new Error("Formatting changed unrelated document content")
    // Reparse to avoid relying on the replaced DOM node's position.
    const resultDoc = parse(main), resultParagraph = paragraphs(resultDoc)[ordinal]!
    if (tracked) {
      const pPr = properties(resultParagraph)
      if (!pPr) throw new Error("Missing paragraph properties")
      let maxId = 0
      for (const node of Array.from(resultDoc.getElementsByTagNameNS(W, "*"))) {
        const value = Number(node.getAttributeNS(W, "id"))
        if (Number.isSafeInteger(value)) maxId = Math.max(maxId, value)
      }
      const change = resultDoc.createElementNS(W, "w:pPrChange")
      change.setAttributeNS(W, "w:id", String(maxId + 1)); change.setAttributeNS(W, "w:author", author); change.setAttributeNS(W, "w:date", new Date().toISOString())
      change.appendChild(properties(before[ordinal]!) ? resultDoc.importNode(properties(before[ordinal]!)!, true) : resultDoc.createElementNS(W, "w:pPr"))
      pPr.appendChild(change)
    }
    const baked = zipSync({ ...source, [MAIN]: main })
    const result = zipSync({ ...source, [MAIN]: strToU8(serializer.serializeToString(resultDoc)) })
    if (tracked) {
      await assertLayeredRoundTrip(result, bytes, baked)
      const dx = await docxodus()
      const revisions = (data: Uint8Array) => {
        const probe = dx.openDocxSession(data, { persistAnchorIds: false })
        try { return probe.listRevisions() } finally { probe.close() }
      }
      const prior = new Set(revisions(bytes).map(row => row.id))
      const added = revisions(result).filter(row => !prior.has(row.id))
      if (added.length !== 1 || added[0]!.type !== "format" || added[0]!.author !== author)
        throw new Error("Formatting revision could not be verified")
      for (const accept of [false, true]) {
        const probe = dx.openDocxSession(result, { persistAnchorIds: false })
        try {
          const decision = accept ? probe.acceptRevision(added[0]!.id) : probe.rejectRevision(added[0]!.id)
          if (!decision.success) throw new Error("Formatting revision cannot be resolved")
          const resolved = paragraphs(parse(unzipSync(probe.save())[MAIN]!))
          const expected = accept ? paragraphs(parse(main)) : before
          if (resolved.length !== expected.length) throw new Error("Formatting resolution changed paragraph boundaries")
          resolved.forEach((paragraph, index) => wordChatParagraphFormatChanged(expected[index]!, paragraph))
        } finally { probe.close() }
      }
    }
    return result
  } finally { opened.session.close() }
}
