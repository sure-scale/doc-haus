import { DOMParser } from "@xmldom/xmldom"
import type { WordChatAuthoredAmendment, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const MAIN = "/word/document.xml"
const elements = (node: Node) => Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1)
function parse(xml: string) {
  if (typeof xml !== "string" || xml.length > 2_000_000 || /<!DOCTYPE/i.test(xml)) throw new Error("Invalid clause OOXML")
  const fail = () => { throw new Error("Invalid clause OOXML") }
  const document = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
  if (document.doctype || document.documentElement?.namespaceURI !== PKG || document.documentElement.localName !== "package") throw new Error("Expected a scoped Word package")
  const parts = new Map<string, Element>()
  for (const part of elements(document.documentElement)) {
    const name = part.getAttributeNS(PKG, "name")
    if (part.namespaceURI !== PKG || part.localName !== "part" || !name || parts.has(name) || parts.size >= 64) throw new Error("Invalid Word package part")
    parts.set(name, part)
  }
  if (!parts.has(MAIN)) throw new Error("Missing clause content")
  return parts
}
function content(parts: Map<string, Element>, table: boolean, section: boolean) {
  const data = elements(parts.get(MAIN)!).filter(node => node.namespaceURI === PKG && node.localName === "xmlData")
  if (data.length !== 1 || elements(data[0]!).length !== 1) throw new Error("Invalid clause document")
  const document = elements(data[0]!)[0]!
  if (document.namespaceURI !== W || document.localName !== "document") throw new Error("Invalid clause document")
  const bodies = elements(document).filter(node => node.namespaceURI === W && node.localName === "body")
  if (bodies.length !== 1) throw new Error("Invalid clause body")
  const children = elements(bodies[0]!)
  if (section && children.at(-1)?.namespaceURI === W && children.at(-1)?.localName === "sectPr") children.pop()
  if (children.length && children.every(node => node.namespaceURI === W && node.localName === "p")) return { nodes: children, shape: children }
  if (table && children.length === 1 && children[0]!.namespaceURI === W && children[0]!.localName === "tbl") {
    const wrapper = children[0]!, cells = Array.from(wrapper.getElementsByTagNameNS(W, "tc"))
    // A paragraph fragment may carry its existing cell wrapper. Never import
    // another cell or guess among additional empty/nonempty table paragraphs.
    if (cells.length !== 1 || wrapper.getElementsByTagNameNS(W, "tbl").length) throw new Error("Choose paragraphs within one ordinary table cell")
    const nodes = elements(cells[0]!).filter(node => node.namespaceURI === W && node.localName === "p")
    if (!nodes.length || nodes.length !== wrapper.getElementsByTagNameNS(W, "p").length) throw new Error("Unsupported table clause content")
    return { nodes, shape: children }
  }
  throw new Error("Choose ordinary paragraphs within one document passage")
}
function text(node: Node, version: "Current" | "Original" = "Current"): string {
  if (node.nodeType !== 1) return ""
  const element = node as Element
  if (element.namespaceURI === W) {
    if (["pPr", "rPr"].includes(element.localName)) return ""
    if (element.localName === (version === "Current" ? "del" : "ins")) return ""
    if (["t", "delText"].includes(element.localName)) return element.textContent || ""
    if (element.localName === "tab") return "\t"
    if (["br", "cr"].includes(element.localName)) return "\u000b"
  }
  return elements(element).map(child => text(child, version)).join("")
}
function canonical(node: Node, paragraph = false): unknown {
  if (node.nodeType === 1) {
    const element = node as Element
    const attributes = Array.from(element.attributes).filter(attr => attr.namespaceURI !== XMLNS && !(
      paragraph && (attr.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(attr.localName)
        || element.namespaceURI === W && element.localName === "p" && attr.namespaceURI === W14 && attr.localName === "paraId")
    )).map(attr => [attr.namespaceURI || "", attr.localName, attr.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)))
    return [element.namespaceURI || "", element.localName, attributes, Array.from(element.childNodes).map(child => canonical(child, paragraph)).filter(child => child !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) {
    const parent = node.parentNode as Element | null
    if (!node.nodeValue?.trim() && !(parent?.namespaceURI === W && ["t", "delText", "instrText"].includes(parent.localName))) return undefined
    return ["text", node.nodeValue]
  }
  if (node.nodeType === 7 || node.nodeType === 8) return [node.nodeType, node.nodeName, node.nodeValue]
  return undefined
}
export function capturedWordChatParagraphs(xml: string): string[] {
  return content(parse(xml), true, true).nodes.map(node => text(node))
}

// Compare accepted character formatting, paragraph properties and named bookmark
// anchors rather than run boundaries: Word may split runs during OOXML import.
function paragraphView(paragraph: Element, version: "Current" | "Original") {
  const runs: Array<{ text: string; format: string }> = []
  const bookmarks = new Map<string, { name: string; start: number; end?: number }>()
  let offset = 0
  const append = (value: string, format: string) => {
    if (!value) return
    const last = runs.at(-1)
    if (last?.format === format) last.text += value
    else runs.push({ text: value, format })
    offset += value.length
  }
  const visit = (node: Element, format = "null") => {
    if (node.namespaceURI === W) {
      if (node.localName === (version === "Current" ? "del" : "ins") || ["pPr", "rPr"].includes(node.localName)) return
      if (node.localName === "r") format = JSON.stringify(canonical(elements(node).find(child => child.namespaceURI === W && child.localName === "rPr") ?? nullNode) ?? null)
      if (["t", "delText", "tab", "br", "cr"].includes(node.localName)) { append(text(node, version), format); return }
      if (node.localName === "bookmarkStart") {
        const id = node.getAttributeNS(W, "id"), name = node.getAttributeNS(W, "name")
        if (!id || !name || bookmarks.has(id)) throw new Error("Invalid clause bookmark")
        bookmarks.set(id, { name, start: offset }); return
      }
      if (node.localName === "bookmarkEnd") {
        const mark = bookmarks.get(node.getAttributeNS(W, "id") || "")
        if (!mark || mark.end !== undefined) throw new Error("Invalid clause bookmark")
        mark.end = offset; return
      }
    }
    for (const child of elements(node)) visit(child, format)
  }
  visit(paragraph)
  const marks = [...bookmarks.values()]
  if (marks.some(mark => mark.end === undefined) || new Set(marks.map(mark => mark.name)).size !== marks.length) throw new Error("Incomplete clause bookmark")
  const properties = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")?.cloneNode(true) as Element | undefined
  if (properties) {
    const run = elements(properties).find(node => node.namespaceURI === W && node.localName === "rPr")
    if (run) {
      for (const mark of elements(run)) if (mark.namespaceURI === W && ["ins", "del"].includes(mark.localName)) run.removeChild(mark)
      if (!elements(run).length && !run.attributes.length) properties.removeChild(run)
    }
  }
  return { properties: canonical(properties && (elements(properties).length || properties.attributes.length) ? properties : nullNode) ?? null, runs,
    bookmarks: marks.sort((a, b) => a.name.localeCompare(b.name)) }
}
const nullNode = { nodeType: 0 } as Node
function paragraphMark(paragraph: Element) {
  const properties = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")
  const run = properties && elements(properties).find(node => node.namespaceURI === W && node.localName === "rPr")
  const marks = run ? elements(run).filter(node => node.namespaceURI === W && ["ins", "del"].includes(node.localName)) : []
  if (marks.length > 1) throw new Error("Ambiguous paragraph revision")
  return marks[0]
}
function clauseView(xml: string, version: "Current" | "Original") {
  return content(parse(xml), true, true).nodes.filter(node => {
    const mark = paragraphMark(node)
    if (mark?.localName !== (version === "Current" ? "del" : "ins")) return true
    if (text(node, version)) throw new Error("Only complete paragraph insertions and deletions are supported")
    return false
  }).map(node => paragraphView(node, version))
}
export function sameImportedWordChat(authored: string, observed: string): boolean {
  const marks = (xml: string) => content(parse(xml), true, true).nodes.map(paragraphMark).filter((mark): mark is Element => Boolean(mark))
    .map(mark => [mark.localName, mark.getAttributeNS(W, "author")])
  return ["Current", "Original"].every(version => JSON.stringify(clauseView(authored, version as "Current" | "Original")) === JSON.stringify(clauseView(observed, version as "Current" | "Original"))) &&
    JSON.stringify(marks(authored)) === JSON.stringify(marks(observed))
}

/** Compare the complete target and every dependency actually imported. Word's
 * selected OOXML includes changing write IDs and unrelated document settings;
 * those do not authorize an edit or form part of the delivered paragraph. */
export function sameWordChatOoxml(before: string, current: string, delivered: string): boolean {
  const names = [...parse(delivered).keys()].sort()
  if (names.some(name => ["/word/settings.xml", "/word/webSettings.xml"].includes(name) || name.startsWith("/docProps/") || /(?:header|footer|webextension)/i.test(name))) throw new Error("Import exceeds the clause scope")
  const signature = (xml: string) => {
    const parts = parse(xml)
    return JSON.stringify(names.map(name => {
      const part = parts.get(name)
      if (!part) throw new Error("A clause formatting dependency changed")
      return name === MAIN
        ? [name, part.getAttributeNS(PKG, "contentType"), content(parts, true, true).shape.map(node => canonical(node, true))]
        : [name, canonical(part)]
    }))
  }
  return signature(before) === signature(current)
}

/** Verify both views and the actual authored revision spans. The server cannot
 * expand a selected subspan merely by returning matching final paragraph text. */
export function reviewedWordChatArtifact(artifact: WordChatAuthoredAmendment, proposalId: string, target: WordChatTarget, expected: string[], captured?: string) {
  if (!artifact || artifact.proposalId !== proposalId || artifact.author !== "doc.haus" || !Array.isArray(artifact.paragraphs) || JSON.stringify(artifact.paragraphs) !== JSON.stringify(expected) || !Array.isArray(artifact.revisions)) throw new Error("The authored amendment differs from the reviewed wording")
  const paragraphs = content(parse(artifact.ooxml), false, false).nodes
  const current = clauseView(artifact.ooxml, "Current"), original = clauseView(artifact.ooxml, "Original")
  if (current.length !== expected.length || original.length !== target.paragraphs.length ||
    JSON.stringify(current.map(p => p.runs.map(r => r.text).join(""))) !== JSON.stringify(expected) ||
    JSON.stringify(original.map(p => p.runs.map(r => r.text).join(""))) !== JSON.stringify(target.paragraphs.map(p => p.text)))
    throw new Error("The authored amendment differs from the reviewed passage")
  if (captured && JSON.stringify(clauseView(captured, "Original")) !== JSON.stringify(clauseView(artifact.ooxml, "Original"))) throw new Error("The authored amendment changes original formatting or bookmark anchors")
  const revisions: Array<Array<{ author: string; type: "Added" | "Deleted"; text: string }>> = []
  let offset = 0
  const start = target.startOffset, end = target.paragraphs.slice(0, -1).reduce((at, p) => at + p.text.length + 1, 0) + target.endOffset
  for (const [index, paragraph] of paragraphs.entries()) {
    const changes: typeof revisions[number] = []
    const visit = (node: Element) => {
      if (node.namespaceURI === W && ["pPr", "rPr"].includes(node.localName)) return
      if (node.namespaceURI === W && ["ins", "del"].includes(node.localName)) {
        const added = node.localName === "ins", wording = text(node, added ? "Current" : "Original")
        if (!wording || node.getAttributeNS(W, "author") !== "doc.haus" || offset < start || offset + (added ? 0 : wording.length) > end || node.getElementsByTagNameNS(W, "ins").length || node.getElementsByTagNameNS(W, "del").length) throw new Error("The authored revision exceeds the approved selection")
        changes.push({ author: "doc.haus", type: added ? "Added" : "Deleted", text: wording })
        if (!added) offset += wording.length
        return
      }
      if (node.namespaceURI === W && ["t", "delText", "tab", "br", "cr"].includes(node.localName)) { offset += text(node).length; return }
      for (const child of elements(node)) visit(child)
    }
    visit(paragraph)
    const mark = paragraphMark(paragraph)
    if (mark) {
      const added = mark.localName === "ins"
      if (index === paragraphs.length - 1 || mark.getAttributeNS(W, "author") !== "doc.haus" || offset < start || offset + (added ? 0 : 1) > end)
        throw new Error("The paragraph revision exceeds the approved selection")
      changes.push({ author: "doc.haus", type: added ? "Added" : "Deleted", text: "\r" })
    }
    if (index < paragraphs.length - 1 && mark?.localName !== "ins") offset++
    revisions.push(changes)
  }
  const signature = (changes: typeof artifact.revisions) => JSON.stringify(changes.map(change => [change.author, change.type, change.text]))
  if (!revisions.flat().length || signature(revisions.flat()) !== signature(artifact.revisions)) throw new Error("The authored revision metadata differs from its Word markup")
  return revisions
}
