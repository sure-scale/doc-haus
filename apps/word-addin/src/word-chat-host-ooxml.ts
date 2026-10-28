import { DOMParser } from "@xmldom/xmldom"
import { validWordChatFormatting, type WordChatAuthoredAmendment, type WordChatParagraphFormat, type WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { reviewedWordChatParagraphFormatting } from "../../../dochaus/lib/word-chat-formatting-ooxml"

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
  const currentProperties = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")
  const changes = currentProperties ? elements(currentProperties).filter(node => node.namespaceURI === W && node.localName === "pPrChange") : []
  if (changes.length > 1) throw new Error("Ambiguous paragraph formatting")
  const old = changes[0] && elements(changes[0]).find(node => node.namespaceURI === W && node.localName === "pPr")
  if (changes[0] && !old) throw new Error("Missing original paragraph properties")
  const properties = (version === "Original" && old ? old : currentProperties)?.cloneNode(true) as Element | undefined
  if (properties) {
    for (const change of elements(properties)) if (change.namespaceURI === W && change.localName === "pPrChange") properties.removeChild(change)
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
  const formats = (xml: string) => content(parse(xml), true, true).nodes.map(paragraph => {
    const pPr = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")
    return pPr ? elements(pPr).filter(node => node.namespaceURI === W && node.localName === "pPrChange").map(node => node.getAttributeNS(W, "author")) : []
  })
  return ["Current", "Original"].every(version => JSON.stringify(clauseView(authored, version as "Current" | "Original")) === JSON.stringify(clauseView(observed, version as "Current" | "Original"))) &&
    JSON.stringify(marks(authored)) === JSON.stringify(marks(observed)) && JSON.stringify(formats(authored)) === JSON.stringify(formats(observed))
}

export function sameWordChatParagraphContent(before: string, after: string): boolean {
  const view = (xml: string) => content(parse(xml), true, true).shape.map(node => canonical(node, true))
  return JSON.stringify(view(before)) === JSON.stringify(view(after))
}

function unchangedFormattingDependencies(before: string, after: string) {
  const captured = parse(before), delivered = parse(after)
  for (const [name, part] of delivered) {
    if (name === MAIN) continue
    if (["/word/settings.xml", "/word/webSettings.xml"].includes(name) || name.startsWith("/docProps/") || /(?:header|footer|webextension)/i.test(name)) throw new Error("Formatting exceeds the paragraph scope")
    const original = captured.get(name)
    if (!original || original.getAttributeNS(PKG, "contentType") !== part.getAttributeNS(PKG, "contentType")) throw new Error("Formatting dependency changed")
    const data = elements(part)[0], prior = elements(original)[0]
    if (!data || !prior) throw new Error("Missing formatting dependency")
    if (name.endsWith(".rels")) {
      const relations = elements(elements(data)[0]!)
      if (name === "/_rels/.rels") {
        if (relations.length !== 1 || relations[0]!.getAttribute("Target") !== "word/document.xml" || relations[0]!.getAttribute("Type") !== "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" || relations[0]!.getAttribute("TargetMode") === "External") throw new Error("Invalid scoped root relationship")
      } else {
        const sourceRelations = elements(elements(prior)[0]!)
        if (relations.some(rel => !sourceRelations.some(old => JSON.stringify(canonical(rel)) === JSON.stringify(canonical(old))))) throw new Error("Formatting relationship changed")
      }
    } else if (data.localName === "binaryData" ? data.textContent?.replace(/\s/g, "") !== prior.textContent?.replace(/\s/g, "")
      : JSON.stringify(elements(data).map(node => canonical(node))) !== JSON.stringify(elements(prior).map(node => canonical(node)))) throw new Error("Formatting dependency changed")
  }
}

/** Compare the complete target and every dependency actually imported. Word's
 * selected OOXML includes changing write IDs and unrelated document settings;
 * those do not authorize an edit or form part of the delivered paragraph. */
const diagnosticName = (value: string) => value.replace(/[^a-zA-Z0-9_./:-]/g, "?").slice(0, 120)
function canonicalDifference(before: unknown, after: unknown, path: string): string | undefined {
  if (JSON.stringify(before) === JSON.stringify(after)) return
  if (before === undefined || after === undefined) return `${path}: XML node ${before === undefined ? "added" : "missing"}`
  if (!Array.isArray(before) || !Array.isArray(after)) return `${path}: XML shape changed`
  if (before[0] === "text" || after[0] === "text") return `${path}: text changed (text omitted)`
  // canonical() represents each element as namespace/name/attributes/children.
  if (before.length === 4 && after.length === 4 && typeof before[0] === "string" && typeof after[0] === "string") {
    if (before[0] !== after[0] || before[1] !== after[1]) return `${path}: element name or namespace changed`
    const element = `${path}/${diagnosticName(before[1] as string)}`
    const oldAttributes = before[2] as string[][], newAttributes = after[2] as string[][]
    const key = (attr: string[]) => JSON.stringify(attr.slice(0, 2))
    for (const attr of oldAttributes) {
      const current = newAttributes.find(candidate => key(candidate) === key(attr))
      if (!current) return `${element}/@${diagnosticName(attr[1]!)}: attribute missing`
      if (attr[2] !== current[2]) return `${element}/@${diagnosticName(attr[1]!)}: attribute changed (value omitted)`
    }
    const added = newAttributes.find(attr => !oldAttributes.some(candidate => key(candidate) === key(attr)))
    if (added) return `${element}/@${diagnosticName(added[1]!)}: attribute added`
    return canonicalDifference(before[3], after[3], element)
  }
  for (let index = 0; index < Math.max(before.length, after.length); index++) {
    const difference = canonicalDifference(before[index], after[index], `${path}[${index + 1}]`)
    if (difference) return difference
  }
  return `${path}: XML metadata changed`
}
export function wordChatOoxmlDifference(before: string, current: string, delivered: string): string | undefined {
  const names = [...parse(delivered).keys()].sort()
  if (names.some(name => ["/word/settings.xml", "/word/webSettings.xml"].includes(name) || name.startsWith("/docProps/") || /(?:header|footer|webextension)/i.test(name))) throw new Error("Import exceeds the clause scope")
  const oldParts = parse(before), newParts = parse(current)
  for (const name of names) {
    const oldPart = oldParts.get(name), newPart = newParts.get(name), path = diagnosticName(name)
    if (!oldPart || !newPart) return `${path}: dependency missing from ${!oldPart ? "reviewed" : "current"} Word capture`
    if (name === MAIN) {
      if (oldPart.getAttributeNS(PKG, "contentType") !== newPart.getAttributeNS(PKG, "contentType")) return `${path}: contentType changed`
      const difference = canonicalDifference(content(oldParts, true, true).shape.map(node => canonical(node, true)), content(newParts, true, true).shape.map(node => canonical(node, true)), path)
      if (difference) return difference
    } else {
      const difference = canonicalDifference(canonical(oldPart), canonical(newPart), path)
      if (difference) return difference
    }
  }
}
export function sameWordChatOoxml(before: string, current: string, delivered: string): boolean {
  return wordChatOoxmlDifference(before, current, delivered) === undefined
}

/** Verify both views and the actual authored revision spans. The server cannot
 * expand a selected subspan merely by returning matching final paragraph text. */
export function reviewedWordChatArtifact(artifact: WordChatAuthoredAmendment, proposalId: string, target: WordChatTarget, expected: string[], captured?: string, formatting?: WordChatParagraphFormat[]) {
  if (!artifact || artifact.proposalId !== proposalId || artifact.author !== "doc.haus" || !Array.isArray(artifact.paragraphs) || JSON.stringify(artifact.paragraphs) !== JSON.stringify(expected) || !Array.isArray(artifact.revisions)) throw new Error("The authored amendment differs from the reviewed wording")
  const paragraphs = content(parse(artifact.ooxml), false, false).nodes
  if (formatting !== undefined) {
    if (!captured || !validWordChatFormatting(target, formatting) || JSON.stringify(artifact.formatting) !== JSON.stringify(formatting)) throw new Error("The formatting differs from the reviewed proposal")
    const before = content(parse(captured), true, true).nodes
    if (JSON.stringify(before.map(node => text(node))) !== JSON.stringify(expected)) throw new Error("The formatted passage changed")
    unchangedFormattingDependencies(captured, artifact.ooxml)
    const revisions = reviewedWordChatParagraphFormatting(before, paragraphs, formatting)
    if (JSON.stringify(revisions.flat()) !== JSON.stringify(artifact.revisions)) throw new Error("The formatting revision metadata differs from its Word markup")
    return revisions
  }
  if (artifact.formatting !== undefined) throw new Error("Unreviewed paragraph formatting")
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
