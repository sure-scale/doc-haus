import type { WordChatParagraphFormat } from "./word-chat-amendment-contract"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const elements = (node: Node) => Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1)
const child = (node: Element | undefined, name: string) => node && elements(node).find(n => n.namespaceURI === W && n.localName === name)
function canonical(node: Node | undefined): unknown {
  if (!node) return null
  if (node.nodeType === 1) {
    const element = node as Element
    return [element.namespaceURI, element.localName, Array.from(element.attributes).filter(a => a.namespaceURI !== XMLNS)
      .map(a => [a.namespaceURI ?? "", a.localName, a.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    Array.from(node.childNodes).map(canonical).filter(n => n !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) {
    const parent = node.parentNode as Element | null
    if (!node.nodeValue?.trim() && !(parent?.namespaceURI === W && ["t", "delText", "instrText"].includes(parent.localName))) return undefined
    return ["text", node.nodeValue]
  }
  if (node.nodeType === 7 || node.nodeType === 8) return [node.nodeType, node.nodeName, node.nodeValue]
  return undefined
}
const same = (a: Node | undefined, b: Node | undefined) => JSON.stringify(canonical(a)) === JSON.stringify(canonical(b))
const properties = (paragraph: Element) => child(paragraph, "pPr")
function normalized(properties: Element | undefined) {
  return properties && (properties.attributes.length || elements(properties).length) ? properties : undefined
}
function withoutProperties(paragraph: Element) {
  const copy = paragraph.cloneNode(true) as Element
  const pPr = properties(copy)
  if (pPr) copy.removeChild(pPr)
  return copy
}
function currentText(node: Element): string {
  if (node.namespaceURI === W) {
    if (["pPr", "rPr", "del"].includes(node.localName)) return ""
    if (["t", "delText"].includes(node.localName)) return node.textContent ?? ""
    if (node.localName === "tab") return "\t"
    if (["br", "cr"].includes(node.localName)) return "\v"
  }
  return elements(node).map(currentText).join("")
}
function remaining(pPr: Element | undefined, op: WordChatParagraphFormat) {
  if (!pPr) return undefined
  const copy = pPr.cloneNode(true) as Element
  const style = child(copy, "pStyle"), spacing = child(copy, "spacing")
  if (op.style !== undefined && style) {
    style.removeAttributeNS(W, "val")
    if (!style.attributes.length && !elements(style).length) copy.removeChild(style)
  }
  if (spacing) {
    for (const [field, attribute] of [["spacingBefore", "before"], ["spacingAfter", "after"], ["lineSpacing", "line"]] as const)
      if (op[field] !== undefined) spacing.removeAttributeNS(W, attribute)
    if (op.lineSpacing !== undefined) spacing.removeAttributeNS(W, "lineRule")
    if (!spacing.attributes.length && !elements(spacing).length) copy.removeChild(spacing)
  }
  return normalized(copy)
}

/** Verify the SDK's untracked result before adding a single tracked property
 * change. Inline XML (including existing revisions/bookmarks) is exact. */
export function wordChatParagraphFormatChanged(before: Element, after: Element, op?: WordChatParagraphFormat): boolean {
  if (!same(withoutProperties(before), withoutProperties(after))) throw new Error("Formatting changes document wording or inline markup")
  const old = properties(before), next = properties(after)
  if (!op) {
    if (!same(normalized(old), normalized(next))) throw new Error("Formatting exceeds the reviewed paragraphs")
    return false
  }
  if (!same(remaining(old, op), remaining(next, op))) throw new Error("Formatting changes unapproved paragraph properties")
  if (op.style !== undefined && child(next, "pStyle")?.getAttributeNS(W, "val") !== op.style) throw new Error("Wrong paragraph style")
  const spacing = child(next, "spacing")
  for (const [field, attribute] of [["spacingBefore", "before"], ["spacingAfter", "after"], ["lineSpacing", "line"]] as const)
    if (op[field] !== undefined && spacing?.getAttributeNS(W, attribute) !== String(op[field]! * (field === "lineSpacing" ? 240 : 20))) throw new Error("Wrong paragraph spacing")
  if (op.lineSpacing !== undefined && spacing?.getAttributeNS(W, "lineRule") !== "auto") throw new Error("Wrong line spacing rule")
  return !same(normalized(old), normalized(next))
}

/** A format-only artifact records exactly the captured old pPr and only the
 * approved new leaves. Existing content/revision XML must remain untouched. */
export function reviewedWordChatParagraphFormatting(before: Element[], after: Element[], formatting: WordChatParagraphFormat[]) {
  if (before.length !== after.length) throw new Error("Formatting changes paragraph boundaries")
  const operations = new Map(formatting.map(op => [op.paragraph, op]))
  const revisions = before.map((original, index) => {
    const paragraph = after[index]!, copy = paragraph.cloneNode(true) as Element, pPr = properties(copy)
    const changes = pPr ? elements(pPr).filter(n => n.namespaceURI === W && n.localName === "pPrChange") : []
    if (changes.length > 1) throw new Error("Ambiguous formatting revision")
    const change = changes[0]
    if (change) pPr!.removeChild(change)
    const changed = wordChatParagraphFormatChanged(original, copy, operations.get(index))
    if (!changed) {
      if (change) throw new Error("Unchanged paragraph has a formatting revision")
      return []
    }
    const old = change && elements(change)
    if (!change || change.getAttributeNS(W, "author") !== "doc.haus" || !/^\d+$/.test(change.getAttributeNS(W, "id") ?? "") ||
      !Number.isFinite(Date.parse(change.getAttributeNS(W, "date") ?? "")) || old?.length !== 1 || old[0]!.namespaceURI !== W || old[0]!.localName !== "pPr" ||
      !same(normalized(properties(original)), normalized(old[0]))) throw new Error("Formatting revision does not preserve the captured properties")
    return [{ author: "doc.haus", type: "Formatted" as const, text: currentText(original) }]
  })
  if (!revisions.flat().length) throw new Error("Formatting makes no change")
  return revisions
}
