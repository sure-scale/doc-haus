import type { WordChatInlineFormat } from "./word-chat-amendment-contract"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XML = "http://www.w3.org/XML/1998/namespace"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const W16DU = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
const elements = (node: Node) => Array.from(node.childNodes).filter((item): item is Element => item.nodeType === 1)
const named = (node: Element, name: string) => elements(node).filter(item => item.namespaceURI === W && item.localName === name)
function canonical(node: Node): unknown {
  if (node.nodeType === 1) {
    const e = node as Element
    return [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.namespaceURI !== XMLNS).map(a => [a.namespaceURI ?? "", a.localName, a.value]).sort((a,b) => JSON.stringify(a).localeCompare(JSON.stringify(b))), Array.from(e.childNodes).map(canonical).filter(x => x !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) {
    const parent = node.parentNode as Element | null
    if (!node.nodeValue?.trim() && !(parent?.namespaceURI === W && ["t","delText","instrText"].includes(parent.localName))) return undefined
    return ["text",node.nodeValue]
  }
  if (node.nodeType === 7 || node.nodeType === 8) return [node.nodeType,node.nodeName,node.nodeValue]
  return undefined
}
const same = (a: Node | undefined, b: Node | undefined) => JSON.stringify(a ? canonical(a) : null) === JSON.stringify(b ? canonical(b) : null)
function emptyProperties(value: Element | undefined) {
  return value && (Array.from(value.attributes).some(a => a.namespaceURI !== XMLNS) || elements(value).length || Array.from(value.childNodes).some(n => n.nodeType !== 1 && (n.nodeType !== 3 || n.nodeValue?.trim()))) ? value : undefined
}
function direct(properties: Element | undefined, field: keyof WordChatInlineFormat): boolean | undefined {
  if (!properties) return undefined
  const name = field === "bold" ? "b" : field === "italic" ? "i" : "u", found = named(properties, name)
  if (found.length > 1) throw Error("inline_duplicate_property")
  if (!found.length) return undefined
  const node = found[0]!
  if (node.childNodes.length || Array.from(node.attributes).some(a => a.namespaceURI !== XMLNS && (a.namespaceURI !== W || a.localName !== "val"))) throw Error("inline_unsupported_property")
  const value = node.getAttributeNS(W,"val")
  if (field === "underline") {
    if (value === "single") return true
    if (value === "none") return false
  } else {
    if (!value || ["1","true","on"].includes(value)) return true
    if (["0","false","off"].includes(value)) return false
  }
  throw Error("inline_unsupported_property")
}
function remainder(properties: Element | undefined, op: WordChatInlineFormat, removeChange: boolean) {
  if (!properties) return undefined
  const copy = properties.cloneNode(true) as Element
  for (const [key,name] of [["bold","b"],["italic","i"],["underline","u"]] as const) if (op[key] !== undefined) for (const node of named(copy,name)) copy.removeChild(node)
  if (removeChange) for (const node of named(copy,"rPrChange")) copy.removeChild(node)
  return emptyProperties(copy)
}
type Token = { type:"text"; value:string; run: Element; properties?: Element; position:number } | { type:"opaque"; value:string; node:Node; position:number }
function visible(node: Element): string {
  if (node.namespaceURI === W) {
    if (["del","pPr","rPr"].includes(node.localName)) return ""
    if (node.localName === "t") return node.textContent ?? ""
    if (node.localName === "tab") return "\t"
    if (["br","cr"].includes(node.localName)) return "\v"
  }
  return elements(node).map(visible).join("")
}
function project(paragraph: Element) {
  if (paragraph.namespaceURI !== W || paragraph.localName !== "p") throw Error("inline_expected_paragraph")
  const tokens: Token[] = []; let position = 0
  for (const node of Array.from(paragraph.childNodes)) {
    if (node.nodeType !== 1) {
      if (node.nodeType === 3 && !node.nodeValue?.trim()) continue
      tokens.push({type:"opaque",value:"",node,position}); continue
    }
    const element = node as Element
    if (element.namespaceURI === W && element.localName === "r") {
      const props = named(element,"rPr")
      if (props.length > 1 || Array.from(element.childNodes).some(child => child.nodeType !== 1 && (child.nodeType !== 3 || Boolean(child.nodeValue?.trim()))) || elements(element).some(child => child.namespaceURI !== W || !["rPr","t"].includes(child.localName))) throw Error("inline_unsupported_run")
      for (const t of named(element,"t")) {
        if (Array.from(t.childNodes).some(child => ![3,4].includes(child.nodeType)) || Array.from(t.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === XML && a.localName === "space" && ["preserve","default"].includes(a.value)))) throw Error("inline_unsupported_text")
        const text = t.textContent ?? ""
        if (!text.length) throw Error("inline_empty_run")
        for (let i=0;i<text.length;i++) tokens.push({type:"text",value:text[i]!,run:element,properties:props[0],position:position++})
      }
      if (!named(element,"t").length) throw Error("inline_unsupported_run")
    } else {
      const text = visible(element); tokens.push({type:"opaque",value:text,node:element,position}); position += text.length
    }
  }
  return {tokens,text:tokens.map(token => token.value).join("")}
}
function runAttributes(run: Element) {
  return JSON.stringify(Array.from(run.attributes).filter(a => a.namespaceURI !== XMLNS).map(a => [a.namespaceURI ?? "",a.localName,a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))))
}
function boundary(text:string, at:number) {
  return at <= 0 || at >= text.length || !(text.charCodeAt(at-1) >= 0xd800 && text.charCodeAt(at-1) <= 0xdbff && text.charCodeAt(at) >= 0xdc00 && text.charCodeAt(at) <= 0xdfff)
}
function wellFormed(text:string) {
  for (let i=0;i<text.length;i++) {
    const code=text.charCodeAt(i)
    if (code>=0xd800 && code<=0xdbff) { const next=text.charCodeAt(++i); if (!(next>=0xdc00 && next<=0xdfff)) return false }
    else if (code>=0xdc00 && code<=0xdfff) return false
  }
  return true
}

/** Compare content and properties across run splits without discarding prior
 * revision carriers or annotations. Missing and empty rPr are equivalent. */
export function sameWordChatInlineFormattingContent(before: Element, after: Element): boolean {
  try {
    if (runAttributes(before) !== runAttributes(after)) return false
    const a=project(before), b=project(after)
    return a.text === b.text && a.tokens.length === b.tokens.length && a.tokens.every((old,i) => {
      const next=b.tokens[i]!
      if (old.type !== next.type || old.value !== next.value || old.position !== next.position) return false
      if (old.type === "opaque" && next.type === "opaque") return same(old.node,next.node)
      return old.type === "text" && next.type === "text" && runAttributes(old.run) === runAttributes(next.run) && same(emptyProperties(old.properties),emptyProperties(next.properties))
    })
  } catch { return false }
}

/** Validate the exact source span and reject selected opaque text or prior
 * formatting revisions before a native operation can be dispatched. */
export function assertWordChatInlineFormattingTarget(before: Element, span: {startOffset:number;endOffset:number;text:string}, op: WordChatInlineFormat) {
  if (!op || typeof op !== "object" || Array.isArray(op) || !Object.keys(op).length || Object.keys(op).some(k => !["bold","italic","underline"].includes(k) || typeof (op as Record<string,unknown>)[k] !== "boolean")) throw Error("inline_invalid_operation")
  const projection = project(before), start = span?.startOffset, end = span?.endOffset
  if (!span || typeof span.text !== "string" || !wellFormed(span.text) || !Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end <= start || end > projection.text.length || !boundary(projection.text,start) || !boundary(projection.text,end) || projection.text.slice(start,end) !== span.text) throw Error("inline_invalid_span")
  for (const token of projection.tokens) {
    if (token.type === "opaque") {
      if (token.value.length && token.position < end && token.position + token.value.length > start) throw Error("inline_protected_span")
    } else if (token.position >= start && token.position < end) {
      if (token.properties && named(token.properties,"rPrChange").length) throw Error("inline_existing_format_revision_in_span")
      for (const key of ["bold","italic","underline"] as const) if (op[key] !== undefined) direct(token.properties,key)
    }
  }
}

/** Verify a physical tracked inline format change. This compares direct run
 * properties; inherited effective formatting requires an independent check. */
export function verifyWordChatInlineFormatting(before: Element, after: Element, span: {startOffset:number;endOffset:number;text:string}, op: WordChatInlineFormat, expectedAuthor?: string, effectiveProof?: WordChatInlineFormat) {
  if (!op || typeof op !== "object" || Array.isArray(op) || !Object.keys(op).length || Object.keys(op).some(k => !["bold","italic","underline"].includes(k) || typeof (op as Record<string,unknown>)[k] !== "boolean")) throw Error("inline_invalid_operation")
  assertWordChatInlineFormattingTarget(before, span, op)
  const a=project(before), b=project(after)
  const start = span?.startOffset, end = span?.endOffset
  if (runAttributes(before) !== runAttributes(after)) throw Error("inline_paragraph_attributes_changed")
  if (!span || typeof span.text !== "string" || !wellFormed(span.text) || !Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end <= start || end > a.text.length || !boundary(a.text,start) || !boundary(a.text,end) || a.text.slice(start,end) !== span.text) throw Error("inline_invalid_span")
  if (a.text !== b.text || a.tokens.length !== b.tokens.length) throw Error("inline_text_or_structure_changed")
  const oldIds = new Set(["ins","del","rPrChange","pPrChange"].flatMap(name => Array.from(before.getElementsByTagNameNS(W,name)).map(e => e.getAttributeNS(W,"id"))))
  const carriers = new Map<Element,{id:string;author:string;date:string;start:number;end:number}>(), ids = new Set<string>()
  for (let i=0;i<a.tokens.length;i++) {
    const old=a.tokens[i]!, next=b.tokens[i]!
    if (old.type !== next.type || old.value !== next.value || old.position !== next.position) throw Error("inline_text_or_structure_changed")
    if (old.type === "opaque" && next.type === "opaque") {
      if (!same(old.node,next.node)) throw Error("inline_protected_content_changed")
      if (old.value.length && old.position < end && old.position + old.value.length > start) throw Error("inline_protected_span")
      continue
    }
    if (old.type !== "text" || next.type !== "text") throw Error("inline_text_or_structure_changed")
    if (runAttributes(old.run) !== runAttributes(next.run)) throw Error("inline_run_attributes_changed")
    const inside = old.position >= start && old.position < end
    if (!inside) { if (!same(emptyProperties(old.properties),emptyProperties(next.properties))) throw Error("inline_unselected_format_changed"); continue }
    if (old.properties && named(old.properties,"rPrChange").length) throw Error("inline_existing_format_revision_in_span")
    const changes = next.properties ? named(next.properties,"rPrChange") : []
    if (changes.length > 1) throw Error("inline_ambiguous_revision")
    for (const key of ["bold","italic","underline"] as const) if (op[key] !== undefined) { const value = direct(next.properties,key); if (value !== op[key] && !(value === undefined && effectiveProof?.[key] === op[key])) throw Error("inline_requested_property_not_confirmed") }
    if (!same(remainder(old.properties,op,false),remainder(next.properties,op,true))) throw Error("inline_unrequested_property_changed")
    const changed = (["bold","italic","underline"] as const).some(key => op[key] !== undefined && direct(old.properties,key) !== direct(next.properties,key))
    const change = changes[0]
    if (!changed && !change) continue
    if (!changed && !(effectiveProof && (["bold","italic","underline"] as const).some(key => op[key] !== undefined && direct(next.properties,key) === undefined && effectiveProof[key] === op[key]))) throw Error("inline_unnecessary_revision")
    if (!change || !next.properties) throw Error("inline_missing_revision")
    const oldProperties = elements(change), id=change.getAttributeNS(W,"id"), author=change.getAttributeNS(W,"author"), date=change.getAttributeNS(W,"date"), utc=change.getAttributeNS(W16DU,"dateUtc")
    if (!/^\d+$/.test(id ?? "") || !author || expectedAuthor !== undefined && author !== expectedAuthor || !Number.isFinite(Date.parse(date ?? "")) || utc && (!Number.isFinite(Date.parse(utc)) || Date.parse(utc) !== Date.parse(date!)) ||
      Array.from(change.attributes).some(attr => attr.namespaceURI !== XMLNS && !(attr.namespaceURI === W && ["id","author","date"].includes(attr.localName)) && !(attr.namespaceURI === W16DU && attr.localName === "dateUtc")) ||
      Array.from(change.childNodes).some(node => node.nodeType !== 1 && (node.nodeType !== 3 || Boolean(node.nodeValue?.trim()))) ||
      oldProperties.length !== 1 || oldProperties[0]!.namespaceURI !== W || oldProperties[0]!.localName !== "rPr" || !same(emptyProperties(old.properties),emptyProperties(oldProperties[0]))) throw Error("inline_revision_old_properties_mismatch")
    const existing=carriers.get(change)
    if (!existing) { if (ids.has(id!) || oldIds.has(id!)) throw Error("inline_revision_id_collision"); ids.add(id!); carriers.set(change,{id:id!,author:author!,date:date!,start:old.position,end:old.position+1}) }
    else existing.end=old.position+1
  }
  return { text:a.text, revisions:[...carriers.values()] }
}
