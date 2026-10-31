import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { assertNativeWordChatInlineDependencies, projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { normalizeNativeWordChatCommentIds } from "./word-chat-native-comments"

type Projection = ReturnType<typeof projectNativeWordChatBody>
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  W14 = "http://schemas.microsoft.com/office/word/2010/wordml",
  W16 = "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
  PKG = "http://schemas.microsoft.com/office/2006/xmlPackage",
  XMLNS = "http://www.w3.org/2000/xmlns/",
  XML = "http://www.w3.org/XML/1998/namespace"
const elements = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const named = (n: Node, name: string) => elements(n).filter((e) => e.namespaceURI === W && e.localName === name)
const fail = (): never => {
  throw Error("Native list insertion could not be verified")
}
const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
function parse(xml: string) {
  if (xml.length > 16_000_000 || /<!DOCTYPE/i.test(xml)) return fail()
  const reject = () => {
    throw Error("Invalid Word XML")
  }
  const doc = new DOMParser({ errorHandler: { warning: reject, error: reject, fatalError: reject } }).parseFromString(
    xml,
    "application/xml",
  )
  if (doc.doctype || doc.documentElement.namespaceURI !== PKG || doc.documentElement.localName !== "package")
    return fail()
  const parts = elements(doc.documentElement).filter(
    (e) => e.namespaceURI === PKG && e.localName === "part" && e.getAttributeNS(PKG, "name") === "/word/document.xml",
  )
  if (parts.length !== 1) return fail()
  const body = Array.from(parts[0]!.getElementsByTagNameNS(W, "body"))
  if (body.length !== 1) return fail()
  return { doc, body: body[0]! }
}
function paragraph(xml: string) {
  const reject = () => {
    throw Error("Invalid paragraph XML")
  }
  const doc = new DOMParser({ errorHandler: { warning: reject, error: reject, fatalError: reject } }).parseFromString(
    xml,
    "application/xml",
  )
  if (doc.doctype || doc.documentElement.namespaceURI !== W || doc.documentElement.localName !== "p") return fail()
  return doc.documentElement
}
function canon(n: Node): unknown {
  if (n.nodeType === 1) {
    const e = n as Element
    for (const a of Array.from(e.attributes))
      if (a.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(a.localName) && !/^[0-9a-fA-F]{8}$/.test(a.value)) return fail()
    return [
      e.namespaceURI,
      e.localName,
      Array.from(e.attributes)
        .filter(
          (a) => a.namespaceURI !== XMLNS && !(a.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(a.localName)),
        )
        .map((a) => [a.namespaceURI, a.localName, a.value])
        .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
      Array.from(e.childNodes)
        .map(canon)
        .filter((x) => x !== undefined),
    ]
  }
  if (n.nodeType === 3 || n.nodeType === 4)
    return n.nodeValue?.trim() || ["t", "delText"].includes((n.parentNode as Element)?.localName) ? ["text", n.nodeValue] : undefined
  if (n.nodeType === 7 || n.nodeType === 8) return [n.nodeType, n.nodeName, n.nodeValue]
  return undefined
}
function only(n: Node, names: string[]) {
  return elements(n).every((e) => e.namespaceURI === W && names.includes(e.localName))
}
function validMeta(e: Element) {
  return Array.from(e.attributes).every(
    (a) =>
      a.namespaceURI === XMLNS ||
      (a.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)) ||
      (a.namespaceURI === W14 && ["paraId", "textId"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)),
  )
}
/** A plain top-level last list item with explicit direct numbering.
 * Caller must independently verify native list membership and last-item identity. */
export function assertNativeWordChatListInsertionTarget(projection: Projection, target: WordChatTarget, position: "end" | "before_final" = "end"): number {
  if (position !== "end" && position !== "before_final") return fail()
  if (
    !target ||
    !target.id ||
    target.paragraphs?.length !== 1 ||
    target.startOffset !== 0 ||
    target.endOffset !== target.text.length ||
    target.paragraphs[0]!.text !== target.text ||
    !target.text.length
  )
    return fail()
  const index = projection.paragraphs.findIndex((p) => p.id === target.paragraphs[0]!.id),
    item = projection.paragraphs[index]
  if (!item || item.text !== target.text || item.nesting !== 0 || !/^body\/p\[\d+\]$/.test(item.path)) return fail()
  const p = paragraph(item.xml), pp = named(p, "pPr"), runs = named(p, "r")
  if (!validMeta(p) || elements(p).length !== 2 || elements(p)[0] !== pp[0] || elements(p)[1] !== runs[0] || pp.length !== 1 || runs.length !== 1 || elements(runs[0]!).length !== 1 || named(runs[0]!, "t").length !== 1 || runs[0]!.textContent !== target.text) return fail()
  const clean = (e: Element): boolean => Array.from(e.childNodes).every(n => n.nodeType === 1 ? clean(n as Element) : n.nodeType === 3 && (e.namespaceURI === W && e.localName === "t" || !n.nodeValue?.trim()))
  const t = named(runs[0]!, "t")[0]!
  if (!clean(p) || Array.from(runs[0]!.attributes).some(a => a.namespaceURI !== XMLNS) || elements(t).length || Array.from(t.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === XML && a.localName === "space" && a.value === "preserve"))) return fail()
  const num = named(pp[0]!, "numPr"), noAttrs = (e: Element) => Array.from(e.attributes).every(a => a.namespaceURI === XMLNS)
  if (!noAttrs(pp[0]!) || elements(pp[0]!).length !== 1 || num.length !== 1 || !noAttrs(num[0]!)) return fail()
  const level = named(num[0]!, "ilvl"), id = named(num[0]!, "numId")
  const valOnly = (e: Element) => Array.from(e.attributes).every(a => a.namespaceURI === XMLNS || a.namespaceURI === W && a.localName === "val") && !elements(e).length
  if (elements(num[0]!).length !== 2 || level.length !== 1 || id.length !== 1 || !valOnly(level[0]!) || !valOnly(id[0]!) || !/^[0-8]$/.test(level[0]!.getAttributeNS(W,"val") || "") || !/^[0-9]+$/.test(id[0]!.getAttributeNS(W,"val") || "") || Number(id[0]!.getAttributeNS(W,"val")) <= 0) return fail()
  // The paragraph can be locally plain yet lie inside a whole-story range or
  // complex field. Track the structural graph up to and across both edges.
  const {body}=parse(projection.ooxml), path=/^body\/p\[(\d+)\]$/.exec(item.path)
  if (!path) return fail()
  const physical=elements(body)[Number(path[1])]
  if (!physical || physical.namespaceURI!==W || physical.localName!=="p") return fail()
  const active=new Set<string>(); let fields=0
  const starts:Record<string,string>={commentRangeStart:"comment",bookmarkStart:"bookmark",permStart:"perm",moveFromRangeStart:"moveFrom",moveToRangeStart:"moveTo",customXmlInsRangeStart:"customIns",customXmlDelRangeStart:"customDel",customXmlMoveFromRangeStart:"customMoveFrom",customXmlMoveToRangeStart:"customMoveTo"}
  const ends:Record<string,string>={commentRangeEnd:"comment",bookmarkEnd:"bookmark",permEnd:"perm",moveFromRangeEnd:"moveFrom",moveToRangeEnd:"moveTo",customXmlInsRangeEnd:"customIns",customXmlDelRangeEnd:"customDel",customXmlMoveFromRangeEnd:"customMoveFrom",customXmlMoveToRangeEnd:"customMoveTo"}
  const walk=(e:Element) => {
    if (e===physical && (active.size || fields)) return fail()
    if (e.namespaceURI===W) {
      const family=starts[e.localName] ?? ends[e.localName]
      if (family) {
        const id=e.getAttributeNS(W,"id"), key=`${family}:${id}`
        if (!id || starts[e.localName] && active.has(key) || ends[e.localName] && !active.has(key)) return fail()
        if (starts[e.localName]) active.add(key); else active.delete(key)
      }
      if (e.localName==="fldChar") {
        const kind=e.getAttributeNS(W,"fldCharType")
        if (kind==="begin") fields++
        else if (kind==="end") { if (!fields) return fail(); fields-- }
      }
    }
    for (const child of elements(e)) walk(child)
    if (e===physical && (active.size || fields)) return fail()
  }
  walk(body); if (active.size || fields) return fail()
  return index
}
/** Verifies main-story structure/carriers, annotation graph and captured
 * style/numbering/font/theme dependencies. Word people/settings/relationship
 * serialization and complete compressed-package identity remain a caller audit. */
export function verifyNativeWordChatListInsertion(
  before: Projection,
  after: Projection,
  target: WordChatTarget,
  text: string,
  position: "end" | "before_final" = "end",
): string[] {
  const index = assertNativeWordChatListInsertionTarget(before, target, position)
  if (
    typeof text !== "string" ||
    !text.trim() ||
    text.length > 24_000 ||
    /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(text) ||
    /[\u0000-\u001f\u007f]/.test(text) ||
    after.paragraphs.length !== before.paragraphs.length + 1
  )
    return fail()
  const beforePosition = position === "before_final"
  const left = after.paragraphs[index], right = after.paragraphs[index + 1]
  const inserted = beforePosition ? left : right, anchor = beforePosition ? right : left
  if (!inserted || !anchor || anchor.text !== target.text || inserted.text !== text ||
    anchor.nesting !== 0 || inserted.nesting !== 0 || anchor.id === inserted.id ||
    (beforePosition ? anchor.id !== target.paragraphs[0]!.id : inserted.id !== target.paragraphs[0]!.id)) return fail()
  for (let i = 0; i < before.paragraphs.length; i++) {
    if (i === index && !beforePosition) continue
    const next = after.paragraphs[i + (beforePosition ? (i >= index ? 1 : 0) : (i > index ? 1 : 0))]
    if (!next || next.id !== before.paragraphs[i]!.id || next.text !== before.paragraphs[i]!.text || next.nesting !== before.paragraphs[i]!.nesting) return fail()
  }
  assertNativeWordChatInlineDependencies(before.ooxml, after.ooxml)
  const b = parse(before.ooxml),
    a = parse(after.ooxml)
  // Reuse the projection's qualified unbound terminal carrier rule. Word may
  // assign that synthetic non-native blank a new paraId on every export.
  const trimUnboundTerminal = (body: Element, projection: Projection) => {
    const children = elements(body),
      section = children.at(-1),
      tail = children.at(-2)
    if (
      section?.namespaceURI !== W ||
      section.localName !== "sectPr" ||
      tail?.namespaceURI !== W ||
      tail.localName !== "p"
    )
      return
    const path = `body/p[${children.length - 2}]`
    if (projection.paragraphs.some((p) => p.path === path)) return
    if (!validMeta(tail) || Array.from(tail.childNodes).some((n) => n.nodeType !== 3 || Boolean(n.nodeValue?.trim())))
      return fail()
    body.removeChild(tail)
  }
  trimUnboundTerminal(b.body, before)
  trimUnboundTerminal(a.body, after)
  const oldNodes = elements(b.body),
    nodes = elements(a.body),
    match = /^body\/p\[(\d+)\]$/.exec(before.paragraphs[index]!.path)
  if (!match) return fail()
  const pos = Number(match[1]),
    oldP = oldNodes[pos],
    leftP = nodes[pos],
    newP = nodes[pos + 1]
  if (
    !oldP ||
    !leftP ||
    !newP ||
    oldP.localName !== "p" ||
    leftP.localName !== "p" ||
    newP.localName !== "p" ||
    oldP.namespaceURI !== W ||
    leftP.namespaceURI !== W ||
    newP.namespaceURI !== W ||
    !validMeta(leftP) ||
    !validMeta(newP)
  )
    return fail()
  const oldParaId = oldP.getAttributeNS(W14, "paraId"),
    newParaId = newP.getAttributeNS(W14, "paraId")
  if (oldParaId !== newParaId || (oldParaId && leftP.getAttributeNS(W14, "paraId") === oldParaId)) return fail()
  const clean = (n: Element): boolean =>
    Array.from(n.childNodes).every((c) =>
      c.nodeType === 1
        ? clean(c as Element)
        : c.nodeType === 3 && ((n.namespaceURI === W && n.localName === "t") || !c.nodeValue?.trim()),
    )
  const pp = named(leftP, "pPr"),
    pr = pp.length === 1 ? named(pp[0]!, "rPr") : [],
    mark = pr.length === 1 ? named(pr[0]!, "ins") : []
  if (
    pp.length !== 1 ||
    !clean(pp[0]!) ||
    elements(pp[0]!).length !== 2 ||
    named(pp[0]!, "numPr").length !== 1 ||
    elements(pp[0]!)[0] !== named(pp[0]!, "numPr")[0] ||
    elements(pp[0]!)[1] !== pr[0] ||
    pr.length !== 1 ||
    elements(pr[0]!).length !== 1 ||
    mark.length !== 1 ||
    mark[0]!.childNodes.length ||
    Array.from(pp[0]!.attributes).some((a) => a.namespaceURI !== XMLNS) ||
    Array.from(pr[0]!.attributes).some((a) => a.namespaceURI !== XMLNS)
  )
    return fail()
  const contentP = beforePosition ? leftP : newP
  const content = named(contentP, "ins"), newProps = named(contentP, "pPr"), oldProps = named(oldP, "pPr")
  if (!clean(contentP) || elements(contentP).length !== 2 || newProps.length !== 1 || content.length !== 1 || elements(contentP)[0] !== newProps[0] || elements(contentP)[1] !== content[0] || oldProps.length !== 1 || (!beforePosition && !equal(canon(newProps[0]!), canon(oldProps[0]!))) || !only(content[0]!, ["r"]) || named(content[0]!, "r").length !== 1) return fail()
  let value = ""
  for (const run of named(content[0]!, "r")) {
    const ts = named(run, "t")
    if (
      !validMeta(run) ||
      elements(run).length !== 1 ||
      ts.length !== 1 ||
      elements(ts[0]!).length ||
      Array.from(ts[0]!.childNodes).some((n) => n.nodeType !== 3) ||
      Array.from(ts[0]!.attributes).some(
        (a) =>
          a.namespaceURI !== XMLNS && !(a.namespaceURI === XML && a.localName === "space" && a.value === "preserve"),
      )
    )
      return fail()
    value += ts[0]!.textContent ?? ""
  }
  if (value !== text) return fail()
  const id = (e: Element) => e.getAttributeNS(W, "id"),
    author = mark[0]!.getAttributeNS(W, "author"),
    date = mark[0]!.getAttributeNS(W, "date"),
    utc = mark[0]!.getAttributeNS(W16, "dateUtc")
  for (const node of [mark[0]!, content[0]!]) {
    const n = id(node)
    if (
      !n ||
      !/^\d+$/.test(n) ||
      !author?.trim() ||
      !date ||
      Number.isNaN(Date.parse(date)) ||
      node.getAttributeNS(W, "author") !== author ||
      node.getAttributeNS(W, "date") !== date ||
      node.getAttributeNS(W16, "dateUtc") !== utc ||
      (utc && utc !== date) ||
      Array.from(node.attributes).some(
        (a) =>
          a.namespaceURI !== XMLNS &&
          !(a.namespaceURI === W && ["id", "author", "date"].includes(a.localName)) &&
          !(a.namespaceURI === W16 && a.localName === "dateUtc"),
      )
    )
      return fail()
  }
  if (id(mark[0]!) === id(content[0]!)) return fail()
  pp[0]!.removeChild(pr[0]!)
  if (!equal(canon(pp[0]!), canon(oldProps[0]!))) return fail()
  const retained = beforePosition ? newP : leftP
  if (beforePosition) {
    a.body.removeChild(leftP)
  } else {
    a.body.removeChild(newP)
    if (oldParaId) {
      const attr = retained.getAttributeNodeNS(W14, "paraId")
      if (!attr) return fail()
      retained.setAttributeNS(W14, attr.name, oldParaId)
    }
  }
  // Word may regenerate textId on the retained anchor in either operation.
  // Rebind only that validated metadata; the full inverse still compares all
  // paragraph content, properties, revisions and other paragraph metadata.
  const oldTextId = oldP.getAttributeNodeNS(W14, "textId")
  const retainedTextId = retained.getAttributeNodeNS(W14, "textId")
  if (Boolean(oldTextId) !== Boolean(retainedTextId)) return fail()
  if (oldTextId && retainedTextId) {
    if (!/^[0-9a-fA-F]{8}$/.test(oldTextId.value) || !/^[0-9a-fA-F]{8}$/.test(retainedTextId.value)) return fail()
    retained.setAttributeNS(W14, retainedTextId.name, oldTextId.value)
  }
  const normalized = normalizeNativeWordChatCommentIds(before.ooxml, new XMLSerializer().serializeToString(a.doc)),
    derived = parse(normalized)
  const carrierNames = new Set(["ins", "del", "pPrChange", "rPrChange"])
  const collect = (body: Element) => {
    const out: Array<{ path: string; node: Element }> = []
    const visit = (e: Element, path: string) => {
      if (e.namespaceURI === W && carrierNames.has(e.localName)) out.push({ path, node: e })
      elements(e).forEach((c, i) => visit(c, `${path}/${c.localName}[${i}]`))
    }
    visit(body, "body")
    return out
  }
  const old = collect(b.body),
    now = collect(derived.body),
    oldIds = old.map((x) => id(x.node)),
    newIds = now.map((x) => id(x.node))
  if (
    old.length !== now.length ||
    oldIds.some((x) => !x || !/^\d+$/.test(x)) ||
    newIds.some((x) => !x || !/^\d+$/.test(x)) ||
    new Set(oldIds).size !== old.length ||
    new Set(newIds).size !== now.length ||
    newIds.includes(id(mark[0]!)) ||
    newIds.includes(id(content[0]!))
  )
    return fail()
  for (let i = 0; i < old.length; i++) {
    const x = old[i]!,
      y = now[i]!
    if (x.path !== y.path || x.node.namespaceURI !== y.node.namespaceURI || x.node.localName !== y.node.localName)
      return fail()
    const attr = y.node.getAttributeNodeNS(W, "id")!
    const name = attr.name
    y.node.removeAttributeNS(W, "id")
    y.node.setAttributeNS(W, name, id(x.node)!)
  }
  for (const body of [b.body, derived.body]) {
    const section = elements(body).at(-1)
    if (section?.namespaceURI === W && section.localName === "sectPr") section.removeAttributeNS(W, "rsidSect")
  }
  if (!equal(canon(b.body), canon(derived.body))) return fail()
  return [author!]
}
