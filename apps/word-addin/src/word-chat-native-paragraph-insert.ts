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
  throw Error("Native paragraph insertion could not be verified")
}
const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
// Word's editing-session ids on paragraphs and runs carry no content; Word
// stamps rsidP/rsidRPr whenever paragraph or run properties change.
const session = ["rsidR", "rsidRDefault", "rsidP", "rsidRPr"]
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
    return [
      e.namespaceURI,
      e.localName,
      Array.from(e.attributes)
        .filter(
          (a) => a.namespaceURI !== XMLNS && !(a.namespaceURI === W && session.includes(a.localName)),
        )
        .map((a) => [a.namespaceURI, a.localName, a.value])
        .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
      Array.from(e.childNodes)
        .map(canon)
        .filter((x) => x !== undefined),
    ]
  }
  if (n.nodeType === 3 || n.nodeType === 4)
    return n.nodeValue?.trim() || (n.parentNode as Element)?.localName === "t" ? ["text", n.nodeValue] : undefined
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
      (a.namespaceURI === W && session.includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)) ||
      (a.namespaceURI === W14 && ["paraId", "textId"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)),
  )
}
// Content whose inverse would need more than this change's own carriers.
const refused = new Set([
  "ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
  "pPrChange", "rPrChange", "numPr", "sectPr", "framePr", "fldChar", "fldSimple", "instrText", "drawing", "pict",
  "object", "commentRangeStart", "commentRangeEnd", "commentReference", "bookmarkStart", "bookmarkEnd", "permStart",
  "permEnd", "hyperlink", "sdt", "customXml", "smartTag", "footnoteReference", "endnoteReference", "vanish",
  "webHidden", "specVanish",
])
/** One complete top-level paragraph of literal runs, with any paragraph and
 * run properties except direct numbering, frames, sections and revisions.
 * Fields, objects, links and annotation anchors are refused. */
export function assertNativeWordChatParagraphInsertionTarget(projection: Projection, target: WordChatTarget): number {
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
  const p = paragraph(item.xml),
    children = elements(p),
    props = named(p, "pPr"),
    runs = named(p, "r")
  if (
    !validMeta(p) ||
    props.length > 1 ||
    (props.length === 1 && children[0] !== props[0]) ||
    !runs.length ||
    children.length !== props.length + runs.length ||
    Array.from(p.getElementsByTagNameNS("*", "*")).some(
      (e) =>
        e.namespaceURI === W
          ? refused.has(e.localName)
          : (e.parentNode as Element | null)?.namespaceURI !== W || (e.parentNode as Element).localName !== "rPr",
    )
  )
    return fail()
  let value = ""
  for (const run of runs) {
    const kids = elements(run),
      ts = named(run, "t"),
      rp = named(run, "rPr")
    if (
      ts.length !== 1 ||
      rp.length > 1 ||
      kids.length !== ts.length + rp.length ||
      (rp.length === 1 && kids[0] !== rp[0]) ||
      elements(ts[0]!).length ||
      Array.from(ts[0]!.childNodes).some((n) => n.nodeType !== 3)
    )
      return fail()
    value += ts[0]!.textContent ?? ""
  }
  if (value !== target.text) return fail()
  return index
}
/** Verifies main-story structure/carriers, annotation graph and captured
 * style/numbering/font/theme dependencies by undoing only this insertion's
 * carriers. `restyled` requires one tracked style change on the new paragraph;
 * the host proves its live style. Word people/settings/relationship
 * serialization and complete compressed-package identity remain a caller audit. */
export function verifyNativeWordChatParagraphInsertion(
  before: Projection,
  after: Projection,
  target: WordChatTarget,
  text: string,
  restyled = false,
): string[] {
  const index = assertNativeWordChatParagraphInsertionTarget(before, target)
  if (
    typeof text !== "string" ||
    !text.trim() ||
    text.length > 24_000 ||
    /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(text) ||
    /[\u0000-\u001f\u007f]/.test(text) ||
    after.paragraphs.length !== before.paragraphs.length + 1
  )
    return fail()
  const left = after.paragraphs[index],
    inserted = after.paragraphs[index + 1]
  if (
    !left ||
    !inserted ||
    left.text !== target.text ||
    inserted.text !== text ||
    left.nesting !== 0 ||
    inserted.nesting !== 0 ||
    inserted.id !== target.paragraphs[0]!.id ||
    left.id === inserted.id
  )
    return fail()
  for (let i = 0; i < before.paragraphs.length; i++)
    if (i !== index) {
      const next = after.paragraphs[i + (i > index ? 1 : 0)]!
      if (
        !next ||
        next.id !== before.paragraphs[i]!.id ||
        next.text !== before.paragraphs[i]!.text ||
        next.nesting !== before.paragraphs[i]!.nesting
      )
        return fail()
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
  const blank = (e: Element) => !elements(e).length && Array.from(e.attributes).every((x) => x.namespaceURI === XMLNS)
  const sig = (list: Element[]) => list.map(canon)
  const is = (name: string) => (e: Element) => e.namespaceURI === W && e.localName === name
  const anchorProps = named(oldP, "pPr")[0],
    anchorMark = anchorProps && named(anchorProps, "rPr")[0],
    base = anchorProps ? elements(anchorProps) : []
  // The anchor text keeps a copy of its properties on Word's new inserted mark.
  const pp = named(leftP, "pPr"),
    pr = pp.length === 1 ? named(pp[0]!, "rPr") : [],
    mark = pr.length === 1 ? named(pr[0]!, "ins") : []
  if (
    pp.length !== 1 ||
    elements(leftP)[0] !== pp[0] ||
    !clean(pp[0]!) ||
    pr.length !== 1 ||
    mark.length !== 1 ||
    mark[0]!.childNodes.length ||
    Array.from(pp[0]!.attributes).some((a) => a.namespaceURI !== XMLNS) ||
    Array.from(pr[0]!.attributes).some((a) => a.namespaceURI !== XMLNS)
  )
    return fail()
  // The new paragraph keeps the original mark: the anchor's properties, or the
  // requested style with a tracked change whose old properties are the anchor's.
  const np = named(newP, "pPr"),
    content = named(newP, "ins"),
    props = np[0],
    changes = props ? named(props, "pPrChange") : []
  if (
    !clean(newP) ||
    np.length > 1 ||
    (props && elements(newP)[0] !== props) ||
    content.length !== 1 ||
    elements(newP).length !== np.length + 1 ||
    !only(content[0]!, ["r"]) ||
    (props && Array.from(props.attributes).some((a) => a.namespaceURI !== XMLNS)) ||
    changes.length > 1
  )
    return fail()
  const current = props ? elements(props).filter((e) => e !== changes[0]) : []
  if (restyled) {
    const change = changes[0],
      prior = change ? named(change, "pPr") : [],
      style = current.filter(is("pStyle")),
      rest = (list: Element[]) => list.filter((e) => !is("pStyle")(e))
    if (
      !change ||
      elements(props!).at(-1) !== change ||
      elements(change).length !== 1 ||
      prior.length !== 1 ||
      Array.from(prior[0]!.attributes).some((a) => a.namespaceURI !== XMLNS) ||
      !equal(sig(elements(prior[0]!)), sig(base.filter((e) => !is("rPr")(e)))) ||
      style.length > 1 ||
      !equal(sig(rest(current)), sig(rest(base))) ||
      equal(sig(style), sig(base.filter(is("pStyle"))))
    )
      return fail()
  } else if (changes.length || !equal(sig(current), sig(base))) return fail()
  // Inserted text may only take the anchor mark's or final run's character properties.
  const lastRun = named(oldP, "r").at(-1),
    inherited = [anchorMark, lastRun && named(lastRun, "rPr")[0]].map((e) => (e ? sig(elements(e)) : []))
  let value = ""
  for (const run of named(content[0]!, "r")) {
    const kids = elements(run),
      rp = named(run, "rPr"),
      ts = named(run, "t")
    if (
      !validMeta(run) ||
      rp.length > 1 ||
      (rp.length === 1 &&
        (kids[0] !== rp[0] ||
          blank(rp[0]!) ||
          Array.from(rp[0]!.attributes).some((a) => a.namespaceURI !== XMLNS) ||
          !inherited.some((x) => equal(x, sig(elements(rp[0]!)))))) ||
      kids.length !== rp.length + 1 ||
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
    utc = mark[0]!.getAttributeNS(W16, "dateUtc"),
    owned = [mark[0]!, content[0]!, ...changes]
  for (const node of owned) {
    const n = id(node),
      when = node.getAttributeNS(W, "date"),
      at = when ? Date.parse(when) - Date.parse(date ?? "") : NaN
    if (
      !n ||
      !/^\d+$/.test(n) ||
      !author?.trim() ||
      !date ||
      Number.isNaN(Date.parse(date)) ||
      node.getAttributeNS(W, "author") !== author ||
      // The style is a second Word action in the same batch; its stamp may tick.
      (node === changes[0] ? !(at >= 0 && at <= 60_000) : when !== date) ||
      (node.getAttributeNS(W16, "dateUtc") || when) !== when ||
      (node !== changes[0] && node.getAttributeNS(W16, "dateUtc") !== utc) ||
      Array.from(node.attributes).some(
        (a) =>
          a.namespaceURI !== XMLNS &&
          !(a.namespaceURI === W && ["id", "author", "date"].includes(a.localName)) &&
          !(a.namespaceURI === W16 && a.localName === "dateUtc"),
      )
    )
      return fail()
  }
  if (new Set(owned.map(id)).size !== owned.length) return fail()
  // Inverse: drop the inserted mark and any containers it alone needed, then the
  // new paragraph; the anchor text and its copied properties must be unchanged.
  pr[0]!.removeChild(mark[0]!)
  if (!anchorMark && blank(pr[0]!)) pp[0]!.removeChild(pr[0]!)
  if (!anchorProps && blank(pp[0]!)) leftP.removeChild(pp[0]!)
  a.body.removeChild(newP)
  if (oldParaId) {
    const attr = leftP.getAttributeNodeNS(W14, "paraId")
    if (!attr) return fail()
    const name = attr.name
    leftP.removeAttributeNS(W14, "paraId")
    leftP.setAttributeNS(W14, name, oldParaId)
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
    owned.some((x) => newIds.includes(id(x)))
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
