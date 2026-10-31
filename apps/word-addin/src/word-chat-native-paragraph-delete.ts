import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { assertNativeWordChatStoryInverse } from "./word-chat-native-paragraph-merge"

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
  throw Error("Native paragraph deletion could not be verified")
}
const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
// Word's editing-session ids on paragraphs and runs carry no content; Word
// stamps rsidP/rsidRPr whenever paragraph or run properties change.
const session = ["rsidR", "rsidRDefault", "rsidP", "rsidRPr"]
const numeric = (s: string) => s.replace(/^0+(?=\d)/, "")
const hex = (v: string | null) => /^[0-9a-fA-F]{8}$/.test(v ?? "")
const blank = (n: Node) => n.nodeType === 3 && !n.nodeValue?.trim()
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
    return n.nodeValue?.trim() || ["t", "delText"].includes((n.parentNode as Element)?.localName)
      ? ["text", n.nodeValue]
      : undefined
  if (n.nodeType === 7 || n.nodeType === 8) return [n.nodeType, n.nodeName, n.nodeValue]
  return undefined
}
function validMeta(e: Element) {
  return Array.from(e.attributes).every(
    (a) =>
      a.namespaceURI === XMLNS ||
      (a.namespaceURI === W && session.includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)) ||
      (a.namespaceURI === W14 && ["paraId", "textId"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)),
  )
}
// Revisions and section breaks inside the target would make the inverse inexact.
const unsupported = new Set(["ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange", "numberingChange", "sectPr"])
/** A plain run child: the text carrier (w:t, or w:delText once deleted) or an
 * attribute-free tab, line break or rendered-page-break cache. */
function plain(e: Element, carrier: "t" | "delText") {
  if (e.namespaceURI !== W) return false
  if (e.localName === carrier)
    return (
      Array.from(e.attributes).every(
        (a) => a.namespaceURI === XMLNS || (a.namespaceURI === XML && a.localName === "space" && a.value === "preserve"),
      ) && Array.from(e.childNodes).every((n) => n.nodeType === 3)
    )
  return (
    ["tab", "br", "cr", "lastRenderedPageBreak"].includes(e.localName) &&
    !e.childNodes.length &&
    Array.from(e.attributes).every((a) => a.namespaceURI === XMLNS)
  )
}
function qualified(p: Element | undefined, text: string) {
  if (!p || p.namespaceURI !== W || p.localName !== "p" || !validMeta(p)) return fail()
  const children = elements(p),
    pp = named(p, "pPr"),
    runs = named(p, "r")
  if (
    pp.length > 1 ||
    (pp.length && children[0] !== pp[0]) ||
    children.length !== pp.length + runs.length ||
    !runs.length ||
    Array.from(p.childNodes).some((n) => n.nodeType !== 1 && !blank(n)) ||
    Array.from(p.getElementsByTagNameNS(W, "*")).some((e) => unsupported.has(e.localName))
  )
    return fail()
  let value = ""
  for (const r of runs) {
    const parts = elements(r),
      props = named(r, "rPr")
    if (
      props.length > 1 ||
      (props.length && parts[0] !== props[0]) ||
      Array.from(r.childNodes).some((n) => n.nodeType !== 1 && !blank(n)) ||
      Array.from(r.attributes).some(
        (a) => a.namespaceURI !== XMLNS && !(a.namespaceURI === W && ["rsidR", "rsidRPr"].includes(a.localName) && hex(a.value)),
      )
    )
      return fail()
    for (const c of parts.slice(props.length)) {
      if (!plain(c, "t")) return fail()
      value += c.localName === "t" ? (c.textContent ?? "") : c.localName === "tab" ? "\t" : c.localName === "lastRenderedPageBreak" ? "" : "\u000b"
    }
  }
  if (value !== text) return fail()
}
/** One complete top-level paragraph of any style, numbering or run formatting
 * whose runs carry only plain text, followed by a top-level text paragraph. */
export function assertNativeWordChatParagraphDeletionTarget(projection: Projection, target: WordChatTarget): number {
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
    item = projection.paragraphs[index],
    next = projection.paragraphs[index + 1]
  if (!item || item.text !== target.text || item.nesting !== 0 || !/^body\/p\[\d+\]$/.test(item.path)) return fail()
  if (!next || !next.id || !next.text || next.nesting !== 0 || !/^body\/p\[\d+\]$/.test(next.path)) return fail()
  qualified(elements(parse(projection.ooxml).body)[Number(/\d+/.exec(item.path)![0])], target.text)
  return index
}

/** Proves Word's deletion carriers on the target, inverts exactly those, and
 * compares the restored paragraph and complete story with the capture. */
export function verifyNativeWordChatParagraphDeletion(
  before: Projection,
  after: Projection,
  target: WordChatTarget,
): string[] {
  const index = assertNativeWordChatParagraphDeletionTarget(before, target)
  if (after.paragraphs.length !== before.paragraphs.length) return fail()
  for (let i = 0; i < before.paragraphs.length; i++) {
    const x = before.paragraphs[i]!,
      y = after.paragraphs[i]
    if (!y || x.id !== y.id || x.path !== y.path || x.nesting !== y.nesting || y.text !== (i === index ? "" : x.text))
      return fail()
  }
  const pos = Number(/\d+/.exec(before.paragraphs[index]!.path)![0]),
    oldP = elements(parse(before.ooxml).body)[pos]!,
    a = parse(after.ooxml),
    actual = elements(a.body)[pos]
  if (!actual || actual.namespaceURI !== W || actual.localName !== "p") return fail()
  // Only the target's evidenced textId and rsidDel may differ. Session ids are
  // handled by the same serialization canonicalization as the insertion verifier.
  if (
    oldP.getAttributeNS(W14, "paraId") !== actual.getAttributeNS(W14, "paraId") ||
    Boolean(oldP.getAttributeNS(W14, "textId")) !== Boolean(actual.getAttributeNS(W14, "textId")) ||
    Array.from(actual.attributes).some(
      (x) =>
        x.namespaceURI !== XMLNS &&
        !(x.namespaceURI === W && [...session, "rsidDel"].includes(x.localName) && hex(x.value)) &&
        !(x.namespaceURI === W14 && ["paraId", "textId"].includes(x.localName) && hex(x.value)),
    )
  )
    return fail()
  const children = elements(actual),
    pp = named(actual, "pPr"),
    rp = pp.length === 1 ? named(pp[0]!, "rPr") : [],
    marks = rp.length === 1 ? named(rp[0]!, "del") : [],
    content = children.slice(pp.length)
  if (
    pp.length !== 1 ||
    children[0] !== pp[0] ||
    marks.length !== 1 ||
    marks[0]!.childNodes.length ||
    !content.length ||
    content.some((e) => e.namespaceURI !== W || e.localName !== "del") ||
    Array.from(actual.childNodes).some((n) => n.nodeType !== 1 && !blank(n))
  )
    return fail()
  const mark = marks[0]!,
    author = mark.getAttributeNS(W, "author"),
    date = mark.getAttributeNS(W, "date"),
    utc = mark.getAttributeNS(W16, "dateUtc"),
    ids: string[] = []
  for (const node of [mark, ...content]) {
    const id = node.getAttributeNS(W, "id")
    if (
      !id ||
      !/^\d+$/.test(id) ||
      !author?.trim() ||
      !date ||
      Number.isNaN(Date.parse(date)) ||
      node.getAttributeNS(W, "author") !== author ||
      node.getAttributeNS(W, "date") !== date ||
      node.getAttributeNS(W16, "dateUtc") !== utc ||
      (utc && utc !== date) ||
      Array.from(node.attributes).some(
        (x) =>
          x.namespaceURI !== XMLNS &&
          !(x.namespaceURI === W && ["id", "author", "date"].includes(x.localName)) &&
          !(x.namespaceURI === W16 && x.localName === "dateUtc"),
      )
    )
      return fail()
    ids.push(numeric(id))
  }
  if (new Set(ids).size !== ids.length) return fail()
  // Inverse: unwrap each deleted run, restoring w:t and dropping rsidDel; every
  // text carrier must be deleted, so a surviving w:t fails as unplain.
  for (const d of content) {
    const runs = named(d, "r")
    if (!runs.length || elements(d).length !== runs.length || Array.from(d.childNodes).some((n) => n.nodeType !== 1 && !blank(n)))
      return fail()
    for (const r of runs) {
      if (r.hasAttributeNS(W, "rsidDel") && !hex(r.getAttributeNS(W, "rsidDel"))) return fail()
      r.removeAttributeNS(W, "rsidDel")
      for (const c of elements(r)) {
        if (c.namespaceURI === W && c.localName === "rPr") continue
        if (!plain(c, "delText")) return fail()
        if (c.localName !== "delText") continue
        const t = a.doc.createElementNS(W, c.prefix ? `${c.prefix}:t` : "t")
        for (const x of Array.from(c.attributes)) t.setAttributeNS(x.namespaceURI, x.name, x.value)
        while (c.firstChild) t.appendChild(c.firstChild)
        r.replaceChild(t, c)
      }
      actual.insertBefore(r, d)
    }
    actual.removeChild(d)
  }
  // Remove the mark, and a container Word added for it.
  rp[0]!.removeChild(mark)
  const priorProps = named(oldP, "pPr"),
    priorRunProps = priorProps.length === 1 ? named(priorProps[0]!, "rPr") : []
  const empty = (e: Element) =>
    !elements(e).length && Array.from(e.attributes).every((x) => x.namespaceURI === XMLNS) && Array.from(e.childNodes).every(blank)
  if (!priorRunProps.length && empty(rp[0]!)) pp[0]!.removeChild(rp[0]!)
  if (!priorProps.length && empty(pp[0]!)) actual.removeChild(pp[0]!)
  actual.removeAttributeNS(W, "rsidDel")
  if (oldP.hasAttributeNS(W14, "textId")) actual.setAttributeNS(W14, "w14:textId", oldP.getAttributeNS(W14, "textId")!)
  if (!equal(canon(oldP), canon(actual))) return fail()
  try {
    assertNativeWordChatStoryInverse(before, after, new XMLSerializer().serializeToString(a.doc), author!, ids)
  } catch (cause) {
    throw Error("Native paragraph deletion could not be verified", { cause })
  }
  return [author!]
}
