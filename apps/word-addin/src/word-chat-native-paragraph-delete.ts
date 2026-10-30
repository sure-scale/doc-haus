import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { assertNativeWordChatInlineDependencies, projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { normalizeNativeWordChatCommentIds } from "./word-chat-native-comments"
import { assertNativeWordChatParagraphInsertionTarget } from "./word-chat-native-paragraph-insert"

type Projection = ReturnType<typeof projectNativeWordChatBody>
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  W14 = "http://schemas.microsoft.com/office/word/2010/wordml",
  W16 = "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
  PKG = "http://schemas.microsoft.com/office/2006/xmlPackage",
  XMLNS = "http://www.w3.org/2000/xmlns/"
const elements = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const named = (n: Node, name: string) => elements(n).filter((e) => e.namespaceURI === W && e.localName === name)
const fail = (): never => {
  throw Error("Native paragraph deletion could not be verified")
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
function canon(n: Node): unknown {
  if (n.nodeType === 1) {
    const e = n as Element
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
      (a.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)) ||
      (a.namespaceURI === W14 && ["paraId", "textId"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)),
  )
}
/** Restricts deletion to a complete plain top-level paragraph with a plain following boundary. */
export function assertNativeWordChatParagraphDeletionTarget(projection: Projection, target: WordChatTarget): number {
  const index = assertNativeWordChatParagraphInsertionTarget(projection, target)
  const next = projection.paragraphs[index + 1]
  if (!next || !next.text) return fail()
  assertNativeWordChatParagraphInsertionTarget(projection, {
    ...target,
    id: next.id,
    text: next.text,
    startOffset: 0,
    endOffset: next.text.length,
    paragraphs: [{ id: next.id, text: next.text }],
  })
  return index
}

/** Checks the captured main story, annotation graph and formatting dependencies;
 * compressed package metadata remains a separate saved-package audit. */
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
  assertNativeWordChatInlineDependencies(before.ooxml, after.ooxml)
  const b = parse(before.ooxml),
    a = parse(after.ooxml)
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
    if (projection.paragraphs.some((p) => p.path === `body/p[${children.length - 2}]`)) return
    if (!validMeta(tail) || Array.from(tail.childNodes).some((n) => n.nodeType !== 3 || Boolean(n.nodeValue?.trim())))
      return fail()
    body.removeChild(tail)
  }
  trimUnboundTerminal(b.body, before)
  trimUnboundTerminal(a.body, after)
  const match = /^body\/p\[(\d+)\]$/.exec(before.paragraphs[index]!.path)
  if (!match) return fail()
  const pos = Number(match[1]),
    oldP = elements(b.body)[pos],
    actual = elements(a.body)[pos]
  if (
    !oldP ||
    !actual ||
    oldP.namespaceURI !== W ||
    actual.namespaceURI !== W ||
    oldP.localName !== "p" ||
    actual.localName !== "p"
  )
    return fail()
  // Only the target's evidenced textId and rsidDel may differ. rsidR/default are
  // handled by the same serialization canonicalization as the insertion verifier.
  if (
    oldP.getAttributeNS(W14, "paraId") !== actual.getAttributeNS(W14, "paraId") ||
    Boolean(oldP.getAttributeNS(W14, "textId")) !== Boolean(actual.getAttributeNS(W14, "textId"))
  )
    return fail()
  for (const attr of Array.from(actual.attributes)) {
    if (attr.namespaceURI === XMLNS) continue
    if (
      attr.namespaceURI === W &&
      ["rsidR", "rsidRDefault", "rsidDel"].includes(attr.localName) &&
      /^[0-9a-fA-F]{8}$/.test(attr.value)
    )
      continue
    if (
      attr.namespaceURI === W14 &&
      ["paraId", "textId"].includes(attr.localName) &&
      /^[0-9a-fA-F]{8}$/.test(attr.value)
    )
      continue
    return fail()
  }
  const clean = (e: Element): boolean =>
    Array.from(e.childNodes).every((n) =>
      n.nodeType === 1
        ? clean(n as Element)
        : n.nodeType === 3 && ((e.namespaceURI === W && e.localName === "delText") || !n.nodeValue?.trim()),
    )
  if (!clean(actual)) return fail()
  const pp = named(actual, "pPr"),
    pr = pp.length === 1 ? named(pp[0]!, "rPr") : [],
    mark = pr.length === 1 ? named(pr[0]!, "del") : [],
    content = named(actual, "del")
  const noAttrs = (e: Element) => Array.from(e.attributes).every((x) => x.namespaceURI === XMLNS)
  if (
    elements(actual).length !== 2 ||
    elements(actual)[0] !== pp[0] ||
    elements(actual)[1] !== content[0] ||
    pp.length !== 1 ||
    pr.length !== 1 ||
    mark.length !== 1 ||
    content.length !== 1 ||
    elements(pp[0]!).length !== 1 ||
    elements(pr[0]!).length !== 1 ||
    !noAttrs(pp[0]!) ||
    !noAttrs(pr[0]!) ||
    mark[0]!.childNodes.length
  )
    return fail()
  const id = (e: Element) => e.getAttributeNS(W, "id"),
    author = mark[0]!.getAttributeNS(W, "author"),
    date = mark[0]!.getAttributeNS(W, "date"),
    utc = mark[0]!.getAttributeNS(W16, "dateUtc")
  for (const node of [mark[0]!, content[0]!]) {
    if (
      !/^\d+$/.test(id(node) || "") ||
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
  }
  if (id(mark[0]!) === id(content[0]!)) return fail()
  const runs = named(content[0]!, "r"),
    oldRun = named(oldP, "r")[0]
  if (elements(content[0]!).length !== 1 || runs.length !== 1 || !oldRun) return fail()
  const run = runs[0]!,
    ts = named(run, "delText")
  if (
    elements(run).length !== 1 ||
    ts.length !== 1 ||
    ts[0]!.textContent !== target.text ||
    elements(ts[0]!).length ||
    Array.from(ts[0]!.childNodes).some((n) => n.nodeType !== 3)
  )
    return fail()
  // Restore the validated run by translating only the evidenced deletion markup;
  // compare all original run/text attributes and significant whitespace.
  const restoredRun = run.cloneNode(true) as Element
  const delRsid = restoredRun.getAttributeNS(W, "rsidDel")
  if (delRsid && !/^[0-9a-fA-F]{8}$/.test(delRsid)) return fail()
  restoredRun.removeAttributeNS(W, "rsidDel")
  const dt = named(restoredRun, "delText")[0]!,
    t = a.doc.createElementNS(W, dt.prefix ? `${dt.prefix}:t` : "t")
  for (const attr of Array.from(dt.attributes)) t.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
  while (dt.firstChild) t.appendChild(dt.firstChild)
  restoredRun.replaceChild(t, dt)
  if (!equal(canon(oldRun), canon(restoredRun))) return fail()
  const markId = id(mark[0]!)!,
    contentId = id(content[0]!)!
  a.body.replaceChild(oldP.cloneNode(true), actual)
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
    newIds.includes(markId) ||
    newIds.includes(contentId)
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
