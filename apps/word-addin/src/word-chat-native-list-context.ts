import { DOMParser } from "@xmldom/xmldom"
import { wordChatListCaptureCurrent, type WordChatListCapture } from "../../../dochaus/lib/word-chat-amendment-contract"
import type { projectNativeWordChatBody } from "./word-chat-host-ooxml"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const elements = (n: Node) => Array.from(n.childNodes).filter((n): n is Element => n.nodeType === 1)
const named = (n: Node, name: string) => elements(n).filter(e => e.namespaceURI === W && e.localName === name)
function parse(xml: string) {
  if (xml.length > 16_000_000 || /<!DOCTYPE/i.test(xml)) throw Error("Invalid list XML")
  const reject = () => { throw Error("Invalid list XML") }
  return new DOMParser({ errorHandler: { warning: reject, error: reject, fatalError: reject } }).parseFromString(xml, "application/xml")
}
function one(nodes: Element[]): Element {
  if (nodes.length !== 1) throw Error("Ambiguous list definition")
  return nodes[0]!
}
const value = (e: Element) => e.getAttributeNS(W, "val")
function numberId(e: Element): string {
  const id = value(e)
  if (typeof id !== "string" || !Number.isSafeInteger(Number(id)) || Number(id) <= 0 || String(Number(id)) !== id)
    throw Error("Invalid numbering reference")
  return id
}
/** Corroborate native list observations with the document's actual numbering
 * instance. Typed labels and a style name cannot establish list membership.
 * This first executor admits level-zero decimal items starting at 1. Unused
 * levels in the document's definition do not change the captured flat list.
 */
export function nativeWordChatDecimalList(projection: ReturnType<typeof projectNativeWordChatBody>, capture: WordChatListCapture): boolean {
  try {
    const paragraphs = projection.paragraphs.map(p => ({ id: p.id, text: p.text }))
    if (!wordChatListCaptureCurrent({ text: "", paragraphs, saved: false, capturedAt: "", truncated: false }, capture)) return false
    let instance: string | undefined
    for (const member of capture.members) {
      const p = projection.paragraphs.find(p => p.id === member.id)
      if (!p || p.nesting !== 0 || !p.path.startsWith("body/p[")) return false
      const root = parse(p.xml).documentElement
      const properties = one(named(root, "pPr")), numbering = one(named(properties, "numPr"))
      if (elements(numbering).length !== 2 || value(one(named(numbering, "ilvl"))) !== "0") return false
      const id = numberId(one(named(numbering, "numId")))
      if (instance !== undefined && instance !== id) return false
      instance = id
    }
    const doc = parse(projection.ooxml)
    const parts = Array.from(doc.getElementsByTagNameNS(PKG, "part"))
      .filter(p => p.getAttributeNS(PKG, "name") === "/word/numbering.xml")
    const numbering = one(Array.from(one(parts).getElementsByTagNameNS(W, "numbering")))
    const num = one(named(numbering, "num").filter(e => e.getAttributeNS(W, "numId") === instance))
    if (elements(num).length !== 1) return false // Overrides/restarts require a distinct reviewed operation.
    const abstractId = value(one(named(num, "abstractNumId")))
    const abstract = one(named(numbering, "abstractNum").filter(e => e.getAttributeNS(W, "abstractNumId") === abstractId))
    const level = one(named(abstract, "lvl").filter(e => e.getAttributeNS(W, "ilvl") === "0"))
    if (level.getAttributeNS(W, "ilvl") !== "0" || value(one(named(level, "numFmt"))) !== "decimal" ||
      value(one(named(level, "start"))) !== "1" || value(one(named(level, "lvlText"))) !== "%1.") return false
    // A partial native collection must not hide another item of this instance.
    const ids = new Set(capture.members.map(p => p.id))
    for (const p of projection.paragraphs) {
      const root = parse(p.xml).documentElement
      const matches = named(root, "pPr").flatMap(pp => named(pp, "numPr")).flatMap(np => named(np, "numId"))
      if (matches.some(e => value(e) === instance) && !ids.has(p.id)) return false
    }
    return true
  } catch { return false }
}
