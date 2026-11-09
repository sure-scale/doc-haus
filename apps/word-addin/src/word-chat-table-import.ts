import { DOMParser } from "@xmldom/xmldom"
import type { WordChatTableRowDeletionProposal } from "../../../dochaus/lib/word-chat-amendment-contract"
import { assertNativeWordChatRowDeletionTarget } from "./word-chat-native-row-delete"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const MAIN = "/word/document.xml"
type Inventory = Array<{ id: string; text: string; tableNestingLevel?: number }>
const children = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const named = (n: Node, name: string) => children(n).filter(e => e.namespaceURI === W && e.localName === name)
const fail = (): never => { throw Error("The table or its imported formatting changed; prepare and review it again") }
function parts(xml: string) {
  if (typeof xml !== "string" || xml.length > 16_000_000 || /<!DOCTYPE|<!ENTITY/i.test(xml)) return fail()
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
  if (doc.doctype || doc.documentElement?.namespaceURI !== PKG || doc.documentElement.localName !== "package") return fail()
  const found = new Map<string, Element>()
  for (const part of children(doc.documentElement)) {
    const name = part.getAttributeNS(PKG, "name")
    if (part.namespaceURI !== PKG || part.localName !== "part" || !name || found.has(name) || found.size >= 64) return fail()
    found.set(name, part)
  }
  if (!found.has(MAIN)) return fail()
  return found
}
function data(part: Element) {
  const nodes = children(part)
  if (nodes.length !== 1 || nodes[0]!.namespaceURI !== PKG || !["xmlData", "binaryData"].includes(nodes[0]!.localName)) return fail()
  return nodes[0]!
}
function body(part: Element) {
  const value = data(part), roots = children(value)
  if (value.localName !== "xmlData" || roots.length !== 1 || roots[0]!.namespaceURI !== W || roots[0]!.localName !== "document") return fail()
  const bodies = named(roots[0]!, "body")
  if (bodies.length !== 1) return fail()
  return bodies[0]!
}
// Prefix differences and indentation are not content changes. Preserve every
// actual attribute, including IDs, widths, spans and revision metadata.
function canonical(node: Node): unknown {
  if (node.nodeType === 1) {
    const e = node as Element
    if (e.namespaceURI === MC) return fail()
    return [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.namespaceURI !== XMLNS)
      .map(a => {
        // Ignorable contains namespace prefixes, not ordinary string values.
        // Its lexical equality alone does not establish the same semantics.
        if (a.namespaceURI === MC) {
          if (a.localName !== "Ignorable") return fail()
          const bindings = a.value.trim().split(/\s+/).map(prefix => {
            const uri = e.lookupNamespaceURI(prefix)
            if (!uri) return fail()
            return [prefix, uri]
          })
          return [a.namespaceURI, a.localName, a.value, bindings]
        }
        if (a.namespaceURI === "http://www.w3.org/2001/XMLSchema-instance" && a.localName === "type") return fail()
        return [a.namespaceURI, a.localName, a.value]
      }).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    Array.from(e.childNodes).map(canonical).filter(n => n !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) return node.nodeValue?.trim() || ["t", "delText", "binaryData"].includes((node.parentNode as Element)?.localName) ? ["text", node.nodeValue] : undefined
  return [node.nodeType, node.nodeName, node.nodeValue]
}
const same = (a: Node, b: Node) => JSON.stringify(canonical(a)) === JSON.stringify(canonical(b))
function relationships(part: Element) {
  const root = children(data(part))
  if (root.length !== 1 || root[0]!.namespaceURI !== REL || root[0]!.localName !== "Relationships") return fail()
  const links = children(root[0]!)
  if (links.some(e => e.namespaceURI !== REL || e.localName !== "Relationship" || !e.getAttribute("Id")) || new Set(links.map(e => e.getAttribute("Id"))).size !== links.length) return fail()
  return links
}
function closedBoundary(main: Element, table: Element) {
  const active = new Set<string>()
  const starts: Record<string, string> = { bookmarkStart: "bookmark", commentRangeStart: "comment", permStart: "perm", moveFromRangeStart: "moveFrom", moveToRangeStart: "moveTo", customXmlInsRangeStart: "customIns", customXmlDelRangeStart: "customDel" }
  const ends: Record<string, string> = { bookmarkEnd: "bookmark", commentRangeEnd: "comment", permEnd: "perm", moveFromRangeEnd: "moveFrom", moveToRangeEnd: "moveTo", customXmlInsRangeEnd: "customIns", customXmlDelRangeEnd: "customDel" }
  const visit = (e: Element, inside = false) => {
    if (e === table && active.size) return fail()
    const within = inside || e === table
    if (e.namespaceURI === W) {
      const family = starts[e.localName] ?? ends[e.localName]
      if (!family && /Range(?:Start|End)$/.test(e.localName)) return fail()
      if (family) {
        const id = e.getAttributeNS(W, "id")
        if (!id || !/^\d+$/.test(id) || within) return fail()
        const key = `${family}:${BigInt(id)}`
        if (starts[e.localName]) { if (active.has(key)) return fail(); active.add(key) }
        else { if (!active.delete(key)) return fail() }
      }
      if (within && ["commentReference", "fldChar", "instrText", "delInstrText"].includes(e.localName)) return fail()
    }
    for (const child of children(e)) visit(child, within)
    if (e === table && active.size) return fail()
  }
  visit(main)
  if (active.size) return fail()
}

/** Pre-dispatch check for a verified, scoped table artifact. This is an offline
 * candidate guard, not a native import route or proof of Word round-tripping.
 * The backend builder must already verify the narrow authored row operation.
 */
export function assertWordChatTableImportPreimage(sourceOoxml: string, currentOoxml: string, inventory: Inventory, proposal: WordChatTableRowDeletionProposal): void {
  assertNativeWordChatRowDeletionTarget(currentOoxml, inventory, proposal)
  const source = parts(sourceOoxml), current = parts(currentOoxml)
  const before = body(source.get(MAIN)!), now = body(current.get(MAIN)!), scoped = children(before)
  const table = named(now, "tbl")[proposal.table.tableIndex]
  if (children(before.parentNode!).length !== 1 || source.get(MAIN)!.getAttributeNS(PKG, "contentType") !== current.get(MAIN)!.getAttributeNS(PKG, "contentType") ||
    !same(before.cloneNode(false), now.cloneNode(false)) || !same(before.parentNode!.cloneNode(false), now.parentNode!.cloneNode(false)) ||
    scoped.length !== 1 || scoped[0]!.namespaceURI !== W || scoped[0]!.localName !== "tbl" || !table || !same(scoped[0]!, table)) return fail()
  closedBoundary(now, table)
  for (const [name, part] of source) {
    if (name === MAIN) continue
    // The paired backend payload imports only formatting dependencies. Never
    // turn a complete capture (comments, settings, add-ins) into a transport.
    if (name !== "/_rels/.rels" && !/^\/word\/(?:_rels\/document\.xml\.rels|styles\.xml|numbering\.xml|fontTable\.xml|theme\/[^/]+\.xml|(?:theme|_rels|fonts|media)\/.+)$/.test(name)) return fail()
    const fresh = current.get(name)
    if (!fresh || part.getAttributeNS(PKG, "contentType") !== fresh.getAttributeNS(PKG, "contentType")) return fail()
    if (name.endsWith(".rels")) {
      const oldLinks = relationships(part), newLinks = relationships(fresh)
      for (const link of oldLinks) {
        if (link.getAttribute("TargetMode") === "External") return fail()
        // Root relationship IDs have no content references. Document and
        // dependency IDs do, and must remain exactly the source identities.
        if (name === "/_rels/.rels") {
          if (oldLinks.length !== 1 || link.getAttribute("Type") !== "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" || !newLinks.some(e => e.getAttribute("Type") === link.getAttribute("Type") && e.getAttribute("Target") === link.getAttribute("Target") && e.getAttribute("TargetMode") !== "External")) return fail()
        } else if (!newLinks.some(e => same(link, e))) return fail()
      }
    } else if (!same(data(part), data(fresh))) return fail()
  }
}
