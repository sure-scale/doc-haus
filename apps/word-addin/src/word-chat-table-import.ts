import { sameTableImportXml } from "./word-chat-table-import-xml"
import { DOMParser } from "@xmldom/xmldom"
import type { WordChatTableRowDeletionProposal } from "../../../dochaus/lib/word-chat-amendment-contract"
import { assertNativeWordChatRowDeletionTarget } from "./word-chat-native-row-delete"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const MAIN = "/word/document.xml"
type Inventory = Array<{ id: string; text: string; tableNestingLevel?: number }>
const children = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const named = (n: Node, name: string) => children(n).filter(e => e.namespaceURI === W && e.localName === name)
const messages = {
  package_xml: "The table's Word package could not be read safely. Refresh and review it again.",
  document_context: "The table's document context does not match the exported file. Refresh and review it again.",
  table_preimage: "The reviewed table or its paragraph identities could not be matched. Refresh and review it again.",
  table_boundary: "The table's annotation or content boundary could not be verified. Refresh and review it again.",
  formatting_dependencies: "The table's imported formatting or relationships could not be verified. Refresh and review it again.",
  range_paragraphs: "Word's table range does not contain the reviewed paragraph identities. Refresh and review it again.",
} as const
export type WordChatTableImportErrorCode = keyof typeof messages
/** Static public notice only; an original parser/native cause stays private.
 * Dispatch uncertainty remains the caller's responsibility, before using this notice. */
export class WordChatTableImportError extends Error {
  constructor(readonly code: WordChatTableImportErrorCode, cause?: unknown) {
    super(messages[code], cause === undefined ? undefined : { cause })
    this.name = "WordChatTableImportError"
  }
}
const fail = (code: WordChatTableImportErrorCode, cause?: unknown): never => { throw new WordChatTableImportError(code, cause) }
function parts(xml: string) {
  if (typeof xml !== "string" || xml.length > 16_000_000 || /<!DOCTYPE|<!ENTITY/i.test(xml)) return fail("package_xml")
  let doc: Document
  try {
    doc = new DOMParser({ errorHandler: { warning: (cause) => fail("package_xml", cause), error: (cause) => fail("package_xml", cause), fatalError: (cause) => fail("package_xml", cause) } }).parseFromString(xml, "application/xml")
  } catch (cause) {
    if (cause instanceof WordChatTableImportError) throw cause
    return fail("package_xml", cause)
  }
  if (doc.doctype || doc.documentElement?.namespaceURI !== PKG || doc.documentElement.localName !== "package") return fail("package_xml")
  const found = new Map<string, Element>()
  for (const part of children(doc.documentElement)) {
    const name = part.getAttributeNS(PKG, "name")
    if (part.namespaceURI !== PKG || part.localName !== "part" || !name || found.has(name) || found.size >= 64) return fail("package_xml")
    found.set(name, part)
  }
  if (!found.has(MAIN)) return fail("package_xml")
  return found
}
function data(part: Element) {
  const nodes = children(part)
  if (nodes.length !== 1 || nodes[0]!.namespaceURI !== PKG || !["xmlData", "binaryData"].includes(nodes[0]!.localName)) return fail("package_xml")
  return nodes[0]!
}
function body(part: Element) {
  const value = data(part), roots = children(value)
  if (value.localName !== "xmlData" || roots.length !== 1 || roots[0]!.namespaceURI !== W || roots[0]!.localName !== "document") return fail("document_context")
  const bodies = named(roots[0]!, "body")
  if (bodies.length !== 1) return fail("document_context")
  return bodies[0]!
}
// Prefix differences and indentation are not content changes. Preserve every
// actual attribute, including IDs, widths, spans and revision metadata.
function same(a: Node, b: Node, code: WordChatTableImportErrorCode) {
  try { return sameTableImportXml(a, b) }
  catch (cause) { return fail(code, cause) }
}
function relationships(part: Element) {
  const root = children(data(part))
  if (root.length !== 1 || root[0]!.namespaceURI !== REL || root[0]!.localName !== "Relationships") return fail("formatting_dependencies")
  const links = children(root[0]!)
  if (links.some(e => e.namespaceURI !== REL || e.localName !== "Relationship" || !e.getAttribute("Id")) || new Set(links.map(e => e.getAttribute("Id"))).size !== links.length) return fail("formatting_dependencies")
  return links
}
function closedBoundary(main: Element, table: Element) {
  const active = new Set<string>()
  const starts: Record<string, string> = { bookmarkStart: "bookmark", commentRangeStart: "comment", permStart: "perm", moveFromRangeStart: "moveFrom", moveToRangeStart: "moveTo", customXmlInsRangeStart: "customIns", customXmlDelRangeStart: "customDel" }
  const ends: Record<string, string> = { bookmarkEnd: "bookmark", commentRangeEnd: "comment", permEnd: "perm", moveFromRangeEnd: "moveFrom", moveToRangeEnd: "moveTo", customXmlInsRangeEnd: "customIns", customXmlDelRangeEnd: "customDel" }
  const visit = (e: Element, inside = false) => {
    if (e === table && active.size) return fail("table_boundary")
    const within = inside || e === table
    if (e.namespaceURI === W) {
      const family = starts[e.localName] ?? ends[e.localName]
      if (!family && /Range(?:Start|End)$/.test(e.localName)) return fail("table_boundary")
      if (family) {
        const id = e.getAttributeNS(W, "id")
        if (!id || !/^\d+$/.test(id) || within) return fail("table_boundary")
        const key = `${family}:${BigInt(id)}`
        if (starts[e.localName]) { if (active.has(key)) return fail("table_boundary"); active.add(key) }
        else { if (!active.delete(key)) return fail("table_boundary") }
      }
      if (within && ["commentReference", "fldChar", "instrText", "delInstrText"].includes(e.localName)) return fail("table_boundary")
    }
    for (const child of children(e)) visit(child, within)
    if (e === table && active.size) return fail("table_boundary")
  }
  visit(main)
  if (active.size) return fail("table_boundary")
}

/** Pre-dispatch check for a verified, scoped table artifact. This is an offline
 * candidate guard, not a native import route or proof of Word round-tripping.
 * The backend builder must already verify the narrow authored row operation.
 */
export function assertWordChatTableImportPreimage(sourceOoxml: string, currentOoxml: string, inventory: Inventory, proposal: WordChatTableRowDeletionProposal): void {
  try { assertNativeWordChatRowDeletionTarget(currentOoxml, inventory, proposal) }
  catch (cause) { return fail("table_preimage", cause) }
  const source = parts(sourceOoxml), current = parts(currentOoxml)
  const before = body(source.get(MAIN)!), now = body(current.get(MAIN)!), scoped = children(before)
  const table = named(now, "tbl")[proposal.table.tableIndex]
  if (children(before.parentNode!).length !== 1 || source.get(MAIN)!.getAttributeNS(PKG, "contentType") !== current.get(MAIN)!.getAttributeNS(PKG, "contentType") ||
    !same(before.cloneNode(false), now.cloneNode(false), "document_context") || !same(before.parentNode!.cloneNode(false), now.parentNode!.cloneNode(false), "document_context")) return fail("document_context")
  if (scoped.length !== 1 || scoped[0]!.namespaceURI !== W || scoped[0]!.localName !== "tbl") return fail("table_boundary")
  if (!table || !same(scoped[0]!, table, "table_preimage")) return fail("table_preimage")
  closedBoundary(now, table)
  for (const [name, part] of source) {
    if (name === MAIN) continue
    // The paired backend payload imports only formatting dependencies. Never
    // turn a complete capture (comments, settings, add-ins) into a transport.
    if (name !== "/_rels/.rels" && !/^\/word\/(?:_rels\/document\.xml\.rels|styles\.xml|numbering\.xml|fontTable\.xml|theme\/[^/]+\.xml|(?:theme|_rels|fonts|media)\/.+)$/.test(name)) return fail("formatting_dependencies")
    const fresh = current.get(name)
    if (!fresh || part.getAttributeNS(PKG, "contentType") !== fresh.getAttributeNS(PKG, "contentType")) return fail("formatting_dependencies")
    if (name.endsWith(".rels")) {
      const oldLinks = relationships(part), newLinks = relationships(fresh)
      for (const link of oldLinks) {
        if (link.getAttribute("TargetMode") === "External") return fail("formatting_dependencies")
        // Root relationship IDs have no content references. Document and
        // dependency IDs do, and must remain exactly the source identities.
        if (name === "/_rels/.rels") {
          if (oldLinks.length !== 1 || link.getAttribute("Type") !== "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" || !newLinks.some(e => e.getAttribute("Type") === link.getAttribute("Type") && e.getAttribute("Target") === link.getAttribute("Target") && e.getAttribute("TargetMode") !== "External")) return fail("formatting_dependencies")
        } else if (!newLinks.some(e => same(link, e, "formatting_dependencies"))) return fail("formatting_dependencies")
      }
    } else if (!same(data(part), data(fresh), "formatting_dependencies")) return fail("formatting_dependencies")
  }
}

/** Match the ordered paragraph identities read from the bound native table's
 * Whole range. The caller must also verify native containment and the complete
 * source preimage. A range OOXML export is a separate serialization, so it is
 * not used as a substitute for that native boundary or the full-body proof. */
export function assertWordChatTableImportRange(rangeParagraphIds: string[], proposal: WordChatTableRowDeletionProposal): void {
  const expected = proposal.table.rows.flat().map(cell => cell.id)
  if (!expected.length || JSON.stringify(expected) !== JSON.stringify(rangeParagraphIds) || new Set(rangeParagraphIds).size !== expected.length) return fail("range_paragraphs")
}

/** After the structural/revision verifier has proved the imported result and
 * relationship graph, retain namespace-aware equality for the actual formatting
 * parts too. Runtime identity mapping never authorizes changed style meaning. */
export function assertWordChatTableImportFormatting(sourceOoxml: string, observedOoxml: string): void {
  const source = parts(sourceOoxml), observed = parts(observedOoxml)
  for (const [name, original] of source) {
    if (name === MAIN || name.endsWith(".rels")) continue
    const current = observed.get(name)
    if (!current || original.getAttributeNS(PKG, "contentType") !== current.getAttributeNS(PKG, "contentType") || !same(data(original), data(current), "formatting_dependencies")) return fail("formatting_dependencies")
  }
}
