import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync } from "fflate"
import { MAX_WORD_AMENDMENT_OOXML_BYTES } from "./docx-flat-opc"
import { docxTableRowDeletionTargets, verifyDocxTableRowDeletion, type DocxTableRowDeletion } from "./docx-table-row-deletion"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const MAIN = "word/document.xml", REL_TYPE = "application/vnd.openxmlformats-package.relationships+xml"
const formatting = new Set(["styles", "theme", "fontTable", "numbering", "font"])
const formattingType = (kind: string, type: string) => kind === "theme" ? type === "application/vnd.openxmlformats-officedocument.theme+xml"
  : kind === "font" ? ["application/vnd.openxmlformats-officedocument.obfuscatedFont", "application/x-font-ttf", "application/vnd.ms-opentype"].includes(type)
  : kind === "image" ? type.startsWith("image/")
  : type === `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`
const serializer = new XMLSerializer()
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const children = (node: Node, local: string) => kids(node).filter(n => n.namespaceURI === W && n.localName === local)
const all = (node: Document | Element) => Array.from(node.getElementsByTagName("*"))
const xml = (node: Node) => serializer.serializeToString(node)
function require(value: unknown, reason: string): asserts value { if (!value) throw new Error(`docx_row_deletion_import:${reason}`) }
function parse(value: Uint8Array | string) {
  const text = typeof value === "string" ? value : strFromU8(value)
  require(!/<!DOCTYPE|<!ENTITY/i.test(text), "unsupported_xml")
  const fail = () => { throw new Error("docx_row_deletion_import:invalid_xml") }
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(text, "application/xml")
  require(doc.documentElement && !doc.doctype, "invalid_xml")
  return doc
}
function body(doc: Document) {
  require(doc.documentElement.namespaceURI === W && doc.documentElement.localName === "document", "invalid_main")
  const found = children(doc.documentElement, "body")
  require(found.length === 1 && kids(doc.documentElement).length === 1, "incomplete_main_boundary")
  return found[0]!
}
function standalone(node: Element) {
  const copy = node.cloneNode(true) as Element
  const present = new Set(Array.from(copy.attributes).filter(a => a.namespaceURI === XMLNS).map(a => a.name))
  for (let parent = node.parentNode; parent?.nodeType === 1; parent = parent.parentNode)
    for (const attr of Array.from((parent as Element).attributes)) if (attr.namespaceURI === XMLNS && !present.has(attr.name)) {
      copy.setAttributeNS(XMLNS, attr.name, attr.value); present.add(attr.name)
    }
  return xml(copy)
}

// An enclosing annotation can have BOTH endpoints outside this table. Inspect
// structural extents in the original story, not just descendants of the row.
function completeTable(doc: Document, index: number) {
  const story = body(doc), table = children(story, "tbl")[index - 1]
  require(table && !table.getElementsByTagNameNS(W, "tbl").length, "table_unavailable_or_nested")
  require(!all(table).some(n => n.namespaceURI === W && (["gridSpan", "vMerge", "hMerge", "gridBefore", "gridAfter", "wBefore", "wAfter"].includes(n.localName))), "merged_geometry_unsupported")
  const extents = new Map<Node, [number, number]>()
  let cursor = 0
  const visit = (node: Node) => {
    const start = cursor++
    if (node.nodeType === 1 || node.nodeType === 9) Array.from(node.childNodes).forEach(visit)
    extents.set(node, [start, cursor++])
  }
  visit(story)
  const [start, end] = extents.get(table)!
  const markers = all(story).filter(n => n.namespaceURI === W && (/Range(?:Start|End)$/.test(n.localName) || ["bookmarkStart", "bookmarkEnd", "permStart", "permEnd"].includes(n.localName)))
  const inside = (node: Node) => { const [a, b] = extents.get(node)!; return a >= start && b <= end }
  require(markers.every(n => !inside(n)), "table_annotation_unsupported")
  const starts = markers.filter(n => /Start$/.test(n.localName))
  for (const first of starts) {
    const id = first.getAttributeNS(W, "id"), lastName = first.localName.replace(/Start$/, "End")
    require(id !== null && /^[+-]?\d+$/.test(id), "invalid_story_range")
    const matches = markers.filter(n => n.localName === lastName && n.getAttributeNS(W, "id") !== null && /^[+-]?\d+$/.test(n.getAttributeNS(W, "id")!) && BigInt(n.getAttributeNS(W, "id")!) === BigInt(id))
    require(matches.length === 1, "ambiguous_story_range")
    const a = extents.get(first)![0], b = extents.get(matches[0]!)![1]
    require(a < b && (b < start || a > end), "annotation_crosses_table")
  }
  require(markers.length === starts.length * 2, "ambiguous_story_range")
  require(!all(table).some(n => n.namespaceURI === W && ["fldChar", "fldSimple", "instrText", "delInstrText", "commentReference", "annotationRef", "proofErr", "footnoteReference", "endnoteReference", "hyperlink", "sectPr"].includes(n.localName)), "table_annotation_or_field_unsupported")
  return table
}
function onlyTable(doc: Document, table: Element) {
  const copy = doc.cloneNode(true) as Document, destination = body(copy)
  while (destination.firstChild) destination.removeChild(destination.firstChild)
  destination.appendChild(copy.importNode(table, true))
  return strToU8(xml(copy))
}
function references(doc: Document) {
  return all(doc).flatMap(n => Array.from(n.attributes).filter(a => a.namespaceURI === O || a.namespaceURI === "urn:schemas-microsoft-com:office:office" && a.localName === "relid").map(a => a.value))
}
function closure(parts: Record<string, Uint8Array>, main: Uint8Array) {
  require(parts["[Content_Types].xml"], "content_types_unavailable")
  const types = parse(parts["[Content_Types].xml"]!)
  require(types.documentElement.namespaceURI === CT && types.documentElement.localName === "Types", "invalid_content_types")
  const typeFor = (name: string) => {
    if (name.endsWith(".rels")) return REL_TYPE
    const override = kids(types.documentElement).filter(n => n.namespaceURI === CT && n.localName === "Override" && n.getAttribute("PartName") === `/${name}`)
    const defaults = kids(types.documentElement).filter(n => n.namespaceURI === CT && n.localName === "Default" && n.getAttribute("Extension") === name.split(".").at(-1))
    require(override.length <= 1 && defaults.length <= 1, "ambiguous_content_type")
    const type = (override[0] ?? defaults[0])?.getAttribute("ContentType")
    require(type && type.length <= 256, "content_type_unavailable")
    return type
  }
  require(typeFor(MAIN) === "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "invalid_main_content_type")
  const selected: Record<string, Uint8Array> = Object.create(null), queue = [MAIN], queued = new Set(queue)
  let bytes = 0
  const add = (name: string, data: Uint8Array) => {
    require(Object.keys(selected).length < 64 && (bytes += data.byteLength) <= MAX_WORD_AMENDMENT_OOXML_BYTES, "dependency_closure_large")
    selected[name] = data
  }
  while (queue.length) {
    const name = queue.shift()!
    if (selected[name]) continue
    require(name === path.posix.normalize(name) && !name.startsWith("/") && !name.startsWith("../") && !/[\\\x00?#%:]/.test(name) && parts[name], "dependency_unavailable")
    const data = name === MAIN ? main : parts[name]!
    add(name, data)
    const relName = path.posix.join(path.posix.dirname(name), "_rels", `${path.posix.basename(name)}.rels`)
    const retained = new Set<string>()
    if (parts[relName]) {
      const rels = parse(parts[relName]!)
      require(rels.documentElement.namespaceURI === REL && rels.documentElement.localName === "Relationships", "invalid_relationships")
      const ids = new Set<string>()
      for (const rel of kids(rels.documentElement)) {
        require(rel.namespaceURI === REL && rel.localName === "Relationship", "invalid_relationships")
        const id = rel.getAttribute("Id"), type = rel.getAttribute("Type") ?? "", kind = type.slice(O.length + 1)
        require(id && !ids.has(id), "ambiguous_relationship")
        ids.add(id)
        if (!type.startsWith(`${O}/`) || !formatting.has(kind) && !(name !== MAIN && kind === "image")) { rel.parentNode!.removeChild(rel); continue }
        require(!rel.getAttribute("TargetMode") || rel.getAttribute("TargetMode") === "Internal", "external_dependency_unsupported")
        const target = rel.getAttribute("Target") ?? ""
        require(target && !/[\\\x00-\x1f?#%:]/.test(target), "invalid_dependency_target")
        const resolved = path.posix.resolve("/", path.posix.dirname(name), target).slice(1)
        require(parts[resolved] && resolved !== MAIN && resolved !== "[Content_Types].xml" && !resolved.endsWith(".rels"), "dependency_unavailable")
        require(formattingType(kind, typeFor(resolved)), "dependency_content_type_unsupported")
        retained.add(id)
        if (!queued.has(resolved)) { require(queued.size < 64, "dependency_closure_large"); queued.add(resolved); queue.push(resolved) }
      }
      if (retained.size) add(relName, strToU8(xml(rels)))
    }
    const type = typeFor(name)
    if (type.endsWith("+xml") || type === "application/xml") require(references(parse(data)).every(id => retained.has(id)), "unsupported_or_dangling_relationship")
  }
  add("_rels/.rels", strToU8(`<Relationships xmlns="${REL}"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`))
  return { selected, typeFor }
}
function flatOpc(parts: Record<string, Uint8Array>, typeFor: (name: string) => string) {
  const flat = parse(`<pkg:package xmlns:pkg="${PKG}"/>`)
  for (const [name, data] of Object.entries(parts)) {
    const type = typeFor(name), part = flat.createElementNS(PKG, "pkg:part")
    part.setAttributeNS(PKG, "pkg:name", `/${name}`); part.setAttributeNS(PKG, "pkg:contentType", type)
    const isXml = type.endsWith("+xml") || type === "application/xml", content = flat.createElementNS(PKG, isXml ? "pkg:xmlData" : "pkg:binaryData")
    if (isXml) content.appendChild(flat.importNode(parse(data).documentElement, true))
    else content.textContent = Buffer.from(data).toString("base64")
    part.appendChild(content); flat.documentElement.appendChild(part)
  }
  const result = xml(flat)
  require(strToU8(result).byteLength <= MAX_WORD_AMENDMENT_OOXML_BYTES, "dependency_closure_large")
  return result
}

/** Pure preparation only: callers must bind a fresh complete parent range and
 * compare its table AND dependency preimage before any native dispatch. The
 * tableIndex is one-based, matching the existing source-bound table targets. */
export async function prepareDocxTableRowDeletionImport(sourceBytes: Uint8Array, authoredBytes: Uint8Array, request: DocxTableRowDeletion) {
  const source = new Uint8Array(sourceBytes), authored = new Uint8Array(authoredBytes), input = { ...request }
  const verified = await verifyDocxTableRowDeletion(source, authored, input)
  const packet = await docxTableRowDeletionTargets(source), target = packet.targets.find(t => t.target_handle === input.target_handle)
  require(target, "target_unavailable")
  const beforeParts = unzipSync(source), afterParts = unzipSync(authored), before = parse(beforeParts[MAIN]!), after = parse(afterParts[MAIN]!)
  const beforeTable = completeTable(before, target.table), table = completeTable(after, target.table)
  const { selected, typeFor } = closure(beforeParts, onlyTable(before, beforeTable))
  const authoredMain = onlyTable(after, table)
  // The full-artifact verifier already proves every non-settings dependency is
  // unchanged. Both preimage and result deliberately use this SAME closure.
  return { beforeOoxml: flatOpc(selected, typeFor), ooxml: flatOpc({ ...selected, [MAIN]: authoredMain }, typeFor),
    beforeTableXml: standalone(beforeTable), tableIndex: target.table, revisions: verified.revisions }
}
