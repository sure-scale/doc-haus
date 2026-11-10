import { DOMParser } from "@xmldom/xmldom"
import { strFromU8 } from "fflate"
import type { WordChatTableRowDeletionProposal } from "../../../dochaus/lib/word-chat-amendment-contract"
import { sameTableImportXml } from "./word-chat-table-import-xml"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage", XMLNS = "http://www.w3.org/2000/xmlns/"
const MAIN = "word/document.xml", REL_TYPE = "application/vnd.openxmlformats-package.relationships+xml"
const LIMIT = 1024 * 1024, formatting = new Set(["styles", "theme", "fontTable", "numbering", "font"])
type Parts = Record<string, Uint8Array>
type Expected = { type: string; data: Element | Uint8Array }
const fail = (reason: string): never => { throw Error(`word_table_import_artifact:${reason}`) }
function require(value: unknown, reason: string): asserts value { if (!value) fail(reason) }
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const named = (node: Node, local: string) => kids(node).filter(n => n.namespaceURI === W && n.localName === local)
const all = (node: Element) => [node, ...Array.from(node.getElementsByTagName("*"))]
const whitespace = (node: Node) => node.nodeType === 3 && !node.nodeValue?.trim()
const isXml = (type: string) => type.endsWith("+xml") || type === "application/xml"
const equalBytes = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((byte, index) => byte === b[index])
function parse(value: string | Uint8Array) {
  const text = typeof value === "string" ? value : strFromU8(value)
  require(!/<!DOCTYPE|<!ENTITY/i.test(text), "unsupported_xml")
  const invalid = () => fail("invalid_xml")
  const doc = new DOMParser({ errorHandler: { warning: invalid, error: invalid, fatalError: invalid } }).parseFromString(text, "application/xml")
  require(doc.documentElement && !doc.doctype, "invalid_xml")
  return doc.documentElement
}
function tableDocument(parts: Parts, index: number) {
  require(parts[MAIN], "main_missing")
  const root = parse(parts[MAIN]!), bodies = named(root, "body")
  require(root.namespaceURI === W && root.localName === "document" && bodies.length === 1 && kids(root).length === 1, "document_boundary")
  const table = named(bodies[0]!, "tbl")[index]
  require(table && !table.getElementsByTagNameNS(W, "tbl").length && !table.getElementsByTagNameNS(W, "sectPr").length, "table_boundary")
  const copy = root.cloneNode(true) as Element, body = named(copy, "body")[0]!
  while (body.firstChild) body.removeChild(body.firstChild)
  body.appendChild(table.cloneNode(true))
  return copy
}
function kindType(kind: string, type: string) {
  if (kind === "theme") return type === "application/vnd.openxmlformats-officedocument.theme+xml"
  if (kind === "font") return ["application/vnd.openxmlformats-officedocument.obfuscatedFont", "application/x-font-ttf", "application/vnd.ms-opentype"].includes(type)
  if (kind === "image") return type.startsWith("image/")
  return type === `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`
}
function resolve(from: string, target: string) {
  require(target && !/[\\\x00-\x1f?#%:]/.test(target), "dependency_target")
  const segments = target.startsWith("/") ? [] : from.split("/").slice(0, -1)
  for (const segment of target.split("/")) {
    if (segment === "..") segments.pop()
    else if (segment && segment !== ".") segments.push(segment)
  }
  return segments.join("/")
}
function expectedClosure(parts: Parts, main: Element) {
  require(parts["[Content_Types].xml"], "types_missing")
  const types = parse(parts["[Content_Types].xml"]!)
  require(types.namespaceURI === CT && types.localName === "Types", "types_invalid")
  const typeFor = (name: string) => {
    if (name.endsWith(".rels")) return REL_TYPE
    const definitions = kids(types).filter(n => n.namespaceURI === CT)
    const overrides = definitions.filter(n => n.localName === "Override" && n.getAttribute("PartName") === `/${name}`)
    const defaults = definitions.filter(n => n.localName === "Default" && n.getAttribute("Extension") === name.split(".").at(-1))
    require(overrides.length <= 1 && defaults.length <= 1, "types_ambiguous")
    const type = (overrides[0] ?? defaults[0])?.getAttribute("ContentType")
    require(type && type.length <= 256, "type_missing")
    return type
  }
  require(typeFor(MAIN) === "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "main_type")
  const selected = new Map<string, Expected>(), queue = [MAIN], queued = new Set(queue)
  const add = (name: string, type: string, data: Element | Uint8Array) => {
    require(!selected.has(name) && selected.size < 64, "closure_large")
    selected.set(name, { type, data })
  }
  while (queue.length) {
    const name = queue.shift()!
    if (selected.has(name)) continue
    require(parts[name] && name === resolve(MAIN, `/${name}`) && !/[\\\x00?#%:]/.test(name), "dependency_missing")
    const type = typeFor(name), data = name === MAIN ? main : isXml(type) ? parse(parts[name]!) : parts[name]!
    add(name, type, data)
    const slash = name.lastIndexOf("/"), relName = `${name.slice(0, slash + 1)}_rels/${name.slice(slash + 1)}.rels`, retained = new Set<string>()
    if (parts[relName]) {
      const rels = parse(parts[relName]!)
      require(rels.namespaceURI === REL && rels.localName === "Relationships", "relationships_invalid")
      const ids = new Set<string>()
      for (const rel of kids(rels)) {
        require(rel.namespaceURI === REL && rel.localName === "Relationship", "relationships_invalid")
        const id = rel.getAttribute("Id"), linkType = rel.getAttribute("Type") ?? "", kind = linkType.slice(O.length + 1)
        require(id && !ids.has(id), "relationship_ambiguous"); ids.add(id)
        if (!linkType.startsWith(`${O}/`) || !formatting.has(kind) && !(name !== MAIN && kind === "image")) { rels.removeChild(rel); continue }
        require(!rel.getAttribute("TargetMode") || rel.getAttribute("TargetMode") === "Internal", "external_dependency")
        const target = resolve(name, rel.getAttribute("Target") ?? "")
        require(parts[target] && target !== MAIN && target !== "[Content_Types].xml" && !target.endsWith(".rels"), "dependency_missing")
        require(kindType(kind, typeFor(target)), "dependency_type")
        retained.add(id)
        if (!queued.has(target)) { require(queued.size < 64, "closure_large"); queued.add(target); queue.push(target) }
      }
      if (retained.size) add(relName, REL_TYPE, rels)
    }
    if (!(data instanceof Uint8Array)) {
      const references = all(data).flatMap(n => Array.from(n.attributes).filter(a => a.namespaceURI === O || a.namespaceURI === "urn:schemas-microsoft-com:office:office" && a.localName === "relid").map(a => a.value))
      require(references.every(id => retained.has(id)), "relationship_dangling")
    }
  }
  add("_rels/.rels", REL_TYPE, parse(`<Relationships xmlns="${REL}"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`))
  return selected
}
function payload(ooxml: string) {
  const root = parse(ooxml)
  require(root.namespaceURI === PKG && root.localName === "package" && Array.from(root.attributes).every(a => a.namespaceURI === XMLNS), "package_invalid")
  const parts = new Map<string, Expected>()
  require(Array.from(root.childNodes).every(n => n.nodeType === 1 || whitespace(n)), "package_content")
  for (const part of kids(root)) {
    require(part.namespaceURI === PKG && part.localName === "part" && parts.size < 64, "part_invalid")
    const name = part.getAttributeNS(PKG, "name"), type = part.getAttributeNS(PKG, "contentType")
    require(name && name.startsWith("/") && type && type.length <= 256 && !parts.has(name.slice(1)), "part_invalid")
    require(Array.from(part.attributes).every(a => a.namespaceURI === XMLNS || a.namespaceURI === PKG && ["name", "contentType"].includes(a.localName)), "part_attributes")
    const data = kids(part)
    require(data.length === 1 && Array.from(part.childNodes).every(n => n.nodeType === 1 || whitespace(n)), "part_content")
    const content = data[0]!
    require(content.namespaceURI === PKG && Array.from(content.attributes).every(a => a.namespaceURI === XMLNS), "part_data")
    let value: Element | Uint8Array
    if (isXml(type)) {
      const roots = kids(content)
      require(content.localName === "xmlData" && roots.length === 1 && Array.from(content.childNodes).every(n => n.nodeType === 1 || whitespace(n)), "part_xml")
      value = roots[0]!
    } else {
      require(content.localName === "binaryData" && Array.from(content.childNodes).every(n => n.nodeType === 3 || n.nodeType === 4), "part_binary")
      const encoded = (content.textContent ?? "").replace(/\s/g, "")
      require(/^(?:[a-zA-Z0-9+/]{4})*(?:[a-zA-Z0-9+/]{2}==|[a-zA-Z0-9+/]{3}=)?$/.test(encoded), "part_binary")
      try { value = Uint8Array.from(atob(encoded), char => char.charCodeAt(0)) } catch { return fail("part_binary") }
    }
    parts.set(name.slice(1), { type, data: value })
  }
  return parts
}
function samePart(expected: Expected, actual: Expected, name: string) {
  require(expected.type === actual.type, "part_type_changed")
  if (expected.data instanceof Uint8Array) require(actual.data instanceof Uint8Array && equalBytes(expected.data, actual.data), "part_bytes_changed")
  else {
    require(!(actual.data instanceof Uint8Array), "part_kind_changed")
    if (name === "_rels/.rels") {
      // The minimal transport root is generated, not the original package's
      // root inventory. Its single unreferenced relationship ID is immaterial.
      const links = kids(actual.data)
      require(actual.data.namespaceURI === REL && actual.data.localName === "Relationships" && links.length === 1 && links[0]!.getAttribute("Id"), "root_relationships")
      const copy = actual.data.cloneNode(true) as Element
      kids(copy)[0]!.setAttribute("Id", "main")
      require(sameTableImportXml(expected.data, copy), "root_relationships")
    } else require(sameTableImportXml(expected.data, actual.data), "part_xml_changed")
  }
}

/** Called only AFTER the full source/authored ZIP oracle proves the reviewed
 * narrow deletion. This independently binds its separately supplied Flat OPC
 * pair; it never grants authority from visible text or caller revision IDs. */
export function verifyWordChatTableImportArtifact(sourceParts: Parts, authoredParts: Parts, pair: { beforeOoxml: string; ooxml: string }, proposal: WordChatTableRowDeletionProposal): void {
  require(pair && typeof pair.beforeOoxml === "string" && typeof pair.ooxml === "string", "pair_missing")
  require(new TextEncoder().encode(pair.beforeOoxml).byteLength + new TextEncoder().encode(pair.ooxml).byteLength <= LIMIT * 2, "pair_large")
  const index = proposal.table.tableIndex
  require(Number.isSafeInteger(index) && index >= 0, "table_index")
  const before = tableDocument(sourceParts, index), after = tableDocument(authoredParts, index), expected = expectedClosure(sourceParts, before)
  for (const [ooxml, main] of [[pair.beforeOoxml, before], [pair.ooxml, after]] as const) {
    const actual = payload(ooxml)
    require(actual.size === expected.size && [...expected.keys()].every(name => actual.has(name)), "part_set_changed")
    for (const [name, part] of expected) samePart(name === MAIN ? { ...part, data: main } : part, actual.get(name)!, name)
  }
}
