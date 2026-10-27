import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, zipSync } from "fflate"
import { parseZip } from "./zip"
import { WORD_ADDIN_ID, WORD_ADDIN_VERSION, WORD_LINK_SETTING, type WordDocumentLink } from "./word-addin-contract"

const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const WE = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
const TP = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
const TP_REL = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
const WE_REL = "http://schemas.microsoft.com/office/2011/relationships/webextension"
const extensionPart = "word/webextensions/dochaus.xml"
const taskpanePart = "word/webextensions/taskpanes.xml"
const taskpaneRels = "word/webextensions/_rels/taskpanes.xml.rels"
export const MAX_WORD_UPLOAD_BYTES = 128 * 1024 * 1024
const MAX_UNPACKED = 256 * 1024 * 1024
const MAX_PARTS = 1000

export class WordPackageError extends Error {
  constructor(message: string, public readonly code = "invalid_docx", public readonly status: 400 | 413 = 400) { super(message) }
}

function xml(bytes: Uint8Array | string) {
  const text = typeof bytes === "string" ? bytes : strFromU8(bytes)
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new WordPackageError("Word XML cannot contain document type or entity declarations.")
  const fail = () => { throw new WordPackageError("The Word file contains malformed XML.") }
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(text, "application/xml")
  if (!doc.documentElement || doc.getElementsByTagName("parsererror").length) fail()
  return doc
}
const serialize = (doc: Document) => strToU8(new XMLSerializer().serializeToString(doc))
const elements = (doc: Document | Element, namespace: string, name: string) => Array.from(doc.getElementsByTagNameNS(namespace, name))
const remove = (node: Node) => node.parentNode?.removeChild(node)
function append(doc: Document, namespace: string, name: string, attrs: Record<string, string>) {
  const node = doc.createElementNS(namespace, name)
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value)
  doc.documentElement.appendChild(node)
  return node
}
function relationships(bytes?: Uint8Array) {
  return xml(bytes ?? `<Relationships xmlns="${REL}"/>`)
}
function uniqueId(doc: Document) {
  const ids = new Set(elements(doc, REL, "Relationship").map(node => node.getAttribute("Id")))
  let index = 1
  while (ids.has(`rIdDocHaus${index}`)) index++
  return `rIdDocHaus${index}`
}
function targetPart(source: string, target: string) {
  // OOXML relationships use URI paths. External or escaping targets never select a package part.
  if (/^[a-z]+:/i.test(target) || /[\\\0]/.test(target)) return ""
  const segments = (target.startsWith("/") ? target.slice(1) : `${source.slice(0, source.lastIndexOf("/") + 1)}${target}`).split("/")
  const out: string[] = []
  for (const segment of segments) {
    if (segment === "..") { if (!out.length) return ""; out.pop() }
    else if (segment && segment !== ".") out.push(segment)
  }
  return out.join("/")
}
function relPart(part: string) {
  const index = part.lastIndexOf("/")
  return `${part.slice(0, index + 1)}_rels/${part.slice(index + 1)}.rels`
}

function readPackage(bytes: Uint8Array) {
  if (bytes.byteLength > MAX_WORD_UPLOAD_BYTES) throw new WordPackageError("The Word file exceeds the 128 MB upload limit.", "word_too_large", 413)
  try {
    const archive = parseZip(bytes)
    if (archive.uncompressed > MAX_UNPACKED || archive.entries.length > MAX_PARTS)
      throw new WordPackageError("The Word file exceeds the package size limit.", "word_too_large", 413)
    const parts: Record<string, Uint8Array> = Object.create(null)
    for (const entry of archive.entries) {
      if (entry.encrypted || entry.unsafePath || /[\\\0]/.test(entry.filename) || entry.filename.split("/").some(segment => !segment || segment === ".") || parts[entry.filename])
        throw new WordPackageError("The Word file contains an unsafe, encrypted or duplicate package path.")
      const content = entry.content()
      if (/\.(xml|rels)$/i.test(entry.filename)) xml(content)
      parts[entry.filename] = content
    }
    if (!parts["word/document.xml"] || !parts["[Content_Types].xml"] || !parts["_rels/.rels"])
      throw new WordPackageError("The upload must be a readable Word document.")
    const document = xml(parts["word/document.xml"])
    if (document.documentElement.localName !== "document" || !["http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"].includes(document.documentElement.namespaceURI ?? ""))
      throw new WordPackageError("The upload must be a readable Word document.")
    return parts
  } catch (error) {
    if (error instanceof WordPackageError) throw error
    throw new WordPackageError("The upload must be a readable Word document.")
  }
}
function pack(parts: Record<string, Uint8Array>) {
  // Stable ZIP metadata makes identical retries idempotent after removing linkage.
  return zipSync(Object.fromEntries(Object.entries(parts).map(([name, bytes]) => [name, [bytes, { mtime: new Date("2000-01-01T00:00:00Z") }]])))
}

/** Removes only this add-in, retaining unrelated extensions, taskpanes and document parts. */
function strip(parts: Record<string, Uint8Array>) {
  const owned = new Set<string>()
  for (const [name, bytes] of Object.entries(parts)) {
    if (!name.endsWith(".xml")) continue
    const doc = xml(bytes)
    const root = doc.documentElement
    if (root.namespaceURI !== WE || root.localName !== "webextension") continue
    const reference = elements(doc, WE, "reference")[0]
    if (reference?.getAttribute("id")?.replace(/[{}]/g, "").toLowerCase() === WORD_ADDIN_ID ||
      elements(doc, WE, "property").some(node => node.getAttribute("name") === WORD_LINK_SETTING)) owned.add(name)
  }
  if (!owned.size) return false
  const removed = new Set(owned)
  for (const [name, bytes] of Object.entries(parts)) {
    if (!name.endsWith(".rels")) continue
    const doc = xml(bytes)
    const source = name === "_rels/.rels" ? "" : name.replace(/(^|\/)_rels\//, "$1").replace(/\.rels$/, "")
    const ids = new Set<string>()
    for (const node of elements(doc, REL, "Relationship")) {
      if (!owned.has(targetPart(source, node.getAttribute("Target") ?? ""))) continue
      ids.add(node.getAttribute("Id") ?? "")
      remove(node)
    }
    if (!ids.size) continue
    parts[name] = serialize(doc)
    const paneBytes = parts[source]
    if (!paneBytes) continue
    const pane = xml(paneBytes)
    for (const ref of elements(pane, TP, "webextensionref"))
      if (ids.has(ref.getAttributeNS(R, "id") ?? "")) remove(ref.parentNode!)
    if (pane.documentElement.namespaceURI === TP && !elements(pane, TP, "taskpane").length) {
      removed.add(source)
      removed.add(name)
    } else parts[source] = serialize(pane)
  }
  for (const name of [...owned]) if (parts[relPart(name)]) removed.add(relPart(name))
  for (const name of removed) delete parts[name]
  // Remove references and content-type entries for deleted doc.haus parts only.
  for (const [name, bytes] of Object.entries(parts)) {
    if (!name.endsWith(".rels") && name !== "[Content_Types].xml") continue
    const doc = xml(bytes)
    let changed = false
    const source = name === "_rels/.rels" ? "" : name.replace(/(^|\/)_rels\//, "$1").replace(/\.rels$/, "")
    for (const node of elements(doc, REL, "Relationship")) if (removed.has(targetPart(source, node.getAttribute("Target") ?? ""))) { remove(node); changed = true }
    for (const node of elements(doc, CT, "Override")) if (removed.has(node.getAttribute("PartName")?.replace(/^\//, "") ?? "")) { remove(node); changed = true }
    if (changed) parts[name] = serialize(doc)
  }
  return true
}

export function stripWordLink(bytes: Uint8Array) {
  const parts = readPackage(bytes)
  return strip(parts) ? pack(parts) : bytes
}

export function linkedWordDocx(bytes: Uint8Array, link: WordDocumentLink) {
  const parts = readPackage(bytes)
  strip(parts)
  // Existing taskpanes are retained. Refuse a nonstandard root location rather than overwrite it.
  const root = relationships(parts["_rels/.rels"])
  // A Word export can retain this empty internal relationship after omitting
  // its taskpane parts. It names no add-in; leave every live or external target alone.
  for (const node of elements(root, REL, "Relationship"))
    if (node.getAttribute("Type") === TP_REL && node.getAttribute("Target") === "NULL" && !parts["NULL"] &&
      (!node.hasAttribute("TargetMode") || node.getAttribute("TargetMode") === "Internal")) remove(node)
  const rootPane = elements(root, REL, "Relationship").find(node => node.getAttribute("Type") === TP_REL)
  if (rootPane && targetPart("", rootPane.getAttribute("Target") ?? "") !== taskpanePart)
    throw new WordPackageError("This Word file uses an unsupported taskpane location.")
  if (!rootPane) append(root, REL, "Relationship", { Id: uniqueId(root), Type: TP_REL, Target: taskpanePart })
  parts["_rels/.rels"] = serialize(root)
  const panes = xml(parts[taskpanePart] ?? `<wetp:taskpanes xmlns:wetp="${TP}"/>`)
  const rels = relationships(parts[taskpaneRels])
  const id = uniqueId(rels)
  append(rels, REL, "Relationship", { Id: id, Type: WE_REL, Target: "dochaus.xml" })
  const pane = append(panes, TP, "wetp:taskpane", { dockstate: "right", visibility: "1", width: "360", row: "0" })
  const ref = panes.createElementNS(TP, "wetp:webextensionref")
  ref.setAttributeNS(R, "r:id", id)
  pane.appendChild(ref)
  parts[taskpanePart] = serialize(panes)
  parts[taskpaneRels] = serialize(rels)
  const extension = xml(`<we:webextension xmlns:we="${WE}" id="{${WORD_ADDIN_ID.toUpperCase()}}"><we:reference id="${WORD_ADDIN_ID}" version="${WORD_ADDIN_VERSION}" store="developer" storeType="Registry"/><we:alternateReferences/><we:properties/><we:bindings/><we:snapshot xmlns:r="${R}"/></we:webextension>`)
  const properties = elements(extension, WE, "properties")[0]!
  for (const [name, value] of [["Office.AutoShowTaskpaneWithDocument", "true"], [WORD_LINK_SETTING, JSON.stringify(link)]]) {
    const property = extension.createElementNS(WE, "we:property")
    property.setAttribute("name", name)
    property.setAttribute("value", value)
    properties.appendChild(property)
  }
  if (parts[extensionPart]) throw new WordPackageError("The Word file already uses the doc.haus extension path.")
  parts[extensionPart] = serialize(extension)
  const types = xml(parts["[Content_Types].xml"])
  for (const [part, contentType] of [[taskpanePart, "application/vnd.ms-office.webextensiontaskpanes+xml"], [extensionPart, "application/vnd.ms-office.webextension+xml"]])
    if (!elements(types, CT, "Override").some(node => node.getAttribute("PartName") === `/${part}`))
      append(types, CT, "Override", { PartName: `/${part}`, ContentType: contentType })
  parts["[Content_Types].xml"] = serialize(types)
  return pack(parts)
}

export function wordLastModifiedBy(bytes: Uint8Array) {
  const parts = readPackage(bytes)
  const core = parts["docProps/core.xml"]
  return core ? elements(xml(core), "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "lastModifiedBy")[0]?.textContent?.trim().replace(/[\r\n]+/g, " ").slice(0, 256) || "unknown" : "unknown"
}
