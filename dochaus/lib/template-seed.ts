import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// The render seed a template or precedent .docx gives (#202): its styles,
// numbering definitions, theme, fonts and page set-up, with one empty paragraph
// for the render to insert into. Its text and everything that carries content
// rather than look are left out: headers and footers (the firm's page
// furniture applies instead, and a precedent's furniture names its own
// parties), comments, notes, custom XML, document properties, glossary and
// macros. Settings that would lock, track or re-link the new document are
// removed. Styles the render writes by id and the template lacks are added from
// the house seed, based on the template's own default paragraph style.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
const RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
const BODY = "word/document.xml"
const DOCUMENT_RELS = "word/_rels/document.xml.rels"
const MAIN_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
// The document relationships that carry a template's look.
const KEPT = new Set(["styles", "numbering", "theme", "fontTable", "settings", "webSettings", "stylesWithEffects"].map(kind => RELATIONSHIP + kind))
// Styles Docxodus and the render write by id.
const RENDER_STYLES = ["Title", "Heading1", "Heading2", "Heading3", "ListParagraph", "Quote", "Header", "Footer"]
// Settings that would protect the document, track its edits, attach it to the
// template file or a data source, or carry the template's document variables;
// note settings name separators of the notes parts left out.
const DROPPED_SETTINGS = new Set(["attachedTemplate", "documentProtection", "writeProtection", "trackRevisions", "mailMerge", "docVars", "footnotePr", "endnotePr"])

/** The render seed `template` gives, or undefined when it is not a Word
 * document with styles of its own. `house` is the house seed, whose styles
 * fill in any the render needs that the template lacks. */
export function templateSeed(template: Uint8Array, house: Uint8Array): Uint8Array | undefined {
  try {
    return seed(unzipSync(template), unzipSync(house))
  } catch {
    return undefined
  }
}

function seed(parts: Record<string, Uint8Array>, house: Record<string, Uint8Array>) {
  if (!parts[BODY] || !parts[DOCUMENT_RELS] || !parts["[Content_Types].xml"]) return
  const relationships = parse(parts[DOCUMENT_RELS])
  const kept = new Map<string, Uint8Array>()
  // The kept parts and, but for settings, the parts they relate to in turn
  // (embedded fonts, picture bullets, theme images).
  const keep = (name: string, follow: boolean) => {
    if (kept.has(name) || !parts[name]) return
    kept.set(name, parts[name])
    const rels = relsName(name)
    if (!follow || !parts[rels]) return
    kept.set(rels, parts[rels])
    for (const relationship of elements(parse(parts[rels]), PACKAGE_RELS, "Relationship"))
      if (relationship.getAttribute("TargetMode") !== "External") keep(resolve(name, relationship.getAttribute("Target") ?? ""), true)
  }
  let styles: string | undefined
  let settings: string | undefined
  for (const relationship of elements(relationships, PACKAGE_RELS, "Relationship")) {
    const type = relationship.getAttribute("Type") ?? ""
    const target = resolve(BODY, relationship.getAttribute("Target") ?? "")
    if (!KEPT.has(type) || relationship.getAttribute("TargetMode") === "External" || !parts[target]) {
      relationship.parentNode!.removeChild(relationship)
      continue
    }
    if (type === RELATIONSHIP + "styles") styles = target
    if (type === RELATIONSHIP + "settings") settings = target
    keep(target, type !== RELATIONSHIP + "settings")
  }
  if (!styles) return

  kept.set(BODY, emptyBody(parse(parts[BODY])))
  kept.set(DOCUMENT_RELS, serialize(relationships))
  kept.set(styles, withRenderStyles(parse(parts[styles]), house["word/styles.xml"]))
  if (settings) kept.set(settings, sanitizedSettings(parse(parts[settings])))
  kept.set("_rels/.rels", strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELATIONSHIP}officeDocument" Target="${BODY}"/></Relationships>`))
  kept.set("[Content_Types].xml", contentTypes(parse(parts["[Content_Types].xml"]), kept))
  return zipSync(Object.fromEntries(kept), { mtime: new Date(1980, 0, 1) })
}

// One empty paragraph, then the template's last section properties without
// its header and footer references and its distinct first page, which the
// page furniture decides.
function emptyBody(document: Document) {
  const body = elements(document, W, "body")[0]
  if (!body) throw new Error("template_seed: no body")
  const section = children(body).findLast(node => node.localName === "sectPr")
  for (const node of Array.from(body.childNodes)) body.removeChild(node)
  body.appendChild(document.createElementNS(W, "w:p"))
  if (section) {
    for (const node of children(section))
      if (["headerReference", "footerReference", "titlePg"].includes(node.localName)) section.removeChild(node)
    body.appendChild(section)
  }
  return serialize(document)
}

function withRenderStyles(styles: Document, houseStyles: Uint8Array | undefined) {
  const root = styles.documentElement
  const own = children(root).filter(node => node.localName === "style")
  const ids = new Set(own.map(node => node.getAttributeNS(W, "styleId")))
  const names = new Set(own.map(node => attribute(children(node).find(entry => entry.localName === "name"))?.toLowerCase()))
  const normal = own.find(node => node.getAttributeNS(W, "type") === "paragraph" && ["1", "true", "on"].includes(node.getAttributeNS(W, "default") ?? ""))?.getAttributeNS(W, "styleId")
  if (houseStyles) for (const style of children(parse(houseStyles).documentElement)) {
    const id = style.getAttributeNS(W, "styleId") ?? ""
    const name = attribute(children(style).find(entry => entry.localName === "name"))?.toLowerCase()
    if (!RENDER_STYLES.includes(id) || ids.has(id) || names.has(name)) continue
    const copy = styles.importNode(style, true) as Element
    const based = children(copy).find(entry => entry.localName === "basedOn")
    if (based && normal) based.setAttributeNS(W, "w:val", normal)
    else if (based) copy.removeChild(based)
    root.appendChild(copy)
  }
  return serialize(styles)
}

function sanitizedSettings(settings: Document) {
  for (const node of children(settings.documentElement))
    if (node.namespaceURI === W && DROPPED_SETTINGS.has(node.localName)) settings.documentElement.removeChild(node)
  return serialize(settings)
}

function contentTypes(types: Document, kept: Map<string, Uint8Array>) {
  for (const override of elements(types, CONTENT_TYPES_NS, "Override")) {
    const name = (override.getAttribute("PartName") ?? "").replace(/^\//, "")
    if (!kept.has(name)) override.parentNode!.removeChild(override)
    else if (name === BODY) override.setAttribute("ContentType", MAIN_TYPE)
  }
  return serialize(types)
}

function relsName(part: string) {
  const at = part.lastIndexOf("/")
  return `${part.slice(0, at + 1)}_rels/${part.slice(at + 1)}.rels`
}

function resolve(from: string, target: string) {
  const segments = target.startsWith("/") ? [] : from.split("/").slice(0, -1)
  for (const segment of target.replace(/^\//, "").split("/")) {
    if (segment === "..") segments.pop()
    else if (segment && segment !== ".") segments.push(segment)
  }
  return segments.join("/")
}

function parse(bytes: Uint8Array) {
  const fail = (message: string) => { throw new Error(`template_seed: ${message}`) }
  return new DOMParser({ errorHandler: { warning: () => {}, error: fail, fatalError: fail } }).parseFromString(strFromU8(bytes), "application/xml")
}

function serialize(document: Document) {
  return strToU8(new XMLSerializer().serializeToString(document))
}

function elements(node: Document | Element, namespace: string, name: string) {
  return Array.from(node.getElementsByTagNameNS(namespace, name))
}

function children(node: Element) {
  return Array.from(node.childNodes).filter((entry): entry is Element => entry.nodeType === 1)
}

function attribute(node: Element | undefined) {
  return node?.getAttributeNS(W, "val") ?? undefined
}
