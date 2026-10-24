import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// The render seed a template or precedent .docx gives (#202): its styles,
// numbering definitions, theme, fonts and page set-up, with one empty paragraph
// for the render to insert into. Its text and everything that carries content
// rather than look are left out: headers and footers (the firm's page
// furniture applies instead, and a precedent's furniture names its own
// parties), comments, notes, custom XML, document properties, glossary and
// macros. Settings that would lock, track or re-link the new document, or
// carry the template's edit history, are removed, as is anything left pointing
// at a part not carried over. The page set-up is the template's main body, the
// section with the most paragraphs (a precedent's cover page does not set the
// draft's margins and page numbers, nor a closing landscape schedule turn it).
// Styles the render writes by id and the template lacks are added from the
// house seed, based on the template's own default paragraph style; a template
// that has the style under another id (Word translates built-in style ids, not
// their names: "berschrift1" is "heading 1") keeps its own, and
// withTemplateStyleIds points the rendered paragraphs at it.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
const RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
const BODY = "word/document.xml"
const DOCUMENT_RELS = "word/_rels/document.xml.rels"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const MAIN_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
// The document relationships that carry a template's look.
const KEPT = new Set(["styles", "numbering", "theme", "fontTable", "settings", "webSettings", "stylesWithEffects"].map(kind => RELATIONSHIP + kind))
// Styles Docxodus and the render write by id.
const RENDER_STYLES = ["Title", "Heading1", "Heading2", "Heading3", "ListParagraph", "Quote", "Header", "Footer"]
// The built-in name of each style id Docxodus writes.
const BUILT_IN_NAMES = new Map([
  ["Title", "title"], ["ListParagraph", "list paragraph"], ["Quote", "quote"], ["Header", "header"], ["Footer", "footer"],
  ...[1, 2, 3, 4, 5, 6].map(level => [`Heading${level}`, `heading ${level}`] as [string, string]),
])
// Settings that would protect the document, track its edits, attach it to the
// template file or a data source, carry the template's document variables or
// edit sessions (rsids), or ask for even-page headers the page furniture does
// not write; note settings name separators of the notes parts left out.
const DROPPED_SETTINGS = new Set(["attachedTemplate", "documentProtection", "writeProtection", "trackRevisions", "mailMerge", "docVars", "footnotePr", "endnotePr", "evenAndOddHeaders", "rsids"])

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

  const ids = new Set(elements(relationships, PACKAGE_RELS, "Relationship").map(relationship => relationship.getAttribute("Id")))
  kept.set(BODY, emptyBody(parse(parts[BODY]), ids))
  kept.set(DOCUMENT_RELS, serialize(relationships))
  kept.set(styles, withRenderStyles(parse(parts[styles]), house["word/styles.xml"]))
  // The settings' own relationships (an attached template, a data source) are not kept.
  if (settings) kept.set(settings, sanitizedSettings(parse(parts[settings])))
  kept.set("_rels/.rels", strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${RELATIONSHIP}officeDocument" Target="${BODY}"/></Relationships>`))
  kept.set("[Content_Types].xml", contentTypes(parse(parts["[Content_Types].xml"]), kept))
  return zipSync(Object.fromEntries(kept), { mtime: new Date(1980, 0, 1) })
}

// One empty paragraph, then the main body's section properties without its
// header and footer references, distinct first page and page numbering, which
// the page furniture decides, and without its vertical alignment and section
// break, which belong to its place in the template. Whatever still points at a
// relationship not kept (the printer settings, a page background picture) goes.
function emptyBody(document: Document, relationships: Set<string | null>) {
  const body = elements(document, W, "body")[0]
  if (!body) throw new Error("template_seed: no body")
  const section = mainSection(body)
  for (const node of Array.from(body.childNodes)) body.removeChild(node)
  body.appendChild(document.createElementNS(W, "w:p"))
  if (section) {
    for (const node of children(section))
      if (["headerReference", "footerReference", "titlePg", "pgNumType", "vAlign", "type"].includes(node.localName)) section.removeChild(node)
    body.appendChild(section)
  }
  withoutDanglingReferences(document.documentElement, relationships)
  return serialize(document)
}

// The properties of the section with the most paragraphs, the first of equals.
// A section's properties close it: in its last paragraph, or the body's own
// for the last section.
function mainSection(body: Element) {
  let main: { section: Element; paragraphs: number } | undefined
  let paragraphs = 0
  for (const node of children(body)) {
    const properties = node.localName === "p" ? children(node).find(entry => entry.localName === "pPr") : undefined
    const section = node.localName === "sectPr" ? node : properties && children(properties).find(entry => entry.localName === "sectPr")
    if (node.localName !== "sectPr") paragraphs += node.localName === "p" ? 1 : elements(node, W, "p").length
    if (!section) continue
    if (!main || paragraphs > main.paragraphs) main = { section, paragraphs }
    paragraphs = 0
  }
  return main?.section
}

function withoutDanglingReferences(root: Element, relationships: Set<string | null>) {
  for (const node of Array.from(root.getElementsByTagName("*")))
    if (node.parentNode && Array.from(node.attributes).some(entry => entry.namespaceURI === R && !relationships.has(entry.value))) node.parentNode.removeChild(node)
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
  withoutDanglingReferences(settings.documentElement, new Set())
  return serialize(settings)
}

/** Points a rendered document's paragraphs at the template's own style where
 * the render wrote a built-in style id (Heading1, Title) that the template
 * keeps under a translated id of the same built-in name. Unchanged bytes when
 * nothing is renamed. */
export function withTemplateStyleIds(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  if (!parts["word/styles.xml"] || !parts[BODY]) return bytes
  const styles = children(parse(parts["word/styles.xml"]).documentElement).filter(node => node.localName === "style" && node.getAttributeNS(W, "type") === "paragraph")
  const ids = new Set(styles.map(node => node.getAttributeNS(W, "styleId")))
  const renamed = new Map<string, string>()
  for (const [id, name] of BUILT_IN_NAMES) {
    if (ids.has(id)) continue
    const own = styles.find(node => attribute(children(node).find(entry => entry.localName === "name"))?.toLowerCase() === name)?.getAttributeNS(W, "styleId")
    if (own) renamed.set(id, own)
  }
  if (!renamed.size) return bytes
  const document = parse(parts[BODY])
  let changed = false
  for (const style of elements(document, W, "pStyle")) {
    const own = renamed.get(attribute(style) ?? "")
    if (!own) continue
    style.setAttributeNS(W, "w:val", own)
    changed = true
  }
  return changed ? zipSync({ ...parts, [BODY]: serialize(document) }, { mtime: new Date(1980, 0, 1) }) : bytes
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
