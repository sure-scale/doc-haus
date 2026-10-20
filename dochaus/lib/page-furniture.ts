import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import type { DraftingPreferences } from "./drafting-preferences"

// Page furniture for documents doc.haus renders: page numbers, the draft line
// and the firm's chosen legend. The planner picks only each document's form
// from a closed set; what the form carries follows here from the firm's
// settings. Legends have legal effect, so they are the firm's own text and are
// never written by a model. Uploaded, counterparty and template files are
// never given furniture, and a document that already has a header or footer
// keeps it.

/** A planned document's form, from its outline "Form:" line. */
export type DocumentForm = "agreement" | "advice" | "letter" | "other"
export const DOCUMENT_FORMS: readonly DocumentForm[] = ["agreement", "advice", "letter", "other"]

/** What a rendered file is: a planned document's form, or the run's internal drafting notes. */
export type FurnitureKind = DocumentForm | "notes"

/** `header` lines are written in order at the top of every page; `pageNumber`
 * goes in the footer; `firstPage: "blank"` leaves the first page without either
 * (a letter is numbered from its second page). */
export type PageFurniture = {
  header: string[]
  pageNumber?: { format: "plain" | "page"; align: "center" | "right" }
  firstPage?: "blank"
}

export const NOTES_MARKING = "Internal drafting notes — not for circulation"

type FurniturePreferences = Pick<DraftingPreferences, "firm" | "dateFormat" | "pageNumbers" | "pageNumberPosition" | "draftLine" | "agreementLegend" | "adviceLegend" | "notesMarking">

/** The furniture a kind of file carries under the firm's settings, or
 * undefined when it carries none. `draftedAt` is the time the draft was made
 * (an ISO timestamp kept in run state, never the render clock); without it
 * there is no draft line. */
export function pageFurniture(kind: FurnitureKind, prefs: FurniturePreferences, draftedAt?: string): PageFurniture | undefined {
  const legend = kind === "agreement" ? prefs.agreementLegend : kind === "advice" ? prefs.adviceLegend : ""
  const date = draftedAt && prefs.draftLine && (kind === "agreement" || kind === "advice") ? formatDraftDate(draftedAt, prefs.dateFormat) : undefined
  const header = [
    ...(legend ? [legend] : []),
    ...(date ? [`${prefs.firm.trim() ? `${prefs.firm.trim()} draft` : "Draft"}: ${date}`] : []),
    ...(kind === "notes" && prefs.notesMarking ? [NOTES_MARKING] : []),
  ]
  const pageNumber = prefs.pageNumbers === "off" ? undefined : { format: prefs.pageNumbers, align: prefs.pageNumberPosition }
  if (!header.length && !pageNumber) return undefined
  return { header, ...(pageNumber ? { pageNumber } : {}), ...(kind === "letter" ? { firstPage: "blank" as const } : {}) }
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

/** The draft date in the firm's date style, read in the host's local time. */
export function formatDraftDate(at: string, style: DraftingPreferences["dateFormat"]) {
  const date = new Date(at)
  if (Number.isNaN(date.getTime())) return undefined
  const [year, month, day] = [date.getFullYear(), date.getMonth(), date.getDate()]
  if (style === "iso") return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`
  if (style === "day-month-year") return `${day} ${MONTHS[month]} ${year}`
  return `${MONTHS[month]} ${day}, ${year}`
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const R = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const XML = "http://www.w3.org/XML/1998/namespace"
const BODY = "word/document.xml"
const RELS = "word/_rels/document.xml.rels"
const TYPES = "[Content_Types].xml"
type Row = { text: string; align: "center" | "right"; page?: string }

/** Writes `furniture` into a rendered body's last section as new header and
 * footer parts: the header lines right-aligned, one paragraph each, and the
 * page number in its own footer paragraph, a PAGE field after any "Page ". No
 * NUMPAGES: the combined memo and body file restarts the body at 1, so a page
 * count would count the memo. A first page left blank gets explicit empty
 * first-page parts, so no reader falls back to the others. A document that
 * already has a header or footer anywhere (a template's or an uploaded file's
 * own) is returned unchanged. Pure; fixed zip entry times keep identical input
 * byte-identical. */
export function applyPageFurniture(docx: Uint8Array, furniture: PageFurniture): Uint8Array {
  if (!furniture.header.length && !furniture.pageNumber) return docx
  const parts = unzipSync(docx)
  const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
  const document = parse(parts[BODY]!)
  if (["headerReference", "footerReference"].some(name => document.getElementsByTagNameNS(W, name).length)) return docx
  const body = document.getElementsByTagNameNS(W, "body")[0]!
  const section = elementChildren(body).findLast(node => node.localName === "sectPr") ?? body.appendChild(document.createElementNS(W, "w:sectPr"))
  const relationships = parse(parts[RELS] ?? strToU8(`<Relationships xmlns="${R}"/>`))
  const types = parse(parts[TYPES]!)
  const ids = new Set(Array.from(relationships.getElementsByTagNameNS(R, "Relationship")).map(row => row.getAttribute("Id")))
  const references: Element[] = []
  const blank = furniture.firstPage === "blank"
  const add = (kind: "header" | "footer", type: "default" | "first", rows: Row[]) => {
    let n = 1
    while (parts[`word/${kind}${n}.xml`]) n++
    parts[`word/${kind}${n}.xml`] = storyPart(kind, rows)
    let id = 1
    while (ids.has(`rId${id}`)) id++
    ids.add(`rId${id}`)
    const relationship = relationships.createElementNS(R, "Relationship")
    relationship.setAttribute("Id", `rId${id}`)
    relationship.setAttribute("Type", `${O}/${kind}`)
    relationship.setAttribute("Target", `${kind}${n}.xml`)
    relationships.documentElement.appendChild(relationship)
    const override = types.createElementNS(CT, "Override")
    override.setAttribute("PartName", `/word/${kind}${n}.xml`)
    override.setAttribute("ContentType", `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`)
    types.documentElement.appendChild(override)
    const reference = document.createElementNS(W, `w:${kind}Reference`)
    reference.setAttributeNS(W, "w:type", type)
    reference.setAttributeNS(O, "r:id", `rId${id}`)
    references.push(reference)
  }
  if (furniture.header.length) {
    add("header", "default", furniture.header.map(text => ({ text, align: "right" })))
    if (blank) add("header", "first", [{ text: "", align: "right" }])
  }
  if (furniture.pageNumber) {
    const { format, align } = furniture.pageNumber
    // The cached result is what a reader that does not lay out pages shows: the
    // number of the first page this footer appears on.
    add("footer", "default", [{ text: format === "page" ? "Page " : "", align, page: blank ? "2" : "1" }])
    if (blank) add("footer", "first", [{ text: "", align }])
  }
  // Header and footer references come first in a section's properties; the
  // title-page switch goes before the elements the schema puts after it.
  const first = section.firstChild
  for (const reference of references) section.insertBefore(reference, first)
  // Declared once on the root, not on every reference.
  if (!document.documentElement.lookupNamespaceURI("r")) document.documentElement.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", O)
  if (blank) {
    const existing = elementChildren(section).find(node => node.localName === "titlePg")
    if (existing) existing.removeAttributeNS(W, "val")
    else section.insertBefore(document.createElementNS(W, "w:titlePg"), elementChildren(section)
      .find(node => ["textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"].includes(node.localName)) ?? null)
  }
  const serializer = new XMLSerializer()
  parts[BODY] = strToU8(serializer.serializeToString(document))
  parts[RELS] = strToU8(serializer.serializeToString(relationships))
  parts[TYPES] = strToU8(serializer.serializeToString(types))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

// One header or footer part, a paragraph per row in the part's style. Text goes
// in as text nodes, never as markup, as one line without the characters XML 1.0
// forbids (a firm name is free text). A row with `page` ends in a PAGE field
// cached as it. Serialized as Docxodus saves a part it has read (its declaration,
// a space before "/>"; xmldom escapes text as it does), so a later engine save,
// as in applying amendments, leaves the part byte-identical.
function storyPart(kind: "header" | "footer", rows: Row[]) {
  const doc = new DOMParser().parseFromString(`<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:${kind === "header" ? "hdr" : "ftr"} xmlns:w="${W}"/>`, "application/xml")
  const make = (parent: Element, name: string, attributes: Record<string, string> = {}) => {
    const node = doc.createElementNS(W, `w:${name}`)
    for (const [key, value] of Object.entries(attributes)) node.setAttributeNS(W, `w:${key}`, value)
    return parent.appendChild(node)
  }
  const text = (paragraph: Element, name: "t" | "instrText", value: string) => {
    const node = make(make(paragraph, "r"), name)
    node.setAttributeNS(XML, "xml:space", "preserve")
    node.appendChild(doc.createTextNode(value.replace(/[^\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu, " ")))
  }
  for (const row of rows) {
    const paragraph = make(doc.documentElement, "p")
    const properties = make(paragraph, "pPr")
    make(properties, "pStyle", { val: kind === "header" ? "Header" : "Footer" })
    make(properties, "jc", { val: row.align })
    if (row.text) text(paragraph, "t", row.text)
    if (row.page === undefined) continue
    make(make(paragraph, "r"), "fldChar", { fldCharType: "begin" })
    text(paragraph, "instrText", " PAGE ")
    make(make(paragraph, "r"), "fldChar", { fldCharType: "separate" })
    text(paragraph, "t", row.page)
    make(make(paragraph, "r"), "fldChar", { fldCharType: "end" })
  }
  return strToU8(new XMLSerializer().serializeToString(doc).replaceAll("/>", " />"))
}

function elementChildren(node: Element) {
  return Array.from(node.childNodes).filter((row): row is Element => row.nodeType === 1 && (row as Element).namespaceURI === W)
}
