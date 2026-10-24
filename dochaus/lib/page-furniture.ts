import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import type { DraftingPreferences } from "./drafting-preferences"
import { furnitureLegend } from "./furniture-legend"

// Page furniture for documents doc.haus renders: page numbers, the draft line
// and the firm's chosen legend. The planner picks only each document's form
// from a closed set; what the form carries follows here from the firm's
// settings. Legends have legal effect, so they are the firm's own text and are
// never written by a model. Uploaded, counterparty and template files are
// never given furniture, and a document that already has a header or footer
// keeps it; a revised original keeps its own, the firm's header lines written
// above its own header text (addPageFurnitureLines).

/** A planned document's form, from its outline "Form:" line. */
export type DocumentForm = "agreement" | "advice" | "letter" | "advice letter" | "other"
export const DOCUMENT_FORMS: readonly DocumentForm[] = ["agreement", "advice", "letter", "advice letter", "other"]

/** A letter, advising or not: numbered from its second page, its planning
 * labels unheaded. */
export function isLetter(form: string | undefined) {
  return form === "letter" || form === "advice letter"
}

/** What a rendered file is: a planned document's form, or the run's internal drafting notes. */
export type FurnitureKind = DocumentForm | "notes"

/** What a request asks of one document's furniture: `legend` is a legend in
 * the request's own words, shown after the firm's, and `pageNumbers: false`
 * leaves the page numbers out. A request never removes the firm's legend or
 * draft line; only the firm's settings do. */
export type RequestedFurniture = { legend?: string; pageNumbers?: false }

/** `header` lines are written in order at the top of every page; `pageNumber`
 * goes in the footer; `firstPage: "unnumbered"` leaves the first page's number
 * out (a letter, or a firm that does not number page 1), keeping its header. */
export type PageFurniture = {
  header: string[]
  pageNumber?: { format: "plain" | "page" | "dashed"; align: "center" | "right" }
  firstPage?: "unnumbered"
}

export const NOTES_MARKING = "Internal drafting notes — not for circulation"

type FurniturePreferences = Pick<DraftingPreferences, "firm" | "dateFormat" | "pageNumbers" | "pageNumberPosition" | "numberFirstPage" | "draftLine" | "draftLineFirm" | "agreementLegend" | "adviceLegend" | "notesMarking">

/** The furniture a kind of file carries under the firm's settings and what the
 * request asks of it, or undefined when it carries none. `draftedAt` is the
 * time the draft was made (an ISO timestamp kept in run state, never the
 * render clock); without it there is no draft line. A requested legend the
 * firm's already reads, ignoring case, is not repeated. */
export function pageFurniture(kind: FurnitureKind, prefs: FurniturePreferences, draftedAt?: string, requested: RequestedFurniture = {}): PageFurniture | undefined {
  const legend = furnitureLegend(kind, prefs)
  const extra = requested.legend && requested.legend.toLowerCase() !== legend.toLowerCase() ? requested.legend : ""
  const drafted = draftedAt && prefs.draftLine && (kind === "agreement" || kind === "advice" || isLetter(kind))
  const date = drafted ? formatDraftDate(draftedAt, prefs.dateFormat) : undefined
  const firm = prefs.draftLineFirm ? prefs.firm.trim() : ""
  const header = [
    ...(legend ? [legend] : []),
    ...(extra ? [extra] : []),
    ...(date ? [`${firm ? `${firm} draft` : "Draft"}: ${date}`] : []),
    ...(kind === "notes" && prefs.notesMarking ? [NOTES_MARKING] : []),
  ]
  const pageNumber = prefs.pageNumbers === "off" || requested.pageNumbers === false ? undefined : { format: prefs.pageNumbers, align: prefs.pageNumberPosition }
  if (!header.length && !pageNumber) return undefined
  const unnumbered = pageNumber && (isLetter(kind) || !prefs.numberFirstPage)
  return { header, ...(pageNumber ? { pageNumber } : {}), ...(unnumbered ? { firstPage: "unnumbered" as const } : {}) }
}

/** The text each page number format prints before and after the PAGE field. */
export const PAGE_NUMBER_AFFIXES = { plain: ["", ""], page: ["Page ", ""], dashed: ["- ", " -"] } as const

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

/** A draft line as pageFurniture writes it: "Draft: " or a firm's "<firm>
 * draft: ", then a date in one of formatDraftDate's styles, and nothing else. */
const DRAFT_LINE = new RegExp(String.raw`^(?:\S.* draft|Draft): (?:\d{4}-\d{2}-\d{2}|\d{1,2} (?:${MONTHS.join("|")}) \d{4}|(?:${MONTHS.join("|")}) \d{1,2}, \d{4})$`)

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
type Row = { text: string; align: "center" | "right"; page?: string; after?: string }

/** Writes `furniture` into a rendered body's last section as new header and
 * footer parts: the header lines right-aligned, one paragraph each, and the
 * page number in its own footer paragraph, a PAGE field after any "Page " or
 * between dashes. No NUMPAGES: the combined memo and body file restarts the
 * body at 1, so a page count would count the memo. An unnumbered first page
 * gets explicit first-page parts, the same header and an empty footer, so no
 * reader falls back to the others. A document that
 * already has a header or footer anywhere (a template's or an uploaded file's
 * own) is returned unchanged. Pure; fixed zip entry times keep identical input
 * byte-identical. */
export function applyPageFurniture(docx: Uint8Array, furniture: PageFurniture): Uint8Array {
  if (!furniture.header.length && !furniture.pageNumber) return docx
  const parts = unzipSync(docx)
  const document = parseXml(parts[BODY]!)
  if (["headerReference", "footerReference"].some(name => document.getElementsByTagNameNS(W, name).length)) return docx
  const body = document.getElementsByTagNameNS(W, "body")[0]!
  const section = elementChildren(body).findLast(node => node.localName === "sectPr") ?? body.appendChild(document.createElementNS(W, "w:sectPr"))
  const stories = storyWriter(parts, document)
  const references: Element[] = []
  const unnumbered = furniture.firstPage === "unnumbered" && furniture.pageNumber !== undefined
  const add = (kind: "header" | "footer", type: "default" | "first", rows: Row[]) => references.push(stories.add(kind, type, rows))
  if (furniture.header.length) {
    const rows = furniture.header.map(text => ({ text, align: "right" as const }))
    add("header", "default", rows)
    if (unnumbered) add("header", "first", rows)
  }
  if (furniture.pageNumber) {
    const { format, align } = furniture.pageNumber
    const [before, after] = PAGE_NUMBER_AFFIXES[format]
    // The cached result is what a reader that does not lay out pages shows: the
    // number of the first page this footer appears on.
    add("footer", "default", [{ text: before, align, page: unnumbered ? "2" : "1", ...(after ? { after } : {}) }])
    if (unnumbered) add("footer", "first", [{ text: "", align }])
  }
  // Header and footer references come first in a section's properties; the
  // title-page switch goes before the elements the schema puts after it.
  const first = section.firstChild
  for (const reference of references) section.insertBefore(reference, first)
  withRelationshipsPrefix(document)
  if (unnumbered) {
    const existing = elementChildren(section).find(node => node.localName === "titlePg")
    if (existing) existing.removeAttributeNS(W, "val")
    else section.insertBefore(document.createElementNS(W, "w:titlePg"), elementChildren(section)
      .find(node => ["textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"].includes(node.localName)) ?? null)
  }
  return stories.save()
}

/** Writes `furniture`'s header lines into a revised original that keeps its
 * own header and footer: new paragraphs, as applyPageFurniture writes them,
 * above the text of each header its last section shows (the default one, the
 * first page's when it has a distinct first page, even pages' when the
 * document sets them apart), and a header of the lines alone in place of one
 * it shows none for. A line a header already shows (ignoring case and
 * spacing) is not repeated, and a draft line takes the text of one the header
 * shows in that same form (a revised doc.haus draft's), so only its date or
 * firm changes. Its footers and its own page numbering stay as they are:
 * the firm's page numbers are not added. An original with no header or footer
 * at all takes applyPageFurniture's. Pure; fixed zip entry times. */
export function addPageFurnitureLines(docx: Uint8Array, furniture: PageFurniture): Uint8Array {
  const parts = unzipSync(docx)
  const document = parseXml(parts[BODY]!)
  if (!["headerReference", "footerReference"].some(name => document.getElementsByTagNameNS(W, name).length)) return applyPageFurniture(docx, furniture)
  const body = document.getElementsByTagNameNS(W, "body")[0]!
  const section = elementChildren(body).findLast(node => node.localName === "sectPr")
  if (!furniture.header.length || !section) return docx
  const stories = storyWriter(parts, document)
  const on = (node: Element | undefined) => node !== undefined && !["0", "false", "off"].includes(node.getAttributeNS(W, "val") ?? "")
  const settings = stories.part(row => row.getAttribute("Type") === `${O}/settings`)
  const shown = ["default" as const,
    ...on(elementChildren(section).find(node => node.localName === "titlePg")) ? ["first" as const] : [],
    ...settings && parts[settings] && on(parseXml(parts[settings]).getElementsByTagNameNS(W, "evenAndOddHeaders")[0]) ? ["even" as const] : []]
  const rows = furniture.header.map(text => ({ text, align: "right" as const }))
  const lines = elementChildren(parseXml(storyPart("header", rows)).documentElement)
  const written = new Set<string>()
  const added: Element[] = []
  for (const type of shown) {
    const reference = elementChildren(section).find(node => node.localName === "headerReference" && (node.getAttributeNS(W, "type") || "default") === type)
    if (!reference) {
      added.push(stories.add("header", type, rows))
      continue
    }
    const name = stories.part(row => row.getAttribute("Id") === reference.getAttributeNS(O, "id"))
    if (!name || !parts[name] || written.has(name)) continue
    written.add(name)
    const header = parseXml(parts[name])
    const own = Array.from(header.getElementsByTagNameNS(W, "p"))
    const opening = header.documentElement.firstChild
    let changed = false
    for (const line of lines) {
      const text = shownText(line)
      if (own.some(paragraph => sameLine(shownText(paragraph), text))) continue
      const dated = DRAFT_LINE.test(text) ? own.find(paragraph => DRAFT_LINE.test(shownText(paragraph).trim())) : undefined
      if (dated) withText(dated, text)
      else header.documentElement.insertBefore(header.importNode(line, true), opening)
      changed = true
    }
    if (changed) parts[name] = strToU8(new XMLSerializer().serializeToString(header).replaceAll("/>", " />"))
  }
  for (const reference of added) section.insertBefore(reference, section.firstChild)
  if (added.length) withRelationshipsPrefix(document)
  return stories.save()
}

const parseXml = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")

function shownText(paragraph: Element) {
  return Array.from(paragraph.getElementsByTagNameNS(W, "t")).map(node => node.textContent ?? "").join("")
}

function sameLine(a: string, b: string) {
  const flat = (text: string) => text.replace(/\s+/g, " ").trim().toLowerCase()
  return flat(a) === flat(b)
}

/** Gives a paragraph `text` in its first text run, keeping its runs' look. */
function withText(paragraph: Element, text: string) {
  const [first, ...rest] = Array.from(paragraph.getElementsByTagNameNS(W, "t"))
  while (first!.firstChild) first!.removeChild(first!.firstChild)
  first!.appendChild(paragraph.ownerDocument.createTextNode(text))
  first!.setAttributeNS(XML, "xml:space", "preserve")
  for (const node of rest) node.parentNode!.removeChild(node)
}

/** Declares the relationships prefix once on the root, not on every reference. */
function withRelationshipsPrefix(document: Document) {
  if (!document.documentElement.lookupNamespaceURI("r")) document.documentElement.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", O)
}

// Header and footer parts added to a package: `add` writes a part of `rows`
// with its relationship and content type and returns the section reference to
// it; `part` names the part a document relationship targets; `save` writes the
// body, relationships and content types back and zips the package.
function storyWriter(parts: Record<string, Uint8Array>, document: Document) {
  const relationships = parseXml(parts[RELS] ?? strToU8(`<Relationships xmlns="${R}"/>`))
  const types = parseXml(parts[TYPES]!)
  const rows = () => Array.from(relationships.getElementsByTagNameNS(R, "Relationship"))
  const ids = new Set(rows().map(row => row.getAttribute("Id")))
  const add = (kind: "header" | "footer", type: "default" | "first" | "even", lines: Row[]) => {
    let n = 1
    while (parts[`word/${kind}${n}.xml`]) n++
    parts[`word/${kind}${n}.xml`] = storyPart(kind, lines)
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
    return reference
  }
  const part = (where: (row: Element) => boolean) => {
    const target = rows().find(where)?.getAttribute("Target")
    return target ? target.startsWith("/") ? target.slice(1) : `word/${target}` : undefined
  }
  const save = () => {
    const serializer = new XMLSerializer()
    parts[BODY] = strToU8(serializer.serializeToString(document))
    parts[RELS] = strToU8(serializer.serializeToString(relationships))
    parts[TYPES] = strToU8(serializer.serializeToString(types))
    return zipSync(parts, { mtime: new Date(1980, 0, 1) })
  }
  return { add, part, save }
}

// One header or footer part, a paragraph per row in the part's style. Text goes
// in as text nodes, never as markup, as one line without the characters XML 1.0
// forbids (a firm name is free text). A row with `page` has a PAGE field cached
// as it, then any `after` text. Serialized as Docxodus saves a part it has read (its declaration,
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
    if (row.after) text(paragraph, "t", row.after)
  }
  return strToU8(new XMLSerializer().serializeToString(doc).replaceAll("/>", " />"))
}

function elementChildren(node: Element) {
  return Array.from(node.childNodes).filter((row): row is Element => row.nodeType === 1 && (row as Element).namespaceURI === W)
}
