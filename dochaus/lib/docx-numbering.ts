import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// Word list numbering for a rendered document's clauses (#202). A draft is
// written in Markdown with its clause numbers as literal text ("## 1.
// Definitions", "1.1 In this Agreement", "(a) ..."). This pass turns those
// literals into Word multilevel list numbering, so a lawyer who inserts a
// clause in Word gets the following clauses renumbered.
//
// The document's own clause list is used when it has one: a template or
// precedent seed whose numbering is bound to heading (or clause) styles keeps
// its numbering definitions and styles. Otherwise a house list is added and
// bound to the heading styles that carry the numbered clauses.
//
// A literal is converted only where the list would number that paragraph with
// the same values at that position (Word's own counters are simulated), so the
// numbers a reader sees, and the references that cite them, do not change; a
// literal that does not fit the sequence stays text. Converting removes the
// literal from the paragraph's text; Docxodus reports the list label instead
// (BlockMetadata.list.generatedLabel). A paragraph that already belongs to the
// clause list and repeats its own number as text (an amendment that copied the
// visible number) loses the repeated literal. The pass is idempotent and leaves
// the package untouched when it changes nothing.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const BODY = "word/document.xml"
const NUMBERING = "word/numbering.xml"
const STYLES = "word/styles.xml"
const CONTENT_TYPES = "[Content_Types].xml"
const DOCUMENT_RELS = "word/_rels/document.xml.rels"
const NUMBERING_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
const NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

// Schema child order of the elements this pass writes into.
const PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
const LVL_ORDER = ["start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText", "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr"]
const STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"]

type Format = "decimal" | "decimalZero" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman"
const FORMATS = new Set<string>(["decimal", "decimalZero", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman"])

type Level = { start: number; format?: Format; refs: number[]; restart?: number; legal: boolean; style?: string }
type ClauseList = { abstract: string; levels: Level[]; numId: string; overrides: Map<string, Map<number, number>> }

// House clause levels: 1. / 1.1 / 1.1.1 / (a) / (i) / (A).
const HOUSE: { format: Format; text: string; left: number }[] = [
  { format: "decimal", text: "%1.", left: 720 },
  { format: "decimal", text: "%1.%2", left: 720 },
  { format: "decimal", text: "%1.%2.%3", left: 1440 },
  { format: "lowerLetter", text: "(%4)", left: 1440 },
  { format: "lowerRoman", text: "(%5)", left: 2160 },
  { format: "upperLetter", text: "(%6)", left: 2880 },
]

// "1.", "1.1", "1.1.1." ... or "(a)", "(iv)", "(A)", "(2)", then the clause text.
const PATH = /^(\d{1,3}(?:\.\d{1,3}){0,8})(\.?)[ \t\u00a0]+(?=\S)/
const TOKEN = /^\(([a-z]{1,4}|[A-Z]{1,4}|\d{1,3})\)[ \t\u00a0]+(?=\S)/

type Literal = { length: number } & ({ kind: "path"; values: number[]; dotted: boolean } | { kind: "token"; token: string })

type Paragraph = {
  node: Element
  style: string
  heading: boolean
  table: boolean
  member?: { abstract: string; numId: string; level: number; fromStyle: boolean }
  // The paragraph's own numPr numId when it is 0 (numbering suppressed).
  suppressed: boolean
  literal?: Literal
}

type Decision = { paragraph: Paragraph; level: number; kind: "convert" | "strip" }

// `house: false` numbers only against a clause list the document already has,
// for a document whose numbering was settled when it was drafted.
export function numberDocxClauses(bytes: Uint8Array, options: { house?: false } = {}): Uint8Array {
  const parts = unzipSync(bytes)
  if (!parts[BODY] || !parts[STYLES]) return bytes
  const document = parse(strFromU8(parts[BODY]))
  const styles = parse(strFromU8(parts[STYLES]))
  const numbering = parts[NUMBERING] ? parse(strFromU8(parts[NUMBERING])) : undefined
  const catalog = readStyles(styles)
  const lists = readNumbering(numbering, catalog)
  const paragraphs = bodyParagraphs(document, catalog, lists)

  const existing = clauseList(lists, catalog, paragraphs)
  const house = existing || options.house === false ? undefined : houseList(lists, paragraphs)
  const list = existing ?? house
  if (!list) return bytes
  const decisions = plan(list, lists, paragraphs)
  if (!decisions.length) return bytes
  if (house && !decisions.some(row => row.kind === "convert" && row.paragraph.heading)) return bytes

  let target = numbering
  if (house) {
    target = numbering ?? addNumberingPart(parts)
    const bound = houseBindings(decisions, paragraphs, catalog)
    writeHouseList(target, list, house.texts, bound)
    for (const [level, style] of bound) bindStyle(catalog.elements.get(style)!, level, list.numId)
    // A paragraph of a newly bound style that is not a clause keeps no number.
    const converted = new Set(decisions.map(row => row.paragraph))
    for (const paragraph of paragraphs)
      if (!converted.has(paragraph) && !paragraph.member && !paragraph.suppressed && !paragraph.table && [...bound.values()].includes(paragraph.style))
        setParagraphNumbering(paragraph.node, 0, "0")
  }
  // A clause takes the style its level is bound to: the document's own list
  // restyles it (the template's clause styles); the house list binds only the
  // style the clauses already use. A bound style supplies the number.
  const levelStyles = boundLevelStyles(list, catalog, target)
  for (const { paragraph, level, kind } of decisions) {
    stripLiteral(paragraph.node, paragraph.literal!.length)
    if (kind === "strip") continue
    const style = levelStyles.get(level)
    if (style && !house && style !== paragraph.style) setParagraphStyle(paragraph.node, style)
    if (style && (!house || style === paragraph.style)) clearParagraphNumbering(paragraph.node)
    else setParagraphNumbering(paragraph.node, level, list.numId)
  }

  const serialize = (doc: Document) => strToU8(new XMLSerializer().serializeToString(doc))
  parts[BODY] = serialize(document)
  if (house) {
    parts[STYLES] = serialize(styles)
    parts[NUMBERING] = serialize(target!)
  }
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

function parse(xml: string) {
  const fail = (message: string) => { throw new Error(`docx_numbering_xml: ${message}`) }
  return new DOMParser({ errorHandler: { warning: () => {}, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
}

function children(node: Element, name: string) {
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1 && (child as Element).namespaceURI === W && (child as Element).localName === name)
}

function child(node: Element | undefined, name: string) {
  return node ? children(node, name)[0] : undefined
}

function value(node: Element | undefined) {
  return node?.getAttributeNS(W, "val") ?? undefined
}

function element(doc: Document, name: string, attributes: Record<string, string> = {}) {
  const node = doc.createElementNS(W, `w:${name}`)
  for (const [key, val] of Object.entries(attributes)) node.setAttributeNS(W, `w:${key}`, val)
  return node
}

// Inserts `node` under `parent` at its schema position, replacing a same-named child.
function place(parent: Element, node: Element, order: string[]) {
  const existing = child(parent, node.localName!)
  if (existing) {
    parent.replaceChild(node, existing)
    return node
  }
  const rank = order.indexOf(node.localName!)
  const after = Array.from(parent.childNodes).find(other => other.nodeType === 1 && order.indexOf((other as Element).localName!) > rank)
  parent.insertBefore(node, after ?? null)
  return node
}

function properties(paragraph: Element) {
  return child(paragraph, "pPr") ?? (paragraph.insertBefore(element(paragraph.ownerDocument!, "pPr"), paragraph.firstChild) as Element)
}

// ---- styles

type Catalog = {
  elements: Map<string, Element>
  defaultStyle: string
  basedOn: (style: string) => string[]
  outline: (style: string) => number | undefined
  numPr: (style: string) => { numId?: string; level?: number } | undefined
  headingLike: (style: string) => boolean
}

function readStyles(styles: Document): Catalog {
  const elements = new Map<string, Element>()
  let defaultStyle = "Normal"
  for (const style of Array.from(styles.getElementsByTagNameNS(W, "style"))) {
    if (style.getAttributeNS(W, "type") !== "paragraph" && style.getAttributeNS(W, "type") !== "numbering") continue
    const id = style.getAttributeNS(W, "styleId") ?? ""
    if (!id) continue
    elements.set(id, style)
    if (style.getAttributeNS(W, "type") === "paragraph" && ["1", "true", "on"].includes(style.getAttributeNS(W, "default") ?? "")) defaultStyle = id
  }
  const basedOn = (style: string) => {
    const chain: string[] = []
    for (let id: string | undefined = style; id && elements.has(id) && !chain.includes(id); id = value(child(elements.get(id), "basedOn"))) chain.push(id)
    return chain
  }
  const outline = (style: string) => {
    for (const id of basedOn(style)) {
      const level = value(child(child(elements.get(id), "pPr"), "outlineLvl"))
      if (level !== undefined) return Number(level) < 9 ? Number(level) : undefined
    }
    return undefined
  }
  const numPr = (style: string) => {
    for (const id of basedOn(style)) {
      const numbered = child(child(elements.get(id), "pPr"), "numPr")
      if (numbered) {
        const level = value(child(numbered, "ilvl"))
        return { numId: value(child(numbered, "numId")), level: level === undefined ? undefined : Number(level) }
      }
    }
    return undefined
  }
  const headingLike = (style: string) => outline(style) !== undefined
    || /^heading [1-9]$/i.test(value(child(elements.get(style), "name")) ?? "")
    || /^Heading[1-9]$/.test(style)
  return { elements, defaultStyle, basedOn, outline, numPr, headingLike }
}

// ---- numbering

type Lists = {
  numToAbstract: Map<string, string>
  abstracts: Map<string, { levels: Level[] }>
  overrides: Map<string, Map<number, number>>
  maxAbstract: number
  maxNum: number
}

function readNumbering(numbering: Document | undefined, catalog: Catalog): Lists {
  const numToAbstract = new Map<string, string>()
  const overrides = new Map<string, Map<number, number>>()
  const definitions = new Map<string, Element>()
  let maxAbstract = -1
  let maxNum = 0
  if (numbering) {
    for (const abstract of Array.from(numbering.getElementsByTagNameNS(W, "abstractNum"))) {
      const id = abstract.getAttributeNS(W, "abstractNumId") ?? ""
      definitions.set(id, abstract)
      maxAbstract = Math.max(maxAbstract, Number(id) || 0)
    }
    for (const num of Array.from(numbering.getElementsByTagNameNS(W, "num"))) {
      const id = num.getAttributeNS(W, "numId") ?? ""
      maxNum = Math.max(maxNum, Number(id) || 0)
      const abstract = value(child(num, "abstractNumId"))
      if (abstract !== undefined) numToAbstract.set(id, abstract)
      const starts = new Map<number, number>()
      for (const override of children(num, "lvlOverride")) {
        const start = value(child(override, "startOverride"))
        if (start !== undefined) starts.set(Number(override.getAttributeNS(W, "ilvl") ?? 0), Number(start))
      }
      if (starts.size) overrides.set(id, starts)
    }
  }
  // A list defined through a numbering style (w:numStyleLink) takes its levels
  // from the abstractNum that carries the matching w:styleLink.
  const resolve = (id: string) => {
    const link = value(child(definitions.get(id), "numStyleLink"))
    if (!link) return id
    const target = catalog.numPr(link)?.numId
    const linked = target === undefined ? undefined : numToAbstract.get(target)
    return linked !== undefined && definitions.has(linked) ? linked : id
  }
  const abstracts = new Map<string, { levels: Level[] }>()
  for (const id of definitions.keys()) {
    const levels: Level[] = []
    for (const lvl of children(definitions.get(resolve(id))!, "lvl")) {
      const format = value(child(lvl, "numFmt")) ?? "decimal"
      const restart = value(child(lvl, "lvlRestart"))
      const legal = child(lvl, "isLgl")
      levels[Number(lvl.getAttributeNS(W, "ilvl") ?? 0)] = {
        start: Number(value(child(lvl, "start")) ?? 1),
        format: FORMATS.has(format) ? format as Format : undefined,
        refs: [...(value(child(lvl, "lvlText")) ?? "").matchAll(/%([1-9])/g)].map(match => Number(match[1])),
        restart: restart === undefined ? undefined : Number(restart),
        legal: legal !== undefined && !["0", "false", "off"].includes(value(legal) ?? ""),
        style: value(child(lvl, "pStyle")),
      }
    }
    abstracts.set(id, { levels })
  }
  return { numToAbstract, abstracts, overrides, maxAbstract, maxNum }
}

// ---- paragraphs

function bodyParagraphs(document: Document, catalog: Catalog, lists: Lists): Paragraph[] {
  return Array.from(document.getElementsByTagNameNS(W, "p")).flatMap(node => {
    let table = false
    for (let up = node.parentNode as Element | null; up && up.nodeType === 1; up = up.parentNode as Element | null) {
      if (up.namespaceURI !== W) continue
      if (["txbxContent", "del", "moveFrom", "footnote", "endnote", "comment"].includes(up.localName!)) return []
      if (up.localName === "tc") table = true
    }
    const pPr = child(node, "pPr")
    const style = value(child(pPr, "pStyle")) ?? catalog.defaultStyle
    const own = child(pPr, "numPr")
    const ownNum = value(child(own, "numId"))
    const ownLevel = value(child(own, "ilvl"))
    const styled = catalog.numPr(style)
    const numId = ownNum ?? styled?.numId
    const suppressed = ownNum === "0"
    let member: Paragraph["member"]
    if (numId !== undefined && numId !== "0" && lists.numToAbstract.has(numId)) {
      const abstract = lists.numToAbstract.get(numId)!
      const bound = lists.abstracts.get(abstract)?.levels.findIndex(level => level?.style === style) ?? -1
      const level = ownLevel !== undefined ? Number(ownLevel) : styled?.level ?? (bound >= 0 ? bound : 0)
      member = { abstract, numId, level, fromStyle: ownNum === undefined }
    }
    const outline = value(child(pPr, "outlineLvl"))
    const heading = (outline !== undefined && Number(outline) < 9) || catalog.headingLike(style)
    const excluded = /^(Title|Subtitle|TOC\d*|TOCHeading)$/i.test(style)
    return [{ node, style, heading, table, member, suppressed, literal: excluded ? undefined : readLiteral(paragraphText(node), heading) }]
  })
}

// The paragraph's leading run text, as Word shows it before any list label.
function textNodes(paragraph: Element) {
  return Array.from(paragraph.getElementsByTagNameNS(W, "*")).filter(node => {
    if (node.localName !== "t" && node.localName !== "tab") return false
    if (node.localName === "tab" && node.parentNode && (node.parentNode as Element).localName !== "r") return false
    for (let up = node.parentNode as Element | null; up && up !== paragraph; up = up.parentNode as Element | null)
      if (up.localName === "p" || up.localName === "del" || up.localName === "moveFrom") return false
    return true
  })
}

function paragraphText(paragraph: Element) {
  if (Array.from(paragraph.getElementsByTagNameNS(W, "fldSimple")).length || Array.from(paragraph.getElementsByTagNameNS(W, "fldChar")).length) return ""
  return textNodes(paragraph).map(node => node.localName === "tab" ? "\t" : node.textContent ?? "").join("")
}

function readLiteral(text: string, heading: boolean): Literal | undefined {
  const path = text.match(PATH)
  if (path) {
    const values = path[1]!.split(".").map(Number)
    // A lone "1" without its dot is too often ordinary text ("5 Year Warranty").
    if (values.length === 1 && !path[2]) return undefined
    // Outside headings a single "1." is a list item the Markdown render already numbered.
    if (values.length === 1 && !heading) return undefined
    return { kind: "path", values, dotted: Boolean(path[2]), length: path[0].length }
  }
  const token = text.match(TOKEN)
  if (token) return { kind: "token", token: token[1]!, length: token[0].length }
  return undefined
}

function stripLiteral(paragraph: Element, length: number) {
  let remaining = length
  for (const node of textNodes(paragraph)) {
    if (remaining <= 0) break
    if (node.localName === "tab") {
      remaining--
      removeEmptyRun(node)
      continue
    }
    const text = node.textContent ?? ""
    const taken = Math.min(text.length, remaining)
    remaining -= taken
    node.textContent = text.slice(taken)
    if (!node.textContent) removeEmptyRun(node)
  }
}

function removeEmptyRun(node: Element) {
  const run = node.parentNode as Element
  run.removeChild(node)
  if (run.localName === "r" && !Array.from(run.childNodes).some(other => other.nodeType === 1 && (other as Element).localName !== "rPr"))
    run.parentNode?.removeChild(run)
}

function setParagraphStyle(paragraph: Element, style: string) {
  place(properties(paragraph), element(paragraph.ownerDocument!, "pStyle", { val: style }), PPR_ORDER)
}

function setParagraphNumbering(paragraph: Element, level: number, numId: string) {
  const doc = paragraph.ownerDocument!
  const numPr = element(doc, "numPr")
  numPr.appendChild(element(doc, "ilvl", { val: String(level) }))
  numPr.appendChild(element(doc, "numId", { val: numId }))
  place(properties(paragraph), numPr, PPR_ORDER)
}

function clearParagraphNumbering(paragraph: Element) {
  const pPr = child(paragraph, "pPr")
  const numPr = child(pPr, "numPr")
  if (numPr) pPr!.removeChild(numPr)
}

// ---- choosing the clause list

// The document's own clause list: the list bound to the most heading-like
// styles (or numbering the most heading paragraphs), else one bound to several
// clause styles. Bullet lists and the plain lists a Markdown render creates
// are not clause lists.
function clauseList(lists: Lists, catalog: Catalog, paragraphs: Paragraph[]): ClauseList | undefined {
  const bindings = new Map<string, Map<number, Set<string>>>()
  const bind = (abstract: string, level: number, style: string) => {
    const levels = bindings.get(abstract) ?? new Map<number, Set<string>>()
    levels.set(level, (levels.get(level) ?? new Set()).add(style))
    bindings.set(abstract, levels)
  }
  for (const [abstract, { levels }] of lists.abstracts)
    levels.forEach((level, index) => { if (level?.style && catalog.elements.has(level.style)) bind(abstract, index, level.style) })
  for (const [style, node] of catalog.elements) {
    if (node.getAttributeNS(W, "type") !== "paragraph") continue
    const numbered = child(child(node, "pPr"), "numPr")
    const numId = value(child(numbered, "numId"))
    const abstract = numId === undefined ? undefined : lists.numToAbstract.get(numId)
    if (abstract === undefined) continue
    const level = value(child(numbered, "ilvl"))
    const bound = lists.abstracts.get(abstract)?.levels.findIndex(row => row?.style === style) ?? -1
    bind(abstract, level !== undefined ? Number(level) : bound >= 0 ? bound : 0, style)
  }
  const scored = [...lists.abstracts].flatMap(([abstract, { levels }]) => {
    if (!levels[0]?.format) return []
    const styles = [...new Set([...(bindings.get(abstract)?.values() ?? [])].flatMap(set => [...set]))]
    const headings = styles.filter(style => catalog.headingLike(style)).length
      + paragraphs.filter(row => row.member?.abstract === abstract && row.heading).length
    if (!headings && styles.length < 2) return []
    return [{ abstract, headings, styles: styles.length }]
  }).sort((a, b) => b.headings - a.headings || b.styles - a.styles || Number(a.abstract) - Number(b.abstract))
  const chosen = scored[0]
  if (!chosen) return undefined
  // The num instance the clause styles use, else the first that names the list.
  const numIds = [...lists.numToAbstract].filter(([, abstract]) => abstract === chosen.abstract).map(([numId]) => numId)
  const styleNum = [...catalog.elements.keys()].map(style => catalog.numPr(style)?.numId).find(numId => numId !== undefined && numIds.includes(numId))
  const numId = styleNum ?? numIds.sort((a, b) => Number(a) - Number(b))[0]
  if (numId === undefined) return undefined
  return { abstract: chosen.abstract, levels: lists.abstracts.get(chosen.abstract)!.levels, numId, overrides: lists.overrides }
}

function houseList(lists: Lists, paragraphs: Paragraph[]) {
  // Each decimal level keeps the dot style most of its literals use.
  const dotted = (depth: number) => {
    const rows = paragraphs.flatMap(row => row.literal?.kind === "path" && row.literal.values.length === depth ? [row.literal.dotted] : [])
    return rows.filter(Boolean).length > rows.length / 2
  }
  const texts = HOUSE.map((level, index) => index === 1 || index === 2 ? level.text + (dotted(index + 1) ? "." : "") : level.text)
  const abstract = String(lists.maxAbstract + 1)
  const levels: Level[] = HOUSE.map((level, index) => ({ start: 1, format: level.format, refs: [...texts[index]!.matchAll(/%([1-9])/g)].map(match => Number(match[1])), legal: false }))
  return { abstract, levels, numId: String(lists.maxNum + 1), overrides: new Map(), texts } satisfies ClauseList & { texts: string[] }
}

// ---- Word's counters

function letterValue(token: string, upper: boolean) {
  if (!(upper ? /^([A-Z])\1*$/ : /^([a-z])\1*$/).test(token)) return undefined
  return token.toLowerCase().charCodeAt(0) - 96 + 26 * (token.length - 1)
}

const ROMAN: [string, number][] = [["m", 1000], ["cm", 900], ["d", 500], ["cd", 400], ["c", 100], ["xc", 90], ["l", 50], ["xl", 40], ["x", 10], ["ix", 9], ["v", 5], ["iv", 4], ["i", 1]]
function romanValue(token: string, upper: boolean) {
  if (!(upper ? /^[IVXLCDM]+$/ : /^[ivxlcdm]+$/).test(token)) return undefined
  let rest = token.toLowerCase()
  let total = 0
  for (const [symbol, amount] of ROMAN) while (rest.startsWith(symbol)) { total += amount; rest = rest.slice(symbol.length) }
  const canonical = (() => { let out = ""; let left = total; for (const [symbol, amount] of ROMAN) while (left >= amount) { out += symbol; left -= amount } return out })()
  return canonical === token.toLowerCase() && total > 0 ? total : undefined
}

function tokenValue(token: string, format: Format) {
  if (format === "decimal" || format === "decimalZero") return /^\d+$/.test(token) ? Number(token) : undefined
  if (format === "lowerLetter" || format === "upperLetter") return letterValue(token, format === "upperLetter")
  return romanValue(token, format === "upperRoman")
}

class Counters {
  values: (number | undefined)[] = []
  seen = new Set<string>()
  constructor(private levels: Level[], private overrides: Map<string, Map<number, number>>) {}
  enter(numId: string) {
    if (this.seen.has(numId)) return
    this.seen.add(numId)
    for (const level of this.overrides.get(numId)?.keys() ?? []) this.values[level] = undefined
  }
  start(level: number, numId: string) {
    return this.overrides.get(numId)?.get(level) ?? this.levels[level]?.start ?? 1
  }
  shown(level: number, numId: string) {
    return this.values[level] ?? this.start(level, numId)
  }
  next(level: number, numId: string) {
    return this.values[level] === undefined ? this.start(level, numId) : this.values[level]! + 1
  }
  advance(level: number, numId: string) {
    this.values[level] = this.next(level, numId)
    for (let deeper = level + 1; deeper < Math.max(this.values.length, this.levels.length); deeper++) {
      const restart = this.levels[deeper]?.restart
      if (restart === 0) continue
      if (restart === undefined || level < restart) this.values[deeper] = undefined
    }
  }
}

// The level a literal names in the clause list at this point, if any: a
// decimal path "1.2.3" is the level whose label is built from levels 1-3,
// with the earlier values as they now stand; "(a)" is a level whose label is
// its own value alone. A level whose count continues is preferred over one
// that would start afresh.
function literalLevel(literal: Literal, list: ClauseList, counters: Counters, numId: string) {
  const decimal = (level?: Level) => level?.format === "decimal" || level?.format === "decimalZero"
  if (literal.kind === "path") {
    const level = literal.values.length - 1
    const definition = list.levels[level]
    if (!definition?.format || definition.refs.join() !== literal.values.map((_, index) => index + 1).join()) return undefined
    if (!definition.legal && !literal.values.every((_, index) => decimal(list.levels[index]))) return undefined
    if (literal.values.slice(0, -1).some((amount, index) => counters.shown(index, numId) !== amount)) return undefined
    return counters.next(level, numId) === literal.values.at(-1) ? level : undefined
  }
  const fits = list.levels.flatMap((definition, level) => level >= 1 && definition?.format && definition.refs.join() === String(level + 1)
    && tokenValue(literal.token, definition.format) === counters.next(level, numId) ? [level] : [])
  return fits.find(level => counters.values[level] !== undefined) ?? fits[0]
}

function literalMatches(literal: Literal, list: ClauseList, counters: Counters, level: number, numId: string) {
  const definition = list.levels[level]
  if (!definition?.format) return false
  if (literal.kind === "path")
    return definition.refs.join() === literal.values.map((_, index) => index + 1).join()
      && literal.values.every((amount, index) => counters.shown(index, numId) === amount)
  return definition.refs.join() === String(level + 1) && tokenValue(literal.token, definition.format) === counters.shown(level, numId)
}

function plan(list: ClauseList, lists: Lists, paragraphs: Paragraph[]) {
  const counters = new Map<string, Counters>()
  const countersFor = (abstract: string) => {
    if (!counters.has(abstract)) {
      const levels = abstract === list.abstract ? list.levels : lists.abstracts.get(abstract)?.levels ?? []
      counters.set(abstract, new Counters(levels, lists.overrides))
    }
    return counters.get(abstract)!
  }
  const decisions: Decision[] = []
  for (const paragraph of paragraphs) {
    if (paragraph.member) {
      const own = countersFor(paragraph.member.abstract)
      own.enter(paragraph.member.numId)
      own.advance(paragraph.member.level, paragraph.member.numId)
      if (paragraph.member.abstract === list.abstract && paragraph.literal && !paragraph.table
        && literalMatches(paragraph.literal, list, own, paragraph.member.level, paragraph.member.numId))
        decisions.push({ paragraph, level: paragraph.member.level, kind: "strip" })
      continue
    }
    if (!paragraph.literal || paragraph.table) continue
    const clause = countersFor(list.abstract)
    const level = literalLevel(paragraph.literal, list, clause, list.numId)
    if (level === undefined) continue
    clause.enter(list.numId)
    clause.advance(level, list.numId)
    decisions.push({ paragraph, level, kind: "convert" })
  }
  return decisions
}

// ---- writing the list

// A house level is bound to a heading style when every clause at that level
// uses it, it carries no other level, and most paragraphs in it are clauses,
// so a heading of that style added later in Word is numbered too.
function houseBindings(decisions: Decision[], paragraphs: Paragraph[], catalog: Catalog) {
  const converted = decisions.filter(row => row.kind === "convert")
  const bound = new Map<number, string>()
  for (const level of new Set(converted.map(row => row.level))) {
    const styles = new Set(converted.filter(row => row.level === level).map(row => row.paragraph.style))
    if (styles.size !== 1) continue
    const style = [...styles][0]!
    if (style === catalog.defaultStyle || !catalog.headingLike(style) || !catalog.elements.has(style)) continue
    if (converted.some(row => row.level !== level && row.paragraph.style === style)) continue
    if (catalog.numPr(style)) continue
    const clauses = converted.filter(row => row.paragraph.style === style).length
    const others = paragraphs.filter(row => row.style === style && !row.table && !converted.some(entry => entry.paragraph === row)).length
    if (clauses > others) bound.set(level, style)
  }
  return bound
}

function addNumberingPart(parts: Record<string, Uint8Array>) {
  const numbering = parse(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"/>`)
  const types = parse(strFromU8(parts[CONTENT_TYPES]!))
  const override = types.createElementNS(CONTENT_TYPES_NS, "Override")
  override.setAttribute("PartName", "/word/numbering.xml")
  override.setAttribute("ContentType", NUMBERING_TYPE)
  types.documentElement.appendChild(override)
  parts[CONTENT_TYPES] = strToU8(new XMLSerializer().serializeToString(types))
  const rels = parse(parts[DOCUMENT_RELS] ? strFromU8(parts[DOCUMENT_RELS]) : `<Relationships xmlns="${PACKAGE_RELS}"/>`)
  const ids = new Set(Array.from(rels.getElementsByTagNameNS(PACKAGE_RELS, "Relationship")).map(row => row.getAttribute("Id")))
  let index = 1
  while (ids.has(`rId${index}`)) index++
  const relationship = rels.createElementNS(PACKAGE_RELS, "Relationship")
  relationship.setAttribute("Id", `rId${index}`)
  relationship.setAttribute("Type", NUMBERING_REL)
  relationship.setAttribute("Target", "numbering.xml")
  rels.documentElement.appendChild(relationship)
  parts[DOCUMENT_RELS] = strToU8(new XMLSerializer().serializeToString(rels))
  return numbering
}

function writeHouseList(numbering: Document, list: ClauseList, texts: string[], bound: Map<number, string>) {
  const root = numbering.documentElement
  const abstract = element(numbering, "abstractNum", { abstractNumId: list.abstract })
  abstract.appendChild(element(numbering, "multiLevelType", { val: "multilevel" }))
  HOUSE.forEach((level, index) => {
    const lvl = element(numbering, "lvl", { ilvl: String(index) })
    lvl.appendChild(element(numbering, "start", { val: "1" }))
    lvl.appendChild(element(numbering, "numFmt", { val: level.format }))
    if (bound.has(index)) lvl.appendChild(element(numbering, "pStyle", { val: bound.get(index)! }))
    lvl.appendChild(element(numbering, "lvlText", { val: texts[index]! }))
    lvl.appendChild(element(numbering, "lvlJc", { val: "left" }))
    const pPr = element(numbering, "pPr")
    pPr.appendChild(element(numbering, "ind", { left: String(level.left), hanging: "720" }))
    lvl.appendChild(pPr)
    abstract.appendChild(lvl)
  })
  // Every abstractNum precedes every num.
  const firstNum = children(root, "num")[0] ?? children(root, "numIdMacAtCleanup")[0]
  root.insertBefore(abstract, firstNum ?? null)
  const num = element(numbering, "num", { numId: list.numId })
  num.appendChild(element(numbering, "abstractNumId", { val: list.abstract }))
  root.insertBefore(num, children(root, "numIdMacAtCleanup")[0] ?? null)
}

function bindStyle(style: Element, level: number, numId: string) {
  const doc = style.ownerDocument!
  const pPr = child(style, "pPr") ?? place(style, element(doc, "pPr"), STYLE_ORDER)
  const numPr = element(doc, "numPr")
  numPr.appendChild(element(doc, "ilvl", { val: String(level) }))
  numPr.appendChild(element(doc, "numId", { val: numId }))
  place(pPr, numPr, PPR_ORDER)
}

// The paragraph style the list numbers at each level, from the level's
// w:pStyle or a style whose numbering names the list.
function boundLevelStyles(list: ClauseList, catalog: Catalog, numbering: Document | undefined) {
  const styles = new Map<number, string>()
  const current = readNumbering(numbering, catalog)
  const levels = current.abstracts.get(list.abstract)?.levels ?? list.levels
  levels.forEach((level, index) => { if (level?.style && catalog.elements.has(level.style)) styles.set(index, level.style) })
  for (const [style, node] of catalog.elements) {
    const own = child(child(node, "pPr"), "numPr")
    const numId = value(child(own, "numId"))
    if (node.getAttributeNS(W, "type") !== "paragraph" || numId === undefined || current.numToAbstract.get(numId) !== list.abstract) continue
    const ilvl = value(child(own, "ilvl"))
    const level = ilvl !== undefined ? Number(ilvl) : Math.max(0, levels.findIndex(row => row?.style === style))
    if (!styles.has(level)) styles.set(level, style)
  }
  return styles
}
