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
// its numbering definitions and styles; of several, the one that can show the
// most of the document's numbers. Otherwise a house list is added and bound to
// the heading styles that carry the numbered clauses.
//
// A literal converts only where the list would show exactly that label at that
// position (Word's own counters are simulated, a skipped level counting as
// started), so the numbers a reader sees, and the references that cite them,
// do not change. Numbers convert by sequence: a sequence opens at a heading
// numbered with the list's first top-level label ("1.") and runs to the next
// such heading (a schedule numbered afresh, which restarts the list). A
// sequence converts whole or stays exactly as typed: one number the list cannot
// show (a gap, an undotted or run-in number, a label the list writes
// differently) keeps all of its numbers typed. Numbers before the first
// sequence (parties, recitals) stay typed. Converting removes the literal from
// the paragraph's text; Docxodus reports the list label instead
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
const STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"]

type Format = "decimal" | "decimalZero" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman"
const FORMATS = new Set<string>(["decimal", "decimalZero", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman"])

type Level = { start: number; format?: Format; text: string; restart?: number; legal: boolean; style?: string }
type ClauseList = { abstract: string; levels: Level[]; numId: string; overrides: Map<string, Map<number, number>> }

// House clause levels: 1. / 1.1 / 1.1.1 / (a) / (i) / (A) / (I), then two
// levels no typed number takes, so Tab in Word still goes deeper.
const HOUSE: { format: Format; text: string; left: number }[] = [
  { format: "decimal", text: "%1.", left: 720 },
  { format: "decimal", text: "%1.%2", left: 720 },
  { format: "decimal", text: "%1.%2.%3", left: 1440 },
  { format: "lowerLetter", text: "(%4)", left: 2160 },
  { format: "lowerRoman", text: "(%5)", left: 2880 },
  { format: "upperLetter", text: "(%6)", left: 3600 },
  { format: "upperRoman", text: "(%7)", left: 4320 },
  { format: "lowerLetter", text: "%8.", left: 5040 },
  { format: "lowerRoman", text: "%9.", left: 5760 },
]

// "1.", "1.1", "1.1.1." ... or "(a)", "(iv)", "(A)", "(2)", then the clause text.
const PATH = /^(\d{1,3}(?:\.\d{1,3}){0,8})(\.?)[ \t ]+(?=\S)/
const TOKEN = /^\(([a-z]{1,4}|[A-Z]{1,4}|\d{1,3})\)[ \t ]+(?=\S)/

// `text` is the number as typed ("1.1", "(a)"); `length` runs to the clause text.
type Literal = { text: string; length: number }

type Paragraph = {
  node: Element
  style: string
  heading: boolean
  table: boolean
  member?: { abstract: string; numId: string; level: number }
  // The paragraph's own numPr numId when it is 0 (numbering suppressed).
  suppressed: boolean
  // A later part of a paragraph an amendment split (`split` option).
  split: boolean
  literal?: Literal
}

type Decision = { paragraph: Paragraph; level: number; kind: "convert" | "strip" | "detach"; numId: string }

/** Turns a document's literal clause numbers into Word list numbering.
 * `house: false` numbers only against a clause list the document already has,
 * for a document whose numbering was settled when it was drafted: a literal
 * converts there only inside a sequence the list already numbers. `split`
 * names the later parts of paragraphs an amendment split: each leaves the
 * clause list, unless it opens with a number of a deeper level of the same
 * list, which it then takes. */
export function numberDocxClauses(bytes: Uint8Array, options: { house?: false; split?: (paragraph: Element) => boolean } = {}): Uint8Array {
  const parts = unzipSync(bytes)
  if (!parts[BODY] || !parts[STYLES]) return bytes
  const document = parse(strFromU8(parts[BODY]))
  const styles = parse(strFromU8(parts[STYLES]))
  const numbering = parts[NUMBERING] ? parse(strFromU8(parts[NUMBERING])) : undefined
  const catalog = readStyles(styles)
  const lists = readNumbering(numbering, catalog)
  const paragraphs = bodyParagraphs(document, catalog, lists, options.split)

  // Of the document's own clause lists the one that numbers the most; the
  // house list only for a document without one.
  const own = clauseLists(lists, catalog, paragraphs)
  const house = own.length || options.house === false ? undefined : houseList(lists, paragraphs)
  const candidates = house ? [house] : own
  const firstRestart = lists.maxNum + (house ? 2 : 1)
  const ranked = candidates.map(list => ({ list, planned: plan(list, paragraphs, options.house === false, firstRestart) }))
  const best = ranked.reduce<(typeof ranked)[number] | undefined>((top, row) => !top || row.planned.score > top.planned.score ? row : top, undefined)
  if (!best || !best.planned.decisions.length) return bytes
  const { list, planned: { decisions, restarts } } = best

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
  for (const numId of restarts) writeRestart(target!, list, lists, numId)
  // A clause takes the style its level is bound to: the document's own list
  // restyles it (the template's clause styles), but a heading keeps a heading
  // style; the house list binds only the style the clauses already use. A
  // style supplies the number only where its own numbering is this level of
  // this list instance; otherwise the paragraph carries it directly.
  const levelStyles = house ? new Map<number, string>() : boundLevelStyles(list, catalog, lists)
  for (const { paragraph, level, kind, numId } of decisions) {
    if (kind === "detach") {
      if (catalog.numPr(paragraph.style)?.numId) setParagraphNumbering(paragraph.node, 0, "0")
      else clearParagraphNumbering(paragraph.node)
      continue
    }
    stripLiteral(paragraph.node, paragraph.literal!.length)
    if (kind === "strip") continue
    let style = paragraph.style
    const bound = levelStyles.get(level)
    if (bound && bound !== style && !(paragraph.heading && !catalog.headingLike(bound))) {
      setParagraphStyle(paragraph.node, bound)
      style = bound
    }
    const styled = catalog.numPr(style)
    if (styled?.numId === numId && (styled.level ?? levelOfStyle(list, style)) === level) clearParagraphNumbering(paragraph.node)
    else setParagraphNumbering(paragraph.node, level, numId)
  }

  const serialize = (doc: Document) => strToU8(new XMLSerializer().serializeToString(doc))
  parts[BODY] = serialize(document)
  if (house) parts[STYLES] = serialize(styles)
  if (house || restarts.length) parts[NUMBERING] = serialize(target!)
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

/** The w:num ids of the document's clause list (the list numberDocxClauses
 * numbers clauses with), and how many paragraphs it numbers. */
export function clauseNumbering(bytes: Uint8Array): { numIds: Set<string>; paragraphs: number } {
  const parts = unzipSync(bytes)
  const none = { numIds: new Set<string>(), paragraphs: 0 }
  if (!parts[BODY] || !parts[STYLES]) return none
  const catalog = readStyles(parse(strFromU8(parts[STYLES])))
  const lists = readNumbering(parts[NUMBERING] ? parse(strFromU8(parts[NUMBERING])) : undefined, catalog)
  const paragraphs = bodyParagraphs(parse(strFromU8(parts[BODY])), catalog, lists)
  const ranked = clauseLists(lists, catalog, paragraphs).map(list => ({ list, members: paragraphs.filter(row => !row.table && row.member?.abstract === list.abstract).length }))
  const best = ranked.reduce<(typeof ranked)[number] | undefined>((top, row) => !top || row.members > top.members ? row : top, undefined)
  if (!best?.members) return none
  return { numIds: new Set([...lists.numToAbstract].filter(([, abstract]) => abstract === best.list.abstract).map(([numId]) => numId)), paragraphs: best.members }
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
  return { elements, defaultStyle, outline, numPr, headingLike }
}

// ---- numbering

type Lists = {
  // Each num's list, a list defined through a numbering style (w:numStyleLink)
  // being the abstractNum that carries the matching w:styleLink.
  numToAbstract: Map<string, string>
  // Each num's own w:abstractNumId, as written.
  numDefinition: Map<string, string>
  abstracts: Map<string, { levels: Level[] }>
  overrides: Map<string, Map<number, number>>
  maxAbstract: number
  maxNum: number
}

function readNumbering(numbering: Document | undefined, catalog: Catalog): Lists {
  const numDefinition = new Map<string, string>()
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
      if (abstract !== undefined) numDefinition.set(id, abstract)
      const starts = new Map<number, number>()
      for (const override of children(num, "lvlOverride")) {
        const start = value(child(override, "startOverride"))
        if (start !== undefined) starts.set(Number(override.getAttributeNS(W, "ilvl") ?? 0), Number(start))
      }
      if (starts.size) overrides.set(id, starts)
    }
  }
  const resolve = (id: string) => {
    const link = value(child(definitions.get(id), "numStyleLink"))
    if (!link) return id
    const target = catalog.numPr(link)?.numId
    const linked = target === undefined ? undefined : numDefinition.get(target)
    return linked !== undefined && definitions.has(linked) && !value(child(definitions.get(linked), "numStyleLink")) ? linked : id
  }
  const numToAbstract = new Map([...numDefinition].map(([numId, abstract]) => [numId, resolve(abstract)]))
  const abstracts = new Map<string, { levels: Level[] }>()
  for (const id of definitions.keys()) {
    if (resolve(id) !== id) continue
    const levels: Level[] = []
    for (const lvl of children(definitions.get(id)!, "lvl")) {
      const format = value(child(lvl, "numFmt")) ?? "decimal"
      const restart = value(child(lvl, "lvlRestart"))
      const legal = child(lvl, "isLgl")
      levels[Number(lvl.getAttributeNS(W, "ilvl") ?? 0)] = {
        start: Number(value(child(lvl, "start")) ?? 1),
        format: FORMATS.has(format) ? format as Format : undefined,
        text: value(child(lvl, "lvlText")) ?? "",
        restart: restart === undefined ? undefined : Number(restart),
        legal: legal !== undefined && !["0", "false", "off"].includes(value(legal) ?? ""),
        style: value(child(lvl, "pStyle")),
      }
    }
    abstracts.set(id, { levels })
  }
  return { numToAbstract, numDefinition, abstracts, overrides, maxAbstract, maxNum }
}

// ---- paragraphs

function bodyParagraphs(document: Document, catalog: Catalog, lists: Lists, split?: (paragraph: Element) => boolean): Paragraph[] {
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
      member = { abstract, numId, level }
    }
    const outline = value(child(pPr, "outlineLvl"))
    const heading = (outline !== undefined && Number(outline) < 9) || catalog.headingLike(style)
    const excluded = /^(Title|Subtitle|TOC\d*|TOCHeading)$/i.test(style)
    return [{ node, style, heading, table, member, suppressed, split: Boolean(split?.(node)), literal: excluded ? undefined : readLiteral(paragraphText(node), heading) }]
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
    // A lone "1" without its dot is too often ordinary text ("5 Year Warranty").
    if (!path[1]!.includes(".") && !path[2]) return undefined
    // Outside headings a single "1." is a list item the Markdown render already numbered.
    if (!path[1]!.includes(".") && !heading) return undefined
    return { text: path[1]! + path[2]!, length: path[0].length }
  }
  const token = text.match(TOKEN)
  if (token) return { text: `(${token[1]})`, length: token[0].length }
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

// The document's own clause lists, best bound first: lists bound to heading-like
// styles (or numbering heading paragraphs), else to several clause styles.
// Bullet lists and the plain lists a Markdown render creates are not clause
// lists. A list is used through the num instance its clause styles use, else
// the first that names it.
function clauseLists(lists: Lists, catalog: Catalog, paragraphs: Paragraph[]): ClauseList[] {
  const bindings = new Map<string, Set<string>>()
  const bind = (abstract: string, style: string) => bindings.set(abstract, (bindings.get(abstract) ?? new Set()).add(style))
  for (const [abstract, { levels }] of lists.abstracts)
    for (const level of levels) if (level?.style && catalog.elements.has(level.style)) bind(abstract, level.style)
  const styleNums = new Set<string>()
  for (const [style, node] of catalog.elements) {
    if (node.getAttributeNS(W, "type") !== "paragraph") continue
    const numId = value(child(child(child(node, "pPr"), "numPr"), "numId"))
    const abstract = numId === undefined ? undefined : lists.numToAbstract.get(numId)
    if (abstract === undefined) continue
    styleNums.add(numId!)
    bind(abstract, style)
  }
  return [...lists.abstracts].flatMap(([abstract, { levels }]) => {
    if (!levels[0]?.format) return []
    const styles = [...bindings.get(abstract) ?? []]
    const headings = styles.filter(style => catalog.headingLike(style)).length
      + paragraphs.filter(row => row.member?.abstract === abstract && row.heading).length
    if (!headings && styles.length < 2) return []
    const numIds = [...lists.numToAbstract].filter(([, id]) => id === abstract).map(([numId]) => numId).sort((a, b) => Number(a) - Number(b))
    const numId = numIds.find(id => styleNums.has(id)) ?? numIds[0]
    if (numId === undefined) return []
    return [{ list: { abstract, levels, numId, overrides: lists.overrides }, headings, styles: styles.length }]
  }).sort((a, b) => b.headings - a.headings || b.styles - a.styles || Number(a.list.abstract) - Number(b.list.abstract)).map(row => row.list)
}

function houseList(lists: Lists, paragraphs: Paragraph[]) {
  // Each decimal level keeps the dot style most of its literals use.
  const dotted = (depth: number) => {
    const rows = paragraphs.flatMap(row => row.literal && /^\d/.test(row.literal.text) && row.literal.text.replace(/\.$/, "").split(".").length === depth ? [row.literal.text.endsWith(".")] : [])
    return rows.filter(Boolean).length > rows.length / 2
  }
  const texts = HOUSE.map((level, index) => index === 1 || index === 2 ? level.text + (dotted(index + 1) ? "." : "") : level.text)
  const abstract = String(lists.maxAbstract + 1)
  const levels: Level[] = HOUSE.map((level, index) => ({ start: 1, format: level.format, text: texts[index]!, legal: false }))
  return { abstract, levels, numId: String(lists.maxNum + 1), overrides: new Map(), texts } satisfies ClauseList & { texts: string[] }
}

// ---- Word's counters and labels

const ROMAN: [string, number][] = [["m", 1000], ["cm", 900], ["d", 500], ["cd", 400], ["c", 100], ["xc", 90], ["l", 50], ["xl", 40], ["x", 10], ["ix", 9], ["v", 5], ["iv", 4], ["i", 1]]

function formatted(amount: number, format: Format) {
  if (format === "decimal") return String(amount)
  if (format === "decimalZero") return String(amount).padStart(2, "0")
  if (format === "lowerLetter" || format === "upperLetter") {
    const letters = String.fromCharCode(97 + (amount - 1) % 26).repeat(Math.floor((amount - 1) / 26) + 1)
    return format === "upperLetter" ? letters.toUpperCase() : letters
  }
  let out = ""
  let left = amount
  for (const [symbol, worth] of ROMAN) while (left >= worth) { out += symbol; left -= worth }
  return format === "upperRoman" ? out.toUpperCase() : out
}

// The label Word shows for `level` with each level's value as `at` gives it.
function labelOf(levels: Level[], level: number, at: (level: number) => number) {
  const definition = levels[level]
  if (!definition?.format || amountInvalid(at(level))) return undefined
  let complete = true
  const label = definition.text.replace(/%([1-9])/g, (_, reference: string) => {
    const index = Number(reference) - 1
    const format = definition.legal ? "decimal" : levels[index]?.format
    if (!format || index > level || amountInvalid(at(index))) {
      complete = false
      return ""
    }
    return formatted(at(index), format)
  })
  return complete ? label.trim() : undefined
}

function amountInvalid(amount: number) {
  return !Number.isInteger(amount) || amount < 1
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
  // A number at a level counts every level above it that has not started yet
  // as started (Word and LibreOffice show its start value and count on from it).
  advance(level: number, numId: string) {
    for (let higher = 0; higher < level; higher++) if (this.values[higher] === undefined) this.values[higher] = this.start(higher, numId)
    this.values[level] = this.next(level, numId)
    for (let deeper = level + 1; deeper < Math.max(this.values.length, this.levels.length); deeper++) {
      const restart = this.levels[deeper]?.restart
      if (restart === 0) continue
      if (restart === undefined || level < restart) this.values[deeper] = undefined
    }
  }
  save() {
    return { values: [...this.values], seen: new Set(this.seen) }
  }
  restore(saved: ReturnType<Counters["save"]>) {
    this.values = saved.values
    this.seen = saved.seen
  }
}

// The level of the list that would show `literal` next at this point; a level
// whose count continues is preferred over one that would start afresh.
function literalLevel(literal: Literal, list: ClauseList, counters: Counters, numId: string) {
  const fits = list.levels.flatMap((_, level) =>
    labelOf(list.levels, level, at => at === level ? counters.next(level, numId) : counters.shown(at, numId)) === literal.text ? [level] : [])
  return fits.find(level => counters.values[level] !== undefined) ?? fits[0]
}

// The label the list shows at `level` for a sequence opened afresh.
function opensSequence(literal: Literal, list: ClauseList) {
  return labelOf(list.levels, 0, level => list.overrides.get(list.numId)?.get(level) ?? list.levels[level]?.start ?? 1) === literal.text
}

type Sequence = { converts: Decision[]; settled: boolean; failed: boolean; saved: ReturnType<Counters["save"]>; numId?: string }

// The decisions numbering the document with `list`, sequence by sequence; the
// score is the paragraphs the list would then number. `settledOnly` converts
// only within sequences the list already numbers.
function plan(list: ClauseList, paragraphs: Paragraph[], settledOnly: boolean, firstRestart: number) {
  const counters = new Counters(list.levels, list.overrides)
  const decisions: Decision[] = []
  const restarts: string[] = []
  let used = false
  let members = 0
  let sequence: Sequence | undefined
  const close = () => {
    if (!sequence) return
    if (sequence.settled) decisions.push(...sequence.converts)
    else if (!sequence.failed && !settledOnly) {
      // A sequence after one the list already numbers restarts it.
      const numId = used ? String(firstRestart + restarts.length) : list.numId
      if (used) restarts.push(numId)
      used = true
      decisions.push(...sequence.converts.map(row => ({ ...row, numId })))
    } else counters.restore(sequence.saved)
    sequence = undefined
  }
  for (const paragraph of paragraphs) {
    if (paragraph.table) continue
    const member = paragraph.member?.abstract === list.abstract ? paragraph.member : undefined
    if (member && !paragraph.split) {
      counters.enter(member.numId)
      if (!sequence || (member.level === 0 && counters.next(0, member.numId) === counters.start(0, member.numId))) {
        close()
        sequence = { converts: [], settled: true, failed: false, saved: counters.save() }
      }
      sequence.settled = true
      sequence.numId = member.numId
      used = true
      members++
      counters.advance(member.level, member.numId)
      if (paragraph.literal && labelOf(list.levels, member.level, at => counters.shown(at, member.numId)) === paragraph.literal.text)
        decisions.push({ paragraph, level: member.level, kind: "strip", numId: member.numId })
      continue
    }
    if (member) {
      // A later part of a split clause: a deeper level of its list, or out of it.
      const numId = sequence?.numId ?? member.numId
      const level = paragraph.literal ? literalLevel(paragraph.literal, list, counters, numId) : undefined
      if (level !== undefined && level > member.level) {
        counters.advance(level, numId)
        decisions.push({ paragraph, level, kind: "convert", numId })
      } else decisions.push({ paragraph, level: 0, kind: "detach", numId: "0" })
      continue
    }
    if (paragraph.member || !paragraph.literal) continue
    if (paragraph.heading && opensSequence(paragraph.literal, list)) {
      close()
      sequence = { converts: [], settled: false, failed: false, saved: counters.save() }
      counters.values = []
      counters.enter(list.numId)
      counters.advance(0, list.numId)
      sequence.converts.push({ paragraph, level: 0, kind: "convert", numId: list.numId })
      continue
    }
    if (!sequence || sequence.failed) continue
    const numId = sequence.numId ?? list.numId
    const level = literalLevel(paragraph.literal, list, counters, numId)
    if (level === undefined) {
      if (!sequence.settled) sequence.failed = true
      continue
    }
    counters.advance(level, numId)
    sequence.converts.push({ paragraph, level, kind: "convert", numId })
  }
  close()
  return { decisions, restarts, score: members + decisions.filter(row => row.kind === "convert").length }
}

// ---- writing the list

// A house level is bound to a heading style when every clause at that level
// uses it, it carries no other level, and after the first such clause most
// paragraphs of that style are clauses, so a heading of that style added later
// in Word is numbered too (headings before the clauses, such as the parties or
// recitals, keep no number).
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
    const first = paragraphs.indexOf(converted.find(row => row.paragraph.style === style)!.paragraph)
    const clauses = converted.filter(row => row.paragraph.style === style).length
    const others = paragraphs.filter((row, at) => at > first && row.style === style && !row.table && !converted.some(entry => entry.paragraph === row)).length
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

// A stable eight-digit hex identifier (FNV-1a) for the house list's w:nsid and w:tmpl.
function hexId(text: string) {
  let hash = 0x811c9dc5
  for (const code of strToU8(text)) hash = Math.imul(hash ^ code, 0x01000193) >>> 0
  return hash.toString(16).toUpperCase().padStart(8, "0")
}

function writeHouseList(numbering: Document, list: ClauseList, texts: string[], bound: Map<number, string>) {
  const root = numbering.documentElement
  const abstract = element(numbering, "abstractNum", { abstractNumId: list.abstract })
  abstract.appendChild(element(numbering, "nsid", { val: hexId(`doc.haus clause list ${texts.join(" ")}`) }))
  abstract.appendChild(element(numbering, "multiLevelType", { val: "multilevel" }))
  abstract.appendChild(element(numbering, "tmpl", { val: hexId(`doc.haus clause template ${texts.join(" ")}`) }))
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

// A num instance of the clause list that starts its first level afresh.
function writeRestart(numbering: Document, list: ClauseList, lists: Lists, numId: string) {
  const num = element(numbering, "num", { numId })
  num.appendChild(element(numbering, "abstractNumId", { val: lists.numDefinition.get(list.numId) ?? list.abstract }))
  const override = element(numbering, "lvlOverride", { ilvl: "0" })
  override.appendChild(element(numbering, "startOverride", { val: String(list.levels[0]?.start ?? 1) }))
  num.appendChild(override)
  numbering.documentElement.insertBefore(num, children(numbering.documentElement, "numIdMacAtCleanup")[0] ?? null)
}

function bindStyle(style: Element, level: number, numId: string) {
  const doc = style.ownerDocument!
  const pPr = child(style, "pPr") ?? place(style, element(doc, "pPr"), STYLE_ORDER)
  const numPr = element(doc, "numPr")
  numPr.appendChild(element(doc, "ilvl", { val: String(level) }))
  numPr.appendChild(element(doc, "numId", { val: numId }))
  place(pPr, numPr, PPR_ORDER)
}

// The level a style numbers when its numbering names no level: the level that
// names the style, else the first.
function levelOfStyle(list: ClauseList, style: string) {
  return Math.max(0, list.levels.findIndex(level => level?.style === style))
}

// The paragraph style of each level of the document's own list: the level's
// w:pStyle, else a style whose numbering names the list at that level.
function boundLevelStyles(list: ClauseList, catalog: Catalog, lists: Lists) {
  const styles = new Map<number, string>()
  list.levels.forEach((level, index) => { if (level?.style && catalog.elements.has(level.style)) styles.set(index, level.style) })
  for (const [style, node] of catalog.elements) {
    const own = child(child(node, "pPr"), "numPr")
    const numId = value(child(own, "numId"))
    if (node.getAttributeNS(W, "type") !== "paragraph" || numId === undefined || lists.numToAbstract.get(numId) !== list.abstract) continue
    const ilvl = value(child(own, "ilvl"))
    const level = ilvl !== undefined ? Number(ilvl) : levelOfStyle(list, style)
    if (!styles.has(level)) styles.set(level, style)
  }
  return styles
}
