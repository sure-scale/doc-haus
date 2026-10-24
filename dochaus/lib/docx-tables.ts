import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { decodeString } from "micromark-util-decode-string"

// The tables of a Word body, read as grids of cell text, and the carry of an
// original's tables into a body rendered from Markdown. The canonical source
// text gives every cell as its own paragraph and an empty cell as blank lines,
// so a revision drafted from it loses the tables' rows and empty cells; and a
// Markdown table renders in the renderer's own table style. A revised body's
// tables are therefore matched to the original's by content and by the
// passage each follows, in order, and each matched table is replaced by the
// original's own table element: as it stands when every cell reads as the
// original's, or with only the cells whose text differs rewritten, and with
// rows added or removed where the drafted table adds or drops them. A table
// whose columns change, or whose original holds content that cannot move
// between packages, keeps the drafted table. A carried table keeps the
// spacing, size and font its paragraphs had from the original's defaults and
// styles. What happened to each original table is returned, so the caller
// records it.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const A = "http://schemas.openxmlformats.org/drawingml/2006/main"
const XML_NS = "http://www.w3.org/XML/1998/namespace"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const BODY = "word/document.xml"
const STYLES = "word/styles.xml"
/** A drafted table is a counterpart of an original one from this share of shared cell text. */
const MATCH = 0.5
/** A counterpart is preferred by up to this much for the passage it follows:
 * all of it when its passage has the original table's words, else the share
 * of those words the two passages have in common. */
const PLACED = 0.25
// Revision marks, notes, comments, fields, bound content controls and
// embedded objects refer to parts or ids of their own package; a table holding
// one is not carried. List numbering is refused separately (refusal).
const UNCARRIED = new Set(["ins", "del", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge", "rPrChange", "pPrChange", "tblPrChange", "trPrChange",
  "tcPrChange", "tblGridChange", "tblPrExChange", "numberingChange", "footnoteReference", "endnoteReference", "commentReference", "commentRangeStart",
  "commentRangeEnd", "fldSimple", "fldChar", "dataBinding", "drawing", "pict", "object", "altChunk", "sectPr", "gridBefore", "gridAfter"])
// Characters a Markdown table row would read as markup, given in pipeTable as
// character references, which the renderer shows as the character itself:
// escapes, emphasis, code and strike marks, column bars, raw HTML and line
// breaks, references, anchor tokens, images and link destinations.
const MARKUP = /[\\*`~|<]|&(?=#?[A-Za-z0-9]+;)|\{(?=#)|!(?=\[)|(?<=\])\(/g

/** One body table: the visible text of each cell by grid position (an empty
 * or merged-over position reads ""), and the text of the body paragraph it
 * follows ("" when it follows another table or opens the body). */
export type DocxTable = { rows: string[][]; follows: string }

/** What became of one original table in the revised body. `table` is its
 * 1-based position among the original's tables. */
export type TableCarry = { table: number; follows: string } & (
  | { outcome: "carried" }
  | { outcome: "edited"; cells: string[]; added: number; removed: string[] }
  | { outcome: "drafted"; reason: string }
  | { outcome: "missing" }
)

type Cell = { node: Element; text: string; span: number; continued: boolean }
type Grid = { node: Element; rows: { node: Element; cells: Cell[] }[] }

/** The top-level tables of a .docx body, in order, including those inside
 * block-level content controls. */
export function readDocxTables(bytes: Uint8Array): DocxTable[] {
  const body = unzipSync(bytes)[BODY]
  if (!body) return []
  return bodyTables(parse(strFromU8(body))).map(({ node, follows }) => ({ rows: gridText(grid(node)), follows }))
}

/** A table as Markdown pipe rows, its first row as the header row, with empty
 * cells kept empty, a cell's lines joined by "<br>" and every character a
 * Markdown row would read as markup given as a character reference, so the
 * rows render back to the cells' own text. */
export function pipeTable(rows: string[][]) {
  const width = Math.max(1, ...rows.map(row => row.length))
  const cell = (text: string) => text.replace(MARKUP, character => `&#${character.codePointAt(0)};`).split("\n").join("<br>")
  const line = (row: string[]) => `| ${Array.from({ length: width }, (_, at) => cell(row[at] ?? "")).join(" | ")} |`
  return [line(rows[0] ?? []), `| ${Array.from({ length: width }, () => "---").join(" | ")} |`, ...rows.slice(1).map(line)].join("\n")
}

/** Each original table's counterpart among the drafted tables, by index: the
 * order-keeping pairing with the most shared cell text, where a pair shares at
 * least half its cell text or follows the same passage with as many columns,
 * and where a counterpart that follows the same passage as its original
 * outweighs one that only reads alike. */
export function matchTables(originals: DocxTable[], drafted: DocxTable[]): (number | undefined)[] {
  const score = originals.map(original => drafted.map(table => pairScore(original, table)))
  const best: number[][] = Array.from({ length: originals.length + 1 }, () => new Array<number>(drafted.length + 1).fill(0))
  const paired = (i: number, j: number) => score[i]![j]! + best[i + 1]![j + 1]!
  for (let i = originals.length - 1; i >= 0; i--)
    for (let j = drafted.length - 1; j >= 0; j--)
      best[i]![j] = Math.max(best[i + 1]![j]!, best[i]![j + 1]!, score[i]![j]! > 0 ? paired(i, j) : 0)
  const pairs: (number | undefined)[] = new Array(originals.length).fill(undefined)
  for (let i = 0, j = 0; i < originals.length && j < drafted.length;) {
    if (score[i]![j]! > 0 && best[i]![j] === paired(i, j)) pairs[i++] = j++
    else if (best[i]![j] === best[i + 1]![j]) i++
    else j++
  }
  return pairs
}

/** The rendered body with each original table carried over its drafted
 * counterpart, and what became of every original table. The carried tables'
 * cells carry no paragraph pointers; the caller gives them pointers. */
export function carryOriginalTables(rendered: Uint8Array, original: Uint8Array): { bytes: Uint8Array; carries: TableCarry[] } {
  const originalParts = unzipSync(original)
  const parts = unzipSync(rendered)
  if (!originalParts[BODY] || !parts[BODY]) return { bytes: rendered, carries: [] }
  const originals = bodyTables(parse(strFromU8(originalParts[BODY]))).map(({ node, follows }) => ({ node, table: { rows: gridText(grid(node)), follows } }))
  if (!originals.length) return { bytes: rendered, carries: [] }
  const target = parse(strFromU8(parts[BODY]))
  const drafted = bodyTables(target).map(({ node, follows }) => ({ node, table: { rows: gridText(grid(node)), follows } }))
  const pairs = matchTables(originals.map(row => row.table), drafted.map(row => row.table))
  const sourceStyles = originalParts[STYLES] ? parse(strFromU8(originalParts[STYLES])) : undefined
  const sourceById = sourceStyles ? stylesById(sourceStyles) : new Map<string, Element>()
  const styles = new Set<string>()
  const placed: Element[] = []
  const carries = originals.map(({ node, table }, index): TableCarry => {
    const at = pairs[index]
    const base = { table: index + 1, follows: table.follows }
    if (at === undefined) return { ...base, outcome: "missing" }
    const counterpart = drafted[at]!
    const refused = refusal(node, sourceById)
    if (refused) return { ...base, outcome: "drafted", reason: refused }
    const carried = carry(target.importNode(node, true) as Element, counterpart.table.rows)
    if ("reason" in carried) return { ...base, outcome: "drafted", reason: carried.reason }
    dropRenderedSpacer(counterpart.node)
    counterpart.node.parentNode!.replaceChild(carried.node, counterpart.node)
    referencedStyles(carried.node).forEach(id => styles.add(id))
    placed.push(carried.node)
    return carried.cells.length || carried.added || carried.removed.length
      ? { ...base, outcome: "edited", cells: carried.cells, added: carried.added, removed: carried.removed }
      : { ...base, outcome: "carried" }
  })
  if (!placed.length) return { bytes: rendered, carries }
  const targetStyles = parts[STYLES] ? parse(strFromU8(parts[STYLES])) : undefined
  if (targetStyles && sourceStyles && copyStyles(targetStyles, sourceStyles, styles)) parts[STYLES] = strToU8(new XMLSerializer().serializeToString(targetStyles))
  const theme = Object.keys(originalParts).find(name => /^word\/theme\/[^/]+\.xml$/.test(name))
  keepOriginalDefaults(placed, formatting(sourceStyles, theme ? parse(strFromU8(originalParts[theme]!)) : undefined), formatting(targetStyles, undefined))
  parts[BODY] = strToU8(new XMLSerializer().serializeToString(target))
  return { bytes: zipSync(parts, { mtime: new Date(1980, 0, 1) }), carries }
}

/** The original table element, already imported into the rendered body, made
 * to read as the drafted rows: rows paired by position when the counts agree,
 * else by their first cell; a cell is rewritten only where its text differs. */
function carry(node: Element, drafted: string[][]): { node: Element; cells: string[]; added: number; removed: string[] } | { reason: string } {
  clean(node)
  const table = grid(node)
  const original = gridText(table)
  if (!table.rows.length) return { reason: "the original table has no rows" }
  const width = original[0]!.length
  if (drafted.some(row => row.length !== width)) return { reason: `the drafted table has ${drafted[0]?.length ?? 0} columns where the original has ${width}` }
  const pairs = original.length === drafted.length ? original.map((_, at) => at) : alignRows(original, drafted)
  if (pairs.every(at => at === undefined)) return { reason: "no row of the drafted table begins as a row of the original does" }
  const cells: string[] = []
  const removed: string[] = []
  for (const [index, row] of table.rows.entries()) {
    const at = pairs[index]
    if (at === undefined) {
      removed.push(original[index]!.find(text => text.trim()) ?? `row ${index + 1}`)
      continue
    }
    const issue = rewriteRow(row.cells, original[index]!, drafted[at]!, `row ${index + 1}`, cells)
    if (issue) return { reason: issue }
  }
  // A drafted row the original lacks takes the form of the original row it
  // follows, or of the first body row when it follows the header row or none.
  let added = 0
  const bodyRow = (table.rows.find((_, index) => index > 0 && pairs[index] !== undefined) ?? table.rows[1] ?? table.rows[0])!.node
  let anchor: Element | undefined
  for (const [at, text] of drafted.entries()) {
    const index = pairs.indexOf(at)
    if (index >= 0) { anchor = table.rows[index]!.node; continue }
    const model = anchor && anchor !== table.rows[0]!.node ? anchor : bodyRow
    const row = model.cloneNode(true) as Element
    // An added row is a body row, never a repeated header row.
    Array.from(row.getElementsByTagNameNS(W, "tblHeader")).forEach(mark => mark.parentNode?.removeChild(mark))
    const cells = wrap(node, row).rows[0]!.cells
    if (cells.some(cell => cell.span > 1 || cell.continued)) return { reason: "a drafted row is added beside merged cells" }
    if (text.slice(cells.length).some(value => value.trim())) return { reason: `a drafted row is added with text in a column the original's row ${table.rows.findIndex(row => row.node === model) + 1} does not have` }
    for (const [column, cell] of cells.entries()) {
      const issue = rewriteCell(cell.node, text[column] ?? "")
      if (issue) return { reason: `a drafted row is added in the form of a row whose cell holds ${issue}` }
    }
    const reference = anchor ? anchor.nextSibling : table.rows.find((_, index) => pairs[index] !== undefined)!.node
    ;(anchor ?? reference as Element).parentNode!.insertBefore(row, reference)
    anchor = row
    added++
  }
  table.rows.forEach((row, index) => { if (pairs[index] === undefined) remove(row.node) })
  return { node, cells, added, removed }
}

function rewriteRow(cells: Cell[], original: string[], drafted: string[], label: string, changed: string[]) {
  let column = 0
  for (const cell of cells) {
    const positions = Array.from({ length: cell.span }, (_, offset) => column + offset)
    column += cell.span
    if (positions.every(at => same(original[at] ?? "", drafted[at] ?? ""))) continue
    // A merged-over position holds no text of its own.
    if (cell.continued || positions.slice(1).some(at => (drafted[at] ?? "").trim()))
      return `the drafted table writes into a merged cell (${label}, column ${positions[0]! + 1})`
    const issue = rewriteCell(cell.node, drafted[positions[0]!] ?? "")
    if (issue) return `the drafted table changes a cell holding ${issue} (${label}, column ${positions[0]! + 1})`
    changed.push(`${label}, column ${positions[0]! + 1}`)
  }
  // A row with fewer cells than the grid has no place for text beyond them.
  const beyond = drafted.findIndex((text, at) => at >= column && text.trim())
  if (beyond >= 0) return `the drafted table writes into a position ${label} does not have (column ${beyond + 1})`
}

/** Original rows paired with drafted rows by their first cell, in order (a longest common subsequence). */
function alignRows(original: string[][], drafted: string[][]) {
  const key = (row: string[]) => loose(row[0] ?? "") || loose(row.join(" "))
  const a = original.map(key)
  const b = drafted.map(key)
  const lengths: number[][] = Array.from({ length: a.length + 1 }, () => new Array<number>(b.length + 1).fill(0))
  for (let i = a.length - 1; i >= 0; i--)
    for (let j = b.length - 1; j >= 0; j--)
      lengths[i]![j] = a[i] === b[j] ? lengths[i + 1]![j + 1]! + 1 : Math.max(lengths[i + 1]![j]!, lengths[i]![j + 1]!)
  const pairs: (number | undefined)[] = new Array(a.length).fill(undefined)
  for (let i = 0, j = 0; i < a.length && j < b.length;) {
    if (a[i] === b[j]) pairs[i++] = j++
    else if (lengths[i + 1]![j]! >= lengths[i]![j + 1]!) i++
    else j++
  }
  return pairs
}

/** How well a drafted table stands for an original one; 0 when it does not. */
function pairScore(original: DocxTable, drafted: DocxTable) {
  const content = shared(original.rows, drafted.rows)
  const passage = wording(original.follows, drafted.follows)
  const placed = passage === 1 && (original.rows[0]?.length ?? 0) === (drafted.rows[0]?.length ?? 0)
  if (content < MATCH && !placed) return 0
  return Math.max(content, placed ? MATCH : 0) + PLACED * passage
}

/** The share of non-empty cell text two tables have in common. */
function shared(a: string[][], b: string[][]) {
  const texts = (rows: string[][]) => rows.flat().map(loose).filter(Boolean)
  const left = texts(a)
  const right = texts(b)
  if (!left.length || !right.length)
    return !left.length && !right.length && a.length === b.length && (a[0]?.length ?? 0) === (b[0]?.length ?? 0) ? 1 : 0
  return inCommon(left, right)
}

/** The share of words two passages have in common; 0 when either has none. */
function wording(a: string, b: string) {
  const words = (text: string) => plain(text).match(/[\p{L}\p{N}]+/gu) ?? []
  const left = words(a)
  const right = words(b)
  return left.length && right.length ? inCommon(left, right) : 0
}

/** The share of two lists' items both hold, counting repeats. */
function inCommon(left: string[], right: string[]) {
  const counts = new Map<string, number>()
  left.forEach(text => counts.set(text, (counts.get(text) ?? 0) + 1))
  const common = right.filter(text => {
    const count = counts.get(text) ?? 0
    if (count) counts.set(text, count - 1)
    return count > 0
  }).length
  return common / Math.max(left.length, right.length)
}

/** Cell text as compared: Unicode-normalized, typographic dashes, quotes and spaces made plain, whitespace collapsed. */
function normalized(text: string) {
  return text.normalize("NFKC").replace(/[\u2010-\u2015\u2212]/g, "-").replace(/[\u2018\u2019\u201A\u201B]/g, "'").replace(/[\u201C\u201D\u201E\u201F]/g, "\"")
    .replace(/\s+/g, " ").trim()
}

/** Whether a drafted cell reads as the original one: as written, or as the
 * renderer shows the original's text written into a row without its
 * character references, which a drafter may leave out. */
function same(original: string, drafted: string) {
  return normalized(original) === normalized(drafted) || normalized(asWritten(original)) === normalized(drafted)
}

/** Text as matched: compared, and in any case. */
export function plain(text: string) {
  return normalized(text).toLowerCase()
}

/** Text as matched between an original and a drafted table, markup aside. */
function loose(text: string) {
  return plain(asWritten(text))
}

/** Text as the renderer shows it when it is written into a Markdown row as it
 * stands: links show their text, an escape its character, emphasis, code and
 * strike marks nothing, a "<br>" a line break and a reference its character. */
function asWritten(text: string) {
  return text.replace(/!?\[([^\]]*)\]\([^)]*\)/g, "$1")
    .replace(/\\([\s\S]?)|[*`]|~~|<br\s*\/?>/gi, (markup, escaped: string | undefined) => escaped ?? (markup.startsWith("<") ? "\n" : ""))
    .replace(/&(?:#[0-9]+|#x[0-9a-f]+|[a-z][a-z0-9]*);/gi, reference => decodeString(reference))
}

function parse(xml: string) {
  return new DOMParser().parseFromString(xml, "application/xml")
}

function children(node: Element | undefined, name: string) {
  if (!node) return []
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1 && (child as Element).namespaceURI === W && (child as Element).localName === name)
}

/** The child elements of the given names, read through the content controls
 * and custom XML elements that may wrap them, in document order. */
function content(node: Element, names: string[]): Element[] {
  return Array.from(node.childNodes).flatMap(child => {
    if (child.nodeType !== 1 || (child as Element).namespaceURI !== W) return []
    const element = child as Element
    if (names.includes(element.localName)) return [element]
    if (element.localName === "sdt") return children(element, "sdtContent").flatMap(inner => content(inner, names))
    if (element.localName === "customXml") return content(element, names)
    return []
  })
}

/** The body's tables, with the text of the paragraph each follows: the
 * nearest non-empty paragraph before it, or "" when another table comes first. */
function bodyTables(document: Document) {
  const body = document.getElementsByTagNameNS(W, "body")[0]
  if (!body) return []
  const blocks = content(body, ["p", "tbl"])
  return blocks.flatMap((node, at) => {
    if (node.localName !== "tbl") return []
    let follows = ""
    for (let before = at - 1; before >= 0; before--) {
      if (blocks[before]!.localName === "tbl") break
      follows = paragraphText(blocks[before]!, blocks[before]!).replace(/\s+/g, " ").trim()
      if (follows) break
    }
    return [{ node, follows }]
  })
}

function grid(node: Element): Grid {
  return { node, rows: content(node, ["tr"]).map(row => ({ node: row, cells: content(row, ["tc"]).map(cell => {
    const properties = children(cell, "tcPr")[0]
    const span = Number(properties && children(properties, "gridSpan")[0]?.getAttributeNS(W, "val")) || 1
    const merge = properties && children(properties, "vMerge")[0]
    return { node: cell, text: cellText(cell), span: Math.max(1, span), continued: Boolean(merge) && merge!.getAttributeNS(W, "val") !== "restart" }
  }) })) }
}

/** Each row's text by grid position; a ragged row is padded with empty positions. */
function gridText(table: Grid) {
  const rows = table.rows.map(row => row.cells.flatMap(cell => [cell.continued ? "" : cell.text, ...new Array<string>(cell.span - 1).fill("")]))
  const width = Math.max(0, ...rows.map(row => row.length))
  return rows.map(row => [...row, ...new Array<string>(width - row.length).fill("")])
}

function cellText(cell: Element) {
  return Array.from(cell.getElementsByTagNameNS(W, "p")).map(paragraph => paragraphText(paragraph, paragraph).trim()).join("\n").trim()
}

/** A paragraph's visible text: deleted and hidden runs and nested paragraphs aside. */
function paragraphText(node: Node, paragraph: Element): string {
  if (node.nodeType !== 1) return ""
  const element = node as Element
  if (element.namespaceURI === W) {
    if (["del", "moveFrom"].includes(element.localName) || element.localName === "p" && element !== paragraph) return ""
    if (element.localName === "r" && hidden(element)) return ""
    if (element.localName === "t") return element.textContent ?? ""
    if (element.localName === "tab") return " "
    if (["br", "cr"].includes(element.localName)) return "\n"
    if (element.localName === "noBreakHyphen") return "-"
  }
  return Array.from(node.childNodes).map(child => paragraphText(child, paragraph)).join("")
}

function hidden(run: Element) {
  return children(children(run, "rPr")[0], "vanish").some(mark => !["0", "false", "off"].includes(mark.getAttributeNS(W, "val") ?? ""))
}

/** Why a table cannot be carried into another package, if it cannot. */
function refusal(table: Element, styles: Map<string, Element>) {
  const elements = Array.from(table.getElementsByTagName("*"))
  if (elements.some(node => node.namespaceURI !== W)) return "the original table holds content outside plain Word text (a drawing, an equation or an Office extension)"
  const refused = elements.find(node => UNCARRIED.has(node.localName))
  if (refused) return `the original table holds ${["fldSimple", "fldChar"].includes(refused.localName) ? "a field" : `a ${refused.localName} element`}, which refers to its own package`
  if (elements.some(node => node.localName === "numPr" && numbered(node))) return "the original table holds list numbering, which refers to its own package"
  if (referencedStyles(table).some(id => styleChain(styles, id).some(style => Array.from(style.getElementsByTagNameNS(W, "numPr")).some(numbered))))
    return "the original table uses a style with list numbering, which refers to its own package"
  if (elements.some(node => node.localName === "hyperlink" && node.getAttributeNS(W, "anchor")))
    return "the original table links to a bookmark, which is not carried"
  if (elements.some(node => Array.from(node.attributes).some(attribute => attribute.namespaceURI === R))) return "the original table links to another part of its package"
}

/** Whether list numbering names a list; numbering 0 turns numbering off. */
function numbered(numbering: Element) {
  const id = children(numbering, "numId")[0]?.getAttributeNS(W, "val")
  return Boolean(id) && id !== "0"
}

/** Leaves only Word markup the rendered body can hold: attributes of other
 * namespaces (paragraph ids, for instance) and bookmarks, whose ids belong to
 * the original package, are removed. */
function clean(table: Element) {
  for (const node of [table, ...Array.from(table.getElementsByTagName("*"))]) {
    for (const attribute of Array.from(node.attributes))
      if (attribute.namespaceURI !== W && attribute.namespaceURI !== XML_NS && !/^xmlns(:|$)/.test(attribute.name)) node.removeAttributeNode(attribute)
    if (node.namespaceURI === W && ["bookmarkStart", "bookmarkEnd"].includes(node.localName)) node.parentNode?.removeChild(node)
  }
}

/** A cell made to read as `text`, or what in it prevents that. Where the text
 * has as many lines as the cell's, only the paragraphs whose lines differ are
 * rewritten, each keeping its own formatting; otherwise each line becomes a
 * paragraph formatted as the cell's paragraph in its place. A cell holding a
 * nested table or a content control is not rewritten. */
function rewriteCell(cell: Element, text: string): string | undefined {
  if (["tbl", "sdt", "customXml"].some(name => cell.getElementsByTagNameNS(W, name).length)) return "a nested table or a content control"
  const paragraphs = children(cell, "p")
  if (!paragraphs.length) return "no paragraph"
  // The cell's lines by paragraph, its leading and trailing blank lines aside, as cellText reads them.
  const lines = paragraphs.flatMap((paragraph, index) => paragraphText(paragraph, paragraph).trim().split("\n").map(line => ({ index, line })))
  let start = 0
  let end = lines.length
  while (start < end && !lines[start]!.line.trim()) start++
  while (end > start && !lines[end - 1]!.line.trim()) end--
  const kept = lines.slice(start, end)
  const drafted = text.split("\n")
  if (kept.length && drafted.length === kept.length) {
    for (const index of new Set(kept.map(line => line.index))) {
      const own = kept.flatMap((line, at) => line.index === index ? [at] : [])
      if (own.some(at => !same(kept[at]!.line, drafted[at]!))) writeParagraph(paragraphs[index]!, own.map(at => drafted[at]!))
    }
    return
  }
  const models = kept.length ? kept.map(line => paragraphs[line.index]!) : [paragraphs[0]!]
  for (const [at, line] of drafted.entries()) {
    const next = models[Math.min(at, models.length - 1)]!.cloneNode(true) as Element
    writeParagraph(next, [line])
    cell.insertBefore(next, paragraphs[0]!)
  }
  paragraphs.forEach(paragraph => cell.removeChild(paragraph))
}

/** A paragraph's content replaced by lines joined by breaks, keeping its
 * paragraph formatting and the formatting of its first visible run, or of its
 * paragraph mark when it has none. */
function writeParagraph(paragraph: Element, lines: string[]) {
  const document = paragraph.ownerDocument!
  const properties = children(paragraph, "pPr")[0]
  const model = Array.from(paragraph.getElementsByTagNameNS(W, "r")).find(run => run.getElementsByTagNameNS(W, "t").length && !hidden(run))
  const runProperties = model ? children(model, "rPr")[0] : children(properties, "rPr")[0]
  for (const child of Array.from(paragraph.childNodes)) if (child !== properties) paragraph.removeChild(child)
  if (!lines.some(line => line.trim())) return
  const run = document.createElementNS(W, "w:r")
  const copied = document.createElementNS(W, "w:rPr")
  for (const property of Array.from(runProperties?.childNodes ?? []))
    if (property.nodeType === 1 && !["vanish", "specVanish"].includes((property as Element).localName)) copied.appendChild(property.cloneNode(true))
  if (copied.childNodes.length) run.appendChild(copied)
  for (const [at, line] of lines.entries()) {
    if (at) run.appendChild(document.createElementNS(W, "w:br"))
    if (!line.trim()) continue
    const t = document.createElementNS(W, "w:t")
    t.setAttributeNS(XML_NS, "xml:space", "preserve")
    t.appendChild(document.createTextNode(line.trim()))
    run.appendChild(t)
  }
  paragraph.appendChild(run)
}

/** A row removed, with a row-level content control it leaves empty. */
function remove(row: Element) {
  const parent = row.parentNode as Element
  parent.removeChild(row)
  if (parent.localName === "sdtContent" && !content(parent, ["tr"]).length) parent.parentNode?.parentNode?.removeChild(parent.parentNode)
}

/** A detached row read in the context of its table, for its cells. */
function wrap(table: Element, row: Element): Grid {
  const holder = table.ownerDocument!.createElementNS(W, "w:tbl")
  holder.appendChild(row)
  const read = grid(holder)
  holder.removeChild(row)
  return read
}

/** The renderer closes each Markdown table with an empty spacing paragraph; a
 * carried table stands as in the original, except that the paragraph stays
 * between two tables, which would otherwise read as one. */
function dropRenderedSpacer(table: Element) {
  const element = (node: Node | null) => {
    while (node && node.nodeType !== 1) node = node.nextSibling
    return node as Element | null
  }
  const paragraph = element(table.nextSibling)
  if (!paragraph || paragraph.namespaceURI !== W || paragraph.localName !== "p" || paragraphText(paragraph, paragraph).trim()) return
  const after = element(paragraph.nextSibling)
  if (after?.namespaceURI === W && after.localName === "tbl") return
  const inner = Array.from(paragraph.childNodes).filter(node => node.nodeType === 1 && (node as Element).localName !== "pPr")
  const spacing = paragraph.getElementsByTagNameNS(W, "spacing")[0]
  if (!inner.length && spacing?.getAttributeNS(W, "before") === "80" && spacing.getAttributeNS(W, "after") === "120") paragraph.parentNode!.removeChild(paragraph)
}

function referencedStyles(table: Element) {
  return ["tblStyle", "pStyle", "rStyle"].flatMap(name => Array.from(table.getElementsByTagNameNS(W, name)).map(node => node.getAttributeNS(W, "val") ?? "")).filter(Boolean)
}

function stylesById(document: Document) {
  return new Map(Array.from(document.getElementsByTagNameNS(W, "style")).map(node => [node.getAttributeNS(W, "styleId") ?? "", node]))
}

/** A style and the styles it is based on, nearest first. */
function styleChain(styles: Map<string, Element>, id: string | undefined) {
  const chain: Element[] = []
  for (let next = id; next && styles.has(next) && !chain.includes(styles.get(next)!);) {
    chain.push(styles.get(next)!)
    next = children(styles.get(next)!, "basedOn")[0]?.getAttributeNS(W, "val") || undefined
  }
  return chain
}

/** The rendered styles with every style the carried tables use that they
 * lack, and the styles those are based on or linked to, copied from the
 * original. A table whose styles number its paragraphs is not carried
 * (refusal); list numbering in a style copied only as a link is left out, as
 * its numbering part is not carried. Whether anything was copied. */
function copyStyles(into: Document, from: Document, wanted: Set<string>) {
  const root = into.documentElement
  if (!root || !wanted.size) return false
  const present = stylesById(into)
  const available = stylesById(from)
  const queue = [...wanted]
  let copied = false
  while (queue.length) {
    const id = queue.shift()!
    const style = present.has(id) ? undefined : available.get(id)
    if (!style) continue
    const imported = into.importNode(style, true) as Element
    clean(imported)
    Array.from(imported.getElementsByTagNameNS(W, "numPr")).forEach(node => node.parentNode?.removeChild(node))
    // The rendered body keeps its own default style of each kind.
    if (Array.from(present.values()).some(node => node.getAttributeNS(W, "type") === imported.getAttributeNS(W, "type") && node.getAttributeNS(W, "default") === "1"))
      imported.removeAttributeNS(W, "default")
    root.appendChild(imported)
    present.set(id, imported)
    copied = true
    for (const name of ["basedOn", "link"]) {
      const next = children(imported, name)[0]?.getAttributeNS(W, "val")
      if (next) queue.push(next)
    }
  }
  return copied
}

// What gives a paragraph its spacing and a run its size and font: document
// defaults, then the table style, then the paragraph style, then a run's
// character style, then direct formatting. Where a package has document
// defaults, a value none of them sets is the format's own default.
type Formatting = { styles: Map<string, Element>; paragraph?: string; defaults?: Element; fonts: { major?: string; minor?: string } }
const SPACING = ["before", "after", "line", "lineRule"] as const
const SPACING_DEFAULT: Record<(typeof SPACING)[number], string> = { before: "0", after: "0", line: "240", lineRule: "auto" }
const SIZES = ["sz", "szCs"] as const
const SIZE_DEFAULT = "20"
const PARAGRAPH_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
  "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
  "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
const RUN_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
  "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em",
  "lang", "eastAsianLayout", "specVanish", "oMath"]

function formatting(styles: Document | undefined, theme: Document | undefined): Formatting {
  const byId = styles ? stylesById(styles) : new Map<string, Element>()
  const paragraph = Array.from(byId.values()).find(style => style.getAttributeNS(W, "type") === "paragraph" && ["1", "true", "on"].includes(style.getAttributeNS(W, "default") ?? ""))
  const font = (name: string) => {
    const scheme = theme?.getElementsByTagNameNS(A, name)[0]
    const latin = scheme && Array.from(scheme.childNodes).find(node => node.nodeType === 1 && (node as Element).localName === "latin") as Element | undefined
    return latin?.getAttribute("typeface") || undefined
  }
  return { styles: byId, paragraph: paragraph?.getAttributeNS(W, "styleId") || undefined, defaults: styles?.getElementsByTagNameNS(W, "docDefaults")[0], fonts: { major: font("majorFont"), minor: font("minorFont") } }
}

/** Each carried paragraph and run given, as direct formatting, the spacing,
 * size and font it had in the original where the rendered package would give
 * it other values and its own formatting does not set them; so a carried row
 * keeps its height and its text its face. A run's own theme font is written as
 * the face the original's theme names, since the rendered package has no
 * theme of its own. Conditional table-style formatting and the eastern and
 * complex-script fonts are not compared. Nor is any other formatting the
 * original's default and paragraph styles give, such as alignment,
 * indentation or colour: a carried paragraph takes it from the rendered
 * package's style of the same name. */
function keepOriginalDefaults(tables: Element[], from: Formatting, into: Formatting) {
  for (const carried of tables) for (const paragraph of Array.from(carried.getElementsByTagNameNS(W, "p"))) {
    let table = paragraph.parentNode as Element | null
    while (table && !(table.namespaceURI === W && table.localName === "tbl")) table = table.parentNode as Element | null
    const tableStyle = children(children(table ?? undefined, "tblPr")[0], "tblStyle")[0]?.getAttributeNS(W, "val") || undefined
    const paragraphStyle = children(children(paragraph, "pPr")[0], "pStyle")[0]?.getAttributeNS(W, "val") || undefined
    const paragraphSources = (look: Formatting) => [...styleChain(look.styles, paragraphStyle ?? look.paragraph), ...styleChain(look.styles, tableStyle)]
      .map(style => children(style, "pPr")[0]).concat(children(children(look.defaults, "pPrDefault")[0], "pPr")[0])
    const direct = children(children(paragraph, "pPr")[0], "spacing")[0]
    const spacing = SPACING.flatMap(name => {
      if (direct?.getAttributeNS(W, name)) return []
      const value = (look: Formatting) => paragraphSources(look).map(pPr => children(pPr, "spacing")[0]?.getAttributeNS(W, name)).find(Boolean) ?? (look.defaults ? SPACING_DEFAULT[name] : undefined)
      const was = value(from)
      return was !== undefined && was !== value(into) ? [[name, was] as const] : []
    })
    if (spacing.some(([name]) => name === "line") && !direct?.getAttributeNS(W, "lineRule") && !spacing.some(([name]) => name === "lineRule"))
      spacing.push(["lineRule", paragraphSources(from).map(pPr => children(pPr, "spacing")[0]?.getAttributeNS(W, "lineRule")).find(Boolean) ?? "auto"])
    if (spacing.length) {
      const element = property(first(paragraph, "pPr"), "spacing", PARAGRAPH_ORDER)
      spacing.forEach(([name, value]) => element.setAttributeNS(W, `w:${name}`, value))
    }
    // Each run, and the paragraph mark, which sets an empty paragraph's height.
    const holders: [Element, () => Element][] = [
      ...Array.from(paragraph.getElementsByTagNameNS(W, "r")).map(run => [run, () => first(run, "rPr")] as [Element, () => Element]),
      [paragraph, () => property(first(paragraph, "pPr"), "rPr", PARAGRAPH_ORDER)],
    ]
    for (const [holder, ensure] of holders) {
      const own = holder === paragraph ? children(children(paragraph, "pPr")[0], "rPr")[0] : children(holder, "rPr")[0]
      const characterStyle = children(own, "rStyle")[0]?.getAttributeNS(W, "val") || undefined
      const runSources = (look: Formatting) => [...styleChain(look.styles, characterStyle), ...styleChain(look.styles, paragraphStyle ?? look.paragraph), ...styleChain(look.styles, tableStyle)]
        .map(style => children(style, "rPr")[0]).concat(children(children(look.defaults, "rPrDefault")[0], "rPr")[0])
      const size = (name: string) => (look: Formatting) => runSources(look).map(rPr => children(rPr, name)[0]?.getAttributeNS(W, "val")).find(Boolean) ?? (look.defaults ? SIZE_DEFAULT : undefined)
      const face = (look: Formatting) => {
        for (const rPr of runSources(look)) {
          const fonts = children(rPr, "rFonts")[0]
          const theme = fonts?.getAttributeNS(W, "asciiTheme")
          if (theme) return look.fonts[theme.startsWith("major") ? "major" : "minor"] ?? ""
          if (fonts?.getAttributeNS(W, "ascii")) return fonts.getAttributeNS(W, "ascii")
        }
      }
      for (const name of SIZES) {
        if (children(own, name).length) continue
        const was = size(name)(from)
        if (was !== undefined && was !== size(name)(into)) property(ensure(), name, RUN_ORDER).setAttributeNS(W, "w:val", was)
      }
      const fonts = children(own, "rFonts")[0]
      if (fonts?.getAttributeNS(W, "asciiTheme")) {
        for (const slot of ["ascii", "hAnsi"]) {
          const theme = fonts.getAttributeNS(W, `${slot}Theme`)
          const named = theme ? from.fonts[theme.startsWith("major") ? "major" : "minor"] : undefined
          if (!named) continue
          // A theme font outranks a named one, so the theme attribute is removed.
          fonts.removeAttributeNS(W, `${slot}Theme`)
          fonts.setAttributeNS(W, `w:${slot}`, named)
        }
        continue
      }
      if (fonts?.getAttributeNS(W, "ascii")) continue
      const was = face(from)
      if (was && was !== face(into)) {
        const element = property(ensure(), "rFonts", RUN_ORDER)
        element.setAttributeNS(W, "w:ascii", was)
        element.setAttributeNS(W, "w:hAnsi", was)
      }
    }
  }
}

/** A property element of `parent`, added in schema order when absent. */
function property(parent: Element, name: string, order: string[]) {
  const found = children(parent, name)[0]
  if (found) return found
  const element = parent.ownerDocument!.createElementNS(W, `w:${name}`)
  const rank = order.indexOf(name)
  const before = Array.from(parent.childNodes).find(node => node.nodeType === 1 && order.indexOf((node as Element).localName) > rank)
  parent.insertBefore(element, before ?? null)
  return element
}

/** The properties element that opens a paragraph or run, added when absent. */
function first(parent: Element, name: "pPr" | "rPr") {
  const found = children(parent, name)[0]
  if (found) return found
  const element = parent.ownerDocument!.createElementNS(W, `w:${name}`)
  parent.insertBefore(element, parent.firstChild)
  return element
}
