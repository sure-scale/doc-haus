import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// The tables of a Word body, read as grids of cell text, and the carry of an
// original's tables into a body rendered from Markdown. The canonical source
// text gives every cell as its own paragraph and an empty cell as blank lines,
// so a revision drafted from it loses the tables' rows and empty cells; and a
// Markdown table renders in the renderer's own table style. A revised body's
// tables are therefore matched to the original's by content, in order, and
// each matched table is replaced by the original's own table element: as it
// stands when every cell reads as the original's, or with only the cells whose
// text differs rewritten, and with rows added or removed where the drafted
// table adds or drops them. A table whose columns change, or whose original
// holds content that cannot move between packages, keeps the drafted table.
// What happened to each original table is returned, so the caller records it.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XML_NS = "http://www.w3.org/XML/1998/namespace"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const BODY = "word/document.xml"
const STYLES = "word/styles.xml"
/** Shown for a "|" inside a cell, which a Markdown table row cannot hold; read back as "|". */
export const PIPE_STAND_IN = "¦"
/** A drafted table is a counterpart of an original one from this share of shared cell text. */
const MATCH = 0.5
// Revision marks, notes, comments, list numbering and embedded objects refer
// to parts or ids of their own package; a table holding one is not carried.
const UNCARRIED = new Set(["ins", "del", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge", "rPrChange", "pPrChange", "tblPrChange", "trPrChange",
  "tcPrChange", "tblGridChange", "tblPrExChange", "numberingChange", "footnoteReference", "endnoteReference", "commentReference", "commentRangeStart",
  "commentRangeEnd", "numPr", "drawing", "pict", "object", "altChunk", "sectPr", "gridBefore", "gridAfter"])

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

/** The top-level tables of a .docx body, in order. */
export function readDocxTables(bytes: Uint8Array): DocxTable[] {
  const body = unzipSync(bytes)[BODY]
  if (!body) return []
  return bodyTables(parse(strFromU8(body))).map(node => ({ rows: gridText(grid(node)), follows: follows(node) }))
}

/** A table as Markdown pipe rows, its first row as the header row, with empty
 * cells kept empty and a cell's lines joined by "<br>". */
export function pipeTable(rows: string[][]) {
  const width = Math.max(1, ...rows.map(row => row.length))
  const line = (row: string[]) => `| ${Array.from({ length: width }, (_, at) => (row[at] ?? "").replaceAll("|", PIPE_STAND_IN).split("\n").join("<br>")).join(" | ")} |`
  return [line(rows[0] ?? []), `| ${Array.from({ length: width }, () => "---").join(" | ")} |`, ...rows.slice(1).map(line)].join("\n")
}

/** Each original table's counterpart among the drafted tables, by index: the
 * order-keeping pairing with the most shared cell text, where a pair shares at
 * least half its cell text or follows the same passage with as many columns. */
export function matchTables(originals: DocxTable[], drafted: DocxTable[]): (number | undefined)[] {
  const score = originals.map(original => drafted.map(table => pairScore(original, table)))
  const best: number[][] = Array.from({ length: originals.length + 1 }, () => new Array<number>(drafted.length + 1).fill(0))
  for (let i = originals.length - 1; i >= 0; i--)
    for (let j = drafted.length - 1; j >= 0; j--)
      best[i]![j] = Math.max(best[i + 1]![j]!, best[i]![j + 1]!, score[i]![j]! >= MATCH ? score[i]![j]! + best[i + 1]![j + 1]! : 0)
  const pairs: (number | undefined)[] = new Array(originals.length).fill(undefined)
  for (let i = 0, j = 0; i < originals.length && j < drafted.length;) {
    if (score[i]![j]! >= MATCH && best[i]![j] === score[i]![j]! + best[i + 1]![j + 1]!) pairs[i++] = j++
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
  const originals = bodyTables(parse(strFromU8(originalParts[BODY]))).map(node => ({ node, table: { rows: gridText(grid(node)), follows: follows(node) } }))
  if (!originals.length) return { bytes: rendered, carries: [] }
  const target = parse(strFromU8(parts[BODY]))
  const drafted = bodyTables(target).map(node => ({ node, table: { rows: gridText(grid(node)), follows: follows(node) } }))
  const pairs = matchTables(originals.map(row => row.table), drafted.map(row => row.table))
  const styles = new Set<string>()
  const carries = originals.map(({ node, table }, index): TableCarry => {
    const at = pairs[index]
    const base = { table: index + 1, follows: table.follows }
    if (at === undefined) return { ...base, outcome: "missing" }
    const counterpart = drafted[at]!
    const refused = uncarried(node)
    if (refused) return { ...base, outcome: "drafted", reason: refused }
    const carried = carry(target.importNode(node, true) as Element, counterpart.table.rows)
    if ("reason" in carried) return { ...base, outcome: "drafted", reason: carried.reason }
    dropRenderedSpacer(counterpart.node)
    counterpart.node.parentNode!.replaceChild(carried.node, counterpart.node)
    referencedStyles(carried.node).forEach(id => styles.add(id))
    return carried.cells.length || carried.added || carried.removed.length
      ? { ...base, outcome: "edited", cells: carried.cells, added: carried.added, removed: carried.removed }
      : { ...base, outcome: "carried" }
  })
  if (!carries.some(row => row.outcome === "carried" || row.outcome === "edited")) return { bytes: rendered, carries }
  parts[BODY] = strToU8(new XMLSerializer().serializeToString(target))
  const copied = copyStyles(parts[STYLES], originalParts[STYLES], styles)
  if (copied) parts[STYLES] = copied
  return { bytes: zipSync(parts, { mtime: new Date(1980, 0, 1) }), carries }
}

/** The original table element, already imported into the rendered body, made
 * to read as the drafted rows: rows paired by position when the counts agree,
 * else by their first cell; a cell is rewritten only where its text differs. */
function carry(node: Element, drafted: string[][]): { node: Element; cells: string[]; added: number; removed: string[] } | { reason: string } {
  clean(node)
  const table = grid(node)
  const original = gridText(table)
  const width = original[0]?.length ?? 0
  if (drafted.some(row => row.length !== width)) return { reason: `the drafted table has ${drafted[0]?.length ?? 0} columns where the original has ${width}` }
  const pairs = original.length === drafted.length ? original.map((_, at) => at) : alignRows(original, drafted)
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
  // A drafted row the original lacks takes the form of the original row it follows.
  let added = 0
  const kept = table.rows.filter((_, index) => pairs[index] !== undefined)
  let anchor: Element | undefined
  for (const [at, text] of drafted.entries()) {
    const index = pairs.indexOf(at)
    if (index >= 0) { anchor = table.rows[index]!.node; continue }
    const model = anchor ?? kept[0]?.node ?? table.rows[0]?.node
    if (!model) return { reason: "the original table has no rows" }
    const row = model.cloneNode(true) as Element
    // An added row is a body row, never a repeated header row.
    Array.from(row.getElementsByTagNameNS(W, "tblHeader")).forEach(mark => mark.parentNode?.removeChild(mark))
    const cells = wrap(node, row).rows[0]!.cells
    if (cells.some(cell => cell.span > 1 || cell.continued)) return { reason: "a drafted row is added beside merged cells" }
    cells.forEach((cell, column) => setCellText(cell.node, readBack(text[column] ?? "")))
    node.insertBefore(row, anchor ? anchor.nextSibling : kept[0]?.node ?? null)
    anchor = row
    added++
  }
  table.rows.forEach((row, index) => { if (pairs[index] === undefined) row.node.parentNode!.removeChild(row.node) })
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
    setCellText(cell.node, readBack(drafted[positions[0]!] ?? ""))
    changed.push(`${label}, column ${positions[0]! + 1}`)
  }
}

/** Original rows paired with drafted rows by their first cell, in order (a longest common subsequence). */
function alignRows(original: string[][], drafted: string[][]) {
  const key = (row: string[]) => plain(row[0] ?? "") || plain(row.join(" "))
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

function pairScore(original: DocxTable, drafted: DocxTable) {
  const columns = (table: DocxTable) => table.rows[0]?.length ?? 0
  const placed = plain(original.follows) !== "" && plain(original.follows) === plain(drafted.follows) && columns(original) === columns(drafted) ? MATCH : 0
  return Math.max(placed, shared(original.rows, drafted.rows))
}

/** The share of non-empty cell text two tables have in common. */
function shared(a: string[][], b: string[][]) {
  const texts = (rows: string[][]) => rows.flat().map(plain).filter(Boolean)
  const left = texts(a)
  const right = texts(b)
  if (!left.length || !right.length)
    return !left.length && !right.length && a.length === b.length && (a[0]?.length ?? 0) === (b[0]?.length ?? 0) ? 1 : 0
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
  return readBack(text).normalize("NFKC").replace(/[\u2010-\u2015\u2212]/g, "-").replace(/[\u2018\u2019\u201A\u201B]/g, "'").replace(/[\u201C\u201D\u201E\u201F]/g, "\"")
    .replace(/\s+/g, " ").trim()
}

function same(a: string, b: string) {
  return normalized(a) === normalized(b)
}

/** Text as matched: compared, and in any case. */
export function plain(text: string) {
  return normalized(text).toLowerCase()
}

function readBack(text: string) {
  return text.replaceAll(PIPE_STAND_IN, "|")
}

function parse(xml: string) {
  return new DOMParser().parseFromString(xml, "application/xml")
}

function children(node: Element, name: string) {
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1 && (child as Element).namespaceURI === W && (child as Element).localName === name)
}

function bodyTables(document: Document) {
  const body = document.getElementsByTagNameNS(W, "body")[0]
  return body ? children(body, "tbl") : []
}

function grid(node: Element): Grid {
  return { node, rows: children(node, "tr").map(row => ({ node: row, cells: children(row, "tc").map(cell => {
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

function paragraphText(node: Node, paragraph: Element): string {
  if (node.nodeType !== 1) return ""
  const element = node as Element
  if (element.namespaceURI === W) {
    if (["del", "moveFrom"].includes(element.localName) || element.localName === "p" && element !== paragraph) return ""
    if (element.localName === "t") return element.textContent ?? ""
    if (element.localName === "tab") return " "
    if (["br", "cr"].includes(element.localName)) return "\n"
    if (element.localName === "noBreakHyphen") return "-"
  }
  return Array.from(node.childNodes).map(child => paragraphText(child, paragraph)).join("")
}

/** The text of the body paragraph a table follows; "" after another table. */
function follows(table: Element) {
  for (let node = table.previousSibling; node; node = node.previousSibling) {
    if (node.nodeType !== 1 || (node as Element).namespaceURI !== W) continue
    const element = node as Element
    if (element.localName === "tbl") return ""
    if (element.localName !== "p" && element.localName !== "sdt") continue
    const text = Array.from(element.getElementsByTagNameNS(W, "p")).concat(element.localName === "p" ? [element] : [])
      .map(paragraph => paragraphText(paragraph, paragraph)).join(" ").replace(/\s+/g, " ").trim()
    if (text) return text
  }
  return ""
}

/** Why a table cannot be carried into another package, if it cannot. */
function uncarried(table: Element) {
  const elements = Array.from(table.getElementsByTagName("*"))
  if (elements.some(node => node.namespaceURI !== W)) return "the original table holds content outside plain Word text (a drawing, an equation or an Office extension)"
  const refused = elements.find(node => UNCARRIED.has(node.localName))
  if (refused) return `the original table holds ${refused.localName === "numPr" ? "list numbering" : `a ${refused.localName} element`}, which refers to its own package`
  if (elements.some(node => Array.from(node.attributes).some(attribute => attribute.namespaceURI === R))) return "the original table links to another part of its package"
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

/** A cell's text replaced, keeping its first paragraph's and first run's formatting. */
function setCellText(cell: Element, text: string) {
  const document = cell.ownerDocument!
  const paragraph = children(cell, "p")[0]
  const paragraphProperties = paragraph && children(paragraph, "pPr")[0]
  const run = paragraph && Array.from(paragraph.getElementsByTagNameNS(W, "r")).find(node => node.getElementsByTagNameNS(W, "t").length) || undefined
  const runProperties = run && children(run, "rPr")[0]
  for (const child of Array.from(cell.childNodes))
    if (!(child.nodeType === 1 && (child as Element).namespaceURI === W && (child as Element).localName === "tcPr")) cell.removeChild(child)
  for (const line of text.split("\n")) {
    const next = document.createElementNS(W, "w:p")
    if (paragraphProperties) next.appendChild(paragraphProperties.cloneNode(true))
    if (line.trim()) {
      const r = document.createElementNS(W, "w:r")
      if (runProperties) r.appendChild(runProperties.cloneNode(true))
      const t = document.createElementNS(W, "w:t")
      t.setAttributeNS(XML_NS, "xml:space", "preserve")
      t.appendChild(document.createTextNode(line.trim()))
      r.appendChild(t)
      next.appendChild(r)
    }
    cell.appendChild(next)
  }
}

/** A detached row read in the context of its table, for its cells. */
function wrap(table: Element, row: Element): Grid {
  const holder = table.ownerDocument!.createElementNS(W, "w:tbl")
  holder.appendChild(row)
  const read = grid(holder)
  holder.removeChild(row)
  return read
}

/** The renderer closes each Markdown table with an empty spacing paragraph; a carried table stands as in the original. */
function dropRenderedSpacer(table: Element) {
  let next = table.nextSibling
  while (next && next.nodeType !== 1) next = next.nextSibling
  const paragraph = next as Element | null
  if (!paragraph || paragraph.namespaceURI !== W || paragraph.localName !== "p" || paragraphText(paragraph, paragraph).trim()) return
  const content = Array.from(paragraph.childNodes).filter(node => node.nodeType === 1 && (node as Element).localName !== "pPr")
  const spacing = paragraph.getElementsByTagNameNS(W, "spacing")[0]
  if (!content.length && spacing?.getAttributeNS(W, "before") === "80" && spacing.getAttributeNS(W, "after") === "120") paragraph.parentNode!.removeChild(paragraph)
}

function referencedStyles(table: Element) {
  return ["tblStyle", "pStyle", "rStyle"].flatMap(name => Array.from(table.getElementsByTagNameNS(W, name)).map(node => node.getAttributeNS(W, "val") ?? "")).filter(Boolean)
}

/** The rendered styles with every style the carried tables use that they
 * lack, and the styles those are based on, copied from the original; list
 * numbering in a copied style is left out, as its numbering part is not carried.
 * Undefined when nothing is copied. */
function copyStyles(target: Uint8Array | undefined, source: Uint8Array | undefined, wanted: Set<string>) {
  if (!target || !source || !wanted.size) return
  const into = parse(strFromU8(target))
  const from = parse(strFromU8(source))
  const root = into.documentElement
  if (!root) return
  const byId = (document: Document) => new Map(Array.from(document.getElementsByTagNameNS(W, "style")).map(node => [node.getAttributeNS(W, "styleId") ?? "", node]))
  const present = byId(into)
  const available = byId(from)
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
  return copied ? strToU8(new XMLSerializer().serializeToString(into)) : undefined
}
