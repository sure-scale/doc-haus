import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"

const PARAGRAPH_RE = /<w:p\b[\s\S]*?<\/w:p>/g
const INLINE_RE = /<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>|<w:r\b[\s\S]*?<\/w:r>/g
const RUN_RE = /<w:r\b[\s\S]*?<\/w:r>/g
const TEXT_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
const HEADING_SPACING = {
  Heading1: { before: "320", after: "160" },
  Heading2: { before: "280", after: "140" },
  Heading3: { before: "220", after: "120" },
}
const TABLE_CELL_PARAGRAPH_PROPERTIES =
  '<w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr>'
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XMLNS = "http://www.w3.org/2000/xmlns/"

/** Generated-main-XML adapter, immediately after SDK save and before source
 * table carry. The SDK can retain a style seed's Word prefix: adapt that
 * generated main XML to the bounded converter's canonical w: convention.
 * This is not a source-table importer. `headingSpacing: false`
 * leaves headings the spacing of their own template styles. */
export function polishMarkdownArtifactsInDocx(bytes: Uint8Array, options: { headingSpacing?: boolean } = {}) {
  return rewriteDocumentXml(bytes, xml => polishMarkdownArtifactsInDocumentXml(xml, options))
}

function rewriteDocumentXml(
  bytes: Uint8Array,
  rewrite: (xml: string) => string,
) {
  const parts = unzipSync(bytes)
  const document = parts["word/document.xml"]
  if (!document) return bytes
  const original = strFromU8(document)
  const polished = rewrite(original)
  const styleAdded = addMissingInlineCodeStyle(parts, polished)
  if (polished === original && !styleAdded) return bytes
  parts["word/document.xml"] = strToU8(polished)
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

function addMissingInlineCodeStyle(parts: Record<string, Uint8Array>, document: string) {
  if (!parts["word/styles.xml"] || !Array.from(new DOMParser().parseFromString(document, "application/xml")
    .getElementsByTagNameNS(W, "rStyle")).some(style => style.getAttributeNS(W, "val") === "Code")) return false
  const styles = strFromU8(parts["word/styles.xml"])
  const parsed = new DOMParser().parseFromString(styles, "application/xml")
  if (Array.from(parsed.getElementsByTagNameNS(W, "style")).some(style => style.getAttributeNS(W, "styleId") === "Code")) return false
  // Table inline-code runs inherit their current formatting. Declare their
  // character style without introducing a font or changing existing styles.
  const root = parsed.documentElement
  if (root.namespaceURI !== W || root.localName !== "styles") return false
  const prefix = root.prefix || "w"
  const style = parsed.createElementNS(W, `${prefix}:style`)
  style.setAttributeNS(W, `${prefix}:type`, "character")
  style.setAttributeNS(W, `${prefix}:styleId`, "Code")
  const name = parsed.createElementNS(W, `${prefix}:name`)
  name.setAttributeNS(W, `${prefix}:val`, "Code")
  style.appendChild(name)
  root.appendChild(style)
  parts["word/styles.xml"] = strToU8(new XMLSerializer().serializeToString(parsed))
  return true
}

export function polishMarkdownArtifactsInDocumentXml(xml: string, options: { headingSpacing?: boolean } = {}) {
  xml = generatedWordSerialization(xml)
  const paragraphs = [...xml.matchAll(PARAGRAPH_RE)]
  let output = ""
  let cursor = 0
  let index = 0

  while (index < paragraphs.length) {
    const paragraph = paragraphs[index]!
    const paragraphXml = paragraph[0]
    const paragraphEnd = paragraph.index! + paragraphXml.length
    const table = readMarkdownTable(paragraphs, index)

    output += xml.slice(cursor, paragraph.index!)

    if (table) {
      output += buildTableXml(table.rows)
      cursor = table.end
      index = table.nextIndex
      continue
    }

    if (paragraphText(paragraphXml).trim() === "---") {
      output += horizontalRuleXml()
      cursor = paragraphEnd
      index++
      continue
    }

    output += options.headingSpacing === false ? paragraphXml : polishHeadingParagraph(paragraphXml)
    cursor = paragraphEnd
    index++
  }

  return output + xml.slice(cursor)
}

/** Namespace-aware adapter for the SDK-generated body, before source carry.
 * Retain original namespace declarations, including aliases used by QName
 * attribute values. Refuse a conflicting w: binding instead of reinterpreting
 * those values when adapting another Word prefix. */
function generatedWordSerialization(xml: string) {
  const document = new DOMParser().parseFromString(xml, "application/xml")
  const words = Array.from(document.getElementsByTagNameNS(W, "*"))
  if (!words.some(element => element.prefix !== "w" || Array.from(element.attributes)
    .some(attribute => attribute.namespaceURI === W && attribute.prefix !== "w"))) return xml
  const elements = Array.from(document.getElementsByTagName("*"))
  if (elements.some(element => Array.from(element.attributes).some(attribute =>
    attribute.namespaceURI === XMLNS && attribute.name === "xmlns:w" && attribute.value !== W)))
    throw new Error("Generated Markdown XML has a conflicting Word namespace prefix.")
  const copy = (node: Node): Node => {
    if (node.nodeType !== 1) return node.cloneNode(true)
    const element = node as Element
    const cloned = document.createElementNS(element.namespaceURI, element.namespaceURI === W ? `w:${element.localName}` : element.tagName)
    for (const attribute of Array.from(element.attributes)) cloned.setAttributeNS(attribute.namespaceURI,
      attribute.namespaceURI === W ? `w:${attribute.localName}` : attribute.name, attribute.value)
    for (const child of Array.from(element.childNodes)) cloned.appendChild(copy(child))
    return cloned
  }
  const root = copy(document.documentElement) as Element
  root.setAttributeNS(XMLNS, "xmlns:w", W)
  document.replaceChild(root, document.documentElement)
  return new XMLSerializer().serializeToString(document)
}

type TableInlineRun = {
  text: string
  properties: string
  hyperlink: string | null
}

type TableCell = TableInlineRun[][]
type TableRow = TableCell[]
type MarkdownTable = TableRow[]

function readMarkdownTable(paragraphs: RegExpMatchArray[], start: number) {
  const header = markdownRow(paragraphs[start]![0])
  if (!header) return undefined

  const separator = markdownRow(paragraphs[start + 1]?.[0] ?? "")
  if (!separator || !isSeparatorRow(separator) || separator.length !== header.length) return undefined

  const rows = [header]
  let index = start + 2
  while (index < paragraphs.length) {
    const row = markdownRow(paragraphs[index]![0])
    if (!row || row.length !== header.length) break
    rows.push(row)
    index++
  }

  // A table of its header row alone, or of one column, is a table when its
  // separator row follows, as in GitHub-flavored Markdown.
  const last = paragraphs[index - 1]!
  return { rows, nextIndex: index, end: last.index! + last[0].length }
}

// A bar-delimited line rendered as a list item, a heading, a quotation or a
// line of code has paragraph properties a table cell would lose. It is not a
// table row, and it stays as text.
function markdownRow(xml: string) {
  const text = paragraphText(xml).trim()
  if (!text.startsWith("|") || !text.endsWith("|") || /<w:pPr\b/.test(xml)) return undefined
  return parseMarkdownRow(paragraphInlineRuns(xml))
}

/** Heading-only mixed-source API. Presentation runs after original tables are
 * carried: only current body paragraph properties may change, never tables or
 * historical pPrChange properties. Expanded names own XML identity here. */
export function polishHeadingSpacing(xml: string) {
  const namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  const document = new DOMParser().parseFromString(xml, "application/xml")
  let changed = false
  const child = (parent: Element, name: string) => Array.from(parent.childNodes).find(node =>
    node.nodeType === 1 && (node as Element).namespaceURI === namespace && (node as Element).localName === name) as Element | undefined
  for (const paragraph of Array.from(document.getElementsByTagNameNS(namespace, "p"))) {
    let ancestor = paragraph.parentNode
    while (ancestor && !(ancestor.nodeType === 1 && (ancestor as Element).namespaceURI === namespace && (ancestor as Element).localName === "tbl")) ancestor = ancestor.parentNode
    if (ancestor) continue
    const properties = child(paragraph, "pPr")
    const style = properties && child(properties, "pStyle")?.getAttributeNS(namespace, "val")
    if (!properties || !style || !Object.hasOwn(HEADING_SPACING, style)) continue
    const spacing = document.createElementNS(namespace, "w:spacing")
    const values = HEADING_SPACING[style as keyof typeof HEADING_SPACING]
    spacing.setAttributeNS(namespace, "w:before", values.before)
    spacing.setAttributeNS(namespace, "w:after", values.after)
    const current = child(properties, "spacing")
    if (current) properties.replaceChild(spacing, current)
    else {
      const afterSpacing = ["ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
      const before = Array.from(properties.childNodes).find(node => node.nodeType === 1
        && (node as Element).namespaceURI === namespace && afterSpacing.includes((node as Element).localName))
      properties.insertBefore(spacing, before ?? null)
    }
    changed = true
  }
  return changed ? new XMLSerializer().serializeToString(document) : xml
}

function polishHeadingParagraph(xml: string) {
  const style = xml.match(/<w:pStyle\b[^>]*w:val="(Heading[123])"[^>]*\/>/)?.[1] as keyof typeof HEADING_SPACING | undefined
  if (!style) return xml
  const values = HEADING_SPACING[style]
  const spacing = `<w:spacing w:before="${values.before}" w:after="${values.after}"/>`
  if (/<w:spacing\b[^>]*\/>/.test(xml)) return xml.replace(/<w:spacing\b[^>]*\/>/, spacing)
  if (xml.includes("<w:pPr>")) return xml.replace("</w:pPr>", `${spacing}</w:pPr>`)
  return xml.replace(/(<w:p\b[^>]*>)/, `$1<w:pPr>${spacing}</w:pPr>`)
}

function parseMarkdownRow(runs: TableInlineRun[]): TableRow | undefined {
  const characters = runs.flatMap((run) =>
    [...run.text.replace(/<br\s*\/?>/gi, "\n")].map((text) => ({
      text,
      properties: run.properties,
      hyperlink: run.hyperlink,
    })),
  )
  const start = characters.findIndex((character) => character.text.trim())
  const end = characters.findLastIndex((character) => character.text.trim())
  if (start < 0 || start === end || characters[start]?.text !== "|" || characters[end]?.text !== "|")
    return undefined
  const cells: TableInlineRun[][] = [[]]
  characters.slice(start + 1, end).forEach((character) => {
    if (character.text === "|") {
      cells.push([])
      return
    }
    cells[cells.length - 1]!.push(character)
  })
  return cells.map((cell) => tableCell(trimInlineWhitespace(cell)))
}

function isSeparatorRow(cells: TableRow) {
  return cells.every((cell) =>
    /^:?-{3,}:?$/.test(
      cell.flatMap((line) => line.map((run) => run.text)).join("").replace(/\s/g, ""),
    )
  )
}

function paragraphText(xml: string) {
  return [...xml.matchAll(TEXT_RE)].map((match) => unescapeXml(match[1] ?? "")).join("")
}

function paragraphInlineRuns(xml: string) {
  const inlineElements = [...xml.matchAll(INLINE_RE)]
  const paragraphPayload = xml
    .replace(/^<w:p\b[^>]*>/, "")
    .replace(/<\/w:p>$/, "")
  const paragraphResidue = inlineElements.reduce(
    (content, element) => content.replace(element[0], ""),
    paragraphPayload,
  ).trim()
  if (paragraphResidue)
    throw new Error(
      "Markdown table polish cannot losslessly conserve an unsupported Word paragraph payload.",
    )
  return inlineElements.flatMap((match) => {
    const inline = match[0]
    if (!inline.startsWith("<w:hyperlink")) return [readInlineRun(inline, null)]
    const hyperlink = inline.match(/^<w:hyperlink\b[^>]*>/)?.[0]
    if (!hyperlink)
      throw new Error("Markdown table hyperlink has no canonical opening element.")
    const runs = [...inline.matchAll(RUN_RE)]
    const residue = runs.reduce(
      (content, run) => content.replace(run[0], ""),
      inline.slice(hyperlink.length).replace(/<\/w:hyperlink>$/, ""),
    ).trim()
    if (residue)
      throw new Error(
        "Markdown table polish cannot losslessly conserve an unsupported Word hyperlink payload.",
      )
    return runs.map((run) =>
      readInlineRun(run[0], hyperlink)
    )
  }).filter((run) => run.text.length)
}

function readInlineRun(xml: string, hyperlink: string | null): TableInlineRun {
  const properties =
    xml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ??
    xml.match(/<w:rPr\b[^>]*\/>/)?.[0] ??
    ""
  const payload = xml
    .replace(/^<w:r\b[^>]*>/, "")
    .replace(/<\/w:r>$/, "")
    .replace(properties, "")
  const textElements = [...payload.matchAll(TEXT_RE)]
  const residue = textElements.reduce(
    (content, element) => content.replace(element[0], ""),
    payload,
  ).trim()
  if (residue)
    throw new Error(
      "Markdown table polish cannot losslessly conserve an unsupported Word run payload.",
    )
  return {
    text: textElements.map((element) => unescapeXml(element[1] ?? "")).join(""),
    properties,
    hyperlink,
  }
}

function trimInlineWhitespace(runs: TableInlineRun[]) {
  const start = runs.findIndex((run) => run.text.trim())
  if (start < 0) return []
  const end = runs.findLastIndex((run) => run.text.trim())
  return runs.slice(start, end + 1).map((run, index, kept) => ({
    ...run,
    text: run.text.slice(
      index === 0 ? run.text.search(/\S/) : 0,
      index === kept.length - 1
        ? run.text.search(/\s*$/)
        : run.text.length,
    ),
  })).filter((run) => run.text)
}

function tableCell(characters: TableInlineRun[]): TableCell {
  return characters.reduce<TableCell>((lines, character) => {
    if (character.text === "\n") return [...lines, []]
    const line = lines[lines.length - 1]!
    const previous = line[line.length - 1]
    if (
      previous &&
      previous.properties === character.properties &&
      previous.hyperlink === character.hyperlink
    ) {
      previous.text += character.text
      return lines
    }
    line.push({ ...character })
    return lines
  }, [[]]).map((line) => line.flatMap(parseLiteralInlineFormatting))
}

function parseLiteralInlineFormatting(run: TableInlineRun) {
  const tokens = [
    ...run.text.matchAll(/\*\*\*([^*\r\n]+)\*\*\*|\*\*([^*\r\n]+)\*\*|\*([^*\r\n]+)\*|`([^`\r\n]+)`/g),
  ]
  let offset = 0
  return tokens.flatMap((token) => {
    const start = token.index ?? 0
    const before = run.text.slice(offset, start)
    offset = start + token[0].length
    const properties = token[1] !== undefined
      ? addItalic(addBold(run.properties))
      : token[2] !== undefined
        ? addBold(run.properties)
        : token[3] !== undefined
          ? addItalic(run.properties)
          : addCodeStyle(run.properties)
    return [
      ...(before ? [{ ...run, text: before }] : []),
      { ...run, text: token[1] ?? token[2] ?? token[3] ?? token[4] ?? "", properties },
    ]
  }).concat(run.text.slice(offset) ? { ...run, text: run.text.slice(offset) } : [])
}

function buildTableXml(rows: MarkdownTable) {
  const cols = rows[0]?.length ?? 1
  const width = Math.floor(9360 / cols)
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>` +
    `<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>` +
    `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>` +
    `<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>` +
    `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>` +
    `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>` +
    `</w:tblBorders><w:tblCellMar>` +
    `<w:top w:w="90" w:type="dxa"/><w:left w:w="90" w:type="dxa"/>` +
    `<w:bottom w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/>` +
    `</w:tblCellMar></w:tblPr><w:tblGrid>` +
    Array.from({ length: cols }, () => `<w:gridCol w:w="${width}"/>`).join("") +
    `</w:tblGrid>` +
    rows.map((row, rowIndex) => buildTableRowXml(row, rowIndex === 0, width)).join("") +
    `</w:tbl><w:p><w:pPr><w:spacing w:before="80" w:after="120"/></w:pPr></w:p>`
  )
}

function buildTableRowXml(row: TableRow, header: boolean, width: number) {
  return `<w:tr><w:trPr><w:cantSplit/>${header ? "<w:tblHeader/>" : ""}</w:trPr>${row.map((cell) => buildTableCellXml(cell, header, width)).join("")}</w:tr>`
}

function buildTableCellXml(cell: TableCell, header: boolean, width: number) {
  const properties =
    `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>` +
    (header ? `<w:shd w:val="clear" w:color="auto" w:fill="EAF2F8"/>` : "") +
    `</w:tcPr>`
  return `<w:tc>${properties}${cell
    .map((line) =>
      `<w:p>${TABLE_CELL_PARAGRAPH_PROPERTIES}${renderCellRuns(line, header)}</w:p>`
    )
    .join("")}</w:tc>`
}

function renderCellRuns(runs: TableInlineRun[], header: boolean) {
  return runs.map((run) => {
    const rendered = cellRun(run, header)
    if (!run.hyperlink) return rendered
    return `${run.hyperlink}${rendered}</w:hyperlink>`
  }).join("") || cellRun({ text: "", properties: "", hyperlink: null }, header)
}

function cellRun(run: TableInlineRun, header: boolean) {
  return `<w:r>${header ? addBold(run.properties) : run.properties}<w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`
}

function horizontalRuleXml() {
  return (
    `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="BFBFBF"/>` +
    `</w:pBdr><w:spacing w:before="120" w:after="160"/></w:pPr></w:p>`
  )
}

function escapeXml(text: string) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}

function unescapeXml(text: string) {
  return text
    .replace(/&#(?:x([0-9a-f]+)|([0-9]+));/gi, (_, hex: string | undefined, decimal: string | undefined) => {
      const codePoint = Number.parseInt(hex ?? decimal ?? "", hex ? 16 : 10)
      if (!isXmlCharacter(codePoint))
        throw new Error("Markdown table contains an invalid XML character reference.")
      return String.fromCodePoint(codePoint)
    })
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, `"`)
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
}

function addBold(properties: string) {
  return addRunProperty(
    properties,
    /<w:b\b[^>]*(?:\/>|>[\s\S]*?<\/w:b>)/,
    "<w:b/>",
  )
}

function addItalic(properties: string) {
  return addRunProperty(
    properties,
    /<w:i\b[^>]*(?:\/>|>[\s\S]*?<\/w:i>)/,
    "<w:i/>",
  )
}

function addCodeStyle(properties: string) {
  return addRunProperty(
    properties,
    /<w:rStyle\b[^>]*w:val="Code"[^>]*\/>/,
    '<w:rStyle w:val="Code"/>',
  )
}

function addRunProperty(
  properties: string,
  present: RegExp,
  element: string,
) {
  if (present.test(properties)) return properties
  if (!properties) return `<w:rPr>${element}</w:rPr>`
  if (/<w:rPr\b[^>]*\/>/.test(properties))
    return properties.replace(/\/>$/, `>${element}</w:rPr>`)
  return properties.replace(/<\/w:rPr>$/, `${element}</w:rPr>`)
}

function isXmlCharacter(codePoint: number) {
  return codePoint === 0x9 ||
    codePoint === 0xa ||
    codePoint === 0xd ||
    (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
    (codePoint >= 0x10000 && codePoint <= 0x10ffff)
}
