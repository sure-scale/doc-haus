import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

const PARAGRAPH_RE = /<w:p\b[\s\S]*?<\/w:p>/g
const INLINE_RE = /<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>|<w:r\b[\s\S]*?<\/w:r>/g
const RUN_RE = /<w:r\b[\s\S]*?<\/w:r>/g
const TEXT_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
const HEADING_SPACING = {
  Heading1: `<w:spacing w:before="320" w:after="160"/>`,
  Heading2: `<w:spacing w:before="280" w:after="140"/>`,
  Heading3: `<w:spacing w:before="220" w:after="120"/>`,
}
const TABLE_CELL_PARAGRAPH_PROPERTIES =
  '<w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr>'

export function polishMarkdownArtifactsInDocx(bytes: Uint8Array) {
  return rewriteDocumentXml(bytes, polishMarkdownArtifactsInDocumentXml)
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
  return zipSync(parts)
}

function addMissingInlineCodeStyle(parts: Record<string, Uint8Array>, document: string) {
  if (!/<w:rStyle\b[^>]*\bw:val\s*=\s*(["'])Code\1/.test(document) || !parts["word/styles.xml"]) return false
  const styles = strFromU8(parts["word/styles.xml"])
  if (/<w:style\b[^>]*\bw:styleId\s*=\s*(["'])Code\1/.test(styles)) return false
  // Table inline-code runs inherit their current formatting. Declare their
  // character style without introducing a font or changing existing styles.
  const defined = styles.replace("</w:styles>", '<w:style w:type="character" w:styleId="Code"><w:name w:val="Code"/></w:style></w:styles>')
  if (defined === styles) return false
  parts["word/styles.xml"] = strToU8(defined)
  return true
}

export function polishMarkdownArtifactsInDocumentXml(xml: string) {
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

    output += polishHeadingParagraph(paragraphXml)
    cursor = paragraphEnd
    index++
  }

  return output + xml.slice(cursor)
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
  const headerXml = paragraphs[start]![0]
  if (!looksLikeMarkdownRow(headerXml)) return undefined
  const header = parseMarkdownRow(paragraphInlineRuns(headerXml))
  if (!header) return undefined

  const separatorXml = paragraphs[start + 1]?.[0] ?? ""
  if (!looksLikeMarkdownRow(separatorXml)) return undefined
  const separator = parseMarkdownRow(
    paragraphInlineRuns(separatorXml),
  )
  if (!separator || !isSeparatorRow(separator) || separator.length !== header.length) return undefined

  const rows = [header]
  let index = start + 2
  while (index < paragraphs.length) {
    if (!looksLikeMarkdownRow(paragraphs[index]![0])) break
    const row = parseMarkdownRow(paragraphInlineRuns(paragraphs[index]![0]))
    if (!row || row.length !== header.length) break
    rows.push(row)
    index++
  }

  if (rows.length === 1) return undefined
  const last = paragraphs[index - 1]!
  return { rows, nextIndex: index, end: last.index! + last[0].length }
}

function looksLikeMarkdownRow(xml: string) {
  const text = paragraphText(xml).trim()
  return text.startsWith("|") && text.endsWith("|")
}

function polishHeadingParagraph(xml: string) {
  const style = xml.match(/<w:pStyle\b[^>]*w:val="(Heading[123])"[^>]*\/>/)?.[1] as keyof typeof HEADING_SPACING | undefined
  if (!style) return xml
  const spacing = HEADING_SPACING[style]
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
  if (start < 0 || characters[start]?.text !== "|" || characters[end]?.text !== "|")
    return undefined
  const cells: TableInlineRun[][] = [[]]
  characters.slice(start + 1, end).forEach((character) => {
    if (character.text === "|") {
      cells.push([])
      return
    }
    cells[cells.length - 1]!.push(character)
  })
  if (cells.length < 2) return undefined
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
  const paragraphProperties =
    xml.match(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/)?.[0] ??
    xml.match(/<w:pPr\b[^>]*\/>/)?.[0] ??
    ""
  if (paragraphProperties)
    throw new Error(
      "Markdown table polish cannot losslessly conserve unsupported Word paragraph properties.",
    )
  const paragraphPayload = xml
    .replace(/^<w:p\b[^>]*>/, "")
    .replace(/<\/w:p>$/, "")
    .replace(paragraphProperties, "")
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
