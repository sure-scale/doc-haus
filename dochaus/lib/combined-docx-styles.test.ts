import { expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { combineMemoAndBody } from "./combined-docx"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const DEFAULTS = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`
  + `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>`
const MARGINS = `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`
const paragraph = (text: string, style = "") => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t>${text}</w:t></w:r></w:p>`
const heading = (look: string, next = "Normal") => `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>`
  + `<w:next w:val="${next}"/><w:link w:val="Heading1Char"/><w:qFormat/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr>${look}</w:rPr></w:style>`
const headingChar = (look: string) => `<w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char"><w:name w:val="Heading 1 Char"/>`
  + `<w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="Heading1"/><w:rPr>${look}</w:rPr></w:style>`

test("a linked paragraph and character style pair is reused or copied as a pair", () => {
  const body = docx(paragraph("Body heading", "Heading1"), DEFAULTS + heading("<w:caps/>") + headingChar("<w:caps/>"))
  const blocks = paragraph("Memo heading", "Heading1") + `<w:p><w:r><w:rPr><w:rStyle w:val="Heading1Char"/></w:rPr><w:t>Run in the heading look</w:t></w:r></w:p>`
  const bodyText = `<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:next w:val="BodyText"/></w:style>`
  // Either member differing from the body's copies both; the copies link to
  // each other and their other references resolve to the copied or body names.
  for (const [ownParagraph, ownCharacter] of [[true, false], [false, true], [true, true]]) {
    const memo = docx(blocks, DEFAULTS + bodyText + (ownParagraph ? heading("<w:b/>", "BodyText") : heading("<w:caps/>")) + headingChar(ownCharacter ? "<w:b/>" : "<w:caps/>"))
    const combined = combineMemoAndBody({ memo, body })
    const styles = styleMap(combined)
    const [first, second] = bodyBlocks(part(combined, "word/document.xml"))
    expect([value(all(first!, "pStyle")[0]), value(all(second!, "rStyle")[0])]).toEqual(["MemoHeading1", "MemoHeading1Char"])
    expect(references(styles.get("MemoHeading1")!)).toEqual(["Normal", ownParagraph ? "MemoBodyText" : "Normal", "MemoHeading1Char"])
    expect(references(styles.get("MemoHeading1Char")!)).toEqual(["DefaultParagraphFont", "", "MemoHeading1"])
    for (const id of ["Heading1", "Heading1Char"]) expect(canonical(styles.get(id)!)).toBe(canonical(styleMap(body).get(id)!))
    expect(danglingStyles(combined)).toEqual([])
  }
  // An identical pair is the body's own.
  const same = combineMemoAndBody({ memo: docx(blocks, DEFAULTS + heading("<w:caps/>") + headingChar("<w:caps/>")), body })
  const [first, second] = bodyBlocks(part(same, "word/document.xml"))
  expect([value(all(first!, "pStyle")[0]), value(all(second!, "rStyle")[0])]).toEqual(["Heading1", "Heading1Char"])
  expect([...styleMap(same).keys()].filter(id => id.startsWith("Memo"))).toEqual([])
})

test("a copied style whose id or name is taken gets a suffix on both, names compared as Word does", () => {
  const list = (indent: string) => `<w:style w:type="paragraph" w:styleId="ListNum"><w:name w:val="List Num"/><w:pPr><w:ind w:left="${indent}"/></w:pPr></w:style>`
  const body = docx(paragraph("Body text."), DEFAULTS + list("720")
    + `<w:style w:type="paragraph" w:customStyle="1" w:styleId="MemoListNum"><w:name w:val="Memo List Num"/></w:style>`
    + `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Callout"><w:name w:val="MEMO QUOTE"/></w:style>`)
  const memo = docx(paragraph("Memo item.", "ListNum") + paragraph("Memo quote.", "Quote"),
    DEFAULTS + list("360") + `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>`)
  const combined = combineMemoAndBody({ memo, body })
  const styles = all(part(combined, "word/styles.xml"), "style")
  const names = styles.map(style => value(all(style, "name")[0]).toLowerCase())
  expect(names.filter((name, index) => names.indexOf(name) !== index)).toEqual([])
  expect(new Set(styles.map(style => style.getAttributeNS(W, "styleId"))).size).toBe(styles.length)
  const used = bodyBlocks(part(combined, "word/document.xml")).slice(0, 2).map(block => value(all(block, "pStyle")[0]))
  expect(used).toEqual(["MemoListNum2", "MemoQuote2"])
  expect(used.map(id => value(all(styleMap(combined).get(id)!, "name")[0]))).toEqual(["Memo List Num 2", "Memo Quote 2"])
})

test("the memo section takes only the body's page size, margins, paper source, text direction and grid", () => {
  const setup = `<w:pgSz w:w="11906" w:h="16838"/>${MARGINS}<w:paperSrc w:first="7" w:other="7"/>`
  const section = `<w:sectPr><w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr><w:type w:val="nextPage"/>${setup}`
    + `<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/></w:pgBorders><w:lnNumType w:countBy="1"/>`
    + `<w:cols w:num="2" w:space="720"/><w:formProt w:val="0"/><w:vAlign w:val="center"/><w:titlePg/><w:textDirection w:val="lrTb"/><w:docGrid w:linePitch="360"/></w:sectPr>`
  const body = docx(paragraph("Body text.") + section, DEFAULTS)
  const memo = docx(paragraph("Memo text.") + `<w:sectPr>${setup}<w:lnNumType w:countBy="1"/><w:cols w:num="2"/></w:sectPr>`, DEFAULTS)
  const combined = combineMemoAndBody({ memo, body })
  const memoSection = all(bodyBlocks(part(combined, "word/document.xml"))[0]!, "sectPr")[0]!
  const names = ["pgSz", "pgMar", "paperSrc", "textDirection", "docGrid"]
  expect(children(memoSection).map(row => row.localName)).toEqual(["type", ...names])
  for (const name of names) expect(canonical(all(memoSection, name)[0]!)).toBe(canonical(all(sectPr(body), name)[0]!))
  // The body keeps its own layout and gains the page-number restart in order.
  expect(children(sectPr(combined)).map(row => row.localName)).toEqual(["footnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders",
    "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "titlePg", "textDirection", "docGrid"])
})

test("the inserted memo section break is explicitly next-page; the body's own section type stays as it was", () => {
  // The memo's own last section break was continuous; it does not carry over.
  const memo = docx(paragraph("Memo text.") + `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`, DEFAULTS)
  for (const type of ["", `<w:type w:val="evenPage"/>`]) {
    const body = docx(paragraph("Body text.") + `<w:sectPr>${type}<w:pgSz w:w="11906" w:h="16838"/>${MARGINS}</w:sectPr>`, DEFAULTS)
    const combined = combineMemoAndBody({ memo, body })
    const memoSection = all(bodyBlocks(part(combined, "word/document.xml"))[0]!, "sectPr")[0]!
    expect(children(memoSection).map(row => row.localName)).toEqual(["type", "pgSz", "pgMar"])
    expect(value(all(memoSection, "type")[0])).toBe("nextPage")
    expect(all(sectPr(combined), "type").map(value)).toEqual(all(sectPr(body), "type").map(value))
  }
})

test("a body whose first section ends inside a block wrapper restarts numbering at that section", () => {
  const cover = paragraph("Cover page.") + `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>${MARGINS}</w:sectPr></w:pPr><w:r><w:t>End of cover.</w:t></w:r></w:p>`
  const sdt = (inner: string) => `<w:sdt><w:sdtPr><w:alias w:val="Cover"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
  const custom = (inner: string) => `<w:customXml w:uri="urn:example" w:element="cover">${inner}</w:customXml>`
  const memo = docx(paragraph("Memo text."), DEFAULTS)
  for (const wrapped of [sdt(cover), custom(cover), custom(sdt(cover))]) {
    const body = docx(wrapped + paragraph("Main body.") + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}</w:sectPr>`, DEFAULTS)
    const combined = combineMemoAndBody({ memo, body })
    const document = part(combined, "word/document.xml")
    const [memoBlock, wrapper] = bodyBlocks(document)
    const first = all(wrapper!, "sectPr")[0]!
    expect(all(first, "pgNumType")[0]?.getAttributeNS(W, "start")).toBe("1")
    expect(all(first, "type").map(value)).toEqual(["nextPage"])
    expect(canonical(sectPr(combined))).toBe(canonical(sectPr(body)))
    // The memo takes the page setup of the section that follows it.
    expect(canonical(all(all(memoBlock!, "sectPr")[0]!, "pgSz")[0]!)).toBe(canonical(all(first, "pgSz")[0]!))
  }
})

// A minimal package: a document, optional styles part, no other parts.
function docx(body: string, styles?: string) {
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`
      + `${styles ? `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` : ""}</Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${styles ? `<Relationship Id="rId1" Type="${O}/styles" Target="styles.xml"/>` : ""}</Relationships>`),
    "word/document.xml": xml(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`),
    ...styles ? { "word/styles.xml": xml(`<w:styles xmlns:w="${W}">${styles}</w:styles>`) } : {},
  }, { mtime: new Date(1980, 0, 1) })
}

function part(bytes: Uint8Array, name: string) {
  return new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[name]!), "application/xml")
}

function all(node: Document | Element, name: string) {
  return Array.from(node.getElementsByTagNameNS(W, name))
}

function children(node: Element) {
  return Array.from(node.childNodes).filter((row): row is Element => row.nodeType === 1)
}

function value(node?: Element) {
  return node?.getAttributeNS(W, "val") ?? ""
}

function canonical(node: Node) {
  return new XMLSerializer().serializeToString(node)
}

function bodyBlocks(doc: Document) {
  return children(all(doc, "body")[0]!).filter(node => node.localName !== "sectPr")
}

function sectPr(bytes: Uint8Array) {
  return children(all(part(bytes, "word/document.xml"), "body")[0]!).find(node => node.localName === "sectPr")!
}

function styleMap(bytes: Uint8Array) {
  return new Map(all(part(bytes, "word/styles.xml"), "style").map(style => [style.getAttributeNS(W, "styleId")!, style]))
}

// A style's basedOn, next and link targets.
function references(style: Element) {
  return ["basedOn", "next", "link"].map(name => value(children(style).find(row => row.localName === name)))
}

function danglingStyles(bytes: Uint8Array) {
  const styles = part(bytes, "word/styles.xml")
  const defined = new Set(all(styles, "style").map(style => style.getAttributeNS(W, "styleId")))
  const document = part(bytes, "word/document.xml")
  return [...["pStyle", "rStyle", "tblStyle"].flatMap(name => all(document, name)),
    ...["basedOn", "next", "link"].flatMap(name => all(styles, name))].map(value).filter(id => !defined.has(id))
}
