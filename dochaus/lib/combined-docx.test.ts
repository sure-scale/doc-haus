import { expect, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { combineMemoAndBody } from "./combined-docx"
import { paragraphTargets, withParagraphPointers } from "./drafting-amendment"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "./render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const MEMO = [
  "# Cover memo",
  "## Points found",
  "The two sources give **different** start dates and use `term-a` inconsistently. [open item: confirm the start date]",
  "1. Memo point one.\n2. Memo point two.\n   - nested aside",
  "- Memo bullet",
  "> A quoted source line.",
  "| Topic | Treatment |\n| --- | --- |\n| Timing | Confirm. |",
  "See [the reference](https://example.com/reference).",
].join("\n\n")
const BODY = [
  "# Services Agreement",
  "## 1. Definitions",
  "The Supplier shall deliver the services.",
  "1. First numbered item.\n2. Second numbered item.",
  "- Bullet one\n- Bullet two",
  "## 2. Term",
  "The term is one year.",
].join("\n\n")

test("the notes come first, then a section break, then the body exactly as registered", async () => {
  const memo = (await renderMarkdownToDocx(MEMO)).bytes
  const body = (await renderMarkdownToDocx(BODY)).bytes
  const combined = combineMemoAndBody({ memo, body })

  const blocks = bodyBlocks(part(combined, "word/document.xml"))
  const split = bodyBlocks(part(memo, "word/document.xml")).length
  expect(blocks.slice(0, split).flatMap(texts)).toEqual(bodyBlocks(part(memo, "word/document.xml")).flatMap(texts))
  expect(blocks.slice(split).flatMap(texts)).toEqual(bodyBlocks(part(body, "word/document.xml")).flatMap(texts))
  // The memo is its own first section, ending at its last paragraph, with the
  // body's page setup and no header or footer. The body starts a new page and
  // numbers its pages from 1; otherwise its blocks are untouched.
  expect(sectionBreaks(combined)).toEqual([split - 1])
  expectMemoSection(combined, body)
  expectPageRestart(combined, body)
  expect(blocks.slice(split).map(canonical)).toEqual(bodyBlocks(part(body, "word/document.xml")).map(canonical))
  expect(keptAll(part(body, "word/styles.xml"), part(combined, "word/styles.xml"), "style")).toBe(true)
  expect(keptAll(part(body, "word/numbering.xml"), part(combined, "word/numbering.xml"), "abstractNum")).toBe(true)
  expect(keptAll(part(body, "word/numbering.xml"), part(combined, "word/numbering.xml"), "num")).toBe(true)

  // Memo headings keep the shared heading styles, except one the body numbers
  // as its clause list, which the memo takes as its own copy; its code style,
  // lists and link carry over without touching the body's definitions.
  const memoBlocks = blocks.slice(0, split)
  expect(memoBlocks.slice(0, 2).map(block => value(all(block, "pStyle")[0]))).toEqual(["Heading1", "MemoHeading2"])
  expect(danglingStyles(combined)).toEqual([])
  expect(numberingProblems(combined)).toEqual([])
  const bodyNums = new Set(blocks.slice(split).flatMap(block => all(block, "numId")).map(value))
  expect(memoBlocks.flatMap(block => all(block, "numId")).map(value).filter(id => bodyNums.has(id))).toEqual([])
  const link = memoBlocks.flatMap(block => all(block, "hyperlink"))[0]!
  const target = Array.from(part(combined, "word/_rels/document.xml.rels").getElementsByTagName("Relationship"))
    .find(row => row.getAttribute("Id") === link.getAttributeNS(O, "id"))
  expect(target?.getAttribute("Target")).toBe("https://example.com/reference")
  expect(target?.getAttribute("TargetMode")).toBe("External")

  // Docxodus re-opens it and both lists start at one, memo before body.
  const markdown = await reopen(combined)
  expect(markdown.indexOf("Cover memo")).toBeLessThan(markdown.indexOf("Services Agreement"))
  expect(markdown).toMatch(/ 1\. Memo point one\./)
  expect(markdown).toMatch(/ 1\. First numbered item\./)
  expect(markdown).toMatch(/ 2\. Second numbered item\./)
})

test("a template-like body keeps its styles, numbering, section and header; memo headings stay unnumbered", async () => {
  const memo = (await renderMarkdownToDocx(MEMO)).bytes
  const body = withParagraphPointers(templateBody())
  const combined = combineMemoAndBody({ memo, body })
  const before = unzipSync(body)
  const after = unzipSync(combined)

  expect(after["word/header1.xml"]).toEqual(before["word/header1.xml"]!)
  expectMemoSection(combined, body)
  expectPageRestart(combined, body)
  expect(all(sectPr(combined), "headerReference")).toHaveLength(1)
  expect(canonical(all(part(combined, "word/styles.xml"), "docDefaults")[0]!)).toBe(canonical(all(part(body, "word/styles.xml"), "docDefaults")[0]!))
  expect(keptAll(part(body, "word/styles.xml"), part(combined, "word/styles.xml"), "style")).toBe(true)
  expect(keptAll(part(body, "word/numbering.xml"), part(combined, "word/numbering.xml"), "abstractNum")).toBe(true)
  expect(keptAll(part(body, "word/numbering.xml"), part(combined, "word/numbering.xml"), "num")).toBe(true)

  const blocks = bodyBlocks(part(combined, "word/document.xml"))
  const split = bodyBlocks(part(memo, "word/document.xml")).length
  expect(blocks.slice(split).map(canonical)).toEqual(bodyBlocks(part(body, "word/document.xml")).map(canonical))
  expect(all(part(combined, "word/document.xml"), "bookmarkStart").map(mark => mark.getAttributeNS(W, "name"))).toContain("_Ref100")

  // The body's Heading1 is auto-numbered; memo headings use copied styles
  // with the memo's own look and no numbering.
  const styles = new Map(all(part(combined, "word/styles.xml"), "style").map(style => [style.getAttributeNS(W, "styleId")!, style]))
  const headings = blocks.slice(0, 2).map(block => value(all(block, "pStyle")[0]))
  expect(headings).toEqual(["MemoHeading1", "MemoHeading2"])
  for (const id of headings) {
    const style = styles.get(id)!
    expect(value(all(style, "name")[0])).toStartWith("Memo ")
    expect(all(style, "numPr")).toHaveLength(0)
    expect(all(style, "outlineLvl")).toHaveLength(1)
    expect(all(style, "b")).toHaveLength(1)
    expect(style.getAttributeNS(W, "default")).toBeFalsy()
  }
  expect(value(all(blocks[split]!, "pStyle")[0])).toBe("Heading1")
  expect(blocks.slice(0, split).flatMap(block => all(block, "pStyle")).map(value)).toContain("MemoQuote")
  expect(danglingStyles(combined)).toEqual([])
  expect(numberingProblems(combined)).toEqual([])

  const markdown = await reopen(combined)
  expect(markdown.indexOf("Cover memo")).toBeLessThan(markdown.indexOf("Interpretation"))
  expect(markdown).toMatch(/ 1\. Memo point one\./)
})

test("paragraph pointers carry over as they come, with unique bookmark ids and names", async () => {
  // The memo also marks _Ref100, which the template body uses too.
  const notes = unzipSync((await renderMarkdownToDocx(MEMO)).bytes)
  notes["word/document.xml"] = strToU8(strFromU8(notes["word/document.xml"]!).replace("</w:t></w:r>", `</w:t></w:r><w:bookmarkStart w:id="0" w:name="_Ref100"/><w:bookmarkEnd w:id="0"/>`))
  const memo = zipSync(notes)
  for (const body of [(await renderMarkdownToDocx(BODY)).bytes, withParagraphPointers(templateBody())]) {
    const combined = combineMemoAndBody({ memo, body })
    const pids = async (bytes: Uint8Array) => (await paragraphTargets(bytes)).map(row => row.pid)
    const expected = [...await pids(memo), ...await pids(body)]
    expect(expected.filter(Boolean).length).toBeGreaterThan(10)
    expect(await pids(combined)).toEqual(expected)
    const marks = all(part(combined, "word/document.xml"), "bookmarkStart")
    expect(marks).toHaveLength(all(part(memo, "word/document.xml"), "bookmarkStart").length + all(part(body, "word/document.xml"), "bookmarkStart").length)
    expect(new Set(marks.map(mark => mark.getAttributeNS(W, "id"))).size).toBe(marks.length)
    expect(new Set(marks.map(mark => mark.getAttributeNS(W, "name"))).size).toBe(marks.length)
    const ends = all(part(combined, "word/document.xml"), "bookmarkEnd").map(mark => mark.getAttributeNS(W, "id"))
    expect(ends.toSorted()).toEqual(marks.map(mark => mark.getAttributeNS(W, "id")).toSorted())
  }
})

test("identical inputs give identical bytes and the inputs are left unchanged", async () => {
  const memo = (await renderMarkdownToDocx(MEMO)).bytes
  const body = withParagraphPointers(templateBody())
  const copies = [memo.slice(), body.slice()]
  const first = combineMemoAndBody({ memo, body })
  expect(combineMemoAndBody({ memo, body })).toEqual(first)
  expect([memo, body]).toEqual(copies)
  const session = (await docxodus()).openDocxSession(first, {})
  try {
    expect(session.save().length).toBeGreaterThan(0)
  } finally {
    session.close()
  }
})

test("a body that opens with a table, or has no or a continuous first section, still starts on its own page", async () => {
  const memo = (await renderMarkdownToDocx(MEMO)).bytes
  for (const section of ["", `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`]) {
    const body = tableBody(section)
    const combined = combineMemoAndBody({ memo, body })
    const blocks = bodyBlocks(part(combined, "word/document.xml"))
    const split = bodyBlocks(part(memo, "word/document.xml")).length
    expect(sectionBreaks(combined)).toEqual([split - 1])
    expect(canonical(blocks[split]!)).toBe(canonical(bodyBlocks(part(body, "word/document.xml"))[0]!))
    const final = sectPr(combined)
    expect(all(final, "type").map(value).filter(type => type !== "nextPage")).toEqual([])
    expect(all(final, "pgNumType")[0]?.getAttributeNS(W, "start")).toBe("1")
    expect(await reopen(combined)).toContain("Parties")
  }
})

test("a memo that would need notes or embedded parts, or has no text, is refused; comment markers are dropped", async () => {
  const memo = (await renderMarkdownToDocx(MEMO)).bytes
  const body = (await renderMarkdownToDocx(BODY)).bytes
  const edit = (insert: string) => {
    const parts = unzipSync(memo)
    parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:t></w:r>", `</w:t></w:r>${insert}`))
    return zipSync(parts)
  }
  expect(() => combineMemoAndBody({ memo: edit(`<w:r><w:footnoteReference w:id="1"/></w:r>`), body })).toThrow("unsupported_memo_dependency")
  expect(() => combineMemoAndBody({ memo: edit(`<w:hyperlink xmlns:r="${O}" r:id="rId404"><w:r><w:t>x</w:t></w:r></w:hyperlink>`), body })).toThrow("unsupported_memo_dependency")
  expect(() => combineMemoAndBody({ memo: edit(`<w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${O}" r:embed="rId1"/></w:drawing></w:r>`), body })).toThrow("unsupported_memo_dependency")
  const empty = unzipSync(memo)
  empty["word/document.xml"] = strToU8(`<w:document xmlns:w="${W}"><w:body><w:p/><w:sectPr/></w:body></w:document>`)
  expect(() => combineMemoAndBody({ memo: zipSync(empty), body })).toThrow("empty_memo")

  const commented = combineMemoAndBody({ memo: edit(`<w:commentRangeStart w:id="0"/><w:r><w:t>kept</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>`), body })
  const xml = strFromU8(unzipSync(commented)["word/document.xml"]!)
  expect(xml).not.toMatch(/comment(RangeStart|RangeEnd|Reference)/)
  expect(xml).toContain("kept")
  expect(danglingStyles(commented)).toEqual([])
})

// A hand-built body in a firm template's shape: a different base font and
// justification, an auto-numbered Heading1 ("Article 1"), no Heading2, Quote
// or Code styles, A4 margins, a header part and a cross-reference bookmark.
// Its numbering nsid deliberately matches the one the renderer uses.
function templateBody() {
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  const paragraph = (text: string, style?: string) =>
    `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t>${text}</w:t></w:r></w:p>`
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${O}/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="${O}/header" Target="header1.xml"/></Relationships>`),
    "word/document.xml": xml(`<w:document xmlns:w="${W}" xmlns:r="${O}"><w:body>`
      + `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Ref100"/><w:r><w:t>Interpretation</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`
      + paragraph("Words defined in the schedule apply.")
      + paragraph("Payment", "Heading1")
      + `<w:p><w:r><w:t xml:space="preserve">Fees are due as stated in Article </w:t></w:r><w:fldSimple w:instr=" REF _Ref100 \\r "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t>.</w:t></w:r></w:p>`
      + `<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1701" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>`
      + `</w:body></w:document>`),
    "word/styles.xml": xml(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>`
      + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>`
      + `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:caps/></w:rPr></w:style>`
      + `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>`
      + `</w:styles>`),
    "word/numbering.xml": xml(`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="0D0CD001"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="Article %1"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`),
    "word/header1.xml": xml(`<w:hdr xmlns:w="${W}">${paragraph("Template header")}</w:hdr>`),
  }, { mtime: new Date(1980, 0, 1) })
}

// A minimal body with no styles, numbering or relationships part that opens
// with a table, under the given (possibly absent) section properties.
function tableBody(section: string) {
  const xml = (text: string) => strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`)
  return zipSync({
    "[Content_Types].xml": xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": xml(`<w:document xmlns:w="${W}"><w:body>`
      + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Parties</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
      + `<w:p><w:r><w:t>Agreed terms follow.</w:t></w:r></w:p>${section}</w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
}

function part(bytes: Uint8Array, name: string) {
  return new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[name]!), "application/xml")
}

function all(node: Document | Element, name: string) {
  return Array.from(node.getElementsByTagNameNS(W, name))
}

function value(node?: Element) {
  return node?.getAttributeNS(W, "val") ?? ""
}

function canonical(node: Node) {
  return new XMLSerializer().serializeToString(node)
}

function bodyBlocks(doc: Document) {
  return Array.from(all(doc, "body")[0]!.childNodes)
    .filter((node): node is Element => node.nodeType === 1 && (node as Element).localName !== "sectPr")
}

function texts(block: Element) {
  return (block.localName === "p" ? [block] : all(block, "p"))
    .map(paragraph => all(paragraph, "t").map(text => text.textContent).join(""))
    .filter(Boolean)
}

function sectPr(bytes: Uint8Array) {
  return Array.from(all(part(bytes, "word/document.xml"), "body")[0]!.childNodes)
    .find((node): node is Element => (node as Element).localName === "sectPr")!
}

// Top-level block indexes whose paragraph ends a section.
function sectionBreaks(bytes: Uint8Array) {
  return bodyBlocks(part(bytes, "word/document.xml")).flatMap((block, index) =>
    block.localName === "p" && all(block, "sectPr").length ? [index] : [])
}

function expectMemoSection(combined: Uint8Array, body: Uint8Array) {
  const blocks = bodyBlocks(part(combined, "word/document.xml"))
  const memo = all(blocks[sectionBreaks(combined)[0]!]!, "sectPr")[0]!
  expect(all(memo, "headerReference").length + all(memo, "footerReference").length).toBe(0)
  for (const name of ["pgSz", "pgMar"]) expect(canonical(all(memo, name)[0]!)).toBe(canonical(all(sectPr(body), name)[0]!))
}

// The body's section gains only a page-number restart.
function expectPageRestart(combined: Uint8Array, body: Uint8Array) {
  const section = sectPr(combined)
  expect(all(sectPr(body), "pgNumType")).toHaveLength(0)
  expect(all(section, "pgNumType")[0]?.getAttributeNS(W, "start")).toBe("1")
  section.removeChild(all(section, "pgNumType")[0]!)
  expect(canonical(section)).toBe(canonical(sectPr(body)))
}

// Every element of this kind in the source part is still present, unchanged.
function keptAll(source: Document, result: Document, name: string) {
  const kept = new Set(all(result, name).map(canonical))
  return all(source, name).length > 0 && all(source, name).every(node => kept.has(canonical(node)))
}

function danglingStyles(bytes: Uint8Array) {
  const styles = part(bytes, "word/styles.xml")
  const defined = new Set(all(styles, "style").map(style => style.getAttributeNS(W, "styleId")))
  const document = part(bytes, "word/document.xml")
  return [...["pStyle", "rStyle", "tblStyle"].flatMap(name => all(document, name)),
    ...["basedOn", "next", "link"].flatMap(name => all(styles, name))].map(value).filter(id => !defined.has(id))
}

// Each used list resolves to a definition, and list definitions stay distinct.
function numberingProblems(bytes: Uint8Array) {
  const numbering = part(bytes, "word/numbering.xml")
  const nums = new Map(all(numbering, "num").map(num => [num.getAttributeNS(W, "numId"), value(all(num, "abstractNumId")[0])]))
  const abstracts = new Set(all(numbering, "abstractNum").map(row => row.getAttributeNS(W, "abstractNumId")))
  const nsids = all(numbering, "nsid").map(value)
  const used = [...all(part(bytes, "word/document.xml"), "numId"), ...all(part(bytes, "word/styles.xml"), "numId")].map(value).filter(id => id !== "0")
  return [
    ...used.filter(id => !abstracts.has(nums.get(id) ?? "")).map(id => `num ${id}`),
    ...(new Set(nsids).size === nsids.length ? [] : ["duplicate nsid"]),
    ...(new Set(nums.values()).size === nums.size ? [] : ["shared abstractNum"]),
  ]
}

async function reopen(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, {})
  try {
    return session.project().markdown
  } finally {
    session.close()
  }
}
