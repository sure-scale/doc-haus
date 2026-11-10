import { expect, test } from "bun:test"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { polishHeadingSpacing, polishMarkdownArtifactsInDocumentXml } from "../../lib/docx-polish"
import { carryOriginalTables } from "../../lib/docx-tables"
import { renderMarkdownToDocx } from "../../lib/render"
import { fillTemplate } from "../../lib/template-fill"
import { presentDraft } from "./draft-presentation"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml")
const main = (bytes: Uint8Array) => parse(strFromU8(unzipSync(bytes)["word/document.xml"]!))
const text = (element: Element) => Array.from(element.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join("")
function shape(node: Node): unknown {
  if (node.nodeType !== 1) return [node.nodeType, node.nodeValue]
  const element = node as Element
  return [element.namespaceURI, element.localName,
    Array.from(element.attributes).filter(attribute => attribute.namespaceURI !== "http://www.w3.org/2000/xmlns/")
      .map(attribute => [attribute.namespaceURI, attribute.localName, attribute.value]).sort(),
    Array.from(element.childNodes).map(shape)]
}

test("generated namespace adaptation preserves expanded names and QName bindings", () => {
  const xml = `<z:document xmlns:z='${W}' xmlns:f='urn:foreign' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' mc:Ignorable='f'><z:body><z:p f:marker='f:token'><z:r><z:t>Literal body</z:t></z:r></z:p><f:p><f:t>Foreign content</f:t></f:p></z:body></z:document>`
  const adapted = parse(polishMarkdownArtifactsInDocumentXml(xml))
  expect(shape(adapted.documentElement)).toEqual(shape(parse(xml).documentElement))
  expect(adapted.documentElement.lookupNamespaceURI("f")).toBe("urn:foreign")
  expect(adapted.documentElement.lookupNamespaceURI("z")).toBe(W)
  expect(adapted.documentElement.prefix).toBe("w")
  expect(() => polishMarkdownArtifactsInDocumentXml(xml.replace("xmlns:f='urn:foreign'", "xmlns:f='urn:foreign' xmlns:w='urn:foreign'")))
    .toThrow("conflicting Word namespace prefix")
})

test("generated full polish keeps its canonical SDK XML convention for house, style and revising seeds", async () => {
  const seed = unzipSync(await Bun.file(new URL("../../templates/_base.docx", import.meta.url)).bytes())
  seed["word/document.xml"] = strToU8(`<z:document xmlns:z='${W}'><z:body><z:p><z:r><z:t>SOURCE BODY MUST NOT LEAK</z:t></z:r></z:p><z:sectPr/></z:body></z:document>`)
  seed["word/styles.xml"] = strToU8(strFromU8(seed["word/styles.xml"]!).replaceAll("w:", "z:").replaceAll("xmlns:w", "xmlns:z"))
  const style = zipSync(seed)
  const literal = "| Code | Literal |\n\n<center>\n~~~\nend\\"
  const content = "# Generated title\n\n## Generated heading\n\n| Field | Value |\n|---|---|\n| Status | Ready |\n\n````\n" + literal + "\n````"
  for (const options of [{ styleTemplate: false as const }, { styleTemplate: style }, { styleTemplate: style, revising: true as const }]) {
    const rendered = await renderMarkdownToDocx(content, undefined, undefined, { ...options, numbering: false })
    const xml = strFromU8(unzipSync(rendered.bytes)["word/document.xml"]!)
    const document = parse(xml)
    expect(xml).not.toContain("SOURCE BODY MUST NOT LEAK")
    expect(document.getElementsByTagNameNS(W, "tbl").length).toBe(1)
    expect(Array.from(document.getElementsByTagNameNS(W, "*")).every(element => element.prefix === "w")).toBe(true)
    expect(Array.from(document.getElementsByTagNameNS(W, "*")).flatMap(element => Array.from(element.attributes))
      .filter(attribute => attribute.namespaceURI === W).every(attribute => attribute.prefix === "w")).toBe(true)
    expect(xml).toContain('<w:pStyle w:val="Heading2"')
    expect(xml).not.toContain("DocHausText")
    const code = Array.from(document.getElementsByTagNameNS(W, "p")).find(paragraph =>
      Array.from(paragraph.getElementsByTagNameNS(W, "rStyle")).some(style => style.getAttributeNS(W, "val") === "Code"))!
    expect(Array.from(code.getElementsByTagNameNS(W, "*")).map(element => element.localName === "t" ? element.textContent
      : element.localName === "br" ? "\n" : "").join("")).toBe(literal)
    const styles = parse(strFromU8(unzipSync(rendered.bytes)["word/styles.xml"]!))
    expect(Array.from(styles.getElementsByTagNameNS(W, "style")).filter(style => style.getAttributeNS(W, "styleId") === "Code")).toHaveLength(1)
  }
  // Template fill does not run full polish: literal restoration itself must
  // accept the SDK's preserved prefix before reopening for placeholder fills.
  const filled = main((await fillTemplate(style, { content: "~~~~\n[insert command]\n\nend\\\n~~~~",
    fills: { "[insert command]": "run | exact" } })).bytes)
  expect(Array.from(filled.getElementsByTagNameNS(W, "t")).map(node => node.textContent).join("")).toBe("run | exactend\\")
  expect(filled.getElementsByTagNameNS(W, "br").length).toBe(2)
})

test("carried source table paragraphs never become presentation labels or promoted headings", async () => {
  const rendered = await renderMarkdownToDocx("# Instrument\n\n## Hidden boundary\n\n### Promoted body\n\nTable lead\n\n| Hidden boundary | Value |\n|---|---|\n| Carried cell | Exact |", undefined, undefined, { numbering: false, styleTemplate: false })
  const parts = unzipSync(rendered.bytes)
  const paragraph = (style: string, value: string) => `<w:p><w:pPr><w:pStyle w:val='${style}'/><w:spacing w:before='91' w:after='37' w:line='277'/></w:pPr><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>${value}</w:t></w:r></w:p>`
  const table = `<w:tbl><w:tblPr><w:tblW w:w='7531' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='13'/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w='3001'/><w:gridCol w:w='4530'/></w:tblGrid>`
    + `<w:tr><w:trPr><w:trHeight w:val='613'/><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w='3001' w:type='dxa'/></w:tcPr>${paragraph("Heading2", "Hidden boundary")}</w:tc><w:tc>${paragraph("Heading3", "Value")}</w:tc></w:tr>`
    + `<w:tr><w:tc>${paragraph("Heading3", "Carried cell")}</w:tc><w:tc>${paragraph("Heading1", "Exact")}</w:tc></w:tr></w:tbl>`
  for (const prefix of ["w", "字"]) {
    const sourceXml = `<w:document xmlns:w='${W}'><w:body><w:p><w:r><w:t>Table lead</w:t></w:r></w:p>${table}<w:sectPr/></w:body></w:document>`
      .replaceAll("w:", `${prefix}:`).replaceAll("xmlns:w", `xmlns:${prefix}`)
    const original = zipSync({ ...parts, "word/document.xml": strToU8(sourceXml) })
    const carried = carryOriginalTables(rendered.bytes, original)
    expect(carried.carries.map(carry => carry.outcome)).toEqual(["carried"])
    const before = shape(main(carried.bytes).getElementsByTagNameNS(W, "tbl")[0]!)
    for (const spacing of [true, false]) {
      const presented = presentDraft(carried.bytes, [{ heading: "Hidden boundary", display_heading: "" }], true, spacing).bytes
      const result = main(presented)
      expect(shape(result.getElementsByTagNameNS(W, "tbl")[0]!)).toEqual(before)
      const bodyHeading = Array.from(result.getElementsByTagNameNS(W, "p")).find(paragraph => text(paragraph) === "Promoted body")!
      expect(bodyHeading.getElementsByTagNameNS(W, "pStyle")[0]!.getAttributeNS(W, "val")).toBe("Heading2")
      expect(bodyHeading.getElementsByTagNameNS(W, "spacing")[0]!.getAttributeNS(W, "before")).toBe(spacing ? "280" : "220")
      const beforeParts = unzipSync(carried.bytes)
      const afterParts = unzipSync(presented)
      expect(Object.keys(afterParts).sort()).toEqual(Object.keys(beforeParts).sort())
      for (const name of Object.keys(beforeParts).filter(name => name !== "word/document.xml")) expect(afterParts[name]).toEqual(beforeParts[name])
    }
  }
})

test("mixed-source heading-only polish uses expanded names and direct current properties", () => {
  for (const [prefix, declaration] of [["字:", `xmlns:字='${W}'`], ["", `xmlns='${W}'`], ["w:", `xmlns:w='${W}'`]]) {
    const attribute = prefix || "a:"
    const xml = `<${prefix}document ${declaration} xmlns:a='${W}' xmlns:f='urn:foreign'><${prefix}body><${prefix}p><${prefix}pPr><${prefix}pStyle ${attribute}val='Heading2'/><${prefix}spacing ${attribute}before='9'/><${prefix}pPrChange ${attribute}id='42'><${prefix}pPr><${prefix}pStyle ${attribute}val='Heading1'/><${prefix}spacing ${attribute}before='999'/></${prefix}pPr></${prefix}pPrChange></${prefix}pPr><${prefix}r><${prefix}t>Body</${prefix}t></${prefix}r></${prefix}p><f:p><f:pPr><f:pStyle f:val='Heading2'/></f:pPr></f:p></${prefix}body></${prefix}document>`
    const result = parse(polishHeadingSpacing(xml))
    const paragraph = result.getElementsByTagNameNS(W, "p")[0]!
    const spacings = paragraph.getElementsByTagNameNS(W, "spacing")
    expect(spacings[0]!.getAttributeNS(W, "before")).toBe("280")
    expect(spacings[1]!.getAttributeNS(W, "before")).toBe("999")
    expect(result.getElementsByTagNameNS("urn:foreign", "spacing").length).toBe(0)
  }
})
