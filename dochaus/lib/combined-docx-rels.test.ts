import { expect, test } from "bun:test"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { combineMemoAndBody } from "./combined-docx"
import { docxodus } from "./docxodus"
import { renderMarkdownToDocx } from "./render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const W16SE = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
const XMLNS = "http://www.w3.org/2000/xmlns/"
// A Word-style notes root: the prefixes are declared once, here.
const ROOT = [`xmlns:w="${W}"`, `xmlns:r="${O}"`, `xmlns:mc="${MC}"`, `xmlns:w14="${W14}"`, `xmlns:w16se="${W16SE}"`,
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`,
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`,
  `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`,
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`,
  `xmlns:v="urn:schemas-microsoft-com:vml"`, `xmlns:o="urn:schemas-microsoft-com:office:office"`, `mc:Ignorable="w14 w16se"`].join(" ")
const IMAGE = `<Relationship Id="rId9" Type="${O}/image" Target="media/image1.png"/>`
const EMOJI = `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w16se"><w16se:symEx w16se:font="Segoe UI Emoji" w16se:char="1F600"/></mc:Choice><mc:Fallback><w:t>:)</w:t></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>Emoji line.</w:t></w:r></w:p>`

test("a memo with an inline image or any other drawing is refused, not copied", async () => {
  const body = await rendered("# Services Agreement\n\nThe term is one year.")
  const inline = (graphic: string) => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Picture 1"/><a:graphic>${graphic}</a:graphic></wp:inline></w:drawing></w:r></w:p>`
  const image = await memoWith(inline(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData>`), IMAGE)
  expect(() => combineMemoAndBody({ memo: image, body })).toThrow("unsupported_memo_dependency")
  // A shape needs no part, but its drawing id may already be the body's.
  const shape = await memoWith(inline(`<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr/><wps:bodyPr/></wps:wsp></a:graphicData>`))
  expect(() => combineMemoAndBody({ memo: shape, body })).toThrow("unsupported_memo_dependency")
})

test("a memo with a VML image, an embedded object or any other element pointing at a part is refused", async () => {
  const body = await rendered("# Services Agreement\n\nThe term is one year.")
  for (const run of [
    `<w:r><w:pict><v:shape id="_x0000_i1025" style="width:10pt;height:10pt"><v:imagedata r:id="rId9" o:title=""/></v:shape></w:pict></w:r>`,
    `<w:r><w:object><o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1" r:id="rId9"/></w:object></w:r>`,
    `<w:r><mc:AlternateContent><mc:Choice Requires="w14"><w14:linked r:id="rId9"/></mc:Choice><mc:Fallback><w:t>x</w:t></mc:Fallback></mc:AlternateContent></w:r>`,
  ]) {
    const memo = await memoWith(`<w:p>${run}</w:p>`, IMAGE)
    expect(() => combineMemoAndBody({ memo, body })).toThrow("unsupported_memo_dependency")
  }
})

test("the memo root's namespaces and ignorable prefixes move onto the combined root", async () => {
  const memo = await memoWith(EMOJI)
  for (const root of [`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}" mc:Ignorable="w14"`]) {
    const combined = combineMemoAndBody({ memo, body: await rendered("# Services Agreement\n\nThe term is one year.", root) })
    const document = part(combined, "word/document.xml").documentElement
    expect(document.getAttributeNS(XMLNS, "w16se")).toBe(W16SE)
    expect(document.getAttributeNS(MC, "Ignorable")?.split(" ")).toEqual(["w14", "w16se"])
    // Requires names a prefix, so it must be in scope where the choice is.
    expect(strFromU8(unzipSync(combined)["word/document.xml"]!).match(/xmlns:w16se=/g)).toHaveLength(1)
    const markdown = await reopen(combined)
    expect(markdown).toContain("Emoji line.")
    expect(markdown).toContain("Services Agreement")
  }
})

test("a body root prefix is never rebound", async () => {
  const memo = await memoWith(EMOJI)
  const shared = `xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14"`
  const root = part(combineMemoAndBody({ memo, body: await rendered("# Services Agreement", shared) }), "word/document.xml").documentElement
  expect(declarations(root)).toMatchObject({ w: W, w14: W14, mc: MC })
  // The memo's w16se means something else in the body: refused, not rebound.
  const clash = await rendered("# Services Agreement", `${shared} xmlns:w16se="urn:example:other"`)
  expect(() => combineMemoAndBody({ memo, body: clash })).toThrow("unsupported_memo_namespace")
})

test("a memo root prefix the body binds elsewhere is left out unless the memo content uses it", async () => {
  const W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
  const declared = async (blocks: string) => {
    const parts = unzipSync(await memoWith(blocks))
    parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace(`mc:Ignorable="w14 w16se"`, `xmlns:w16cid="${W16CID}" mc:Ignorable="w14 w16se w16cid"`))
    return zipSync(parts, { mtime: new Date(1980, 0, 1) })
  }
  const body = await rendered("# Services Agreement", `xmlns:w="${W}" xmlns:w16cid="urn:example:other"`)
  const combined = combineMemoAndBody({ memo: await declared(`<w:p><w:r><w:t>Plain line.</w:t></w:r></w:p>`), body })
  const root = part(combined, "word/document.xml").documentElement
  expect(declarations(root)).toMatchObject({ w14: W14, w16cid: "urn:example:other" })
  expect(root.getAttributeNS(MC, "Ignorable")?.split(" ")).toEqual(["w14", "w16se"])
  expect(await reopen(combined)).toContain("Plain line.")
  // Used on an element or attribute, or named by a markup-compatibility value.
  for (const blocks of [
    `<w:p w16cid:durableId="1"><w:r><w:t>x</w:t></w:r></w:p>`,
    `<w:p><mc:AlternateContent><mc:Choice Requires="w16cid"><w:r><w:t>x</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>x</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>`,
    `<w:p mc:Ignorable="w16cid"><w:r><w:t>x</w:t></w:r></w:p>`,
  ]) {
    const memo = await declared(blocks)
    expect(() => combineMemoAndBody({ memo, body })).toThrow("unsupported_memo_namespace")
  }
})

test("copied styles and lists bring the memo part roots' namespaces and ignorable prefixes", async () => {
  const W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
  const notes = unzipSync((await renderMarkdownToDocx("# Notes\n\n1. One\n2. Two\n\nText.")).bytes)
  const word = (name: string, prefix: string, uri: string, xml: string) =>
    xml.replace(new RegExp(`<w:${name}([^>]*)>`), `<w:${name}$1 xmlns:mc="${MC}" xmlns:${prefix}="${uri}" mc:Ignorable="${prefix}">`)
  notes["word/styles.xml"] = strToU8(word("styles", "w14", W14, strFromU8(notes["word/styles.xml"]!))
    .replace(/(w:styleId="Heading1".*?)<\/w:rPr>/s, `$1<w14:ligatures w14:val="standard"/></w:rPr>`))
  notes["word/numbering.xml"] = strToU8(word("numbering", "w15", W15, strFromU8(notes["word/numbering.xml"]!))
    .replace(/<w:abstractNum /g, `<w:abstractNum w15:restartNumberingAfterBreak="0" `))
  const combined = combineMemoAndBody({ memo: zipSync(notes), body: (await renderMarkdownToDocx("# Agreement\n\n1. Alpha\n2. Beta")).bytes })
  for (const [name, prefix, uri] of [["word/styles.xml", "w14", W14], ["word/numbering.xml", "w15", W15]] as const) {
    const root = part(combined, name).documentElement
    expect(root.getAttributeNS(XMLNS, prefix)).toBe(uri)
    expect(root.getAttributeNS(MC, "Ignorable")?.split(" ")).toContain(prefix)
    // Declared once, on the root, so the Ignorable list resolves.
    expect(strFromU8(unzipSync(combined)[name]!).match(new RegExp(`xmlns:${prefix}=`, "g"))).toHaveLength(1)
  }
  expect(strFromU8(unzipSync(combined)["word/styles.xml"]!)).toMatch(/w:styleId="MemoHeading1".*?<w14:ligatures/s)
  expect(strFromU8(unzipSync(combined)["word/numbering.xml"]!)).toContain("w15:restartNumberingAfterBreak")
  expect(await reopen(combined)).toContain("Agreement")
})

test("a memo's own section breaks keep their page setup, type and columns, but no header, so one added in Word never blocks it", async () => {
  const body = await rendered("# Services Agreement\n\nThe term is one year.")
  const setup = `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`
  const section = `<w:sectPr><w:headerReference w:type="default" r:id="rId8"/><w:footerReference w:type="default" r:id="rId8"/>`
    + `<w:type w:val="continuous"/>${setup}<w:cols w:num="2"/><w:titlePg/><w:printerSettings r:id="rId8"/></w:sectPr>`
  const memo = await memoWith(`<w:p><w:pPr>${section}</w:pPr><w:r><w:t>Landscape part.</w:t></w:r></w:p><w:p><w:r><w:t>Last memo line.</w:t></w:r></w:p>`,
    `<Relationship Id="rId8" Type="${O}/header" Target="header1.xml"/>`)
  const combined = combineMemoAndBody({ memo, body })
  const sections = Array.from(part(combined, "word/document.xml").getElementsByTagNameNS(W, "sectPr"))
  expect(sections).toHaveLength(3)
  const [inner, memoEnd] = sections
  const names = (section: Element) => Array.from(section.childNodes).filter(row => row.nodeType === 1).map(row => (row as Element).localName)
  // A continuous two-column part stays one: its type and columns are layout,
  // not a dependency on another part.
  expect(names(inner!)).toEqual(["type", "pgSz", "pgMar", "cols"])
  expect(inner!.getElementsByTagNameNS(W, "type")[0]!.getAttributeNS(W, "val")).toBe("continuous")
  expect(inner!.getElementsByTagNameNS(W, "cols")[0]!.getAttributeNS(W, "num")).toBe("2")
  expect(inner!.getElementsByTagNameNS(W, "pgSz")[0]!.getAttributeNS(W, "orient")).toBe("landscape")
  // The break before the body still takes only the body's page setup.
  expect(names(memoEnd!)).toEqual(["type", "pgSz", "pgMar"])
  expect(memoEnd!.getElementsByTagNameNS(W, "type")[0]!.getAttributeNS(W, "val")).toBe("nextPage")
  for (const memoSection of [inner!, memoEnd!])
    expect(["headerReference", "footerReference", "printerSettings"].flatMap(name => Array.from(memoSection.getElementsByTagNameNS(W, name)))).toEqual([])
  expect(strFromU8(unzipSync(combined)["word/_rels/document.xml.rels"]!)).not.toContain(`${O}/header`)
  const markdown = await reopen(combined)
  expect(markdown).toContain("Landscape part.")
  expect(markdown.indexOf("Last memo line.")).toBeLessThan(markdown.indexOf("Services Agreement"))
})

test("a memo that binds WordprocessingML to another prefix is rewritten in place, with no attribute doubled", async () => {
  const body = await rendered("# Services Agreement\n\nThe term is one year.")
  // A body paragraph pointer the memo also uses, so the memo's is renamed.
  const pointer = strFromU8(unzipSync(body)["word/document.xml"]!).match(/w:name="(_dh_p[0-9a-f]{12})"/)![1]!
  const prefixed = (xml: string) => xml.replace(/xmlns:w=/g, "xmlns:x=").replace(/(<\/?|\s)w:/g, "$1x:")
  const notes = unzipSync((await renderMarkdownToDocx("# Notes\n\n1. One\n2. Two\n\nText.")).bytes)
  // A changed, already custom heading style is copied under a new id.
  notes["word/styles.xml"] = strToU8(prefixed(strFromU8(notes["word/styles.xml"]!))
    .replace(/x:styleId="Heading1"(.*?)<\/x:rPr>/s, `x:styleId="Heading1" x:customStyle="1"$1<x:u x:val="single"/></x:rPr>`))
  notes["word/numbering.xml"] = strToU8(prefixed(strFromU8(notes["word/numbering.xml"]!)))
  notes["word/document.xml"] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:document xmlns:x="${W}"><x:body>`
    + `<x:p><x:pPr><x:pStyle x:val="Heading1"/></x:pPr><x:bookmarkStart x:id="7" x:name="${pointer}"/><x:r><x:t>Cover memo</x:t></x:r><x:bookmarkEnd x:id="7"/></x:p>`
    + `<x:p><x:pPr><x:numPr><x:ilvl x:val="0"/><x:numId x:val="1"/></x:numPr></x:pPr><x:r><x:t>Listed point.</x:t></x:r></x:p>`
    + `<x:p><x:fldSimple x:instr=" REF ${pointer} \\h "><x:r><x:t>1</x:t></x:r></x:fldSimple><x:hyperlink x:anchor="${pointer}"><x:r><x:t>jump</x:t></x:r></x:hyperlink></x:p>`
    + `<x:sectPr/></x:body></x:document>`)
  const combined = combineMemoAndBody({ memo: zipSync(notes, { mtime: new Date(1980, 0, 1) }), body })

  for (const name of ["word/document.xml", "word/styles.xml", "word/numbering.xml"]) {
    for (const node of Array.from(part(combined, name).getElementsByTagName("*"))) {
      const keys = Array.from(node.attributes).map(row => `${row.namespaceURI} ${row.localName}`)
      expect(keys.filter((key, index) => keys.indexOf(key) !== index)).toEqual([])
    }
  }
  const document = part(combined, "word/document.xml")
  const [heading, listed, links] = Array.from(document.getElementsByTagNameNS(W, "p"))
  const renamed = heading!.getElementsByTagNameNS(W, "bookmarkStart")[0]!.getAttributeNS(W, "name")
  expect(renamed).toMatch(/^_dh_p[0-9a-f]{12}$/)
  expect(renamed).not.toBe(pointer)
  expect(links!.getElementsByTagNameNS(W, "fldSimple")[0]!.getAttributeNS(W, "instr")).toBe(` REF ${renamed} \\h `)
  expect(links!.getElementsByTagNameNS(W, "hyperlink")[0]!.getAttributeNS(W, "anchor")).toBe(renamed)
  expect(heading!.getElementsByTagNameNS(W, "pStyle")[0]!.getAttributeNS(W, "val")).toBe("MemoHeading1")
  const numId = listed!.getElementsByTagNameNS(W, "numId")[0]!.getAttributeNS(W, "val")
  expect(Array.from(part(combined, "word/numbering.xml").getElementsByTagNameNS(W, "num")).map(num => num.getAttributeNS(W, "numId"))).toContain(numId)
  const markdown = await reopen(combined)
  for (const text of ["Cover memo", "Listed point.", "Services Agreement"]) expect(markdown).toContain(text)
})

async function rendered(markdown: string, root?: string) {
  const parts = unzipSync((await renderMarkdownToDocx(markdown)).bytes)
  if (root) parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace(/<w:document[^>]*>/, `<w:document ${root}>`))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

// The rendered notes with a Word-style document part holding the given blocks,
// and an optional extra relationship with its media part.
async function memoWith(blocks: string, relationship?: string) {
  const parts = unzipSync(await rendered("# Cover memo"))
  parts["word/document.xml"] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${ROOT}><w:body><w:p><w:r><w:t>Cover memo</w:t></w:r></w:p>${blocks}<w:sectPr/></w:body></w:document>`)
  if (relationship) {
    parts["word/_rels/document.xml.rels"] = strToU8(strFromU8(parts["word/_rels/document.xml.rels"]!).replace("</Relationships>", `${relationship}</Relationships>`))
    parts["word/media/image1.png"] = new Uint8Array([0x89, 0x50, 0x4e, 0x47])
  }
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

function part(bytes: Uint8Array, name: string) {
  return new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[name]!), "application/xml")
}

function declarations(node: Element) {
  return Object.fromEntries(Array.from(node.attributes).filter(row => row.namespaceURI === XMLNS && row.prefix === "xmlns").map(row => [row.localName, row.value]))
}

async function reopen(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, {})
  try {
    return session.project().markdown
  } finally {
    session.close()
  }
}
