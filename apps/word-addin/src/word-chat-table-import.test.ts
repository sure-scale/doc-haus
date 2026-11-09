import { expect, test } from "bun:test"
import { assertWordChatTableImportPreimage, assertWordChatTableImportRange, assertWordChatTableImportFormatting, WordChatTableImportError, type WordChatTableImportErrorCode } from "./word-chat-table-import"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", P = "http://schemas.microsoft.com/office/2006/xmlPackage"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const row = (values: string[]) => `<w:tr>${values.map(value => `<w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr>${p(value)}</w:tc>`).join("")}</w:tr>`
const table = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>${row(["Item", "Detail"])}${row(["Remove", "Entry"])}</w:tbl>`
const part = (name: string, xml: string, type = "application/xml") => `<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`
const rel = (id: string, kind: string, target: string) => `<Relationship Id="${id}" Type="${R}/${kind}" Target="${target}"/>`
const rels = (links: string) => `<Relationships xmlns="${REL}">${links}</Relationships>`
const style = part("/word/styles.xml", `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
const root = part("/_rels/.rels", rels(rel("root", "officeDocument", "word/document.xml")))
const links = part("/word/_rels/document.xml.rels", rels(rel("format", "styles", "styles.xml")))
const pack = (body: string, dependencies = root + links + style) => `<pkg:package xmlns:pkg="${P}">${part("/word/document.xml", `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`)}${dependencies}</pkg:package>`
const scoped = pack(table)
const fresh = pack(p("Outside") + table + p("After") + "<w:sectPr/>")
const inventory = ["Outside", "Item", "Detail", "Remove", "Entry", "After"].map((text, i) => ({id:`paragraph-${i}`, text, tableNestingLevel: i > 0 && i < 5 ? 1 : 0}))
const proposal = {kind:"table_row_delete" as const, id:"reviewed",documentId:"document",explanation:"Remove row",openItems:[],createdAt:"now",
  table:{tableIndex:0, rows:[inventory.slice(1,3),inventory.slice(3,5)]},rowIndex:1,cells:["Remove","Entry"]}

test("checks the complete parent preimage while leaving outside content and dependencies outside the import", () => {
  assertWordChatTableImportPreimage(scoped, fresh, inventory, proposal)
  const outsideComment = part("/word/comments.xml", `<w:comments xmlns:w="${W}"><w:comment w:id="9" w:author="Reviewer">${p("Keep")}</w:comment></w:comments>`)
  const fullLinks = part("/word/_rels/document.xml.rels", rels(rel("format", "styles", "styles.xml") + rel("review", "comments", "comments.xml")))
  const full = pack(p("Different outside") + table + p("After") + "<w:sectPr/>", root + fullLinks + style + outsideComment)
  assertWordChatTableImportPreimage(scoped, full, inventory.map((x,i) => i === 0 ? {...x,text:"Different outside"} : x), proposal)
  expect(() => assertWordChatTableImportPreimage(pack(table, root + links + style + outsideComment), full, inventory.map((x,i) => i === 0 ? {...x,text:"Different outside"} : x), proposal)).toThrow()
})

test("same-text table formatting and imported defaults cannot drift before dispatch", () => {
  expect(() => assertWordChatTableImportPreimage(scoped, fresh.replace('w:w="1800"', 'w:w="2400"'), inventory, proposal)).toThrow()
  expect(() => assertWordChatTableImportPreimage(scoped, fresh.replace('<w:tblPr/>', '<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>'), inventory, proposal)).toThrow()
  expect(() => assertWordChatTableImportPreimage(scoped, fresh.replace('w:val="22"', 'w:val="28"'), inventory, proposal)).toThrow()
  expect(() => assertWordChatTableImportPreimage(scoped, fresh.replace('Id="format"','Id="renamed"'), inventory, proposal)).toThrow()
  expect(() => assertWordChatTableImportPreimage(scoped, pack(p("Outside")+table+p("After"), root+links), inventory, proposal)).toThrow()
})

test("scoped table imports cannot leak a section or adjacent paragraph, or cross annotations", () => {
  for (const extra of ["<w:sectPr/>", p("Leak")]) expect(() => assertWordChatTableImportPreimage(pack(table+extra), fresh, inventory, proposal)).toThrow()
  const crossing = fresh.replace(p("Outside"), '<w:p><w:bookmarkStart w:id="9" w:name="outside"/><w:r><w:t>Outside</w:t></w:r></w:p>').replace(p("After"), '<w:p><w:r><w:t>After</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>')
  expect(() => assertWordChatTableImportPreimage(scoped, crossing, inventory, proposal)).toThrow()
  const unknown = crossing.replaceAll("w:bookmarkStart", "w:customXmlMoveFromRangeStart").replaceAll("w:bookmarkEnd", "w:customXmlMoveFromRangeEnd")
  expect(() => assertWordChatTableImportPreimage(scoped, unknown, inventory, proposal)).toThrow()
  const enclosed = table.replace(p("Item"), '<w:p><w:bookmarkStart w:id="9" w:name="inside"/><w:r><w:t>Item</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>')
  expect(() => assertWordChatTableImportPreimage(pack(enclosed), pack(p("Outside")+enclosed+p("After")), inventory, proposal)).toThrow()
})

test("expanded XML names allow prefix changes without ignoring identity or inherited context", () => {
  const prefixed = fresh.replaceAll("w:", "word:").replaceAll("xmlns:w=", "xmlns:word=")
  assertWordChatTableImportPreimage(scoped, prefixed, inventory, proposal)
  assertWordChatTableImportPreimage(scoped, fresh.replace('Id="root"','Id="different-root"'), inventory, proposal)
  expect(() => assertWordChatTableImportPreimage(scoped, fresh.replace("<w:body>", '<w:body xml:space="preserve">'), inventory, proposal)).toThrow()
  const identified = table.replace("<w:p>", '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="A1234567">')
  expect(() => assertWordChatTableImportPreimage(pack(identified), fresh, inventory, proposal)).toThrow()
})

test("compatibility prefix bindings and document siblings cannot hide a wider import", () => {
  const compatibility = (xml: string, uri: string) => xml.replace('<w:document ', `<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:extension="${uri}" mc:Ignorable="extension" `)
  assertWordChatTableImportPreimage(compatibility(scoped,"urn:extension:original"), compatibility(fresh,"urn:extension:original"), inventory, proposal)
  expect(() => assertWordChatTableImportPreimage(compatibility(scoped,"urn:extension:original"), compatibility(fresh,"urn:extension:different"), inventory, proposal)).toThrow()
  const dependencyCompatibility = (xml: string, uri: string) => xml.replace('<w:styles ', `<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="${uri}" mc:Ignorable="x" `)
  expect(() => assertWordChatTableImportPreimage(dependencyCompatibility(scoped,"urn:extension:original"), dependencyCompatibility(fresh,"urn:extension:different"), inventory, proposal)).toThrow()
  expect(() => assertWordChatTableImportFormatting(dependencyCompatibility(scoped,"urn:extension:original"), dependencyCompatibility(fresh,"urn:extension:different"))).toThrow()
  const choice = (xml: string) => xml.replace("</w:styles>", '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:extension"><mc:Choice Requires="x"/></mc:AlternateContent></w:styles>')
  expect(() => assertWordChatTableImportPreimage(choice(scoped), choice(fresh), inventory, proposal)).toThrow()
  const background = (xml: string) => xml.replace("<w:body>", '<w:background w:color="FFFFFF"/><w:body>')
  expect(() => assertWordChatTableImportPreimage(background(scoped), background(fresh), inventory, proposal)).toThrow()
})

test("native Whole-table ranges require every reviewed paragraph exactly once in order", () => {
  const ids = proposal.table.rows.flat().map(cell => cell.id)
  assertWordChatTableImportRange(ids, proposal)
  for (const actual of [[], [...ids,"outside"], ["outside",...ids], ids.slice(1), ids.slice(0,-1), [...ids].reverse(), [ids[0]!,ids[0]!,...ids.slice(2)]])
    expect(() => assertWordChatTableImportRange(actual, proposal)).toThrow(WordChatTableImportError)
})

test("categorical import notices distinguish failed boundaries without exposing captured content", () => {
  const ids = proposal.table.rows.flat().map(cell => cell.id)
  const enclosed = table.replace(p("Item"), '<w:p><w:bookmarkStart w:id="9" w:name="inside"/><w:r><w:t>Item</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>')
  const cases: Array<[WordChatTableImportErrorCode, () => void]> = [
    ["package_xml", () => assertWordChatTableImportPreimage("<!DOCTYPE package><private-capture/>", fresh, inventory, proposal)],
    ["document_context", () => assertWordChatTableImportPreimage(scoped, fresh.replace("<w:body>", '<w:body xml:space="preserve">'), inventory, proposal)],
    ["table_preimage", () => assertWordChatTableImportPreimage(scoped, fresh.replace('w:w="1800"', 'w:w="2400"'), inventory, proposal)],
    ["table_boundary", () => assertWordChatTableImportPreimage(pack(enclosed), pack(p("Outside")+enclosed+p("After")), inventory, proposal)],
    ["formatting_dependencies", () => assertWordChatTableImportPreimage(scoped, fresh.replace('w:val="22"', 'w:val="28"'), inventory, proposal)],
    ["range_paragraphs", () => assertWordChatTableImportRange([...ids,"private-capture"], proposal)],
    ["formatting_dependencies", () => assertWordChatTableImportFormatting(scoped, fresh.replace('w:val="22"', 'w:val="28"'))],
  ]
  const notices = new Set<string>()
  for (const [code, reject] of cases) {
    let caught: unknown
    try { reject() } catch (error) { caught = error }
    expect(caught).toBeInstanceOf(WordChatTableImportError)
    const error = caught as WordChatTableImportError
    expect(error.code).toBe(code)
    expect(error.message).toBe(new WordChatTableImportError(code).message)
    expect(error.message).not.toContain("private-capture")
    expect(error.message).not.toContain("styles.xml")
    expect(error.message).not.toContain("<")
    notices.add(error.message)
  }
  expect(notices.size).toBe(6)
})

test("wrapped native binding and unsupported XML causes stay private behind static notices", () => {
  let binding: unknown
  try { assertWordChatTableImportPreimage(scoped, fresh, inventory.map((p,i) => i === 1 ? {...p,id:"private-capture"} : p), proposal) }
  catch (error) { binding = error }
  expect(binding).toBeInstanceOf(WordChatTableImportError)
  expect((binding as WordChatTableImportError).code).toBe("table_preimage")
  expect((binding as WordChatTableImportError).cause).toBeInstanceOf(Error)
  expect((binding as WordChatTableImportError).message).not.toContain("private-capture")
  const choice = (xml: string) => xml.replace("</w:styles>", '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback/></mc:AlternateContent></w:styles>')
  let unsupported: unknown
  try { assertWordChatTableImportPreimage(choice(scoped), choice(fresh), inventory, proposal) }
  catch (error) { unsupported = error }
  expect(unsupported).toBeInstanceOf(WordChatTableImportError)
  expect((unsupported as WordChatTableImportError).code).toBe("formatting_dependencies")
  expect((unsupported as WordChatTableImportError).cause).toBeInstanceOf(Error)
  const privateCause = new Error("private-capture <w:p>secret text</w:p> /private/styles.xml")
  const safe = new WordChatTableImportError("package_xml", privateCause)
  expect(safe.cause).toBe(privateCause)
  expect(safe.message).not.toContain("private-capture")
  expect(safe.message).not.toContain("secret text")
  expect(safe.message).not.toContain("styles.xml")
})
