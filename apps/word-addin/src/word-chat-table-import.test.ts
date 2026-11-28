import { describe, expect, test } from "bun:test"
import { assertWordChatTableImportPreimage, assertWordChatTableImportRange, assertWordChatTableImportFormatting, WordChatTableImportError, type WordChatTableImportErrorCode } from "./word-chat-table-import"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strToU8, unzipSync, zipSync } from "fflate"
import { firstTableImportXmlDifference, sameTableImportXml, sameLiveTableImportPreimageXml } from "./word-chat-table-import-xml"
import { verifyWordChatTableImportArtifact } from "./word-chat-table-import-artifact"

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
    ["package_xml", () => assertWordChatTableImportPreimage(scoped, fresh+"<!ENTITY private-capture>", inventory, proposal)],
    ["document_context", () => assertWordChatTableImportPreimage(scoped, fresh.replace("<w:body>", '<w:body xml:space="preserve">'), inventory, proposal)],
    ["live_table_preimage", () => assertWordChatTableImportPreimage(scoped, fresh.replace('w:w="1800"', 'w:w="2400"'), inventory, proposal)],
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
  expect((binding as WordChatTableImportError).code).toBe("target_validation")
  expect((binding as WordChatTableImportError).detail).toEqual({category:"target_validation",path:"/target",reason:"target_binding"})
  expect((binding as WordChatTableImportError).cause).toBeInstanceOf(Error)
  expect((binding as WordChatTableImportError).message).not.toContain("private-capture")
  const choice = (xml: string) => xml.replace("</w:styles>", '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback/></mc:AlternateContent></w:styles>')
  let unsupported: unknown
  try { assertWordChatTableImportPreimage(choice(scoped), choice(fresh), inventory, proposal) }
  catch (error) { unsupported = error }
  expect(unsupported).toBeInstanceOf(WordChatTableImportError)
  expect((unsupported as WordChatTableImportError).code).toBe("xml_comparison_unsupported")
  expect((unsupported as WordChatTableImportError).detail).toMatchObject({category:"unsupported_xml",reason:"compatibility_element"})
  expect((unsupported as WordChatTableImportError).cause).toBeInstanceOf(Error)
  const privateCause = new Error("private-capture <w:p>secret text</w:p> /private/styles.xml")
  const safe = new WordChatTableImportError("package_xml", privateCause)
  expect(safe.cause).toBe(privateCause)
  expect(safe.message).not.toContain("private-capture")
  expect(safe.message).not.toContain("secret text")
  expect(safe.message).not.toContain("styles.xml")
})

test("target diagnostics classify exact static projector and target reasons without captured values", () => {
  const cases = [
    ["paragraph_inventory",()=>assertWordChatTableImportPreimage(scoped,fresh,[],proposal)],
    ["paragraph_projection",()=>assertWordChatTableImportPreimage(scoped,fresh,inventory.map((p,i)=>i===1?{...p,text:"private-capture"}:p),proposal)],
    ["story_structure",()=>assertWordChatTableImportPreimage(scoped,fresh.replace("<w:body>",'<w:body><w:private-capture/>'),inventory,proposal)],
    ["table_geometry",()=>assertWordChatTableImportPreimage(scoped,fresh,inventory,{...proposal,table:{...proposal.table,tableIndex:99}})],
    ["target_qualification",()=>assertWordChatTableImportPreimage(scoped,fresh.replace(row(["Remove","Entry"]),row(["Remove","Entry"]).replace("<w:tr>","<w:tr><w:trPr><w:tblHeader/></w:trPr>")),inventory,proposal)],
  ] as const
  for(const [reason,run] of cases){
    let error:WordChatTableImportError|undefined
    try{run()}catch(cause){error=cause as WordChatTableImportError}
    expect(error?.code).toBe("target_validation")
    expect(error?.detail).toEqual({category:"target_validation",path:"/target",reason})
    expect(error?.message).not.toContain("private-capture")
  }
})

test("strict live-table differences expose allowlisted paths but never values or unknown XML names", () => {
  const cases = [
    [scoped.replace('w:w="1800"','w:w="private-width"'),"attribute_value","/w:tbl[1]/w:tblGrid[2]/w:gridCol[1]/@w:w"],
    [scoped.replace("<w:tbl>",'<w:tbl w:rsidR="private-revision">'),"attribute_presence","/w:tbl[1]/@w:rsidR"],
    [scoped.replace("<w:tbl>",'<w:tbl private-capture="private-value">'),"attribute_presence","/w:tbl[1]/@attribute"],
    [scoped.replace("Remove","private-text"),"text","/w:tbl[1]/w:tr[4]/w:tc[1]/w:p[2]/w:r[1]/w:t[1]/node[1]"],
    [scoped.replace("</w:tbl>",'<private-capture xmlns="urn:private-namespace">private-text</private-capture></w:tbl>'),"node_count","/w:tbl[1]/element[5]"],
  ] as const
  for(const [source,category,path] of cases){
    let error:WordChatTableImportError|undefined
    try{assertWordChatTableImportPreimage(source,fresh,inventory,proposal)}catch(cause){error=cause as WordChatTableImportError}
    expect(error?.code).toBe("live_table_preimage")
    expect(error?.detail).toEqual({category,path})
    expect(error?.diagnosticMessage).toContain(`${category} at ${path}`)
    expect(error?.message).toBe(new WordChatTableImportError("live_table_preimage").message)
    expect(error?.diagnosticMessage).not.toContain("private-")
    expect(error?.message).not.toContain("private-")
  }
})

test("XML difference reporting retains QName semantics and hides namespace binding values", () => {
  const parse=(xml:string)=>new DOMParser().parseFromString(xml,"application/xml").documentElement
  const a=parse(`<w:p xmlns:w="${W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:privatePrefix="urn:private-before" mc:Ignorable="privatePrefix">${p("private-text")}</w:p>`)
  const b=parse(`<other:p xmlns:other="${W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:privatePrefix="urn:private-after" mc:Ignorable="privatePrefix">${p("private-text").replaceAll("w:","other:")}</other:p>`)
  expect(sameTableImportXml(a,b)).toBe(false)
  expect(firstTableImportXmlDifference(a,b)).toEqual({category:"namespace_binding",path:"/w:p[1]/@mc:Ignorable"})
  const same=parse(`<other:p xmlns:other="${W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:privatePrefix="urn:private-before" mc:Ignorable="privatePrefix">${p("private-text").replaceAll("w:","other:")}</other:p>`)
  expect(sameTableImportXml(a,same)).toBe(true)
  expect(firstTableImportXmlDifference(a,same)).toBeUndefined()
})

test("prototype-key namespace URIs stay comparable and cannot become diagnostic names", () => {
  const parse=(xml:string)=>new DOMParser().parseFromString(xml,"application/xml").documentElement
  for(const uri of ["constructor","__proto__","toString"]){
    const a=parse(`<x:private-capture xmlns:x="${uri}" x:private-attribute="private-value">private-text</x:private-capture>`)
    const same=parse(`<y:private-capture xmlns:y="${uri}" y:private-attribute="private-value">private-text</y:private-capture>`)
    const changed=parse(`<y:private-capture xmlns:y="${uri}" y:private-attribute="private-changed">private-text</y:private-capture>`)
    expect(sameTableImportXml(a,same)).toBe(true)
    expect(firstTableImportXmlDifference(a,same)).toBeUndefined()
    expect(sameTableImportXml(a,changed)).toBe(false)
    expect(firstTableImportXmlDifference(a,changed)).toEqual({category:"attribute_value",path:"/element[1]/@attribute"})
  }
})

test("attribute presence diagnostics identify the actual added or removed expanded name", () => {
  const parse=(attrs:string)=>new DOMParser().parseFromString(`<w:tbl xmlns:w="${W}" ${attrs}/>`,"application/xml").documentElement
  const before=parse('w:val="private-same"'),after=parse('w:id="private-new" w:val="private-same"')
  expect(sameTableImportXml(before,after)).toBe(false)
  expect(firstTableImportXmlDifference(before,after)).toEqual({category:"attribute_presence",path:"/w:tbl[1]/@w:id"})
  expect(firstTableImportXmlDifference(after,before)).toEqual({category:"attribute_presence",path:"/w:tbl[1]/@w:id"})
})

test("only valid addition-session export variation is ignored by the live source preimage", () => {
  const reviewed = table.replace(p("Item"),'<w:p><w:del w:id="7" w:author="Reviewer" w:date="2026-02-01T00:00:00Z"><w:r><w:delText>Older</w:delText></w:r></w:del><w:ins w:id="8" w:author="Reviewer" w:date="2026-02-01T00:00:00Z"><w:r><w:t>Item</w:t></w:r></w:ins></w:p>')
    .replaceAll("<w:tr>",'<w:tr w:rsidR="1234ABCD">')
    .replaceAll("<w:p>",'<w:p w:rsidR="1234ABCD" w:rsidRDefault="1234ABCD">')
    .replaceAll("<w:r>",'<w:r w:rsidR="1234ABCD">')
  const live = reviewed.replaceAll('1234ABCD','ABCD5678')
  const source=pack(reviewed),current=(value:string)=>pack(p("Outside")+value+p("After")+"<w:sectPr/>")
  assertWordChatTableImportPreimage(source,current(live),inventory,proposal)
  assertWordChatTableImportPreimage(source,current(live.replaceAll(/ w:rsidR(?:Default)?="ABCD5678"/g,"")),inventory,proposal)
  assertWordChatTableImportPreimage(pack(reviewed.replaceAll(/ w:rsidR(?:Default)?="1234ABCD"/g,"")),current(live),inventory,proposal)
  const parse=(value:string)=>new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${value}</w:body>`,"application/xml").documentElement
  expect(sameTableImportXml(parse(reviewed),parse(live))).toBe(false)
  expect(sameLiveTableImportPreimageXml(parse(reviewed),parse(live))).toBe(true)
  for(const changed of [
    live.replace('w:w="1800"','w:w="2400"'),
    live.replace("Item","Changed"),
    live.replace("Older","Changed old content"),
    live.replace('w:id="7"','w:id="9"'),
    live.replace('w:author="Reviewer"','w:author="Other"'),
    live.replace('w:date="2026-02-01T00:00:00Z"','w:date="2026-02-02T00:00:00Z"'),
    live.replace('w:rsidR="ABCD5678"','w:rsidR="not-hex"'),
    live.replace('w:rsidRDefault="ABCD5678"','w:rsidRDefault="not-hex"'),
    live.replace('<w:p ', '<w:p w:rsidP="ABCD5678" '),
    live.replace('<w:r ', '<w:r w:rsidRPr="ABCD5678" '),
    live.replace('<w:tbl>', '<w:tbl w:rsidR="ABCD5678">'),
    live.replace('<w:tr w:rsidR="ABCD5678">','<w:tr w:rsidR="ABCD5678" w:rsidDel="12345678">'),
  ]) expect(()=>assertWordChatTableImportPreimage(source,current(changed),inventory,proposal)).toThrow(WordChatTableImportError)
  let error:WordChatTableImportError|undefined
  try{assertWordChatTableImportPreimage(source,current(live.replace('w:w="1800"','w:w="2400"')),inventory,proposal)}catch(cause){error=cause as WordChatTableImportError}
  expect(error?.code).toBe("live_table_preimage")
  expect(error?.detail).toEqual({category:"attribute_value",path:"/w:tbl[1]/w:tblGrid[2]/w:gridCol[1]/@w:w"})
})

test("bare formatting failures report fixed families and reasons without paths, IDs or values", () => {
  const check=(sourceDependencies:string,currentDependencies=root+links+style)=>()=>assertWordChatTableImportPreimage(pack(table,sourceDependencies),pack(p("Outside")+table+p("After"),currentDependencies),inventory,proposal)
  const cases = [
    ["dependency","part_not_allowed",check(root+links+style+part("/word/private-capture.xml","<private-capture/>"))],
    ["dependency","part_missing",check(root+links+style,root+links)],
    ["dependency","content_type",check(root+links+style.replace('pkg:contentType="application/xml"','pkg:contentType="private-value"'))],
    ["root","relationship_data",check(root.replace("</pkg:part>","<pkg:xmlData/></pkg:part>")+links+style)],
    ["root","relationship_root",check(part("/_rels/.rels","<private-capture/>")+links+style)],
    ["root","relationship_entry",check(root.replace('Id="root"',"")+links+style)],
    ["root","relationship_ids",check(part("/_rels/.rels",rels(rel("root","officeDocument","word/document.xml").repeat(2)))+links+style)],
    ["root","external_relationship",check(root.replace('Id="root"','Id="root" TargetMode="External"')+links+style)],
    ["root","root_shape",check(part("/_rels/.rels",rels(rel("one","officeDocument","word/document.xml")+rel("two","officeDocument","word/document.xml")))+links+style)],
    ["root","root_link_match",check(root+links+style,root.replace('Target="word/document.xml"','Target="private-value.xml"')+links+style)],
    ["root","relationship_target_equivalent",check(root+links+style,root.replace('Id="root"','Id="private-id"').replace('Target="word/document.xml"','Target="/word/document.xml"')+links+style)],
    ["document","relationship_id_mismatch",check(root+links.replace(`${R}/styles`,`${R}/image`)+style,root+links.replace(`${R}/styles`,`${R}/image`).replace('Id="format"','Id="private-id"')+style)],
    ["document","relationship_link_match",check(root+links+style,root+links.replace('Target="styles.xml"','Target="private-value.xml"')+style)],
    ["dependency","relationship_link_match",check(root+links+style+part("/word/theme/_rels/theme1.xml.rels",rels(rel("theme","image","../media/image1.png"))),root+links+style+part("/word/theme/_rels/theme1.xml.rels",rels(rel("theme","image","../media/private-value.png"))))],
  ] as const
  for(const [family,reason,run] of cases){
    let error:WordChatTableImportError|undefined
    try{run()}catch(cause){error=cause as WordChatTableImportError}
    expect(error?.code).toBe("formatting_dependencies")
    expect(error?.detail).toEqual({category:"formatting_dependencies",family,reason})
    expect(error?.message).toBe(new WordChatTableImportError("formatting_dependencies").message)
    expect(error?.diagnosticMessage).toContain(reason)
    expect(error?.diagnosticMessage).not.toContain("private-")
    expect(error?.diagnosticMessage).not.toContain(".xml")
    expect(error?.detail).not.toHaveProperty("path")
  }
})

test("only unique unreferenced implicit document-formatting links may renew their IDs", () => {
  const formats=[
    ["styles","styles.xml",`<w:styles xmlns:w="${W}"/>`],
    ["theme","theme/theme1.xml",'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'],
    ["fontTable","fontTable.xml",`<w:fonts xmlns:w="${W}"/>`],
    ["numbering","numbering.xml",`<w:numbering xmlns:w="${W}"/>`],
  ] as const
  for(const [kind,target,xml] of formats){
    const dependencies=(id:string)=>root+part("/word/_rels/document.xml.rels",rels(rel(id,kind,target)))+part(`/word/${target}`,xml)
    assertWordChatTableImportPreimage(pack(table,dependencies("old")),pack(p("Outside")+table+p("After"),dependencies("new")),inventory,proposal)
  }
  const dependencies=(linkXml=rel("format","styles","styles.xml"),format=style)=>root+part("/word/_rels/document.xml.rels",rels(linkXml))+format
  const source=pack(table,dependencies()),freshLink=rel("renewed","styles","styles.xml")
  const current=(linkXml=freshLink,format=style,content=table)=>pack(p("Outside")+content+p("After"),dependencies(linkXml,format))
  for(const changed of [
    current(rel("renewed","styles","other.xml")),
    current(rel("renewed","theme","styles.xml")),
    current(freshLink.replace('/>',' TargetMode="External"/>')),
    current(freshLink.replace('/>',' TargetMode="Internal"/>')),
    current(freshLink.replace('/>',' extra="changed"/>')),
    current(freshLink.replace('/>','><extra/></Relationship>')),
    current(""),
    current(freshLink, ""),
    current(freshLink+rel("ambiguous","styles","styles.xml")),
    current(freshLink,style.replace('w:val="22"','w:val="24"')),
    current(freshLink,style,table.replace('w:w="1800"','w:w="2400"')),
  ]) expect(()=>assertWordChatTableImportPreimage(source,changed,inventory,proposal)).toThrow(WordChatTableImportError)
  expect(()=>assertWordChatTableImportPreimage(pack(table,dependencies(rel("format","styles","styles.xml")+rel("ambiguous","styles","styles.xml"))),current(),inventory,proposal)).toThrow(WordChatTableImportError)
  for(const id of ["format","renewed"]){
    for(const reference of [`xmlns:r="${R}" r:embed="${id}"`,`xmlns:r="${R}" r:unknown="${id}"`,`xmlns:o="urn:schemas-microsoft-com:office:office" o:relid="${id}"`]){
      const referenced=table.replace("<w:p>",`<w:p ${reference}>`)
      expect(()=>assertWordChatTableImportPreimage(pack(referenced,dependencies()),current(freshLink,style,referenced),inventory,proposal)).toThrow(WordChatTableImportError)
      const ownerReference=(xml:string)=>xml.replace("<w:document ",`<w:document ${reference} `)
      expect(()=>assertWordChatTableImportPreimage(ownerReference(source),ownerReference(current()),inventory,proposal)).toThrow(WordChatTableImportError)
    }
  }
  const dependency=(id:string)=>part("/word/theme/_rels/theme1.xml.rels",rels(rel(id,"styles","../styles.xml")))
  expect(()=>assertWordChatTableImportPreimage(pack(table,root+links+style+dependency("old")),pack(p("Outside")+table+p("After"),root+links+style+dependency("new")),inventory,proposal)).toThrow(WordChatTableImportError)
})

describe("table import artifact verification", () => {
  const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
  const MAIN = "word/document.xml", MAIN_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
  const xml = (node: Node) => new XMLSerializer().serializeToString(node)
  const parse = (value: string) => new DOMParser().parseFromString(value, "application/xml")
  const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
  const artifactRow = (texts: string[]) => `<w:tr>${texts.map(text => `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`).join("")}</w:tr>`
  let cached: ReturnType<typeof makeFixture> | undefined
  async function makeFixture() {
    // Real server authoring is fixture setup, not part of the browser module's
    // compilation graph. Keep server-only ES2024 libraries out of Word's build.
    const writerPath = "../../../dochaus/lib/docx-table-row-deletion", builderPath = "../../../dochaus/lib/docx-table-row-deletion-import"
    type Request = { source_sha256: string; target_handle: string; author: string }
    const { authorDocxTableRowDeletion, docxTableRowDeletionTargets } = await import(writerPath) as {
      authorDocxTableRowDeletion: (source: Uint8Array, input: Request) => Promise<{ bytes: Uint8Array }>
      docxTableRowDeletionTargets: (source: Uint8Array) => Promise<{ source_sha256: string; targets: Array<{ row: number; target_handle: string }> }>
    }
    const { prepareDocxTableRowDeletionImport } = await import(builderPath) as {
      prepareDocxTableRowDeletionImport: (source: Uint8Array, authored: Uint8Array, input: Request) => Promise<{ beforeOoxml: string; ooxml: string }>
    }
    const fontType = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
    const source = zipSync({
      "[Content_Types].xml": strToU8(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="odttf" ContentType="${fontType}"/><Override PartName="/${MAIN}" ContentType="${MAIN_TYPE}"/>${["styles", "settings", "fontTable", "numbering"].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml"/>`).join("")}</Types>`),
      "_rels/.rels": strToU8(rels(rel("original-root", "officeDocument", MAIN))),
      "word/_rels/document.xml.rels": strToU8(rels(rel("s", "styles", "styles.xml") + rel("f", "fontTable", "fontTable.xml") + rel("n", "numbering", "numbering.xml") + rel("settings", "settings", "settings.xml"))),
      "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="23"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
      "word/fontTable.xml": strToU8(`<w:fonts xmlns:w="${W}" xmlns:r="${R}"><w:font w:name="Example"><w:embedRegular r:id="font" w:fontKey="{11223344-5566-7788-99AA-BBCCDDEEFF00}"/></w:font></w:fonts>`),
      "word/_rels/fontTable.xml.rels": strToU8(rels(rel("font", "font", "fonts/example.odttf"))),
      "word/fonts/example.odttf": new Uint8Array([7, 8, 9]),
      "word/numbering.xml": strToU8(`<w:numbering xmlns:w="${W}" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml"><w:numPicBullet w:numPicBulletId="1"><w:pict><v:shape><v:imagedata r:id="picture"/></v:shape></w:pict></w:numPicBullet></w:numbering>`),
      "word/_rels/numbering.xml.rels": strToU8(rels(rel("picture", "image", "media/picture.png"))),
      "word/media/picture.png": new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=", "base64")),
      "word/settings.xml": strToU8(`<w:settings xmlns:w="${W}"/>`),
      [MAIN]: strToU8(`<w:document xmlns:w="${W}"><w:body>${p("Outside opening")}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="1700"/></w:tblGrid>${artifactRow(["Kind", "Meaning"])}${artifactRow(["Remove", "Entry"])}</w:tbl>${p("Outside ending")}<w:sectPr/></w:body></w:document>`),
    })
    const packet = await docxTableRowDeletionTargets(source), target = packet.targets.find(t => t.row === 2)!
    const request = { source_sha256: packet.source_sha256, target_handle: target.target_handle, author: "doc.haus" }
    const authored = await authorDocxTableRowDeletion(source, request), pair = await prepareDocxTableRowDeletionImport(source, authored.bytes, request)
    const paragraphs = ["Kind", "Meaning", "Remove", "Entry"].map((text, i) => ({ id: `cell-${i}`, text }))
    const artifactProposal = { kind: "table_row_delete" as const, id: "reviewed", documentId: "document", explanation: "Remove row", openItems: [], createdAt: "now",
      table: { tableIndex: 0, rows: [paragraphs.slice(0, 2), paragraphs.slice(2)] }, rowIndex: 1, cells: ["Remove", "Entry"] }
    return { source: unzipSync(source), authored: unzipSync(authored.bytes), pair, proposal: artifactProposal }
  }
  const fixture = () => cached ??= makeFixture()
  function mutate(ooxml: string, name: string, change: (root: Element) => void) {
    const doc = parse(ooxml), targetPart = kids(doc.documentElement).find(node => node.getAttributeNS(P, "name") === `/${name}`)!
    change(kids(kids(targetPart)[0]!)[0]!); return xml(doc)
  }
  function drop(ooxml: string, name: string) {
    const doc = parse(ooxml), targetPart = kids(doc.documentElement).find(node => node.getAttributeNS(P, "name") === `/${name}`)!
    targetPart.parentNode!.removeChild(targetPart); return xml(doc)
  }

  test("binds a real paired import to the verified source and authored complete table and implicit dependencies", async () => {
    const f = await fixture()
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, f.pair, f.proposal)).not.toThrow()
    const pair = { ...f.pair, ooxml: mutate(f.pair.ooxml, "_rels/.rels", targetRoot => kids(targetRoot)[0]!.setAttribute("Id", "another-root")) }
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, pair, f.proposal)).not.toThrow()
  })

  test("forged after tables cannot change header wording, geometry or deletion carriers", async () => {
    const f = await fixture()
    for (const change of [
      (targetRoot: Element) => targetRoot.getElementsByTagNameNS(W, "t")[0]!.appendChild(targetRoot.ownerDocument!.createTextNode(" forged")),
      (targetRoot: Element) => targetRoot.getElementsByTagNameNS(W, "gridCol")[0]!.setAttributeNS(W, "w:w", "1900"),
      (targetRoot: Element) => targetRoot.getElementsByTagNameNS(W, "del")[0]!.setAttributeNS(W, "w:author", "Someone else"),
    ]) expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, ooxml: mutate(f.pair.ooxml, MAIN, change) }, f.proposal)).toThrow("part_xml_changed")
  })

  test("source preimages and document or body shells cannot be forged even when table text matches", async () => {
    const f = await fixture()
    const changedHeader = mutate(f.pair.beforeOoxml, MAIN, targetRoot => targetRoot.getElementsByTagNameNS(W, "t")[0]!.appendChild(targetRoot.ownerDocument!.createTextNode(" altered")))
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, beforeOoxml: changedHeader }, f.proposal)).toThrow()
    for (const field of ["beforeOoxml", "ooxml"] as const) {
      const body = mutate(f.pair[field], MAIN, targetRoot => targetRoot.getElementsByTagNameNS(W, "body")[0]!.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve"))
      const shell = mutate(f.pair[field], MAIN, targetRoot => targetRoot.setAttributeNS(W, "w:conformance", "strict"))
      for (const value of [body, shell]) expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: value }, f.proposal)).toThrow()
    }
  })

  test("section properties, adjacent paragraphs and arbitrary parts never enter either payload", async () => {
    const f = await fixture()
    for (const field of ["beforeOoxml", "ooxml"] as const) {
      for (const name of ["p", "sectPr"]) {
        const leaked = mutate(f.pair[field], MAIN, targetRoot => targetRoot.getElementsByTagNameNS(W, "body")[0]!.appendChild(targetRoot.ownerDocument!.createElementNS(W, `w:${name}`)))
        expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: leaked }, f.proposal)).toThrow()
      }
      const doc = parse(f.pair[field]), extra = doc.createElementNS(P, "pkg:part")
      extra.setAttributeNS(P, "pkg:name", "/word/comments.xml"); extra.setAttributeNS(P, "pkg:contentType", "application/xml")
      const data = doc.createElementNS(P, "pkg:xmlData"); data.appendChild(doc.createElementNS(W, "w:comments")); extra.appendChild(data); doc.documentElement.appendChild(extra)
      expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: xml(doc) }, f.proposal)).toThrow("part_set_changed")
    }
  })

  test("omitting implicit defaults or transitive font and image dependencies cannot forge a smaller allowed closure", async () => {
    const f = await fixture()
    for (const name of ["word/styles.xml", "word/fonts/example.odttf", "word/media/picture.png", "word/_rels/numbering.xml.rels"]) {
      let pair = { beforeOoxml: drop(f.pair.beforeOoxml, name), ooxml: drop(f.pair.ooxml, name) }
      if (name === "word/styles.xml") {
        const removeLink = (targetRoot: Element) => { const link = kids(targetRoot).find(n => n.getAttribute("Id") === "s")!; targetRoot.removeChild(link) }
        pair = { beforeOoxml: mutate(pair.beforeOoxml, "word/_rels/document.xml.rels", removeLink), ooxml: mutate(pair.ooxml, "word/_rels/document.xml.rels", removeLink) }
      }
      expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, pair, f.proposal)).toThrow("part_set_changed")
    }
  })

  test("changed formatting definitions, binary data, relationship modes and content types refuse", async () => {
    const f = await fixture()
    for (const field of ["beforeOoxml", "ooxml"] as const) {
      const styleEdit = mutate(f.pair[field], "word/styles.xml", targetRoot => targetRoot.getElementsByTagNameNS(W, "sz")[0]!.setAttributeNS(W, "w:val", "31"))
      const external = mutate(f.pair[field], "word/_rels/numbering.xml.rels", targetRoot => kids(targetRoot)[0]!.setAttribute("TargetMode", "External"))
      const missing = mutate(f.pair[field], "word/_rels/numbering.xml.rels", targetRoot => kids(targetRoot)[0]!.setAttribute("Target", "media/missing.png"))
      for (const value of [styleEdit, external, missing]) expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: value }, f.proposal)).toThrow()
      const doc = parse(f.pair[field]), font = kids(doc.documentElement).find(node => node.getAttributeNS(P, "name") === "/word/fonts/example.odttf")!
      kids(font)[0]!.textContent = "AAAA"
      expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: xml(doc) }, f.proposal)).toThrow("part_bytes_changed")
      font.setAttributeNS(P, "pkg:contentType", "application/xml")
      expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { ...f.pair, [field]: xml(doc) }, f.proposal)).toThrow()
    }
  })

  test("unresolvable source dependencies and oversized pairs are refused", async () => {
    const f = await fixture(), bad = { ...f.source }
    bad["word/_rels/numbering.xml.rels"] = strToU8(rels(rel("picture", "image", "media/missing.png")))
    expect(() => verifyWordChatTableImportArtifact(bad, f.authored, f.pair, f.proposal)).toThrow("dependency_missing")
    const external = parse(rels(rel("picture", "image", "media/picture.png"))); kids(external.documentElement)[0]!.setAttribute("TargetMode", "External")
    bad["word/_rels/numbering.xml.rels"] = strToU8(xml(external))
    expect(() => verifyWordChatTableImportArtifact(bad, f.authored, f.pair, f.proposal)).toThrow("external_dependency")
    expect(() => verifyWordChatTableImportArtifact(f.source, f.authored, { beforeOoxml: "x".repeat(1024 * 1024 + 1), ooxml: "x".repeat(1024 * 1024) }, f.proposal)).toThrow("pair_large")
  })
})
