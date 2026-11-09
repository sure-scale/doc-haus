import { expect, test } from "bun:test"
import { assertWordChatTableImportPreimage, assertWordChatTableImportRange, assertWordChatTableImportFormatting, WordChatTableImportError, type WordChatTableImportErrorCode } from "./word-chat-table-import"
import { DOMParser } from "@xmldom/xmldom"
import { firstTableImportXmlDifference, sameTableImportXml, sameLiveTableImportPreimageXml } from "./word-chat-table-import-xml"

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
