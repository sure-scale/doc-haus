import { expect, test } from "bun:test"
import { verifyNativeWordChatInsertedRow } from "./word-chat-native-row-verification"
const PKG="http://schemas.microsoft.com/office/2006/xmlPackage", W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const wrap=(body:string, comment="Text", rsid="00112233")=>`<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="1" w:author="Reviewer"><w:p w:rsidR="${rsid}" w:rsidRDefault="${rsid}"><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/settings.xml"><pkg:xmlData><w:settings><w:rsids><w:rsid w:val="${rsid}"/></w:rsids></w:settings></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles><w:style w:styleId="Normal"/></w:styles></pkg:xmlData></pkg:part></pkg:package>`
const stamp='w:author="Reviewer" w:date="2026-01-01T00:00:00Z"'
const cell=(text:string, mark:number, content:number, change=false)=>`<w:tc><w:tcPr/><w:p><w:pPr><w:rPr><w:ins w:id="${mark}" ${stamp}/><w:b/></w:rPr></w:pPr><w:ins w:id="${content}" ${stamp}><w:r>${change?`<w:rPr><w:rPrChange w:id="9" ${stamp}><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>`:""}<w:t>${text}</w:t></w:r></w:ins></w:p></w:tc>`
const old='<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>Old two</w:t></w:r></w:p></w:tc></w:tr>'
const row=`<w:tr><w:trPr><w:ins w:id="1" ${stamp}/></w:trPr>${cell("New",2,3,true)}${cell("Value",4,5)}</w:tr>`
const body=(extra:string)=>`<w:tbl><w:tblPr/><w:tblGrid/>${old}${extra}</w:tbl>`
const before=wrap(body("")), after=wrap(body(row),"Text","AABBCCDD")
const verify=(xml:string)=>verifyNativeWordChatInsertedRow(before,xml,0,1,["New","Value"])
test("native inserted row verifies plain text, tracked carriers and narrow volatile dependencies",()=>{
  expect(verify(after)).toEqual({authors:["Reviewer"]})
  for(const changed of [after.replace('w:id="1" '+stamp,'w:id="1" w:author="Other" w:date="2026-01-01T00:00:00Z"'),after.replace('>New<','>Wrong<'),after.replace('<w:ins w:id="1"','<w:del w:id="1"').replace('/></w:trPr>','/></w:trPr>'),after.replace('<w:rPr><w:rPrChange','<w:rPr><w:b/><w:rPrChange'),after.replace('w:id="9"','w:id="3"'),after.replace('>Text<','>Changed<'),after.replace('w:styleId="Normal"','w:styleId="Other"'),after.replace('AABBCCDD','arbitrary'),after.replace('<w:tcPr/><w:p><w:pPr>','<w:tcPr/><w:drawing/><w:p><w:pPr>'),after.replace('<w:tcPr/><w:p><w:pPr>','<w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr><w:p><w:pPr>'),after.replace('<w:pPr><w:rPr>','<w:pPr w:foo="1"><w:rPr>')]) expect(()=>verify(changed)).toThrow()
})

const space=(text:string)=>/^\s|\s$/.test(text)?' xml:space="preserve"':""
const fcell=(runs:Array<[string,string]>, mark:number, content:number)=>`<w:tc><w:tcPr/><w:p><w:pPr><w:rPr><w:ins w:id="${mark}" ${stamp}/></w:rPr></w:pPr><w:ins w:id="${content}" ${stamp}>${runs.map(([props,text])=>`<w:r>${props?`<w:rPr>${props}</w:rPr>`:""}<w:t${space(text)}>${text}</w:t></w:r>`).join("")}</w:ins></w:p></w:tc>`
const frow=(a:Array<[string,string]>, b:Array<[string,string]>)=>wrap(body(`<w:tr><w:trPr><w:ins w:id="1" ${stamp}/></w:trPr>${fcell(a,2,3)}${fcell(b,4,5)}</w:tr>`),"Text","AABBCCDD")
const span=(text:string, bold=false, italic=false, underline=false)=>({text,bold,italic,underline})
const cells=["Fee: 100","Note"], planned=[[span("Fee:",true),span(" 100")],[span("Note",false,true,true)]]
const formatted=(xml:string, runs:typeof planned|undefined=planned)=>verifyNativeWordChatInsertedRow(before,xml,0,1,cells,runs)
const BI='<w:b/><w:bCs/>', IU='<w:i/><w:iCs/><w:u w:val="single"/>'
test("formatted row admits on values only on the planned spans, across Word run splits",()=>{
  expect(formatted(frow([[BI,"Fee:"],[""," 100"]],[[IU,"Note"]]))).toEqual({authors:["Reviewer"]})
  // Word may split a span, drop the complex-script twin, spell on/off explicitly or record a formatting change.
  expect(formatted(frow([['<w:b/>',"Fe"],['<w:b w:val="1"/><w:bCs w:val="true"/>',"e:"],['<w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="0"/><w:u w:val="none"/>'," 1"],["","00"]],
    [['<w:i/><w:u w:val="single"/>',"No"],[`<w:i w:val="on"/><w:iCs/><w:u w:val="single"/><w:rPrChange w:id="9" ${stamp}><w:rPr/></w:rPrChange>`,"te"]]))).toEqual({authors:["Reviewer"]})
  for(const [a,b] of [
    [[[BI,"Fee: "],["","100"]],[[IU,"Note"]]],                       // on at the wrong extent
    [[["","Fee:"],[""," 100"]],[[IU,"Note"]]],                        // missing on
    [[[BI,"Fee:"],['<w:i/>'," 100"]],[[IU,"Note"]]],                  // on inside an off span
    [[[BI+'<w:sz w:val="24"/>',"Fee:"],[""," 100"]],[[IU,"Note"]]],   // extra run property
    [[['<w:rStyle w:val="Strong"/>',"Fee:"],[""," 100"]],[[IU,"Note"]]], // style instead of direct flag
    [[['<w:b/><w:bCs w:val="0"/>',"Fee:"],[""," 100"]],[[IU,"Note"]]], // complex-script twin disagrees
    [[['<w:b w:val="maybe"/>',"Fee:"],[""," 100"]],[[IU,"Note"]]],    // unknown toggle value
    [[[BI,"Fee:"],[""," 100"]],[['<w:i/><w:u w:val="double"/>',"Note"]]], // other underline
    [[[BI,"Fee:"],[""," 100"]],[['<w:i/><w:u/>',"Note"]]],            // underline without a value
    [[['<w:b w:color="FF0000"/>',"Fee:"],[""," 100"]],[[IU,"Note"]]], // extra attribute
    [[[BI,"Fee:"],[""," 10O"]],[[IU,"Note"]]],                        // text drift
  ] as Array<[Array<[string,string]>,Array<[string,string]>]>) expect(()=>formatted(frow(a,b))).toThrow()
  // The physical row must match the reviewed plan, not merely be well-formed.
  const exact=frow([[BI,"Fee:"],[""," 100"]],[[IU,"Note"]])
  expect(()=>verifyNativeWordChatInsertedRow(before,exact,0,1,cells)).toThrow()
  expect(()=>formatted(exact,[[span("Fee:",true),span(" 100")],[span("Note",false,true)]])).toThrow()
  expect(()=>formatted(exact,[[span("Fee:",true),span(" 100")]])).toThrow()
})
test("plain row keeps one off span per cell, with explicit off values and split runs",()=>{
  const plain=(xml:string)=>verifyNativeWordChatInsertedRow(before,xml,0,1,["New","Value"])
  expect(plain(frow([['<w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="false"/><w:u w:val="none"/>',"Ne"],["","w"]],[["","Value"]]))).toEqual({authors:["Reviewer"]})
  for(const props of ['<w:b/>','<w:u w:val="single"/>','<w:i w:val="1"/>','<w:sz w:val="24"/>']) expect(()=>plain(frow([[props,"New"]],[["","Value"]]))).toThrow()
})
