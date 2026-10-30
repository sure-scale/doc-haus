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
