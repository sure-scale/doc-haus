import { expect, test } from "bun:test"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { nativeWordChatDecimalList } from "./word-chat-native-list-context"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const p = (text: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const body = p("One") + p("Two")
const numbering = '<w:numbering><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="7"/></w:num></w:numbering>'
const flat = (b=body,n=numbering) => `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${b}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData>${n}</pkg:xmlData></pkg:part></pkg:package>`
const inv = [{id:"a",text:"One",tableNestingLevel:0},{id:"b",text:"Two",tableNestingLevel:0}]
const capture = {id: 17, format:"decimal" as const, members:inv.map((p,i)=>({id:p.id,text:p.text,level:0,label:`${i+1}.`}))}
test("native decimal binding corroborates real definitions and rejects typed or partial membership",()=>{
  const check=(b=body,n=numbering,c=capture)=>nativeWordChatDecimalList(projectNativeWordChatBody(flat(b,n),inv),c)
  expect(check()).toBe(true)
  expect(check(body.replaceAll('<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>',""))).toBe(false)
  expect(check(body.replace('<w:numId w:val="5"/>','<w:numId w:val="6"/>'))).toBe(false)
  expect(check(body,numbering.replace('w:val="decimal"','w:val="lowerLetter"'))).toBe(false)
  expect(check(body,numbering.replace('<w:abstractNumId w:val="7"/>','<w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>'))).toBe(false)
  expect(check(body,numbering.replace('w:val="%1."','w:val="Article %1"'))).toBe(false)
  expect(check(body,numbering.replace('</w:abstractNum>', '<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl></w:abstractNum>'))).toBe(true)
  expect(check(body,numbering.replace('</w:abstractNum>', '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'))).toBe(false)
  expect(check(body,numbering,{...capture,members:[capture.members[0]!]})).toBe(false)
})
test("pending inserted mark and text may remain while list definition binding stays explicit",()=>{
  const after = body.replace('</w:numPr></w:pPr>', '</w:numPr><w:rPr><w:ins w:id="1" w:author="Reviewer"/></w:rPr></w:pPr>') + p("Three").replace('<w:r><w:t>Three</w:t></w:r>','<w:ins w:id="2" w:author="Reviewer"><w:r><w:t>Three</w:t></w:r></w:ins>')
  const inventory=[...inv,{id:"c",text:"Three",tableNestingLevel:0}]
  const c={...capture,members:inventory.map((p,i)=>({id:p.id,text:p.text,level:0,label:`${i+1}.`}))}
  expect(nativeWordChatDecimalList(projectNativeWordChatBody(flat(after),inventory),c)).toBe(true)
})
