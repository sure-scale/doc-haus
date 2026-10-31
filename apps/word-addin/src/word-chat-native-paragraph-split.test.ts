import { expect, test } from "bun:test"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { assertNativeWordChatParagraphSplitTarget, verifyNativeWordChatParagraphSplit } from "./word-chat-native-paragraph-split"
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main", W14="http://schemas.microsoft.com/office/word/2010/wordml", PKG="http://schemas.microsoft.com/office/2006/xmlPackage"
const flat=(body:string,comment="Keep",style="same")=>`<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="3" w:author="Earlier"><w:p><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles w:val="${style}"/></pkg:xmlData></pkg:part></pkg:package>`
const source='<w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:t>One two</w:t></w:r></w:p>'
const mark=(id:string,author="Reviewer")=>`w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"`
const left=`<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:rPr><w:ins ${mark("2")}/></w:rPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>`
const right=`<w:p w14:paraId="11111111" w14:textId="12345678"><w:del ${mark("3")}><w:r w:rsidDel="00112233"><w:delText xml:space="preserve"> </w:delText></w:r></w:del><w:r><w:t>two</w:t></w:r></w:p>`
const prior=(id:string)=>`<w:p><w:bookmarkStart w:id="1" w:name="Keep"/><w:commentRangeStart w:id="3"/><w:ins w:id="${id}" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins><w:commentRangeEnd w:id="3"/><w:bookmarkEnd w:id="1"/><w:r><w:commentReference w:id="3"/></w:r></w:p>`
const inv=[{id:"a",text:"One two",tableNestingLevel:0},{id:"b",text:"Prior",tableNestingLevel:0}],afterInv=[{id:"new",text:"One",tableNestingLevel:0},{id:"a",text:"two",tableNestingLevel:0},inv[1]!]
const target={id:"target",text:"One two",paragraphs:[inv[0]!],startOffset:0,endOffset:7,capturedAt:"now",saved:false}
const before=projectNativeWordChatBody(flat(source+prior("2")),inv)
const after=(body:string,comment="Keep",style="same", inventory=afterInv)=>projectNativeWordChatBody(flat(body,comment,style),inventory)
const good=left+right+prior("4")
test("split proves carriers, full inverse and old revision mapping",()=>{
 expect(assertNativeWordChatParagraphSplitTarget(before,target)).toBe(0)
 expect(verifyNativeWordChatParagraphSplit(before,after(good),target,"One\rtwo")).toEqual(["Reviewer"])
 for(const bad of [good.replace(mark("3"),mark("02")),good.replace(mark("3"),mark("3","Other")),good.replace('> </w:delText>','>x</w:delText>'),good.replace('<w:r><w:t>two','<w:r><w:rPr><w:b/></w:rPr><w:t>two'),good.replace('w:name="Keep"','w:name="Changed"'),good.replace('w:id="4" w:author="Earlier"','w:id="3" w:author="Earlier"')])expect(()=>verifyNativeWordChatParagraphSplit(before,after(bad),target,"One\rtwo")).toThrow()
 expect(()=>verifyNativeWordChatParagraphSplit(before,after(good,"Changed"),target,"One\rtwo")).toThrow()
 expect(()=>verifyNativeWordChatParagraphSplit(before,after(good,"Keep","changed"),target,"One\rtwo")).toThrow()
 expect(()=>verifyNativeWordChatParagraphSplit(before,after(good,"Keep","same",[{...afterInv[0]!,id:"a"},{...afterInv[1]!,id:"new"},afterInv[2]!]),target,"One\rtwo")).toThrow()
 expect(()=>verifyNativeWordChatParagraphSplit(before,after(good),target,"One two")).toThrow()
})
test("target accepts plain multiple runs and refuses partial selection and crossing annotation",()=>{
 expect(()=>assertNativeWordChatParagraphSplitTarget(before,{...target,endOffset:4})).toThrow()
 const multi=source.replace('<w:t>One two</w:t>','<w:t>One</w:t></w:r><w:r><w:t xml:space="preserve"> two</w:t>')
 expect(assertNativeWordChatParagraphSplitTarget(projectNativeWordChatBody(flat(multi+prior('2')),inv),target)).toBe(0)
 const cross='<w:p><w:bookmarkStart w:id="8" w:name="x"/><w:r><w:t>Before</w:t></w:r></w:p>'+source+'<w:p><w:bookmarkEnd w:id="8"/><w:r><w:t>After</w:t></w:r></w:p>'
 expect(()=>assertNativeWordChatParagraphSplitTarget(projectNativeWordChatBody(flat(cross),[{id:'x',text:'Before',tableNestingLevel:0},inv[0]!,{id:'y',text:'After',tableNestingLevel:0}]),target)).toThrow()
})

test("mixed runs preserve the boundary mark and deleted space in the same cell",()=>{
 const meta=(id:string)=>mark(id)
 const cell=(content:string, width="100")=>`<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="${width}"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`
 const src='<w:p w14:paraId="11111111" w14:textId="77777777"><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>One</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> two</w:t></w:r></w:p>'
 const l=`<w:p w14:paraId="22222222" w14:textId="77777777"><w:pPr><w:spacing w:after="120"/><w:rPr><w:ins ${meta("2")}/><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>One</w:t></w:r></w:p>`
 const r=`<w:p w14:paraId="11111111" w14:textId="12345678"><w:pPr><w:spacing w:after="120"/></w:pPr><w:del ${meta("3")}><w:r w:rsidDel="00112233"><w:rPr><w:i/></w:rPr><w:delText xml:space="preserve"> </w:delText></w:r></w:del><w:r><w:rPr><w:i/></w:rPr><w:t>two</w:t></w:r></w:p>`
 const original=[{id:"a",text:"One two",tableNestingLevel:1},{id:"b",text:"Prior",tableNestingLevel:0}], changed=[{id:"new",text:"One",tableNestingLevel:1},{id:"a",text:"two",tableNestingLevel:1},original[1]!]
 const b=projectNativeWordChatBody(flat(cell(src)+prior("4")),original), capture=(xml:string)=>projectNativeWordChatBody(flat(xml+prior("4")),changed)
 expect(verifyNativeWordChatParagraphSplit(b,capture(cell(l+r)),target,"One\rtwo")).toEqual(["Reviewer"])
 const bad=[cell(l.replace('<w:ins '+meta("2")+'/><w:i/>','<w:ins '+meta("2")+'/><w:b/>')+r),cell(l+r.replace('<w:rPr><w:i/></w:rPr><w:delText','<w:rPr><w:b/></w:rPr><w:delText')),cell(l+r.replace('w:after="120"','w:after="121"')),cell(l+r,"200")]
 for(const xml of bad)expect(()=>verifyNativeWordChatParagraphSplit(b,capture(xml),target,"One\rtwo")).toThrow()
})

test("qualified run session IDs may change without masking formatting or foreign attributes",()=>{
 const sourceMeta=source.replace('<w:r>', '<w:r w:rsidR="1234ABCD" w:rsidRPr="00112233">')
 const base=projectNativeWordChatBody(flat(sourceMeta+prior("2")),inv)
 const leftMeta=left.replace('<w:r><w:t>One','<w:r w:rsidR="ABCDEF01" w:rsidRPr="AABBCCDD"><w:t>One')
 const rightMeta=right.replace('<w:r w:rsidDel="00112233">','<w:r w:rsidDel="00112233" w:rsidR="ABCDEF01">').replace('<w:r><w:t>two','<w:r w:rsidRPr="00112233"><w:t>two')
 const candidate=leftMeta+rightMeta+prior("4")
 expect(verifyNativeWordChatParagraphSplit(base,after(candidate),target,"One\rtwo")).toEqual(["Reviewer"])
 for(const xml of [candidate.replace('w:rsidR="ABCDEF01"','w:rsidR="invalid"'),candidate.replace('w:rsidR="ABCDEF01"','foreign="ABCDEF01"'),candidate.replace('<w:t>two','<w:rPr><w:b/></w:rPr><w:t>two')])expect(()=>verifyNativeWordChatParagraphSplit(base,after(xml),target,"One\rtwo")).toThrow()
 expect(()=>assertNativeWordChatParagraphSplitTarget(projectNativeWordChatBody(flat(sourceMeta.replace('1234ABCD','bad')+prior("2")),inv),target)).toThrow()
})
