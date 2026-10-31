import { expect, test } from "bun:test"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { assertNativeWordChatParagraphMergeTarget, verifyNativeWordChatParagraphMerge } from "./word-chat-native-paragraph-merge"
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main", W14="http://schemas.microsoft.com/office/word/2010/wordml", PKG="http://schemas.microsoft.com/office/2006/xmlPackage"
const flat=(body:string,comment="Keep",style="same")=>`<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}" xmlns:w14="${W14}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="3" w:author="Earlier"><w:p><w:r><w:t>${comment}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles w:val="${style}"/></pkg:xmlData></pkg:part></pkg:package>`
const left='<w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:rPr><w:b/></w:rPr><w:t>Left</w:t></w:r></w:p>',right='<w:p w14:paraId="22222222" w14:textId="77777777"><w:r><w:t>Right</w:t></w:r></w:p>'
const mark=(id:string,author="Reviewer")=>`w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"`
const changedLeft=`<w:p w14:paraId="11111111" w14:textId="12345678" w:rsidDel="00112233"><w:pPr><w:rPr><w:del ${mark("2")}/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Left</w:t></w:r></w:p>`
const changedRight=`<w:p w14:paraId="22222222" w14:textId="12345678"><w:ins ${mark("3")}><w:r><w:t xml:space="preserve"> </w:t></w:r></w:ins><w:r><w:t>Right</w:t></w:r></w:p>`
const prior=(id:string)=>`<w:p><w:bookmarkStart w:id="1" w:name="Keep"/><w:commentRangeStart w:id="3"/><w:ins w:id="${id}" w:author="Earlier" w:date="2025-01-01"><w:r><w:t>Prior</w:t></w:r></w:ins><w:commentRangeEnd w:id="3"/><w:bookmarkEnd w:id="1"/><w:r><w:commentReference w:id="3"/></w:r></w:p>`
const inventory=[{id:"a",text:"Left",tableNestingLevel:0},{id:"b",text:"Right",tableNestingLevel:0},{id:"c",text:"Prior",tableNestingLevel:0}], afterInventory=[inventory[0]!,{...inventory[1]!,text:" Right"},inventory[2]!]
const target={id:"target",text:"Left\rRight",paragraphs:inventory.slice(0,2),startOffset:0,endOffset:5,capturedAt:"now",saved:false}
const before=projectNativeWordChatBody(flat(left+right+prior("2")),inventory)
const after=(body:string,comment="Keep",style="same")=>projectNativeWordChatBody(flat(body,comment,style),afterInventory)
const good=changedLeft+changedRight+prior("4")
test("merge proves the whole inverse and qualified prior review ID mapping",()=>{
 expect(assertNativeWordChatParagraphMergeTarget(before,target)).toEqual([0,1]);expect(verifyNativeWordChatParagraphMerge(before,after(good),target)).toEqual(["Reviewer"])
 for(const bad of [good.replace(mark("3"),mark("02")),good.replace(mark("3"),mark("3","Other")),good.replace("<w:b/>","<w:i/>"),good.replace('w:author="Earlier"','w:author="Other"'),good.replace('w:name="Keep"','w:name="Changed"'),good.replace('w:id="4" w:author="Earlier"','w:id="3" w:author="Earlier"')])expect(()=>verifyNativeWordChatParagraphMerge(before,after(bad),target)).toThrow()
 expect(()=>verifyNativeWordChatParagraphMerge(before,after(good,"Changed"),target)).toThrow();expect(()=>verifyNativeWordChatParagraphMerge(before,after(good,"Keep","altered"),target)).toThrow()
 expect(()=>assertNativeWordChatParagraphMergeTarget(before,{...target,endOffset:4})).toThrow()
})
test("crossing annotations, incompatible paragraph properties and old target revisions refuse",()=>{
 const cases=[flat('<w:p><w:bookmarkStart w:id="8" w:name="cross"/><w:r><w:t>Prior</w:t></w:r></w:p>'+left+right+'<w:p><w:bookmarkEnd w:id="8"/><w:r><w:t>End</w:t></w:r></w:p>'),flat(left+right.replace('<w:r>','<w:pPr><w:spacing w:after="20"/></w:pPr><w:r>')),flat(left.replace('<w:r>','<w:ins w:id="9" w:author="X" w:date="2025-01-01"><w:r>').replace('</w:r></w:p>','</w:r></w:ins></w:p>')+right)]
 for(const xml of cases){try{const inv=xml.includes('cross')?[{id:"z",text:"Prior",tableNestingLevel:0},...inventory.slice(0,2),{id:"z2",text:"End",tableNestingLevel:0}]:inventory.slice(0,2);expect(()=>assertNativeWordChatParagraphMergeTarget(projectNativeWordChatBody(xml,inv),target)).toThrow()}catch(e){throw e}}
})

test("a new people part needs one unique internal relationship and preserves other relationships",()=>{
 const rn="http://schemas.openxmlformats.org/package/2006/relationships",pn="http://schemas.microsoft.com/office/word/2012/wordml",people="http://schemas.microsoft.com/office/2011/relationships/people",theme="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
 const part=(name:string,xml:string)=>`<pkg:part pkg:name="${name}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`
 const rel=(id:string,type:string,to:string,extra="")=>`<Relationship Id="${id}" Type="${type}" Target="${to}"${extra}/>`
 const rels=(x:string)=>part('/word/_rels/document.xml.rels',`<Relationships xmlns="${rn}">${x}</Relationships>`)
 const person=part('/word/people.xml',`<w15:people xmlns:w15="${pn}"><w15:person w15:author="Reviewer"><w15:presenceInfo w15:providerId="x" w15:userId="y"/></w15:person></w15:people>`)
 const withParts=(xml:string,x:string)=>xml.replace('</pkg:package>',x+'</pkg:package>')
 const b=projectNativeWordChatBody(withParts(flat(left+right+prior('2')),rels(rel('r1',theme,'theme.xml'))),inventory)
 const candidate=(link:string)=>projectNativeWordChatBody(withParts(flat(good),rels(rel('r2',theme,'theme.xml')+link)+person),afterInventory)
 const link=rel('r1',people,'people.xml');expect(verifyNativeWordChatParagraphMerge(b,candidate(link),target)).toEqual(['Reviewer'])
 for(const bad of [rel('r1',people,'other.xml'),rel('r1',people,'people.xml',' TargetMode="External"'),link+rel('r3',people,'people.xml'),rel('r2',people,'people.xml')])expect(()=>verifyNativeWordChatParagraphMerge(b,candidate(bad),target)).toThrow()
})
