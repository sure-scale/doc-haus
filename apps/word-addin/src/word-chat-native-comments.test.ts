import {expect,test} from "bun:test"
import {normalizeNativeWordChatCommentIds as normalize} from "./word-chat-native-comments"
const PKG="http://schemas.microsoft.com/office/2006/xmlPackage",W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const input=(id:string,content="Existing comment",author="Reviewer",rsid="00112233",extra="")=>`<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t>Anchor</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>${extra}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/comments.xml"><pkg:xmlData><w:comments><w:comment w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:p w:rsidR="${rsid}" w:rsidRDefault="${rsid}"><w:r><w:t>${content}</w:t></w:r></w:p></w:comment></w:comments></pkg:xmlData></pkg:part></pkg:package>`
test("comment IDs map only with exact matching anchors and comment payload",()=>{
  const before=input('1'),after=input('2','Existing comment','Reviewer','AABBCCDD'),out=normalize(before,after)
  expect(out).toMatch(/<w:comment\b[^>]*w:id="1"/);expect(out).toContain('w:commentRangeStart w:id="1"');expect(out).toContain('w:rsidR="00112233"')
  for(const drift of [input('2','Changed'),input('2','Existing comment','Other'),after.replace('<w:commentRangeStart','<w:r><w:t>X</w:t></w:r><w:commentRangeStart'),after.replace('commentReference w:id="2"','commentReference w:id="3"'),after.replace('w:rsidR="AABBCCDD"','w:rsidR="invalid"'),after.replace('<w:comment w:id="2"','<w:comment w:id="2" w:initials="X"'),after.replace('</w:comments>','<w:comment w:id="2"/></w:comments>')])expect(()=>normalize(before,drift)).toThrow()
  const outside=after.replace('</pkg:package>',`<pkg:part pkg:name="/word/header1.xml"><pkg:xmlData><w:hdr><w:commentReference w:id="2"/></w:hdr></pkg:xmlData></pkg:part></pkg:package>`)
  expect(()=>normalize(before,outside)).toThrow()
})

test("comments without optional rsid attributes retain their payload and anchors",()=>{
 const without=(xml:string)=>xml.replaceAll(' w:rsidR="00112233"',"").replaceAll(' w:rsidRDefault="00112233"',"")
 const before=without(input("1")),after=without(input("2"))
 expect(normalize(before,after)).toContain('w:commentRangeStart w:id="1"')
 expect(()=>normalize(before,after.replace('Existing comment','Changed comment'))).toThrow()
})


test("bookmark remaps preserve name and actual character range across run splits",()=>{
 const base=input("1").replace('<w:commentRangeStart w:id="1"/>','<w:bookmarkStart w:id="0" w:name="Existing"/><w:commentRangeStart w:id="1"/>').replace('<w:commentRangeEnd w:id="1"/>','<w:commentRangeEnd w:id="1"/><w:bookmarkEnd w:id="0"/>')
 const revised=base.replaceAll('w:id="1"','w:id="2"').replaceAll('w:id="0"','w:id="3"').replace('<w:r><w:t>Anchor</w:t></w:r>','<w:r><w:t>Anc</w:t></w:r><w:r><w:t>hor</w:t></w:r>')
 const output=normalize(base,revised);expect(output).toContain('w:bookmarkStart w:name="Existing" w:id="0"');expect(output).toContain('w:bookmarkEnd w:id="0"')
 for(const bad of [revised.replace('w:name="Existing"','w:name="Other"'),revised.replace('<w:bookmarkEnd w:id="3"/>','<w:bookmarkEnd w:id="8"/>'),revised.replace('<w:bookmarkEnd w:id="3"/>','<w:bookmarkEnd w:id="3"/><w:bookmarkEnd w:id="3"/>'),revised.replace('<w:bookmarkEnd w:id="3"/>','<w:r><w:t>X</w:t></w:r><w:bookmarkEnd w:id="3"/>')])expect(()=>normalize(base,bad)).toThrow()
})
