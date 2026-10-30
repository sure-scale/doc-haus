import { expect, test } from "bun:test"
import { DOMParser } from "@xmldom/xmldom"
import { sameWordChatInlineFormattingContent, verifyWordChatInlineFormatting as verify } from "./word-chat-inline-formatting-ooxml"
import type { WordChatInlineFormat } from "./word-chat-amendment-contract"

const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const parse=(body:string)=>new DOMParser().parseFromString(`<w:p xmlns:w="${W}" w:rsidR="ABC"><w:pPr><w:pStyle w:val="Normal"/></w:pPr>${body}</w:p>`,"application/xml").documentElement!
const run=(text:string,properties="")=>`<w:r><w:rPr>${properties}</w:rPr><w:t>${text}</w:t></w:r>`
const rev=(old:string,id="7",author="Reviewer")=>`<w:rPrChange w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:rPr>${old}</w:rPr></w:rPrChange>`
const verifyWordChatInlineFormatting=(before:Element,after:Element,start:number,end:number,op:WordChatInlineFormat,author?:string) => verify(before,after,{startOffset:start,endOffset:end,text:Array.from(before.getElementsByTagNameNS(W,"t")).map(t=>t.textContent).join("").slice(start,end)},op,author)

test("resolution comparison tolerates splits but preserves all physical formatting and markers",()=>{
  const props='<w:i/><w:color w:val="112233"/>', bookmark='<w:bookmarkStart w:id="4" w:name="Keep"/>'
  const original=parse(bookmark+run("Alpha",props))
  expect(sameWordChatInlineFormattingContent(original,parse(bookmark+run("Al",props)+run("pha",props)))).toBe(true)
  expect(sameWordChatInlineFormattingContent(original,parse(bookmark+run("Alpha",'<w:i/><w:color w:val="FFFFFF"/>')))).toBe(false)
  expect(sameWordChatInlineFormattingContent(original,parse(bookmark.replace('Keep','Other')+run("Alpha",props)))).toBe(false)
  expect(sameWordChatInlineFormattingContent(parse('<w:r><w:t>Text</w:t></w:r>'),parse(run("Text")))).toBe(true)
})

test("tracked inline formatting verifies split runs, exact old properties and Unicode offsets",()=>{
  const props='<w:i/><w:color w:val="112233"/>', before=parse(run("A😀BC",props))
  const after=parse(run("A",props)+run("😀",`<w:b/>${props}${rev(props)}`)+run("BC",props))
  expect(verifyWordChatInlineFormatting(before,after,1,3,{bold:true},"Reviewer")).toEqual({text:"A😀BC",revisions:[{id:"7",author:"Reviewer",date:"2026-01-01T00:00:00Z",start:1,end:3}]})
  expect(()=>verifyWordChatInlineFormatting(before,after,1,2,{bold:true})).toThrow("inline_invalid_span")
  expect(()=>verifyWordChatInlineFormatting(before,parse(run("A",props)+run("😀",`<w:b/>${props}${rev('<w:i/>')}`)+run("BC",props)),1,3,{bold:true})).toThrow("inline_revision_old_properties_mismatch")
  expect(()=>verifyWordChatInlineFormatting(before,parse(run("A",props)+run("😀",`<w:b/><w:i/><w:color w:val="FFFFFF"/>${rev(props)}`)+run("BC",props)),1,3,{bold:true})).toThrow("inline_unrequested_property_changed")
  expect(()=>verifyWordChatInlineFormatting(before,after,1,3,{bold:true},"Someone else")).toThrow("inline_revision_old_properties_mismatch")
})

test("direct off, unchanged properties and separate mixed-run revisions are checked",()=>{
  const before=parse(run("One",'<w:b/>')+run("Two",'<w:b/><w:i/>'))
  const after=parse(run("One",`<w:b w:val="0"/>${rev('<w:b/>','10')}`)+run("Two",`<w:b w:val="false"/><w:i/>${rev('<w:b/><w:i/>','11')}`))
  expect(verifyWordChatInlineFormatting(before,after,0,6,{bold:false}).revisions.map(r=>r.id)).toEqual(["10","11"])
  expect(verifyWordChatInlineFormatting(before,before,0,6,{bold:true}).revisions).toEqual([])
  expect(()=>verifyWordChatInlineFormatting(before,parse(run("One",'<w:b/>')+run("Two",'<w:b/><w:i/>')),0,6,{bold:false})).toThrow("inline_requested_property_not_confirmed")
  expect(()=>verifyWordChatInlineFormatting(before,parse(run("One",`<w:b w:val="0"/>${rev('<w:b/>','10')}`)+run("Two",`<w:b w:val="0"/><w:i/>${rev('<w:b/><w:i/>','10')}`)),0,6,{bold:false})).toThrow("inline_revision_id_collision")
})

test("prior revisions, paragraph properties, annotations and unselected formatting cannot drift",()=>{
  const prior='<w:ins w:id="20" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Old</w:t></w:r></w:ins>', bookmark='<w:bookmarkStart w:id="4" w:name="Keep"/>'
  const before=parse(`${prior}${bookmark}${run("Target")}${run("Tail",'<w:i/>')}`)
  const after=parse(`${prior}${bookmark}${run("Target",`<w:u w:val="single"/>${rev('','21')}`)}${run("Tail",'<w:i/>')}`)
  expect(verifyWordChatInlineFormatting(before,after,3,9,{underline:true}).revisions).toHaveLength(1)
  expect(()=>verifyWordChatInlineFormatting(before,after,0,3,{underline:true})).toThrow("inline_protected_span")
  expect(()=>verifyWordChatInlineFormatting(before,parse(`${prior.replace('Prior','Other')}${bookmark}${run("Target",`<w:u w:val="single"/>${rev('','21')}`)}${run("Tail",'<w:i/>')}`),3,9,{underline:true})).toThrow("inline_protected_content_changed")
  expect(()=>verifyWordChatInlineFormatting(before,parse(`${prior}${bookmark.replace('Keep','Changed')}${run("Target",`<w:u w:val="single"/>${rev('','21')}`)}${run("Tail",'<w:i/>')}`),3,9,{underline:true})).toThrow("inline_protected_content_changed")
  expect(()=>verifyWordChatInlineFormatting(before,parse(`${prior}${bookmark}${run("Target",`<w:u w:val="single"/>${rev('','21')}`)}${run("Tail",'<w:b/>')}`),3,9,{underline:true})).toThrow("inline_unselected_format_changed")
  const altered=after.cloneNode(true) as Element; altered.getElementsByTagNameNS(W,"pStyle")[0]!.setAttributeNS(W,"w:val","Heading1")
  expect(()=>verifyWordChatInlineFormatting(before,altered,3,9,{underline:true})).toThrow("inline_protected_content_changed")
  const oldFormat = `<w:i/>${rev('','30','Earlier')}`, oldWithFormat=parse(`${run("Prior",oldFormat)}${run("Target")}`)
  const retainedFormat=parse(`${run("Prior",oldFormat)}${run("Target",`<w:b/>${rev('','31')}`)}`)
  expect(verifyWordChatInlineFormatting(oldWithFormat,retainedFormat,5,11,{bold:true}).revisions).toHaveLength(1)
  const changedFormat=parse(`${run("Prior",oldFormat.replace('Earlier','Different'))}${run("Target",`<w:b/>${rev('','31')}`)}`)
  expect(()=>verifyWordChatInlineFormatting(oldWithFormat,changedFormat,5,11,{bold:true})).toThrow("inline_unselected_format_changed")
})

test("native effective proof permits absent direct value but never a conflicting value",()=>{
  const before=parse(run("Text",'<w:b/>')), span={startOffset:0,endOffset:4,text:"Text"}
  const absent=parse(run("Text",rev('<w:b/>')))
  expect(()=>verify(before,absent,span,{bold:false})).toThrow("inline_requested_property_not_confirmed")
  expect(verify(before,absent,span,{bold:false},undefined,{bold:false}).revisions).toHaveLength(1)
  expect(()=>verify(before,absent,span,{bold:false},undefined,{bold:true})).toThrow("inline_requested_property_not_confirmed")
  const conflicting=parse(run("Text",`<w:b/>${rev('<w:b/>')}`))
  expect(()=>verify(before,conflicting,span,{bold:false},undefined,{bold:false})).toThrow("inline_requested_property_not_confirmed")
})
