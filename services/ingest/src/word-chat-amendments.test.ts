import { expect, test } from "bun:test"
import { strToU8, zipSync } from "fflate"
import { authorWordChatTableRowDeletion, authorWordChatAmendment } from "./word-chat-amendments"
import { targetWordChatAmendment, type WordChatProposal } from "../../../dochaus/lib/word-chat-amendment-contract"

function proposal(): WordChatProposal {
  return { id:"reviewed-proposal", documentId:"document", before:"terms", after:"terms", explanation:"Remove emphasis", openItems:[], createdAt:"now",
    inlineFormatting:{bold:false,underline:true}, target:{id:"native-selection",text:"terms",paragraphs:[{id:"native-paragraph",text:"Keep these terms."}],
      startOffset:11,endOffset:16,capturedAt:"now",saved:false} }
}

test("inline preparation returns the exact reviewed native property command without simulating a DOCX", async () => {
  const reviewed = proposal()
  const plan = await authorWordChatAmendment("", reviewed)
  expect(plan).toEqual({kind:"inline_format",proposalId:reviewed.id,target:reviewed.target,inlineFormatting:{bold:false,underline:true}})
  if (!("kind" in plan)) throw new Error("Expected native inline plan")
  plan.target.paragraphs[0]!.text = "Changed by recipient"
  expect(reviewed.target.paragraphs[0]!.text).toBe("Keep these terms.")
})

test("inline preparation cannot turn malformed properties or expanded wording into a native command", async () => {
  const reviewed = proposal()
  for (const changed of [
    {...reviewed,after:"new terms"},
    {...reviewed,formatting:[{paragraph:0,style:"Heading1" as const}]},
    {...reviewed,inlineFormatting:{}},
    {...reviewed,inlineFormatting:{bold:"false"} as never},
    {...reviewed,target:{...reviewed.target,endOffset:17}},
  ]) await expect(authorWordChatAmendment("",changed)).rejects.toThrow("word_proposal_changed")
})


test("paragraph insertion preparation returns a copy of the bound reviewed native plan", async () => {
  const text="Existing paragraph.", target={id:"host-id",text,paragraphs:[{id:"native",text}],startOffset:0,endOffset:text.length,capturedAt:"now",saved:false}
  const reviewed:WordChatProposal={id:"reviewed",documentId:"document",before:text,after:text+"\rAdded literal paragraph.",target,explanation:"Add",openItems:[],createdAt:"now"}
  const plan=await authorWordChatAmendment("unused native capture",reviewed)
  expect(plan).toEqual({kind:"paragraph_insert",proposalId:"reviewed",target,text:"Added literal paragraph."})
  if (!("kind" in plan)) throw new Error("Expected native plan")
  plan.target.paragraphs[0]!.text="changed"
  expect(reviewed.target.paragraphs[0]!.text).toBe(text)
})


test("paragraph insertion plan preserves the reviewed new paragraph Normal style", async () => {
  const text="Existing paragraph.", target={id:"host-id",text,paragraphs:[{id:"native",text}],startOffset:0,endOffset:text.length,capturedAt:"now",saved:false}
  const formatting=[{paragraph:1,style:"Normal" as const}]
  const proposal:WordChatProposal={id:"reviewed",documentId:"document",before:text,after:text+"\rAdded paragraph.",target,formatting,explanation:"Add",openItems:[],createdAt:"now"}
  const plan=await authorWordChatAmendment("unused",proposal)
  expect(plan).toEqual({kind:"paragraph_insert",proposalId:"reviewed",target,text:"Added paragraph.",formatting})
  if (!("kind" in plan) || plan.kind !== "paragraph_insert" || !plan.formatting) throw new Error("Expected styled insertion plan")
  plan.formatting[0]!.style="Title"
  expect(proposal.formatting![0]!.style).toBe("Normal")
})


test("explicit paragraph deletion prepares a source-bound command without importing DOCX",async()=>{
  const text="A full paragraph.",target={id:"selection",text,paragraphs:[{id:"native",text}],startOffset:0,endOffset:text.length,capturedAt:"now",saved:false}
  const proposal:WordChatProposal={id:"reviewed",documentId:"document",before:text,after:"",paragraphAction:"delete",target,explanation:"Remove",openItems:[],createdAt:"now"}
  const plan=await authorWordChatAmendment("unused",proposal)
  expect(plan).toEqual({kind:"paragraph_delete",proposalId:"reviewed",target})
  if (!("kind" in plan)) throw Error("Expected plan")
  plan.target.paragraphs[0]!.text="changed"; expect(proposal.target.paragraphs[0]!.text).toBe(text)
  for (const bad of [{...proposal,after:"new"},{...proposal,paragraphAction:"clear" as never},{...proposal,formatting:[{paragraph:0,style:"Normal" as const}]},{...proposal,inlineFormatting:{bold:true}},{...proposal,target:{...target,startOffset:1}}])
    await expect(authorWordChatAmendment("unused",bad)).rejects.toThrow("word_proposal_changed")
})


test("list plan requires the admitted snapshot and preserves its binding",async()=>{
  const text="Second", paragraph={id:"b",text}, list={id:3,format:"decimal" as const,members:[{id:"a",text:"First",level:0,label:"1."},{...paragraph,level:0,label:"2."}]}
  const target={id:"b",text,paragraphs:[paragraph],startOffset:0,endOffset:text.length,capturedAt:"now",saved:false}
  const snapshot={text:"First\rSecond",paragraphs:[{id:"a",text:"First"},paragraph],capturedAt:"now",saved:false,truncated:false,list}
  const resolved=targetWordChatAmendment(snapshot,text)!
  const proposal:WordChatProposal={id:"reviewed",documentId:"doc",before:text,after:text+"\rThird",target:resolved,paragraphAction:"append_list_item",list,explanation:"Add",openItems:[],createdAt:"now"}
  const plan=await authorWordChatAmendment("current capture",proposal,snapshot)
  expect(plan).toEqual({kind:"list_insert",proposalId:"reviewed",target:resolved,text:"Third",list})
  await expect(authorWordChatAmendment("current capture",proposal)).rejects.toThrow("word_proposal_changed")
  await expect(authorWordChatAmendment("current capture",proposal,{...snapshot,list:{...list,id:4}})).rejects.toThrow("word_proposal_changed")
  await expect(authorWordChatAmendment("current capture",{...proposal,after:text+"\rThird\rFourth"},snapshot)).rejects.toThrow("word_proposal_changed")
})


test("row deletion oracle binds all current table rows and exact requested target", async () => {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const p = (v:string) => `<w:tc><w:p><w:r><w:t>${v}</w:t></w:r></w:p></w:tc>`
  const source = zipSync({
    "[Content_Types].xml":strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    "_rels/.rels":strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels":strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/></Relationships>`),
    "word/settings.xml":strToU8(`<w:settings xmlns:w="${W}"/>`),
    "word/document.xml":strToU8(`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${p("Header")}${p("Value")}</w:tr><w:tr>${p("Delete")}${p("This")}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`),
  })
  const paragraphs = ["Header","Value","Delete","This"].map((text,i)=>({id:`p${i}`,text})), table = {tableIndex:0,rows:[paragraphs.slice(0,2),paragraphs.slice(2)]}
  const snapshot = {text:paragraphs.map(p=>p.text).join("\r"),paragraphs,table,capturedAt:"now",saved:false,truncated:false}
  const proposal = {kind:"table_row_delete" as const,id:"proposal",documentId:"document",explanation:"Remove",openItems:[],createdAt:"now",table,rowIndex:1,cells:["Delete","This"]}
  const oracle = await authorWordChatTableRowDeletion(snapshot,proposal,source)
  expect(oracle.kind).toBe("table_row_delete"); expect(oracle.rowIndex).toBe(1); expect(oracle.revisions.map(r=>r.family)).toEqual(["row_delete"])
  await expect(authorWordChatTableRowDeletion(snapshot,{...proposal,cells:["Delete","Changed"]},source)).rejects.toThrow("word_row_proposal_changed")
  const stale = {...table, rows:[[{...paragraphs[0]!,text:"Changed"},paragraphs[1]!],table.rows[1]!]}
  await expect(authorWordChatTableRowDeletion({...snapshot,table:stale,paragraphs:[{...paragraphs[0]!,text:"Changed"},...paragraphs.slice(1)]},{...proposal,table:stale},source)).rejects.toThrow("word_row_source_changed")
})


test("paragraph merge preparation binds the reviewed boundary without authoring a document", async () => {
  const paragraphs=[{id:"first",text:"First."},{id:"second",text:"Second."}],context={text:"First.\rSecond.",paragraphs,capturedAt:"now",saved:false,truncated:false}
  const target=targetWordChatAmendment(context,context.text)!
  const proposal:WordChatProposal={id:"merge",documentId:"doc",before:context.text,after:"First. Second.",explanation:"Join",openItems:[],createdAt:"now",target,paragraphAction:"merge"}
  expect(await authorWordChatAmendment("",proposal,context)).toEqual({kind:"paragraph_merge",proposalId:"merge",target})
  await expect(authorWordChatAmendment("",{...proposal,after:"First.Second."},context)).rejects.toThrow("word_proposal_changed")
  await expect(authorWordChatAmendment("",proposal,{...context,paragraphs:[paragraphs[1]!,paragraphs[0]!]})).rejects.toThrow("word_proposal_changed")
})


test("paragraph split preparation returns a pinned native plan and refuses drift", async () => {
  const paragraphs=[{id:"p",text:"First. Second."}],context={text:paragraphs[0]!.text,paragraphs,capturedAt:"now",saved:false,truncated:false},target=targetWordChatAmendment(context,context.text)!
  const proposal:WordChatProposal={id:"split",documentId:"doc",before:context.text,after:"First.\rSecond.",explanation:"Split",openItems:[],createdAt:"now",target,paragraphAction:"split"}
  expect(await authorWordChatAmendment("",proposal,context)).toEqual({kind:"paragraph_split",proposalId:"split",target})
  await expect(authorWordChatAmendment("",{...proposal,after:"First.\rChanged."},context)).rejects.toThrow("word_proposal_changed")
  await expect(authorWordChatAmendment("",proposal,{...context,paragraphs:[{...paragraphs[0]!,text:"Changed"}]})).rejects.toThrow("word_proposal_changed")
})
