import { expect, test } from "bun:test"
import { authorWordChatAmendment } from "./word-chat-amendments"
import type { WordChatProposal } from "../../../dochaus/lib/word-chat-amendment-contract"

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
