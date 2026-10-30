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
