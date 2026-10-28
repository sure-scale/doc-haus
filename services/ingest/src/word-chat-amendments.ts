import { authorWordChatTarget } from "../../../dochaus/lib/docx-flat-opc"
import type { WordChatAuthoredAmendment, WordChatProposal } from "../../../dochaus/lib/word-chat-amendment-contract"

export { MAX_WORD_AMENDMENT_OOXML_BYTES, WordAmendmentError } from "../../../dochaus/lib/docx-flat-opc"

/** Pure preview authoring against the live target snapshot. The route must
 * resolve its stored proposal and current editing lease before calling this.
 * Neither this function nor the shared Docxodus writer publishes a version. */
export async function authorWordChatAmendment(ooxml: string, proposal: WordChatProposal): Promise<WordChatAuthoredAmendment> {
  if (!proposal.id || proposal.before !== proposal.target.text || proposal.before === proposal.after)
    throw new Error("word_proposal_changed")
  const authored = await authorWordChatTarget(ooxml, proposal.target, proposal.after)
  return { ...authored, proposalId: proposal.id }
}
