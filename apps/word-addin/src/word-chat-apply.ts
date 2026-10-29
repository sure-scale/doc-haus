import type { WordChatParagraphFormat, WordChatProposal, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { validWordChatFormatting } from "../../../dochaus/lib/word-chat-amendment-contract"

export type WordChatApplyEligibility = { supported: true } | { supported: false; reason: string }

/** Transport availability only. Live Word identity, tables, protection and the
 * authored artifact still need their existing checks before any mutation. */
export function wordChatApplyEligibility(target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[]): WordChatApplyEligibility {
  // Keep scoped imports unavailable while their document-wide numbering effects
  // are unresolved. Only the existing native paragraph-formatting path remains.
  return formatting !== undefined && replacement === target.text && validWordChatFormatting(target, formatting) &&
    formatting.some(op => op.paragraph === target.paragraphs.length - 1)
    ? { supported: true }
    : { supported: false, reason: "Word cannot safely apply this change yet. Make it directly in Word." }
}

/** Refuse known-unavailable transport before the caller claims editing, records
 * a native BEGIN or marks the Word buffer changed. Recheck even for cached UI. */
export async function applyWordChatIfEligible<T>(proposal: WordChatProposal, apply: () => Promise<T>): Promise<T> {
  const eligibility = wordChatApplyEligibility(proposal.target, proposal.after, proposal.formatting)
  if (!eligibility.supported) throw Object.assign(new Error(eligibility.reason), { mutationState: "unchanged" as const })
  return apply()
}
