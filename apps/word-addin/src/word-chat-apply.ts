import type { WordChatParagraphFormat, WordChatProposal, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { validWordChatFormatting } from "../../../dochaus/lib/word-chat-amendment-contract"

export type WordChatApplyEligibility = { supported: true } | { supported: false; reason: string }

/** Experimental native wording transport. It is deliberately not reachable
 * from the shipped pane: Word's text engine has been exercised on isolated
 * synthetic copies, but immediate rendering, the full Docxodus/Apply/sync
 * product path and Windows remain unqualified. */
export function nativeWordChatTextEligible(target: WordChatTarget, replacement: string): boolean {
  return replacement !== target.text && target.paragraphs.length === 1 &&
    target.startOffset <= target.endOffset && !/[\r\n\t\v]/.test(replacement)
}

/** Transport availability only. Live Word identity, target uniqueness,
 * protection, comments/revisions, formatting and the authored artifact must
 * still pass their existing checks before any mutation. */
export function wordChatApplyEligibility(target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[], nativeText = false): WordChatApplyEligibility {
  if (formatting !== undefined) {
    // Native paragraph properties remain the only enabled formatting transport.
    return replacement === target.text && validWordChatFormatting(target, formatting) &&
      formatting.some(op => op.paragraph === target.paragraphs.length - 1)
      ? { supported: true }
      : { supported: false, reason: "Word cannot safely apply this formatting change yet. Make it directly in Word." }
  }
  if (nativeText && nativeWordChatTextEligible(target, replacement)) return { supported: true }
  if (replacement === target.text) return { supported: false, reason: "Choose a wording, formatting or structural change to apply." }
  return { supported: false, reason: "Word cannot safely apply this change yet. Make it directly in Word." }
}

/** Refuse known-unavailable transport before the caller claims editing, records
 * a native BEGIN or marks the Word buffer changed. Recheck even for cached UI. */
export async function applyWordChatIfEligible<T>(proposal: WordChatProposal, apply: () => Promise<T>): Promise<T> {
  const eligibility = wordChatApplyEligibility(proposal.target, proposal.after, proposal.formatting)
  if (!eligibility.supported) throw Object.assign(new Error(eligibility.reason), { mutationState: "unchanged" as const })
  return apply()
}
