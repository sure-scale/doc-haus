import type { WordChatInlineFormat, WordChatListCapture, WordChatParagraphFormat, WordChatProposal, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { targetWordChatAmendment, validWordChatFormatting, validWordChatInlineFormatting, validWordChatParagraphInsertion, validWordChatParagraphDeletion, validWordChatParagraphMerge, validWordChatParagraphSplit, validWordChatListInsertion } from "../../../dochaus/lib/word-chat-amendment-contract"

export type WordChatApplyEligibility = { supported: true } | { supported: false; reason: string }

export function nativeWordChatListInsertionEligible(target: WordChatTarget, replacement: string, action: unknown, list?: WordChatListCapture,
  formatting?: WordChatParagraphFormat[], inlineFormatting?: WordChatInlineFormat): boolean {
  if (!list) return false
  const paragraphs = list.members.map(p => ({ id: p.id, text: p.text }))
  const snapshot = { text: paragraphs.map(p => p.text).join("\r"), paragraphs, list, selection: target,
    saved: target.saved, capturedAt: target.capturedAt, truncated: false }
  const canonical = targetWordChatAmendment(snapshot, target.text)
  return !!canonical && canonical.startOffset === target.startOffset && canonical.endOffset === target.endOffset &&
    validWordChatListInsertion(snapshot, canonical, replacement, action, formatting, inlineFormatting)
}

/** One complete paragraph with its existing space replaced by a mark. */
export function nativeWordChatParagraphSplitEligible(target: WordChatTarget, replacement: string, action: unknown, formatting?: WordChatParagraphFormat[], inlineFormatting?: WordChatInlineFormat): boolean {
  const snapshot = {text:target.paragraphs.map(p=>p.text).join("\r"),paragraphs:target.paragraphs,selection:target,saved:target.saved,capturedAt:target.capturedAt,truncated:false}
  const canonical=targetWordChatAmendment(snapshot,target.text)
  return !!canonical && canonical.startOffset===target.startOffset && canonical.endOffset===target.endOffset && validWordChatParagraphSplit(snapshot,canonical,replacement,action,formatting,inlineFormatting)
}

/** Two complete adjacent paragraphs, with only their boundary joined. */
export function nativeWordChatParagraphMergeEligible(target: WordChatTarget, replacement: string, action: unknown, formatting?: WordChatParagraphFormat[], inlineFormatting?: WordChatInlineFormat): boolean {
  const snapshot = { text:target.paragraphs.map(p => p.text).join("\r"), paragraphs:target.paragraphs,
    selection:target, saved:target.saved, capturedAt:target.capturedAt, truncated:false }
  const canonical = targetWordChatAmendment(snapshot, target.text)
  return !!canonical && canonical.startOffset === target.startOffset && canonical.endOffset === target.endOffset &&
    validWordChatParagraphMerge(snapshot, canonical, replacement, action, formatting, inlineFormatting)
}

/** Explicit reviewed mark deletion, distinct from clearing paragraph text. */
export function nativeWordChatParagraphDeletionEligible(target: WordChatTarget, replacement: string, action: unknown, formatting?: WordChatParagraphFormat[], inlineFormatting?: WordChatInlineFormat): boolean {
  const snapshot = { text: target.paragraphs.map(p => p.text).join("\r"), paragraphs: target.paragraphs,
    selection: target, saved: target.saved, capturedAt: target.capturedAt, truncated: false }
  const canonical = targetWordChatAmendment(snapshot, target.text)
  return !!canonical && canonical.startOffset === target.startOffset && canonical.endOffset === target.endOffset &&
    validWordChatParagraphDeletion(snapshot, canonical, replacement, action, formatting, inlineFormatting)
}

/** A complete unchanged anchor followed by one new paragraph. The live host
 * independently verifies its structure and resolves its native identity. */
export function nativeWordChatParagraphInsertionEligible(target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[], inlineFormatting?: WordChatInlineFormat): boolean {
  const snapshot = { text: target.paragraphs.map(p => p.text).join("\r"), paragraphs: target.paragraphs,
    selection: target, saved: target.saved, capturedAt: target.capturedAt, truncated: false }
  const canonical = targetWordChatAmendment(snapshot, target.text)
  return !!canonical && canonical.startOffset === target.startOffset && canonical.endOffset === target.endOffset &&
    validWordChatParagraphInsertion(snapshot, canonical, replacement, formatting, inlineFormatting)
}

/** Native wording transport for ordinary single-paragraph changes. Exact
 * authored hunks, formatting and fresh Word scope are verified by the host. */
export function nativeWordChatTextEligible(target: WordChatTarget, replacement: string): boolean {
  const paragraph = target.paragraphs[0]?.text
  const boundary = (offset: number) => paragraph !== undefined && Number.isSafeInteger(offset) && offset >= 0 && offset <= paragraph.length &&
    !(offset > 0 && offset < paragraph.length && /[\uD800-\uDBFF]/.test(paragraph[offset - 1]!) && /[\uDC00-\uDFFF]/.test(paragraph[offset]!))
  return target.text.length > 0 && replacement !== target.text && target.paragraphs.length === 1 &&
    boundary(target.startOffset) && boundary(target.endOffset) && paragraph!.slice(target.startOffset, target.endOffset) === target.text &&
    (replacement.length > 0 || target.paragraphs[0]!.text.slice(0, target.startOffset).length + target.paragraphs[0]!.text.slice(target.endOffset).length > 0) &&
    target.startOffset <= target.endOffset && !/[\r\n\t\v]/.test(target.text) && !/[\r\n\t\v]/.test(replacement)
}

/** Transport availability only. Live Word identity, target uniqueness,
 * protection, comments/revisions, formatting and the authored artifact must
 * still pass their existing checks before any mutation. */
export function wordChatApplyEligibility(target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[], nativeText = true, inlineFormatting?: WordChatInlineFormat, paragraphAction?: WordChatProposal["paragraphAction"], list?: WordChatListCapture): WordChatApplyEligibility {
  if (paragraphAction === "split") return nativeText && nativeWordChatParagraphSplitEligible(target,replacement,paragraphAction,formatting,inlineFormatting)
    ? {supported:true}
    : {supported:false,reason:"Review one complete paragraph split at an existing space before applying it."}
  if (paragraphAction === "merge") return nativeText && nativeWordChatParagraphMergeEligible(target, replacement, paragraphAction, formatting, inlineFormatting)
    ? { supported: true }
    : { supported: false, reason: "Review the complete adjacent paragraphs before merging them." }
  if (paragraphAction === "append_list_item") return nativeText && nativeWordChatListInsertionEligible(target, replacement, paragraphAction, list, formatting, inlineFormatting)
    ? { supported: true }
    : { supported: false, reason: "Prepare a new item at the end of the captured automatic list before applying it." }
  if (paragraphAction !== undefined) return nativeText && nativeWordChatParagraphDeletionEligible(target, replacement, paragraphAction, formatting, inlineFormatting)
    ? { supported: true }
    : { supported: false, reason: "Review deletion of one complete paragraph, with its paragraph mark, before applying it." }
  if (inlineFormatting !== undefined) {
    const snapshot = { text: target.paragraphs.map(p => p.text).join("\r"), paragraphs: target.paragraphs, selection: target,
      saved: target.saved, capturedAt: target.capturedAt, truncated: false }
    const canonical = targetWordChatAmendment(snapshot, target.text)
    return canonical && validWordChatInlineFormatting(snapshot, canonical, inlineFormatting, replacement, formatting) &&
      canonical.startOffset === target.startOffset && canonical.endOffset === target.endOffset
      ? { supported: true }
      : { supported: false, reason: "Prepare character formatting for one exact passage with its wording unchanged." }
  }
  if (nativeText && nativeWordChatParagraphInsertionEligible(target, replacement, formatting)) return { supported: true }
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
  const eligibility = wordChatApplyEligibility(proposal.target, proposal.after, proposal.formatting, true, proposal.inlineFormatting, proposal.paragraphAction, proposal.list)
  if (!eligibility.supported) throw Object.assign(new Error(eligibility.reason), { mutationState: "unchanged" as const })
  return apply()
}
