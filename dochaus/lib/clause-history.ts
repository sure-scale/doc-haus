import type { ClauseHistoryParagraph, ClauseHistoryRequest, ClauseHistoryRound } from "./clause-history-contract"
import type { ComparisonParagraph, ParagraphComparison } from "./version-compare-contract"

export type ClauseCorrespondence = { status: "verified"; paragraphs: ComparisonParagraph[]; evidence: NonNullable<ClauseHistoryRound["evidence"]> }
  | { status: "unresolved"; reason: string; paragraphs: [] }

export const historyParagraph = (paragraph: ComparisonParagraph): ClauseHistoryParagraph => ({
  anchorId: paragraph.anchorId, scope: paragraph.location.scope, paragraph: paragraph.location.paragraph, text: paragraph.text,
})
// A native paragraph terminator is outside its wording. No preview, case,
// punctuation, figure or Unicode normalization establishes clause identity.
const capturedText = (text: string) => text.replace(/\r$/, "").replace(/\r\n/g, "\n")
const unresolved = (reason: string): ClauseCorrespondence => ({ status: "unresolved", reason, paragraphs: [] })

export function resolveHistoryAnchor(paragraphs: ComparisonParagraph[], request: ClauseHistoryRequest["paragraph"]): ClauseCorrespondence {
  const candidates = paragraphs.filter(paragraph => paragraph.location.scope === request.scope && capturedText(paragraph.text) === capturedText(request.text)
    && (!request.anchorId || paragraph.anchorId === request.anchorId))
  return candidates.length === 1 && capturedText(request.text).length > 0
    ? { status: "verified", paragraphs: candidates, evidence: "anchor-version" }
    : unresolved(candidates.length > 1 ? "This wording occurs more than once in the saved version. Select an identified saved paragraph to disambiguate it."
      : "The captured paragraph does not match one complete paragraph in the acknowledged saved version. Unsaved wording is not saved history.")
}

const sameSet = (left: string[], right: string[]) => left.length === right.length && left.every(id => right.includes(id))
// A split or merge may move only paragraph boundaries. This comparison proves
// complete wording equality, not that similar legal language means the same thing.
const joined = (paragraphs: ComparisonParagraph[]) => paragraphs.map(paragraph => paragraph.text).join(" ").replace(/[\t\r\n ]+/g, " ").trim()
function literalGroups(paragraphs: ComparisonParagraph[], wording: string, story: string) {
  const groups: ComparisonParagraph[][] = []
  for (let start = 0; start < paragraphs.length; start++) {
    if (paragraphs[start]!.story !== story) continue
    const group: ComparisonParagraph[] = []
    for (let at = start; at < paragraphs.length && paragraphs[at]!.story === story; at++) {
      group.push(paragraphs[at]!)
      const text = joined(group)
      if (text === wording) groups.push([...group])
      if (text.length >= wording.length) break
    }
  }
  return groups
}

/** Map one verified paragraph/group across exact versions of the SAME logical
 * document. #201 supplies complete inventories and diff edges, never ordinals
 * or similarity as identity. Uncertain relations stop here rather than chaining. */
export function correspondHistoryParagraphs(comparison: ParagraphComparison, source: ComparisonParagraph[]): ClauseCorrespondence {
  if (!source.length || source.some(paragraph => !comparison.left.some(candidate => candidate.anchorId === paragraph.anchorId)))
    return unresolved("The clause's source paragraphs could not be located in this comparison.")
  const sourceIds = source.map(paragraph => paragraph.anchorId)
  const structural = comparison.changes.filter(change => change.left.some(id => sourceIds.includes(id)) && (change.left.length > 1 || change.right.length > 1))
  if (structural.length) {
    const edge = structural.length === 1 ? structural[0] : undefined
    const target = edge ? comparison.right.filter(paragraph => edge.right.includes(paragraph.anchorId)) : []
    if (edge && sameSet(edge.left, sourceIds) && target.length === edge.right.length && target.length > 0
      && source.every(paragraph => paragraph.story === source[0]!.story) && target.every(paragraph => paragraph.story === source[0]!.story)
      && joined(source).length > 0 && joined(source) === joined(target)
      && source.every(paragraph => comparison.left.filter(other => other.story === paragraph.story && other.text === paragraph.text).length === 1)
      && target.every(paragraph => comparison.right.filter(other => other.story === paragraph.story && other.text === paragraph.text).length === 1))
      return { status: "verified", paragraphs: target, evidence: "exact-split-merge" }
    return unresolved("The comparison contains a split or merge whose complete clause correspondence is not established.")
  }
  // #201 sometimes represents a split as a modification plus a separate
  // insertion rather than one many-sided edge. Verify ALL literal wording
  // across paragraph boundaries before trusting the surviving first paraId.
  const fullText = joined(source)
  const groups = fullText && source.every(paragraph => paragraph.story === source[0]!.story)
    ? literalGroups(comparison.right, fullText, source[0]!.story) : []
  if (groups.length === 1 && (source.length > 1 || groups[0]!.length > 1)
    && literalGroups(comparison.left, fullText, source[0]!.story).length === 1)
    return { status: "verified", paragraphs: groups[0]!, evidence: "exact-split-merge" }
  const native = source.flatMap(paragraph => {
    if (!paragraph.nativeId || comparison.left.filter(other => other.story === paragraph.story && other.nativeId === paragraph.nativeId).length !== 1) return []
    const matches = comparison.right.filter(other => other.story === paragraph.story && other.nativeId === paragraph.nativeId)
    return matches.length === 1 ? matches : []
  })
  const adjacent = (inventory: ComparisonParagraph[], selected: ComparisonParagraph[], ids: string[]) => selected.some(paragraph => {
    const at = inventory.findIndex(candidate => candidate.anchorId === paragraph.anchorId)
    return [inventory[at - 1], inventory[at + 1]].some(neighbour => neighbour?.story === paragraph.story && ids.includes(neighbour.anchorId))
  })
  const changedNative = native.length === source.length && native.some((paragraph, index) => paragraph.text !== source[index]!.text)
  const boundaryChanges = changedNative && comparison.changes.some(change =>
    !change.left.length && adjacent(comparison.right, native, change.right) || !change.right.length && adjacent(comparison.left, source, change.left))
  if (boundaryChanges) return unresolved("A neighbouring insertion or deletion may split or merge this edited paragraph. Its full clause correspondence is unresolved.")
  if (native.length === source.length && new Set(native.map(paragraph => paragraph.anchorId)).size === native.length)
    return { status: "verified", paragraphs: native, evidence: "word-paragraph-id" }
  const exact = source.flatMap(paragraph => {
    if (!paragraph.text || comparison.left.filter(other => other.story === paragraph.story && other.text === paragraph.text).length !== 1) return []
    const matches = comparison.right.filter(other => other.story === paragraph.story && other.text === paragraph.text)
    return matches.length === 1 ? matches : []
  })
  if (exact.length === source.length && new Set(exact.map(paragraph => paragraph.anchorId)).size === exact.length)
    return { status: "verified", paragraphs: exact, evidence: "exact-wording" }
  return unresolved("The same clause cannot be verified in this version. Repeated, changed, moved or missing wording is not matched by paragraph number or a similar preview.")
}
