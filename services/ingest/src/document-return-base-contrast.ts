import type { ReturnBaseContrast, ReturnBasePassage } from "../../../dochaus/lib/document-return-tie-break-contract"
import type { ComparisonParagraph, ParagraphComparison } from "../../../dochaus/lib/version-compare-contract"

export type BaseContrastPart = { part: number; incoming: ReturnBasePassage; left: ReturnBasePassage; right: ReturnBasePassage }
type Source = { documentId: string; versionId: string; current: boolean; associationCorrected?: boolean }
const passage = (row: ComparisonParagraph): ReturnBasePassage => ({ scope: row.location.scope, story: row.story, paragraph: row.location.paragraph, text: row.text })
const unique = (rows: ComparisonParagraph[], row: ComparisonParagraph) => rows.filter(other => other.story === row.story && other.text === row.text).length === 1
function boundaryChange(comparison: ParagraphComparison, row: ComparisonParagraph, side: "left" | "right") {
  const rows = comparison[side], at = rows.findIndex(other => other.anchorId === row.anchorId)
  const neighbours = [rows[at - 1], rows[at + 1]].filter(other => other?.story === row.story).map(other => other!.anchorId)
  return comparison.changes.some(change => (!change.left.length || !change.right.length) && change[side].some(id => neighbours.includes(id)))
}
/** A unique comparison edge retrieves passages for a semantic question. It is
 * never a stable paragraph identity or proof of historical descent. */
function incomingPassage(comparison: ParagraphComparison, source: ComparisonParagraph) {
  const own = comparison.left.filter(row => row.story === source.story && row.text === source.text)
  if (own.length !== 1) return
  const exact = comparison.right.filter(row => row.story === source.story && row.text === source.text)
  if (exact.length === 1 && unique(comparison.right, exact[0]!)) return exact[0]
  if (exact.length > 1) return
  const edges = comparison.changes.filter(change => change.left.includes(own[0]!.anchorId))
  if (edges.length !== 1 || edges[0]!.kind !== "modified" || edges[0]!.left.length !== 1 || edges[0]!.right.length !== 1) return
  const target = comparison.right.find(row => row.anchorId === edges[0]!.right[0])
  if (!target || target.story !== source.story || !unique(comparison.right, target)
    || boundaryChange(comparison, own[0]!, "left") || boundaryChange(comparison, target, "right")) return
  return target
}

export function planBaseContrast(leftVersionId: string, rightVersionId: string, retained: ParagraphComparison, leftIncoming: ParagraphComparison, rightIncoming: ParagraphComparison) {
  const parts: BaseContrastPart[] = [], unresolved: string[] = []
  let changedPassages = 0
  for (const change of retained.changes) {
    const left = retained.left.filter(row => change.left.includes(row.anchorId)), right = retained.right.filter(row => change.right.includes(row.anchorId))
    if (left.length === 1 && right.length === 1 && left[0]!.text === right[0]!.text) continue
    changedPassages++
    if (change.kind !== "modified" || left.length !== 1 || right.length !== 1 || left[0]!.story !== right[0]!.story
      || !unique(retained.left, left[0]!) || !unique(retained.right, right[0]!)
      || boundaryChange(retained, left[0]!, "left") || boundaryChange(retained, right[0]!, "right")) {
      unresolved.push(`Changed passage ${changedPassages}: repeated, missing or many-sided retained correspondence.`); continue
    }
    const fromLeft = incomingPassage(leftIncoming, left[0]!), fromRight = incomingPassage(rightIncoming, right[0]!)
    if (!fromLeft || !fromRight || fromLeft.story !== fromRight.story || fromLeft.text !== fromRight.text) {
      unresolved.push(`Changed passage ${changedPassages}: incoming correspondence is missing or ambiguous.`); continue
    }
    const triple = { incoming: passage(fromLeft), left: passage(left[0]!), right: passage(right[0]!) }
    // Keep each complete paragraph, including its qualifications. No clipping.
    if (Object.values(triple).some(row => row.text.trim().length < 8 || row.text.length > 2000)) {
      unresolved.push(`Changed passage ${changedPassages}: complete qualified wording exceeds the bounded quotation contract.`); continue
    }
    parts.push({ part: changedPassages, ...triple })
  }
  if (!changedPassages) unresolved.push("No distinguishing retained wording was found; shared identity cannot identify a base.")
  return { leftVersionId, rightVersionId, coverage: { changedPassages, processedPassages: 0 }, unresolved, parts }
}

/** The model must copy the COMPLETE supplied paragraphs, so a supporting quote
 * cannot silently discard an exception from the retrieved passage. */
export function checkBaseContrast(text: string, part: BaseContrastPart): Omit<ReturnBaseContrast["parts"][number], "model"> {
  const value = JSON.parse(text) as Record<string, unknown>
  if (!value || Object.keys(value).sort().join(",") !== "evidence,limitations,reason,support" || !["left", "right", "neither", "uncertain"].includes(String(value.support))
    || typeof value.reason !== "string" || !value.reason.trim() || value.reason.length > 1200 || !Array.isArray(value.limitations) || value.limitations.length > 3
    || value.limitations.some(item => typeof item !== "string" || !item.trim() || item.length > 300)
    || value.support !== "uncertain" && value.limitations.length) throw new Error("The base contrast did not follow its bounded evidence contract.")
  const evidence = value.evidence as Record<string, unknown> | null
  if (evidence !== null && (!evidence || Object.keys(evidence).sort().join(",") !== "incomingQuote,leftQuote,rightQuote"
    || evidence.incomingQuote !== part.incoming.text || evidence.leftQuote !== part.left.text || evidence.rightQuote !== part.right.text)
    || value.support !== "uncertain" && !evidence) throw new Error("The base contrast did not copy its complete exact quotation triple.")
  return { part: part.part, support: value.support as ReturnBaseContrast["parts"][number]["support"], reason: value.reason.trim(), limitations: value.limitations as string[],
    ...(evidence ? { evidence: { incoming: part.incoming, left: part.left, right: part.right } } : {}) }
}

export function supportedBaseIds(candidates: Source[], allSources: Source[], judgments: { versionId: string; relation: string }[], contrasts: ReturnBaseContrast[]) {
  return candidates.filter(candidate => {
    if (candidate.current || candidate.associationCorrected || judgments.find(row => row.versionId === candidate.versionId)?.relation !== "same-document") return false
    // Contradictory document-identity judgments are not cured by excluding current.
    if (candidates.some(other => other.documentId === candidate.documentId && judgments.find(row => row.versionId === other.versionId)?.relation !== "same-document")) return false
    const rivals = allSources.filter(source => source.documentId === candidate.documentId && source.versionId !== candidate.versionId)
    if (!rivals.length) return false
    return rivals.every(rival => {
      const contrast = contrasts.find(row => row.leftVersionId === candidate.versionId && row.rightVersionId === rival.versionId || row.rightVersionId === candidate.versionId && row.leftVersionId === rival.versionId)
      if (!contrast || contrast.unresolved.length || !contrast.coverage.changedPassages || contrast.coverage.processedPassages !== contrast.coverage.changedPassages || contrast.parts.length !== contrast.coverage.changedPassages) return false
      const own = contrast.leftVersionId === candidate.versionId ? "left" : "right"
      return contrast.parts.some(part => part.support === own && part.evidence) && contrast.parts.every(part => (part.support === own || part.support === "neither") && !part.limitations.length && part.evidence)
    })
  }).map(candidate => candidate.versionId)
}
