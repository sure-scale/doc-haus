import { normalize } from "../../lib/quote-verify"
import { checklistFindingRecovery, type ChecklistFinding, type ChecklistItem } from "./checklist-contract"
import type { ChecklistItemRecord } from "./checklist-flow"

export type ChecklistFindingClass = "departure" | "source_conflict" | "not_inspectable" | "missing_inspection" | "omission"
export type ChecklistFindingClassification =
  | { class: "source_conflict"; conflicting_item_ids: string[] }
  | { class: Exclude<ChecklistFindingClass, "source_conflict"> }

/** Why one adverse finding blocks, from its verified structured fields only;
 * never from wording. A finding that does not block delivery (recovery "none")
 * has no class. First match wins: a disclosed departure, a conflict with a
 * sibling item's sources, the indeterminate cause, else an omission. An
 * indeterminate finding that records no cause is a missing inspection: it is
 * reinspected, never repaired as an omission. */
export function classifyChecklistFinding(item: ChecklistItem, finding: ChecklistFinding, siblings: readonly Pick<ChecklistItemRecord, "id" | "item">[]): ChecklistFindingClassification | null {
  if (checklistFindingRecovery(item, finding) === "none") return null
  if (finding.departure_disclosure?.length) return { class: "departure" }
  const conflicting = conflictingItemIDs(item, finding, siblings)
  if (conflicting.length) return { class: "source_conflict", conflicting_item_ids: conflicting }
  if (finding.indeterminate_cause) return { class: finding.indeterminate_cause }
  if (finding.status === "indeterminate" || finding.action_status === "indeterminate") return { class: "missing_inspection" }
  return { class: "omission" }
}

/** A conflicting ref maps to a sibling citing the same document when either
 * normalized quote contains the other. The finding's own item, or an identical
 * copy under another ID (a comparison component of one requirement), is never
 * its own conflict. Sibling order, each ID once. */
function conflictingItemIDs(item: ChecklistItem, finding: ChecklistFinding, siblings: readonly Pick<ChecklistItemRecord, "id" | "item">[]): string[] {
  const refs = finding.conflicting_source_refs ?? []
  if (!refs.length) return []
  const own = JSON.stringify(item)
  const contains = (a: string, b: string) => {
    const [left, right] = [normalize(a), normalize(b)]
    return !!left && !!right && (left.includes(right) || right.includes(left))
  }
  return [...new Set(siblings.filter(sibling => sibling.id !== finding.item_id && JSON.stringify(sibling.item) !== own &&
    refs.some(ref => sibling.item.source_refs.some(source => source.document === ref.document && contains(source.quote, ref.quote))))
    .map(sibling => sibling.id))]
}
