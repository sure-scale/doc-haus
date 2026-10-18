import { buildChecklistDocumentContext, type ChecklistDocumentContext } from "./checklist-document-context"
import type { ChecklistComparisonAssignment, ChecklistItem } from "./checklist-contract"

const MAX_CHARACTERS = 4500
const MAX_PARAGRAPHS = 12

/** Explicit finite comparison assignments over the actual accepted Word body.
 * These are scopes in the existing refinement store, not generated duties. */
export function checklistTargetAssignments(directory: string, item: ChecklistItem) {
  const context = buildChecklistDocumentContext(directory, { item_id: "comparison", item })
  const matches = context.artifacts.filter(row => row.value.aliases.includes(item.target.document))
  if (matches.length !== 1 || !matches[0]!.value.paragraphs.length) throw Error("Full draft comparison needs one readable exact target with accepted paragraphs.")
  const target = matches[0]!
  const groups: ChecklistComparisonAssignment["ranges"][] = []
  let group: ChecklistComparisonAssignment["ranges"] = [], characters = 0
  for (const paragraph of target.value.paragraphs) {
    const length = paragraph.paragraph_text.length
    for (let start = 0; start < Math.max(1, length);) {
      const maximum = Math.min(length, start + MAX_CHARACTERS)
      const boundary = paragraph.paragraph_text.lastIndexOf(" ", maximum - 1)
      let end = maximum < length && boundary > start + MAX_CHARACTERS / 2 ? boundary + 1 : maximum
      // Preserve Unicode code points when a long indivisible token is split.
      if (end < length && /[\uDC00-\uDFFF]/.test(paragraph.paragraph_text[end]!)) end--
      if (group.length && (characters + end - start > MAX_CHARACTERS || group.length >= MAX_PARAGRAPHS)) {
        groups.push(group); group = []; characters = 0
      }
      group.push({ paragraph_index: paragraph.paragraph_index, start_char: start, end_char: end })
      characters += end - start
      start = end || 1
    }
  }
  if (group.length) groups.push(group)
  return {
    target: { id: target.value.id, artifact: target.value.artifact, sha256: target.value.sha256 },
    assignments: [
      ...groups.map(ranges => ({ id: `body:${ranges.map(range => `${range.paragraph_index}@${range.start_char}:${range.end_char}`).join(",")}`, assignment: { kind: "target_body" as const, ranges } })),
      { id: "process-records", assignment: { kind: "process_records" as const, ranges: [] } },
      // Exactly one whole-target unit: document-wide properties are never split by range.
      { id: "document-properties", assignment: { kind: "document_properties" as const, ranges: [] } },
    ],
  }
}

/** Replace lexical samples only for explicitly selected comparison children.
 * Neighbors and other documents remain available through targeted native reads. */
export function assignChecklistTargetContext(context: ChecklistDocumentContext, assignment: ChecklistComparisonAssignment) {
  const target = context.artifacts.filter(row => row.value.aliases.includes(context.input.item.target.document))
  if (target.length !== 1) throw Error("Comparison target is unavailable or ambiguous.")
  const body = assignment.ranges.flatMap(range => {
    const paragraph = target[0]!.value.paragraphs.find(row => row.paragraph_index === range.paragraph_index)
    if (!paragraph || range.end_char < range.start_char || range.end_char > paragraph.paragraph_text.length) throw Error("Assigned current paragraph range is unavailable.")
    const text = paragraph.paragraph_text.slice(range.start_char, range.end_char)
    // The range stays assigned; a whitespace-only slice is not evidence.
    return text.trim() ? [{ document: context.input.item.target.document, locator: `native PIDs ${paragraph.pids.join(", ") || "unavailable"}; accepted paragraph_index ${range.paragraph_index}; UTF-16 chars ${range.start_char}-${range.end_char}`, text, kind: "body" as const }] : []
  })
  // The whole-target unit receives each exact artifact property string, not
  // passages, and the delivery rows of the files that deliver the target
  // (its own files and a combined file that contains it).
  const { delivery, ...rest } = context
  const whole = assignment.kind === "document_properties"
  const delivered = new Set(delivery?.map(row => JSON.stringify(row)))
  const current_evidence = !whole ? body
    : [...(target[0]!.value.annotations ?? []).map((annotation, index) => ({ document: annotation.artifact, locator: `read-record ${target[0]!.path} /annotations/${index}/properties`, text: annotation.properties, kind: "artifact" as const })),
      ...context.input.current_evidence.filter(row => row.kind === "artifact" && delivered.has(row.text))]
  return { ...rest, ...(whole && delivery ? { delivery } : {}), input: { ...context.input, assignment, current_evidence } }
}
