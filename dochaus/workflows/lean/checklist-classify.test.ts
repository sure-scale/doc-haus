import { expect, test } from "bun:test"
import { checklistFindingRecovery, type ChecklistFinding, type ChecklistItem, type ChecklistSourceReference } from "./checklist-contract"
import { classifyChecklistFinding } from "./checklist-classify"

const notice: ChecklistSourceReference = { document: "instructions.txt", locator: "paragraph 2", quote: "The supplier must give thirty days' notice before termination." }
const item: ChecklistItem = {
  check: "Check the termination clause requires thirty days' notice.",
  basis: "directive",
  source_refs: [notice],
  qualifications: [],
  target: { document: "agreement.docx", role: "operative agreement", treatment: "operative" },
}
const refined: ChecklistItem = { ...item, required_action: "adopt" }
const immediate: ChecklistSourceReference = { document: "instructions.txt", locator: "paragraph 5", quote: "Either party may terminate “with immediate effect” on written notice." }
const sibling = { id: "source-1-item-2", item: { ...item, check: "Check termination may take immediate effect.", source_refs: [immediate] } }
const other = { id: "source-2-item-1", item: { ...item, check: "Check the schedule repeats the immediate termination right.", source_refs: [{ ...immediate, document: "schedule.txt" }] } }
const self = { id: "source-1-item-1", item }
const siblings = [self, sibling, other]
const departure = [{ document: "drafting-notes.md", locator: "line 4", text: "The draft deliberately omits the notice period.", kind: "artifact" as const }]
const finding: ChecklistFinding = { item_id: "source-1-item-1", status: "unsatisfied", reason: "The clause states no notice period.", evidence: [], source_refs: [notice], inspection_scope: "Termination clause.", limitation: "" }
const indeterminate: ChecklistFinding = { ...finding, status: "indeterminate", reason: "The needed passage was not established." }
// The sibling quote reworded only by whitespace and quote marks the shared normalizer folds.
const conflict = [{ document: "instructions.txt", locator: "paragraph 5", quote: "may terminate  \"with immediate\n effect\"" }]

test("non-adverse findings have no class, exactly where delivery does not block", () => {
  const cases: [ChecklistItem, ChecklistFinding][] = [
    [item, { ...finding, status: "satisfied" }],
    [item, { ...finding, status: "not_applicable" }],
    // An informational open finding stays non-blocking even with a disclosure.
    [item, { ...finding, status: "open", departure_disclosure: departure, conflicting_source_refs: conflict }],
    [refined, { ...finding, status: "open", action_status: "fulfilled", open_decisions: ["Client approval."] }],
    [item, finding],
    [item, indeterminate],
    [refined, { ...finding, status: "open", action_status: "not_fulfilled", open_decisions: [] }],
    [refined, { ...finding, status: "satisfied" }],
  ]
  for (const [subject, value] of cases)
    expect(classifyChecklistFinding(subject, value, siblings) === null).toBe(checklistFindingRecovery(subject, value) === "none")
  expect(classifyChecklistFinding(item, { ...finding, status: "satisfied" }, siblings)).toBeNull()
  expect(classifyChecklistFinding(item, cases[2]![1], siblings)).toBeNull()
})

test("classifies each adverse class from structured fields only", () => {
  expect(classifyChecklistFinding(item, { ...finding, departure_disclosure: departure }, siblings)).toEqual({ class: "departure" })
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: conflict }, siblings)).toEqual({ class: "source_conflict", conflicting_item_ids: [sibling.id] })
  expect(classifyChecklistFinding(item, { ...indeterminate, indeterminate_cause: "not_inspectable" }, siblings)).toEqual({ class: "not_inspectable" })
  expect(classifyChecklistFinding(item, { ...indeterminate, indeterminate_cause: "missing_inspection" }, siblings)).toEqual({ class: "missing_inspection" })
  expect(classifyChecklistFinding(item, finding, siblings)).toEqual({ class: "omission" })
  // Wording that describes a departure or a conflict is not a structured disclosure.
  expect(classifyChecklistFinding(item, { ...finding, reason: "The draft deliberately departs; sources conflict." }, siblings)).toEqual({ class: "omission" })
  expect(classifyChecklistFinding(refined, { ...finding, status: "open", action_status: "not_fulfilled", open_decisions: [], departure_disclosure: departure }, siblings)).toEqual({ class: "departure" })
})

test("an indeterminate finding with no cause is a missing inspection, not an omission", () => {
  expect(classifyChecklistFinding(item, indeterminate, siblings)).toEqual({ class: "missing_inspection" })
  expect(classifyChecklistFinding(refined, { ...finding, status: "open", action_status: "indeterminate", open_decisions: [] }, siblings)).toEqual({ class: "missing_inspection" })
})

test("first match wins: departure, then a mapped conflict, then the indeterminate cause", () => {
  expect(classifyChecklistFinding(item, { ...indeterminate, departure_disclosure: departure, conflicting_source_refs: conflict, indeterminate_cause: "not_inspectable" }, siblings)).toEqual({ class: "departure" })
  expect(classifyChecklistFinding(item, { ...indeterminate, conflicting_source_refs: conflict, indeterminate_cause: "not_inspectable" }, siblings)).toEqual({ class: "source_conflict", conflicting_item_ids: [sibling.id] })
  expect(classifyChecklistFinding(item, { ...indeterminate, conflicting_source_refs: conflict, indeterminate_cause: "missing_inspection" }, siblings)).toEqual({ class: "source_conflict", conflicting_item_ids: [sibling.id] })
})

test("a conflict ref that maps to no sibling falls through", () => {
  const unmapped = [{ document: "instructions.txt", locator: "paragraph 9", quote: "Notices must be in writing." }]
  // Same quote, other document: the schedule sibling cites schedule.txt only.
  const elsewhere = [{ ...conflict[0]!, document: "letter.txt" }]
  for (const refs of [unmapped, elsewhere]) {
    expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: refs }, siblings)).toEqual({ class: "omission" })
    expect(classifyChecklistFinding(item, { ...indeterminate, conflicting_source_refs: refs, indeterminate_cause: "missing_inspection" }, siblings)).toEqual({ class: "missing_inspection" })
    expect(classifyChecklistFinding(item, { ...indeterminate, conflicting_source_refs: refs, indeterminate_cause: "not_inspectable" }, siblings)).toEqual({ class: "not_inspectable" })
  }
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: [] }, siblings)).toEqual({ class: "omission" })
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: conflict }, [])).toEqual({ class: "omission" })
})

test("maps a ref by document and normalized containment in either direction", () => {
  // The ref quote lies inside the sibling quote (whitespace and curly quotes differ).
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: conflict }, [sibling])).toEqual({ class: "source_conflict", conflicting_item_ids: [sibling.id] })
  // The sibling quote lies inside a longer ref quote.
  const longer = [{ document: "instructions.txt", locator: "paragraphs 5-6", quote: "Either   party may terminate “with immediate effect” on written notice. Clause 4 survives." }]
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: longer }, [sibling])).toEqual({ class: "source_conflict", conflicting_item_ids: [sibling.id] })
  // Every mapped sibling once, in sibling order; the matching document decides which.
  const both = [...conflict, { ...conflict[0]!, document: "schedule.txt" }, ...longer]
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: both }, [other, sibling])).toEqual({ class: "source_conflict", conflicting_item_ids: [other.id, sibling.id] })
  // Overlap that is not containment does not map.
  const overlap = [{ document: "instructions.txt", locator: "paragraph 5", quote: "on written notice. Clause 4 survives." }]
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: overlap }, [sibling])).toEqual({ class: "omission" })
})

test("the finding's own item is never its own conflicting sibling", () => {
  const own = [{ ...notice, quote: "must give thirty days’ notice" }]
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: own }, [self])).toEqual({ class: "omission" })
  // Another component of the same unchanged requirement is the same item under another ID.
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: own }, [self, { id: "source-1-item-1/unit-2", item }])).toEqual({ class: "omission" })
  // A different requirement citing the same passage is a sibling.
  const shared = { id: "source-1-item-3", item: { ...item, check: "Check the notice period applies to both parties." } }
  expect(classifyChecklistFinding(item, { ...finding, conflicting_source_refs: own }, [self, shared])).toEqual({ class: "source_conflict", conflicting_item_ids: [shared.id] })
})
