import { afterEach, expect, test } from "bun:test"
import type { ChecklistFinding, ChecklistItem } from "./checklist-contract"
import { checklistRefinementSourceBasis } from "./checklist-flow"
import { finalGate } from "./gate"
import { writeLeanPlan } from "./plan"
import { checklistOutputState } from "./stages"
import { readLeanState, updateLeanStage, writeLeanText } from "./state"
import { cleanupMatters, matter } from "./test-support"

// The checklist's saved output as the final gate reads it. Integrity failures
// (a missing, failed or unreinspected finding) keep the output invalid and
// block; a completed adverse finding is disclosed with the review draft, with
// the class its verified fields establish. Files are written directly: no
// model, runner or repair round is involved.

afterEach(cleanupMatters)

const ref = (quote: string) => ({ document: "request", locator: "request", quote })
const requirement = (check: string, quote: string, extra: Partial<ChecklistItem> = {}): ChecklistItem =>
  ({ check, basis: "directive", source_refs: [ref(quote)], qualifications: [], target: { document: "d1", role: "Supply note", treatment: "operative" }, ...extra })
const ITEMS = [
  { id: "c1", item: requirement("The note states the delivery period.", "Deliver within thirty days of the order.") },
  { id: "c2", item: requirement("The note states when payment falls due.", "Payment is due on delivery.") },
]
const finding = (item_id: string, status: ChecklistFinding["status"], extra: Partial<ChecklistFinding> = {}): ChecklistFinding => ({
  item_id, status, reason: `Focused result ${status}.`, evidence: [], source_refs: [], inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", ...extra,
})
type Row = { item_id: string; status?: "succeeded" | "failed"; finding?: ChecklistFinding; reinspection?: unknown }

function saveChecks(directory: string, rows: Row[]) {
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: rows.map(row => row.item_id),
    checks: rows.map(row => ({ input_sha: "sha", status: "succeeded", issues: [], ...row })) }))
}

function checklistMatter(rows: Row[] = ITEMS.map(row => ({ item_id: row.id, finding: finding(row.id, "satisfied") }))) {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items: ITEMS }] }))
  saveChecks(directory, rows)
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  return directory
}

const unresolvedItems = (gate: ReturnType<typeof finalGate>) => gate.items.filter(item => item.kind === "checklist_unresolved")

test("an unsatisfied omission finding is disclosed first and makes a review draft with findings", () => {
  const directory = checklistMatter([{ item_id: "c1", finding: finding("c1", "unsatisfied") }, { item_id: "c2", finding: finding("c2", "open") }])
  const state = checklistOutputState(directory)
  expect(state).toMatchObject({ valid: true, unresolved: [{ item_id: "c1", class: "omission" }], unresolvedParents: [], openFindings: [{ item_id: "c2" }] })
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(unresolvedItems(gate)).toHaveLength(1)
  expect(gate.disclosed[0]).toMatchObject({ stage: "checklist-check", kind: "checklist_unresolved", status: "disclosed",
    ref: { checklist_item: "c1", checklist_class: "omission", document: "d1" },
    detail: "Checklist requirement unresolved (omission not fixed by the repair round): The note states the delivery period. Finding: Focused result unsatisfied. Limitation: Comments not inspected.",
    resolution: "not established as satisfied; disclosed with the review draft" })
  expect(gate.items[0]).toEqual(gate.disclosed[0]!)
  expect(gate.disclosed[1]).toMatchObject({ kind: "checklist_open", ref: { checklist_item: "c2" } })
})

test("each verified disclosure class is carried to the disclosed item", () => {
  const cases = [
    { finding: finding("c1", "indeterminate", { indeterminate_cause: "not_inspectable" }), label: "cannot be established from document text", class: "not_inspectable" },
    { finding: finding("c1", "indeterminate", { indeterminate_cause: "missing_inspection" }), label: "not inspected", class: "missing_inspection" },
    { finding: finding("c1", "unsatisfied", { conflicting_source_refs: [ref("Payment is due on delivery.")] }), label: "conflicts with another source requirement", class: "source_conflict" },
    { finding: finding("c1", "unsatisfied", { departure_disclosure: [{ document: "d1", locator: "p1", text: "The period is left out on purpose.", kind: "body" }] }), label: "deliberate departure", class: "departure" },
  ] as const
  for (const row of cases) {
    const directory = checklistMatter([{ item_id: "c1", finding: row.finding }, { item_id: "c2", finding: finding("c2", "satisfied") }])
    expect(checklistOutputState(directory)).toMatchObject({ valid: true, unresolved: [{ item_id: "c1", class: row.class }] })
    const gate = finalGate(directory)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
    expect(unresolvedItems(gate)).toEqual([expect.objectContaining({ status: "disclosed", ref: { checklist_item: "c1", checklist_class: row.class, document: "d1" },
      detail: expect.stringContaining(`(${row.label})`) })])
  }
  const conflict = checklistMatter([{ item_id: "c1", finding: cases[2].finding }, { item_id: "c2", finding: finding("c2", "satisfied") }])
  expect(checklistOutputState(conflict).unresolved).toEqual([expect.objectContaining({ class: "source_conflict", conflicting_item_ids: ["c2"] })])
  expect(unresolvedItems(finalGate(conflict))[0]!.detail).toContain("Conflicting checklist items: c2.")
})

test("a failed check row, a failed reinspection or a missing finding stays invalid and blocks", () => {
  const directory = checklistMatter()
  expect(checklistOutputState(directory).valid).toBe(true)
  expect(finalGate(directory).status).toBe("pass")
  const satisfied = { item_id: "c2", finding: finding("c2", "satisfied") }
  const reinspection = (status: "succeeded" | "failed") => ({ canonical_input_sha: "sha", original_finding: finding("c1", "unsatisfied"), status, issues: status === "failed" ? ["No reply."] : [] })
  // The same adverse finding with a completed reinspection is disclosed, not blocking.
  saveChecks(directory, [{ item_id: "c1", finding: finding("c1", "unsatisfied"), reinspection: reinspection("succeeded") }, satisfied])
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft_with_findings" })
  for (const broken of [
    { item_id: "c1", status: "failed" } as const,
    { item_id: "c1", finding: finding("c1", "unsatisfied"), reinspection: reinspection("failed") },
    { item_id: "c1" },
  ]) {
    saveChecks(directory, [broken, satisfied])
    expect(checklistOutputState(directory)).toEqual({ valid: false, open: false, openFindings: [], unresolved: [], unresolvedParents: [] })
    const gate = finalGate(directory)
    expect(gate).toMatchObject({ status: "blocked", delivery: null, blocking: [{ stage: "checklist-check", kind: "omission", status: "open" }] })
    expect(unresolvedItems(gate)).toEqual([])
  }
})

test("an informational open finding stays a checklist_open disclosure of a review draft", () => {
  const directory = checklistMatter([{ item_id: "c1", finding: finding("c1", "open", { open_decisions: ["Whether the period runs from dispatch"] }) }, { item_id: "c2", finding: finding("c2", "satisfied") }])
  expect(checklistOutputState(directory)).toMatchObject({ valid: true, open: true, openFindings: [{ item_id: "c1" }], unresolved: [], unresolvedParents: [] })
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate.disclosed).toEqual([expect.objectContaining({ kind: "checklist_open", ref: { checklist_item: "c1", document: "d1" },
    detail: "Checklist requirement left open: The note states the delivery period. Finding: Focused result open. Limitation: Comments not inspected. Open decisions: Whether the period runs from dispatch." })])
})

test("adverse coverage of a refined parent is disclosed, not blocking", () => {
  const child = { id: "c1.u1.1", item: requirement("The note adopts the delivery period.", "Deliver within thirty days of the order.", { required_action: "adopt" }) }
  const directory = checklistMatter([
    // Fulfilled adoption without operative evidence cannot certify the parent.
    { item_id: child.id, finding: finding(child.id, "satisfied", { action_status: "fulfilled" }) },
    { item_id: "c2", finding: finding("c2", "satisfied") },
  ])
  writeLeanText(directory, "checklist/refinements.json", JSON.stringify({ parents: [{ parent_id: "c1", parent: ITEMS[0]!.item, scope: "cited_sources",
    source_basis: checklistRefinementSourceBasis(directory), expected_units: ["u1"], units: [{ unit_id: "u1", input_sha: "sha", status: "succeeded",
      result: { parent_id: "c1", unit_id: "u1", disposition: "refined", reason: "Focused duty.", items: [child.item], source_refs: [ref("Deliver within thirty days of the order.")], limitations: [] },
      children: [child], issues: [] }] }] }))
  const state = checklistOutputState(directory)
  expect(state).toMatchObject({ valid: true, unresolved: [], unresolvedParents: [{ parent_id: "c1", status: "indeterminate", issues: [expect.stringContaining("adoption lacks current operative or artifact evidence")] }] })
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(unresolvedItems(gate)).toEqual([gate.disclosed[0]!])
  expect(gate.disclosed[0]).toMatchObject({ status: "disclosed", ref: { checklist_item: "c1", document: "d1" },
    detail: expect.stringContaining("The note states the delivery period.") })
  expect(gate.disclosed[0]!.ref.checklist_class).toBeUndefined()
})

test("a clean checklist still delivers execution-ready output", () => {
  const directory = checklistMatter()
  expect(checklistOutputState(directory)).toEqual({ valid: true, open: false, openFindings: [], unresolved: [], unresolvedParents: [] })
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "execution_ready", blocking: [], disclosed: [] })
})
