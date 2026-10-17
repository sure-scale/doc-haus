import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { renderMarkdownToDocx } from "../../lib/render"
import { checklistTargetAssignments, assignChecklistTargetContext } from "./checklist-target-context"
import { buildChecklistDocumentContext } from "./checklist-document-context"
import { buildChecklistRefinementContext } from "./checklist-refinement-context"
import { buildChecklistCheckerPrompt, type ChecklistCheckerInput, type ChecklistItem } from "./checklist-contract"
import { checklistRefinableParents, readEffectiveChecklistItems, reconcileChecklistParentCoverage, runChecklistCheckFlow, runChecklistExtractionFlow, runChecklistRefinementFlow } from "./checklist-flow"
import { writeLeanPlan } from "./plan"
import { readLeanJson, writeLeanBytes, writeLeanText } from "./state"
import { checklistOutputState, type StageSessionPort } from "./stages"

const dirs: string[] = []
afterEach(() => dirs.splice(0).forEach(dir => rmSync(dir, { recursive: true, force: true })))
const request = "Use the supplied terms. Payment is due after acceptance. Acceptance remains subject to written confirmation. Missing confirmation must remain visibly unresolved."
const ref = { document: "request", locator: "request", quote: request }
const item: ChecklistItem = { check: "Check the draft follows the supplied terms.", basis: "directive", source_refs: [ref], qualifications: ["Preserve unresolved confirmation."], target: { document: "d1", role: "Requested terms", treatment: "operative" } }
async function setup(markdown?: string) {
  const dir = mkdtempSync(path.join(tmpdir(), "target-comparison-")); dirs.push(dir)
  writeLeanPlan(dir, [{ title: "Terms", purpose: "Requested terms", file: "terms.docx", parts: [{ heading: "Terms", scope: "Record terms" }] }])
  writeLeanText(dir, "request.md", request)
  await save(dir, markdown ?? ["# Terms", "Acceptance remains subject to written confirmation.", ...Array.from({ length: 31 }, (_, index) => `Paragraph ${index}: confirmation remains unresolved.`), "| Heading | Condition |\n| --- | --- |\n| Payment | After acceptance |"].join("\n\n"))
  return dir
}
async function save(dir: string, markdown: string) { writeLeanBytes(dir, "d1/body.docx", (await renderMarkdownToDocx(markdown)).bytes) }

// Scripted native boundaries test routing, canonical evidence validation and
// reconciliation. They do not substitute for a real legal-model acceptance run.
function port() {
  const inputs: ChecklistCheckerInput[] = []
  let created = 0
  const sessions: StageSessionPort = {
    async createSession() { return `s${++created}` },
    async prompt(call) {
      if (call.text.includes("SOURCE_INPUT:")) return JSON.stringify({ items: [item, { ...item, check: "Check the confirmation remains visible." }] })
      if (call.text.includes("REFINEMENT_INPUT:")) {
        const packet = JSON.parse(call.text.split("REFINEMENT_INPUT:\n\n")[1]!)
        return JSON.stringify({ parent_id: packet.parent_id, unit_id: packet.source_unit.id, disposition: "refined", reason: "Focused assessment.", items: [{ ...item, required_action: "assess" }], source_refs: [ref], limitations: [] })
      }
      const input: ChecklistCheckerInput = JSON.parse(call.text.split("CHECK_INPUT:\n\n")[1]!)
      inputs.push(input)
      const unsupported = input.current_evidence.some(row => row.text.includes("An additional charge is payable."))
      const condition = input.current_evidence.some(row => row.text.includes("Payment is immediate and unconditional."))
        ? buildChecklistDocumentContext(call.directory, { item_id: input.item_id, item: input.item }).artifacts[0]!.value.paragraphs.find(row => row.paragraph_text === "Acceptance remains subject to written confirmation.") : undefined
      const added = unsupported || !!condition
      return JSON.stringify({ item_id: input.item_id, status: input.assignment ? added ? "unsatisfied" : "satisfied" : input.item.required_action ? "satisfied" : input.item.check === item.check ? "indeterminate" : "satisfied",
        ...(input.item.required_action ? { action_status: "fulfilled", open_decisions: [] } : {}),
        reason: added ? "The assigned material addition conflicts with or lacks support in the supplied terms." : input.assignment ? "The assigned comparison is complete under the supplied terms." : "A complete current comparison is required.",
        evidence: [...input.current_evidence, ...(condition ? [{ document: "d1", locator: `accepted paragraph_index ${condition.paragraph_index}; targeted cross-boundary dependency`, text: condition.paragraph_text, kind: "body" }] : [])], source_refs: [ref], inspection_scope: input.assignment ? JSON.stringify(input.assignment) : "Initial hints only.", limitation: "Scripted isolated contract test." })
    },
  }
  return { sessions, model: null, inputs, count: () => created }
}
async function select(dir: string, deps: ReturnType<typeof port>, parent_id: string) {
  const selection = buildChecklistRefinementContext(dir, { parent_id, scope: "target_body" })
  await runChecklistRefinementFlow(dir, deps, selection)
  return selection
}

test("every actual accepted paragraph including headings and tables is assigned once within generic bounds, with one process-records and one document-properties unit", async () => {
  const dir = await setup(), context = buildChecklistDocumentContext(dir, { item_id: "p", item })
  const target = context.artifacts[0]!.value
  const units = checklistTargetAssignments(dir, item).assignments
  expect(units.filter(unit => unit.assignment.kind === "process_records")).toHaveLength(1)
  expect(units.filter(unit => unit.assignment.kind === "document_properties")).toEqual([{ id: "document-properties", assignment: { kind: "document_properties", ranges: [] } }])
  const ranges = units.flatMap(unit => unit.assignment.ranges)
  expect(ranges.map(range => range.paragraph_index)).toEqual(target.paragraphs.map(row => row.paragraph_index))
  expect(target.paragraphs[0]!.paragraph_text).toBe("Terms")
  expect(target.paragraphs.some(row => row.paragraph_text === "After acceptance")).toBe(true)
  for (const unit of units) {
    expect(unit.assignment.ranges.length).toBeLessThanOrEqual(12)
    expect(unit.assignment.ranges.reduce((sum, range) => sum + range.end_char - range.start_char, 0)).toBeLessThanOrEqual(4500)
    const assigned = assignChecklistTargetContext(context, unit.assignment)
    expect(assigned.input.item).toBe(item)
    if (unit.assignment.kind === "process_records") expect(assigned.input.current_evidence).toEqual([])
    else if (unit.assignment.kind === "document_properties") {
      // Exact artifact property strings only: no passage is assigned to the whole-target unit.
      expect(assigned.input.current_evidence).toEqual(target.annotations!.map((annotation, index) => ({ document: annotation.artifact, locator: `read-record ${context.artifacts[0]!.path} /annotations/${index}/properties`, text: annotation.properties, kind: "artifact" })))
      expect(assigned.input.current_evidence.length).toBeGreaterThan(0)
      for (const row of assigned.input.current_evidence) expect(JSON.parse(row.text).accepted_body_text_measures.paragraphs).toBe(target.paragraphs.length)
    } else expect(assigned.input.current_evidence.map(row => row.text)).toEqual(unit.assignment.ranges.map(range => target.paragraphs.find(row => row.paragraph_index === range.paragraph_index)!.paragraph_text))
  }
})

test("long paragraphs are split completely without increasing the job bound; missing or ambiguous targets fail", async () => {
  const dir = await setup("# Terms\n\n" + "A long current paragraph. ".repeat(600))
  const context = buildChecklistDocumentContext(dir, { item_id: "p", item })
  const ranges = checklistTargetAssignments(dir, item).assignments.flatMap(unit => unit.assignment.ranges).filter(range => range.paragraph_index === 1)
  expect(ranges.length).toBeGreaterThan(2)
  expect(ranges[0]!.start_char).toBe(0)
  expect(ranges.at(-1)!.end_char).toBe(context.artifacts[0]!.value.paragraphs[1]!.paragraph_text.length)
  expect(ranges.slice(1).every((range, index) => range.start_char === ranges[index]!.end_char)).toBe(true)
  expect(() => checklistTargetAssignments(dir, { ...item, target: { ...item.target, document: "Terms" } })).toThrow("exact target")
})

test("explicit selection preserves the consumed parent, unrelated settled check and unchanged child cache", async () => {
  const dir = await setup(), deps = port()
  const extraction = await runChecklistExtractionFlow(dir, deps), parent = extraction.items[0]!
  await runChecklistCheckFlow(dir, deps)
  await runChecklistCheckFlow(dir, deps, { reinspectAdverse: true })
  const before = readLeanJson(dir, "checklist/checks.json") as { checks: unknown[] }
  const original = before.checks[0]
  expect(original).toHaveProperty("reinspection")
  const calls = deps.count()
  const selection = await select(dir, deps, parent.id)
  expect(deps.count()).toBe(calls) // deterministic assignment, no bookkeeping model job
  const result = await runChecklistCheckFlow(dir, deps)
  expect(result.parent_coverage[0]!.status).toBe("satisfied")
  expect(deps.inputs.slice(-selection.units.length).every(input => input.assignment)).toBe(true)
  const saved = readLeanJson(dir, "checklist/checks.json") as { checks: unknown[] }
  expect(saved.checks).toContainEqual(original)
  expect(saved.checks).toContainEqual(before.checks[1])
  expect(deps.count()).toBe(calls + selection.units.length)
  await select(dir, deps, parent.id)
  const count = deps.count()
  await runChecklistCheckFlow(dir, deps)
  expect(deps.count()).toBe(count)
  expect((readLeanJson(dir, "checklist/checks.json") as { checks: unknown[] }).checks).toEqual(saved.checks)
})

test("a distant unsupported material addition and cross-boundary qualification remain blocking after a passing comparison", async () => {
  for (const addition of ["An additional charge is payable.", "Payment is immediate and unconditional."]) {
    const dir = await setup(), deps = port()
    const parent = (await runChecklistExtractionFlow(dir, deps)).items[0]!
    await runChecklistCheckFlow(dir, deps)
    await select(dir, deps, parent.id)
    expect((await runChecklistCheckFlow(dir, deps)).parent_coverage[0]!.status).toBe("satisfied")
    await save(dir, ["# Terms", "Acceptance remains subject to written confirmation.", ...Array.from({ length: 30 }, () => "Written confirmation remains unresolved."), addition].join("\n\n"))
    expect(readEffectiveChecklistItems(dir).complete).toBe(false)
    await select(dir, deps, parent.id)
    const result = await runChecklistCheckFlow(dir, deps)
    expect(result.parent_coverage[0]!.status).toBe("unsatisfied")
    const adverse = result.checks.find(row => row.finding?.status === "unsatisfied")!
    expect(adverse.finding!.evidence.some(row => row.text === addition)).toBe(true)
    const input = deps.inputs.findLast(row => row.item_id === adverse.item_id)!
    if (addition.includes("unconditional")) {
      expect(input.current_evidence.some(row => row.text === "Acceptance remains subject to written confirmation.")).toBe(false)
      expect(adverse.finding!.evidence.some(row => row.locator.includes("targeted cross-boundary dependency"))).toBe(true)
    }
    expect(input.item.qualifications).toEqual(item.qualifications)
    expect(input.item.source_refs[0]!.quote).toContain("Acceptance remains subject to written confirmation")
  }
})

test("unchanged selection preserves each consumed adverse child assessment", async () => {
  const dir = await setup("# Terms\n\nAn additional charge is payable."), deps = port()
  const parent = (await runChecklistExtractionFlow(dir, deps)).items[0]!
  await runChecklistCheckFlow(dir, deps)
  await select(dir, deps, parent.id)
  await runChecklistCheckFlow(dir, deps)
  const reviewed = await runChecklistCheckFlow(dir, deps, { reinspectAdverse: true })
  expect(reviewed.checks.find(row => row.finding?.status === "unsatisfied")!.reinspection?.status).toBe("succeeded")
  const count = deps.count()
  await select(dir, deps, parent.id)
  const same = await runChecklistCheckFlow(dir, deps, { reinspectAdverse: true })
  expect(deps.count()).toBe(count)
  expect(same.checks).toEqual(reviewed.checks)
  expect(same.parent_coverage[0]!.status).toBe("unsatisfied")
})

test("scope changes retain inactive child findings and consumed budgets in the existing checks pool", async () => {
  const dir = await setup("# Terms\n\nAn additional charge is payable."), deps = port()
  const parent = (await runChecklistExtractionFlow(dir, deps)).items[0]!
  await runChecklistCheckFlow(dir, deps)
  await select(dir, deps, parent.id)
  await runChecklistCheckFlow(dir, deps)
  const reviewed = await runChecklistCheckFlow(dir, deps, { reinspectAdverse: true })
  const adverse = reviewed.checks.find(row => row.finding?.status === "unsatisfied")!
  expect(adverse.reinspection?.status).toBe("succeeded")
  await runChecklistRefinementFlow(dir, deps, { parent_id: parent.id, scope: "cited_sources", units: [{ id: "original-authority", passages: [ref], role: "Original instruction" }], instructions: [ref], targets: [{ document: "d1", role: "Requested terms" }], retrieval: "Exact original authority." })
  await runChecklistCheckFlow(dir, deps)
  expect((readLeanJson(dir, "checklist/checks.json") as { checks: unknown[] }).checks).toContainEqual(adverse)
  writeLeanText(dir, "checklist/report.md", "Focused current findings.")
  expect(checklistOutputState(dir).valid).toBe(true) // inactive adverse comparison has no current delivery authority
  await select(dir, deps, parent.id)
  const count = deps.count()
  const same = await runChecklistCheckFlow(dir, deps, { reinspectAdverse: true })
  expect(deps.count()).toBe(count)
  expect(same.checks.find(row => row.item_id === adverse.item_id)).toEqual(adverse)
  expect(same.parent_coverage[0]!.status).toBe("unsatisfied")
  expect(checklistOutputState(dir).valid).toBe(false)
})

test("missing body scope, missing process result and uninspected dependency cannot certify the parent", async () => {
  const dir = await setup(), deps = port(), parent = (await runChecklistExtractionFlow(dir, deps)).items[0]!
  await select(dir, deps, parent.id)
  const result = await runChecklistCheckFlow(dir, deps)
  expect(result.parent_coverage[0]!.status).toBe("satisfied")
  const process = readEffectiveChecklistItems(dir).items.find(row => row.assignment?.kind === "process_records")!
  expect(reconcileChecklistParentCoverage(dir, result.checks.filter(row => row.item_id !== process.id))[0]!.status).toBe("indeterminate")
  const disputed = result.checks.map((row, index) => index ? row : { ...row, finding: { ...row.finding!, status: "indeterminate" as const, limitation: "Cross-boundary source exception remains uninspected." } })
  expect(reconcileChecklistParentCoverage(dir, disputed)[0]!.status).toBe("indeterminate")
  const saved = readLeanJson(dir, "checklist/refinements.json") as { parents: { units: unknown[] }[] }
  saved.parents[0]!.units.shift()
  writeLeanText(dir, "checklist/refinements.json", JSON.stringify(saved))
  expect(reconcileChecklistParentCoverage(dir, result.checks)[0]!.status).toBe("indeterminate")
})

test("assignment instructions separate recorded method from substantive effects and retain cross-boundary source obligations", async () => {
  const dir = await setup(), context = buildChecklistDocumentContext(dir, { item_id: "p", item })
  for (const unit of checklistTargetAssignments(dir, item).assignments) {
    const prompt = buildChecklistCheckerPrompt(assignChecklistTargetContext(context, unit.assignment).input)
    expect(prompt).toContain("Decide only this component")
    if (unit.assignment.kind === "target_body") {
      expect(prompt).toContain("EVERY assigned")
      expect(prompt).toContain("including across assignment boundaries")
      expect(prompt).toContain("A boilerplate label never exempts a material operative effect")
      expect(prompt).toContain("no missing fact or approval")
      expect(prompt).toContain("separately required source checks remain necessary")
    } else if (unit.assignment.kind === "process_records") expect(prompt).toContain("do not require a complete body comparison here")
    else {
      expect(prompt).toContain("Decide ONLY document-wide properties of the named target")
      expect(prompt).toContain("Report not_applicable when the parent states no document-wide property")
      expect(prompt).not.toContain("EVERY assigned")
    }
    if (unit.assignment.kind === "document_properties") expect(prompt).not.toContain("Only the separate document-properties assignment decides it")
    else expect(prompt).toContain("If the parent concerns only such a property, report not_applicable for this component rather than indeterminate")
  }
})

test("range units defer a document-wide parent to its one document-properties finding without relaxing indeterminate aggregation", async () => {
  const dir = await setup(), deps = port(), parent = (await runChecklistExtractionFlow(dir, deps)).items[0]!
  await select(dir, deps, parent.id)
  const result = await runChecklistCheckFlow(dir, deps)
  const whole = readEffectiveChecklistItems(dir).items.filter(row => row.assignment?.kind === "document_properties")
  expect(whole).toHaveLength(1)
  expect(deps.inputs.filter(input => input.item_id === whole[0]!.id).flatMap(input => input.current_evidence).every(row => row.kind === "artifact" && "accepted_body_text_measures" in JSON.parse(row.text))).toBe(true)
  const deferred = (status: "satisfied" | "indeterminate" | "not_applicable") => result.checks.map(row => ({ ...row, finding: { ...row.finding!, status: row.item_id === whole[0]!.id ? status : "not_applicable" as const } }))
  expect(reconcileChecklistParentCoverage(dir, deferred("satisfied"))[0]!.status).toBe("satisfied")
  expect(reconcileChecklistParentCoverage(dir, deferred("indeterminate"))[0]!.status).toBe("indeterminate")
  expect(reconcileChecklistParentCoverage(dir, result.checks.filter(row => row.item_id !== whole[0]!.id))[0]!.status).toBe("indeterminate")
})

test("a saved comparison without the document-properties unit is stale until the full-draft check is selected again", async () => {
  const dir = await setup(), deps = port(), parent = (await runChecklistExtractionFlow(dir, deps)).items[0]!
  await runChecklistCheckFlow(dir, deps)
  const selection = await select(dir, deps, parent.id)
  const saved = readLeanJson(dir, "checklist/refinements.json") as { parents: { expected_units: string[]; units: { unit_id: string }[] }[] }
  saved.parents[0]!.expected_units = saved.parents[0]!.expected_units.filter(id => id !== "document-properties")
  saved.parents[0]!.units = saved.parents[0]!.units.filter(unit => unit.unit_id !== "document-properties")
  writeLeanText(dir, "checklist/refinements.json", JSON.stringify(saved))
  const stale = await runChecklistCheckFlow(dir, deps)
  expect(stale.status).toBe("incomplete")
  expect(stale.parent_coverage[0]!.status).toBe("indeterminate")
  const count = deps.count()
  await select(dir, deps, parent.id)
  expect(deps.count()).toBe(count) // deterministic re-selection, no model call
  const current = await runChecklistCheckFlow(dir, deps)
  // Unchanged range and process children keep their completed checks; only the new unit is checked.
  expect(deps.count()).toBe(count + 1)
  expect(deps.inputs.at(-1)!.assignment).toEqual({ kind: "document_properties", ranges: [] })
  expect(current.status).toBe("succeeded")
  expect(current.parent_coverage[0]!.status).toBe("satisfied")
  expect(current.parent_coverage[0]!.children).toHaveLength(selection.units.length)
})

test("selection projection uses only original structured adverse findings and tolerates unavailable records", async () => {
  const dir = await setup(), deps = port()
  expect(checklistRefinableParents(dir)).toEqual([])
  const parent = (await runChecklistExtractionFlow(dir, deps)).items[0]!
  await runChecklistCheckFlow(dir, deps)
  expect(checklistRefinableParents(dir)).toEqual([{ parent_id: parent.id, check: parent.item.check, target: parent.item.target, finding: { status: "indeterminate", reason: "A complete current comparison is required." }, scopes: ["target_body"] }])
  writeLeanText(dir, "checklist/checks.json", "{broken")
  expect(checklistRefinableParents(dir)).toEqual([])
})
