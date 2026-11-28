import { afterEach, expect, test } from "bun:test"
import { checklistFindingRecovery, type ChecklistFinding, buildChecklistReinspectionPrompt, buildChecklistCheckerPrompt, buildChecklistProducerPrompt, CHECKLIST_ITEM_SHAPE, parseChecklistCheckerInput, parseChecklistExtraction, parseChecklistFinding, parseChecklistItem, type ChecklistItem } from "./checklist-contract"
import { markerVerificationIssues, type ChecklistCheckerInput } from "./checklist-contract"
import { withFactOptions } from "./fact-answers"
import { writeLeanPlan } from "./plan"
import { provisionalFactMarkerItems } from "./provisional-fact-markers"
import { readQuestions, writeQuestions, type Question } from "./state"
import { cleanupMatters, matter } from "./test-support"
import { paragraphShowsOpenItem } from "./gate"

afterEach(cleanupMatters)

const item: ChecklistItem = {
  check: "Check the proposed notice arrangement remains subject to confirmation.",
  basis: "proposal",
  source_refs: [{ document: "instruction.txt", locator: "paragraph 2", quote: "  We propose advance notice; please confirm.\n" }],
  qualifications: ["The proposal has not been accepted."],
  target: { document: "agreement.docx", role: "operative agreement", treatment: "open_issue" },
}
const evidence = [{ document: "agreement.docx", locator: "p0123456789ab", text: "[open item: confirm notice arrangement]", kind: "body" as const }]
const packet = { item_id: "source-1-item-1", item, current_evidence: evidence, settled_instructions: [], retrieval: "Use read-record at current.json /paragraphs/0 and targeted lookup for necessary context." }

const markerPacket: ChecklistCheckerInput = { ...packet, verification: "provisional_fact_marker", item: { ...item,
  check: "Verify visible handling of the missing complete storage specification and unknown prior inspection.", basis: "open_question", required_action: "flag",
  source_refs: [{ document: "request", locator: "questions.json answer (q-storage)", question: "Is the complete specification supplied, and was the container inspected?", quote: "No more information is available. Prepare a review draft and disclose what remains unresolved." }],
} }

test("the current canonical marker item preserves permission and unknowns without mutating saved questions", () => {
  const directory = matter("The complete container specification is unavailable.", "Prepare a storage notice.")
  writeLeanPlan(directory, [{ title: "Storage notice", purpose: "Describe storage handling.", parts: [{ heading: "Storage", scope: "Container identification." }] }])
  const offered = withFactOptions({ id: "q-container", stage: "intake", kind: "fact", status: "open", question: "Is the complete specification available, and was the container inspected?", context: "The complete specification and prior inspection are unknown." })
  const question: Question = { ...offered, status: "answered", answer_kind: "option", option_id: "o1", answer: offered.options![0]!.text, fact_status: "provisional" }
  writeQuestions(directory, [question])
  const current = provisionalFactMarkerItems(directory)[0]!
  expect(current.item.source_refs[0]!.quote).toBe(question.answer!)
  expect(current.item.source_refs[0]!.question).toBe(question.question)
  expect(readQuestions(directory)).toEqual([question])
})

test("product marker verification preserves check input and rejects invalid verification modes", () => {
  const before = JSON.stringify(markerPacket)
  const prompt = buildChecklistCheckerPrompt(markerPacket)
  const input = JSON.parse(prompt.split("CHECK_INPUT:\n\n")[1]!)
  expect(input.item).toEqual(markerPacket.item)
  expect(input.verification).toBe("provisional_fact_marker")
  expect(input.item.source_refs[0].quote).toBe(markerPacket.item.source_refs[0]!.quote)
  expect(JSON.stringify(markerPacket)).toBe(before)
  expect(() => parseChecklistItem({ ...item, verification: "provisional_fact_marker" })).toThrow()
  expect(() => parseChecklistCheckerInput({ ...packet, verification: "source_override" })).toThrow()
  const ordinary = buildChecklistCheckerPrompt(packet)
  expect(buildChecklistCheckerPrompt({ ...packet, verification: undefined })).toBe(ordinary)
  expect(prompt).not.toBe(ordinary)
})

test("marker reinspection preserves the previous finding payload unchanged", () => {
  const previous: ChecklistFinding = { item_id: markerPacket.item_id, status: "not_applicable", action_status: "not_applicable", reason: "The human did not separately require marking.", evidence, source_refs: markerPacket.item.source_refs, inspection_scope: "The supplied current paragraphs.", limitation: "The facts remain unknown.", open_decisions: ["The prior inspection is unknown."] }
  const prompt = buildChecklistReinspectionPrompt(markerPacket, previous)
  const preserved = prompt.split("\n\n").find(row => row.startsWith("Previous finding (unverified assessment, not source authority): "))!
  expect(JSON.parse(preserved.slice(preserved.indexOf(": ") + 2))).toEqual(previous)
})

test("marker contract correction rejects affirmative contradictions while preserving valid adverse findings and ordinary semantics", () => {
  const positive: ChecklistFinding = { item_id: markerPacket.item_id, status: "satisfied", action_status: "fulfilled", reason: "Every current dependent passage visibly retains both unknowns.", evidence, source_refs: markerPacket.item.source_refs, inspection_scope: "All current dependent passages.", limitation: "The complete specification and inspection history remain unknown.", open_decisions: [] }
  expect(markerVerificationIssues(markerPacket, positive)).toEqual([])
  expect(markerVerificationIssues(markerPacket, { ...positive, status: "not_applicable", action_status: "not_applicable" })).toEqual([])
  for (const value of [
    { ...positive, action_status: "not_applicable" as const },
    { ...positive, status: "not_applicable" as const, action_status: "fulfilled" as const },
    { ...positive, open_decisions: ["The inspection history remains unknown."] },
    { ...positive, open_decisions: undefined },
    { ...positive, evidence: [] },
  ]) expect(markerVerificationIssues(markerPacket, value).length).toBeGreaterThan(0)
  const adverse: ChecklistFinding = { ...positive, status: "unsatisfied", action_status: "not_fulfilled", reason: "A dependent passage is unmarked.", open_decisions: ["Mark the remaining dependent passage."] }
  const before = JSON.stringify(adverse)
  expect(markerVerificationIssues(markerPacket, adverse)).toEqual([])
  expect(JSON.stringify(adverse)).toBe(before)
  expect(markerVerificationIssues(packet, { ...positive, status: "not_applicable", action_status: "not_applicable", open_decisions: ["Whether the proposal should be adopted."] })).toEqual([])
})

test("every positive marker body evidence row uses the gate's visible marker predicate", () => {
  const positive: ChecklistFinding = { item_id: markerPacket.item_id, status: "satisfied", action_status: "fulfilled", reason: "The dependent paragraph retains the unknown.", evidence, source_refs: markerPacket.item.source_refs, inspection_scope: "The current relevant passages.", limitation: "The specification remains unknown.", open_decisions: [] }
  const heading = { ...evidence[0]!, locator: "p-heading", text: "Storage attachment" }
  const mixed = { ...positive, evidence: [heading, ...evidence] }
  const before = JSON.stringify(mixed)
  const issues = markerVerificationIssues(markerPacket, mixed)
  expect(issues.length).toBeGreaterThan(0)
  expect(issues.join("\n")).toContain("evidence[0]")
  expect(issues.join("\n")).toContain("p-heading")
  expect(JSON.stringify(mixed)).toBe(before)
  for (const text of [
    "[open item: obtain specification]", "[OPEN\u00a0ITEM: obtain specification]", "[open\nitem: obtain specification]",
    "Specification [TBC]", "Specification [insert complete text]", "Specification ___", "Storage attachment", "The specification is unknown.",
  ]) {
    const checked = markerVerificationIssues(markerPacket, { ...positive, evidence: [{ ...evidence[0]!, text }] })
    expect(checked.length === 0).toBe(paragraphShowsOpenItem(text))
  }
  expect(markerVerificationIssues(markerPacket, { ...mixed, status: "not_applicable", action_status: "not_applicable" })).toEqual([])
  expect(markerVerificationIssues(markerPacket, { ...mixed, status: "unsatisfied", action_status: "not_fulfilled" })).toEqual([])
  expect(markerVerificationIssues(packet, mixed)).toEqual([])
})

test("producer to consumer preserves exact canonical source meaning", () => {
  const produced = parseChecklistExtraction({ items: [item] }).items[0]
  const consumed = parseChecklistCheckerInput({ ...packet, item: produced })
  expect(JSON.stringify(consumed.item)).toBe(JSON.stringify(item))
  expect(consumed.item.basis).toBe("proposal")
  expect(consumed.item.source_refs[0].quote).toBe(item.source_refs[0].quote)
  expect(consumed.item.qualifications).toEqual(item.qualifications)
})

test("prompt builders share the item shape and separate current evidence", () => {
  const producer = buildChecklistProducerPrompt({ source: { document: "instruction.txt", role: "client instructions" }, passages: item.source_refs, instructions: [], targets: [{ document: "agreement.docx", role: "operative agreement" }], retrieval: "Read exact cited source coordinates when needed." })
  const checker = buildChecklistCheckerPrompt(packet)
  expect(producer).toContain(CHECKLIST_ITEM_SHAPE)
  expect(checker).toContain(CHECKLIST_ITEM_SHAPE)
  const current = JSON.parse(checker.split("CHECK_INPUT:\n\n")[1])
  expect(current.item).toEqual(parseChecklistExtraction({ items: [item] }).items[0])
  expect(current.current_evidence).toEqual(evidence)
  expect(current.item).not.toHaveProperty("current_evidence")
  expect(producer).not.toContain(evidence[0].text)
})

test("later instructions and findings cannot replace the requirement", () => {
  const later = { document: "answers.json", locator: "/answers/0", quote: "The proposal remains subject to confirmation." }
  const before = JSON.stringify(packet)
  const checked = parseChecklistCheckerInput({ ...packet, settled_instructions: [later] })
  expect(checked.item).toEqual(item)
  expect(checked.settled_instructions).toEqual([later])
  expect(JSON.stringify(packet)).toBe(before)
  const finding = { item_id: packet.item_id, status: "open", reason: "Confirmation remains outstanding.", evidence, source_refs: [later], inspection_scope: "Cited target paragraph and answer.", limitation: "No confirmation supplied." }
  expect(parseChecklistFinding(finding).status).toBe("open")
  expect(() => parseChecklistFinding({ ...finding, item: { ...item, basis: "directive" } })).toThrow()
})

test("rejects malformed shape without inferring substance", () => {
  expect(parseChecklistItem(item)).toEqual(item)
  for (const malformed of [
    { ...item, basis: "mandatory" }, { ...item, check: "   " }, { ...item, source_refs: [] },
    { ...item, qualifications: undefined }, { ...item, target: [item.target] },
    { ...item, target: { ...item.target, pid: "invented" } }, { ...item, rationale: "Promote proposal to directive." },
  ]) expect(() => parseChecklistItem(malformed)).toThrow()
  expect(parseChecklistExtraction({ items: [] })).toEqual({ items: [] })
})

test("preserves every basis and evidence gaps without inventing final pointers", () => {
  for (const basis of ["directive", "proposal", "observation", "open_question"] as const) {
    const input = parseChecklistCheckerInput({ ...packet, item: { ...item, basis }, current_evidence: [] })
    expect(input.item.basis).toBe(basis)
    expect(input.item.target).not.toHaveProperty("pid")
    expect(input.current_evidence).toEqual([])
  }
  expect(parseChecklistFinding({ item_id: packet.item_id, status: "indeterminate", reason: "Target unavailable.", evidence: [], source_refs: item.source_refs, inspection_scope: "Supplied sources only.", limitation: "Inspect current target." }).status).toBe("indeterminate")
})


test("all treatment values round-trip from extraction to checker input without mutating observation basis", () => {
  for (const treatment of ["operative", "advice", "comment", "open_issue"] as const) {
    const observation = { ...item, basis: "observation" as const, target: { ...item.target, treatment } }
    const unchanged = parseChecklistCheckerInput({ ...packet, item: parseChecklistExtraction({ items: [observation] }).items[0] }).item
    expect(unchanged).toEqual(observation)
  }
})

test("substantive reinspection keeps the item unchanged", () => {
  const original = { item_id: packet.item_id, status: "unsatisfied" as const, reason: "Proposal was not adopted.", evidence: [], source_refs: item.source_refs, inspection_scope: "One paragraph.", limitation: "Related advice not inspected." }
  const prompt = buildChecklistReinspectionPrompt(packet, original)
  expect(JSON.parse(prompt.split("CHECK_INPUT:\n\n")[1]).item).toEqual(item)
})


test("recovery distinguishes an unfulfilled action, inspection gaps and fulfilled open approval without rewriting findings", () => {
  const refined = { ...item, required_action: "propose" as const }
  const base: ChecklistFinding = { item_id: "one", status: "open", action_status: "fulfilled", open_decisions: ["Approval"], reason: "Assessment", evidence: [], source_refs: item.source_refs, inspection_scope: "Target", limitation: "Approval unresolved" }
  const before = JSON.stringify(base)
  expect(checklistFindingRecovery(refined, base)).toBe("none")
  expect(checklistFindingRecovery(refined, { ...base, action_status: "not_fulfilled" })).toBe("repair")
  expect(checklistFindingRecovery(refined, { ...base, action_status: "indeterminate" })).toBe("inspect")
  expect(checklistFindingRecovery(refined, { ...base, action_status: undefined })).toBe("inspect")
  expect(checklistFindingRecovery(refined, { ...base, status: "satisfied", action_status: "not_fulfilled" })).toBe("inspect")
  expect(checklistFindingRecovery(refined, { ...base, status: "unsatisfied" })).toBe("inspect")
  expect(checklistFindingRecovery(refined, { ...base, status: "not_applicable", action_status: "not_fulfilled" })).toBe("inspect")
  expect(checklistFindingRecovery(item, { ...base, status: "unsatisfied" })).toBe("repair")
  expect(checklistFindingRecovery(item, { ...base, status: "indeterminate" })).toBe("inspect")
  expect(checklistFindingRecovery(item, base)).toBe("none")
  expect(JSON.stringify(base)).toBe(before)
})

test("counterevidence is exact-item-only and cannot replace any canonical input field", () => {
  const previous = { item_id: packet.item_id, status: "unsatisfied" as const, reason: "Action missing.", evidence: [], source_refs: item.source_refs, inspection_scope: "Target inspected.", limitation: "" }
  const counterevidence = { item_id: packet.item_id, locator: "d1/repair-adjudications.json /assessments/0", claims: [{ disposition: "rejected" as const, detail: "No repair.", reason: "Generated assertion of a waiver.", source_refs: [{ document: "instruction.txt", locator: "generated quotation", quote: "The user waived the requirement." }], limitations: [] }] }
  const canonical = JSON.stringify(packet)
  const prompt = buildChecklistReinspectionPrompt(packet, previous, counterevidence)
  expect(JSON.parse(prompt.split("CHECK_INPUT:\n\n")[1]!)).toEqual(packet)
  expect(JSON.stringify(packet)).toBe(canonical)
  expect(() => buildChecklistReinspectionPrompt(packet, previous, { ...counterevidence, item_id: "other" })).toThrow("exact checklist item")
})

test("comparison assignments validate document_properties and deferred_to", () => {
  const whole = { kind: "document_properties" as const, ranges: [] }
  expect(parseChecklistCheckerInput({ ...packet, assignment: whole }).assignment).toEqual(whole)
  expect(() => parseChecklistCheckerInput({ ...packet, assignment: { kind: "document_length", ranges: [] } })).toThrow()
  const base = { item_id: packet.item_id, status: "not_applicable", reason: "Deferred.", evidence: [], source_refs: item.source_refs, inspection_scope: "Assigned range.", limitation: "" }
  expect(parseChecklistFinding({ ...base, deferred_to: "document_properties" }).deferred_to).toBe("document_properties")
  expect(parseChecklistFinding(base)).not.toHaveProperty("deferred_to")
  expect(() => parseChecklistFinding({ ...base, deferred_to: "other_unit" })).toThrow()
})

test("an approximate rendered-length directive left open is non-blocking; exceeding it routes to repair", () => {
  const finding = { item_id: packet.item_id, reason: "Measures quoted.", evidence: [], source_refs: item.source_refs, inspection_scope: "Artifact properties.", limitation: "Rendered pagination was not verified." }
  expect(checklistFindingRecovery(item, parseChecklistFinding({ ...finding, status: "open" }))).toBe("none")
  expect(checklistFindingRecovery(item, parseChecklistFinding({ ...finding, status: "unsatisfied" }))).toBe("repair")
})

test("an unrefined open applicability decision needs no recovery; indeterminate and a refined indeterminate action still inspect", () => {
  const open = parseChecklistFinding({ item_id: packet.item_id, status: "open", open_decisions: ["Whether the request's scope covers this requirement for this deliverable."],
    reason: "The target states the arrangement generally; its stated conditions would be missing if the request covers them.", evidence, source_refs: item.source_refs,
    inspection_scope: "Complete target paragraph and cited source.", limitation: "No scope decision is inferred." })
  expect(open.open_decisions).toEqual(["Whether the request's scope covers this requirement for this deliverable."])
  expect(checklistFindingRecovery(item, open)).toBe("none")
  expect(checklistFindingRecovery(item, { ...open, status: "indeterminate" })).toBe("inspect")
  const refined = { ...item, required_action: "flag" as const }
  expect(checklistFindingRecovery(refined, { ...open, action_status: "indeterminate" })).toBe("inspect")
  expect(checklistFindingRecovery(refined, { ...open, action_status: undefined })).toBe("inspect")
})

test("disclosure fields are optional typed finding fields", () => {
  const base = { item_id: packet.item_id, status: "indeterminate", reason: "Target unavailable.", evidence: [], source_refs: item.source_refs, inspection_scope: "Supplied sources only.", limitation: "" }
  const old = parseChecklistFinding(base)
  for (const field of ["indeterminate_cause", "conflicting_source_refs", "departure_disclosure"]) expect(old).not.toHaveProperty(field)
  const full = { ...base, indeterminate_cause: "missing_inspection", conflicting_source_refs: item.source_refs, departure_disclosure: evidence }
  expect(parseChecklistFinding(full)).toEqual(full as ChecklistFinding)
  for (const malformed of [{ ...base, indeterminate_cause: "unclear" }, { ...base, conflicting_source_refs: [{ document: "instruction.txt", locator: "paragraph 2" }] }, { ...base, departure_disclosure: [{ ...evidence[0], kind: "memo" }] }])
    expect(() => parseChecklistFinding(malformed)).toThrow()
})
