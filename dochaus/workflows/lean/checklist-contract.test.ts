import { expect, test } from "bun:test"
import { checklistFindingRecovery, checklistTextMeasuresMeaning, type ChecklistFinding, buildChecklistReinspectionPrompt, buildChecklistCheckerPrompt, buildChecklistProducerPrompt, CHECKLIST_ITEM_SHAPE, parseChecklistCheckerInput, parseChecklistExtraction, parseChecklistFinding, parseChecklistItem, type ChecklistItem } from "./checklist-contract"

const item: ChecklistItem = {
  check: "Check the proposed notice arrangement remains subject to confirmation.",
  basis: "proposal",
  source_refs: [{ document: "instruction.txt", locator: "paragraph 2", quote: "  We propose advance notice; please confirm.\n" }],
  qualifications: ["The proposal has not been accepted."],
  target: { document: "agreement.docx", role: "operative agreement", treatment: "open_issue" },
}
const evidence = [{ document: "agreement.docx", locator: "p0123456789ab", text: "[open item: confirm notice arrangement]", kind: "body" as const }]
const packet = { item_id: "source-1-item-1", item, current_evidence: evidence, settled_instructions: [], retrieval: "Use read-record at current.json /paragraphs/0 and targeted lookup for necessary context." }

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


test("producer and checker share treatment meanings without promoting source authority", () => {
  const producer = buildChecklistProducerPrompt({ source: { document: "instruction.txt", role: "client instructions" }, passages: item.source_refs, instructions: [], targets: [{ document: "agreement.docx", role: "operative agreement" }], retrieval: "Use targeted lookup." })
  const checker = buildChecklistCheckerPrompt(packet)
  const definitions = producer.split("\n\n").find(paragraph => paragraph.startsWith("Treatment describes presentation"))!
  expect(definitions).toBeDefined()
  expect(checker.split("\n\n")).toContain(definitions)
  expect(definitions).toContain("advice means explanatory or advisory prose, including memo commentary")
  expect(definitions).toContain("comment means a native document annotation")
  expect(definitions).toContain("Require native comments only when source instructions support that presentation")
  expect(definitions).toContain("preserve the ambiguity")
  for (const treatment of ["operative", "advice", "comment", "open_issue"] as const) {
    const observation = { ...item, basis: "observation" as const, target: { ...item.target, treatment } }
    const unchanged = parseChecklistCheckerInput({ ...packet, item: parseChecklistExtraction({ items: [observation] }).items[0] }).item
    expect(unchanged).toEqual(observation)
    expect(buildChecklistCheckerPrompt({ ...packet, item: observation })).toContain("an observation alone imposes no revision or disclosure duty")
  }
})


test("substantive reinspection keeps the item unchanged and prior assessment untrusted", () => {
  const original = { item_id: packet.item_id, status: "unsatisfied" as const, reason: "Proposal was not adopted.", evidence: [], source_refs: item.source_refs, inspection_scope: "One paragraph.", limitation: "Related advice not inspected." }
  const prompt = buildChecklistReinspectionPrompt(packet, original)
  expect(prompt).toContain("may be right or wrong")
  expect(prompt).toContain("A proposal does not require adoption")
  expect(prompt).toContain("Perform targeted additional inspection")
  expect(prompt).toContain("preserve indeterminate")
  expect(prompt).toContain("Previous finding (unverified assessment, not source authority)")
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
  expect(prompt).toContain("Independently verify its source quotations and reasoning")
  expect(prompt).toContain("an inapplicable limb does not excuse an applicable unfulfilled action")
  expect(() => buildChecklistReinspectionPrompt(packet, previous, { ...counterevidence, item_id: "other" })).toThrow("exact checklist item")
})

test("comparison assignments include one whole-target document-properties kind with its own prompt branch", () => {
  const whole = { kind: "document_properties" as const, ranges: [] }
  expect(parseChecklistCheckerInput({ ...packet, assignment: whole }).assignment).toEqual(whole)
  expect(() => parseChecklistCheckerInput({ ...packet, assignment: { kind: "document_length", ranges: [] } })).toThrow()
  const prompt = buildChecklistCheckerPrompt({ ...packet, assignment: whole })
  expect(prompt).toContain("other assignments cover the remaining paragraphs, process records and document-wide properties")
  expect(prompt).toContain("Decide ONLY document-wide properties of the named target from its supplied exact artifact properties, including accepted_body_text_measures")
  expect(prompt).toContain("Report not_applicable when the parent states no document-wide property")
  expect(prompt).not.toContain("Only the separate document-properties assignment decides it")
  for (const kind of ["target_body", "process_records"] as const) {
    const range = buildChecklistCheckerPrompt({ ...packet, assignment: { kind, ranges: [] } })
    expect(range).toContain("Only the separate document-properties assignment decides it. If the parent concerns only such a property, report not_applicable for this component rather than indeterminate")
    expect(range).not.toContain("Decide ONLY document-wide properties")
  }
  expect(buildChecklistCheckerPrompt(packet)).not.toContain("document-properties assignment")
  expect(prompt).not.toContain("deferred_to")
  expect(buildChecklistCheckerPrompt(packet)).not.toContain("deferred_to")
  expect(buildChecklistCheckerPrompt({ ...packet, assignment: { kind: "target_body", ranges: [] } })).toContain('Whenever this component defers such a property, also return "deferred_to":"document_properties"; otherwise omit that field.')
  const base = { item_id: packet.item_id, status: "not_applicable", reason: "Deferred.", evidence: [], source_refs: item.source_refs, inspection_scope: "Assigned range.", limitation: "" }
  expect(parseChecklistFinding({ ...base, deferred_to: "document_properties" }).deferred_to).toBe("document_properties")
  expect(parseChecklistFinding(base)).not.toHaveProperty("deferred_to")
  expect(() => parseChecklistFinding({ ...base, deferred_to: "other_unit" })).toThrow()
})

test("text measures can disprove but never establish a rendered-length directive", () => {
  const whole = buildChecklistCheckerPrompt({ ...packet, assignment: { kind: "document_properties", ranges: [] } })
  const instructions = whole.split("CHECK_INPUT:")[0]!
  expect(instructions).toContain(checklistTextMeasuresMeaning)
  expect(checklistTextMeasuresMeaning).toContain("Never report such a directive satisfied from text measures: report unsatisfied when the measures clearly exceed it; otherwise report an approximate directive open, quoting the exact measures and stating that rendered pagination was not verified")
  expect(checklistTextMeasuresMeaning).toContain("a strict page or layout limit that the measures cannot decide remains indeterminate")
  expect(instructions).toContain("Any other property they cannot establish remains indeterminate.")
  for (const kind of ["target_body", "process_records"] as const)
    expect(buildChecklistCheckerPrompt({ ...packet, assignment: { kind, ranges: [] } })).not.toContain(checklistTextMeasuresMeaning)
  // An approximate rendered-length directive left open is non-blocking; exceeding it routes to repair.
  const finding = { item_id: packet.item_id, reason: "Measures quoted.", evidence: [], source_refs: item.source_refs, inspection_scope: "Artifact properties.", limitation: "Rendered pagination was not verified." }
  expect(checklistFindingRecovery(item, parseChecklistFinding({ ...finding, status: "open" }))).toBe("none")
  expect(checklistFindingRecovery(item, parseChecklistFinding({ ...finding, status: "unsatisfied" }))).toBe("repair")
})

// Applicability left unsettled by the request's scope after inspection is a
// decision for the lawyer, not missing evidence; directed requirements stay decided.
test("an inspected scope-only applicability question is an open decision in plain, component and reinspection prompts", () => {
  const rule = "If the needed evidence was inspected and the only unresolved question is whether the request's scope covers this requirement for this deliverable, report open, name that decision in open_decisions, and state what the target contains and which elements would be missing. Use indeterminate only when evidence or inspection is missing."
  const directed = "A requirement directed by the request or a current human instruction is applicable: decide its compliance."
  const component = "Every component applies the applicability rule below exactly as a whole-item check does. Content in unassigned ranges or records is neither evidence nor a gap for this component"
  const binding = "A term binding the parties in a source is source content, not by itself a directive for the deliverable to cover; unless the request or a current human instruction directs that coverage, whether the deliverable must cover it is an applicability question."
  const previous = { item_id: packet.item_id, status: "indeterminate" as const, reason: "Applicability was not established.", evidence: [], source_refs: item.source_refs, inspection_scope: "Assigned range.", limitation: "" }
  const plain = buildChecklistCheckerPrompt(packet)
  const components = (["target_body", "process_records", "document_properties"] as const).map(kind => buildChecklistCheckerPrompt({ ...packet, assignment: { kind, ranges: [] } }))
  const reinspections = [packet, { ...packet, assignment: { kind: "target_body" as const, ranges: [] } }].map(input => buildChecklistReinspectionPrompt(input, previous))
  for (const prompt of [plain, ...components, ...reinspections]) {
    const instructions = prompt.split("CHECK_INPUT:")[0]!
    expect(instructions).toContain(rule)
    expect(instructions).toContain(directed)
    expect(instructions).toContain(binding)
    expect(instructions).toContain("open preserves a substantive unresolved fact or decision, or an unsettled applicability decision")
    expect(instructions).not.toContain("use indeterminate when that relationship cannot be established")
  }
  expect(plain).not.toContain(component)
  for (const prompt of [...components, reinspections[1]!]) expect(prompt).toContain(component)
  expect(buildChecklistProducerPrompt({ source: { document: "instruction.txt", role: "client instructions" }, passages: item.source_refs, instructions: [], targets: [{ document: "agreement.docx", role: "operative agreement" }], retrieval: "Use targeted lookup." })).toContain(binding)
  // Every finding may name its open decisions; only a refined item also reports action status.
  expect(plain).toContain('"limitation":"...","open_decisions":[]}')
  expect(plain).not.toContain('"action_status"')
  expect(buildChecklistCheckerPrompt({ ...packet, item: { ...item, required_action: "propose" } })).toContain('"limitation":"...","action_status":"fulfilled|not_fulfilled|indeterminate|not_applicable","open_decisions":[]}')
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
