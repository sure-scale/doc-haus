import { afterEach, describe, expect, test } from "bun:test"
import { rmSync } from "node:fs"
import type { ChecklistCheckerInput, ChecklistCurrentEvidence, ChecklistFinding, ChecklistItem, ChecklistProducerInput } from "./checklist-contract"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { finalGate } from "./gate"
import { LEAN_COVERAGE_SCHEMA, coverageItems, coveragePlanHash, coverageScopeQuestion, type CoverageAct, type LeanCoverage } from "./package-coverage"
import { PACKAGE_SCOPE_QUESTION, packageScopeQuestionID, scopeQuestionShows } from "./package-scope"
import { leanSteps, readLeanPlan, writeLeanPlan, type PlannedDocument } from "./plan"
import { enqueueLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { answerQuestion, leanFile, readLeanJson, readLeanState, readQuestions, replaceStageQuestions, updateLeanStage, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

// An instruction (a directive item) the checklist check left open may be
// carried by no planned document. Disclosure alone cannot settle that, so the
// gate, as a backstop to the plan-side coverage check, asks the lawyer about
// all of them in one package-scope question after the walk. It pauses nothing:
// the run completes as a review draft and each instruction stays disclosed
// until a later check settles it. The question is keyed on the verbatim
// instructions. An answered package-scope question settles the instructions it
// showed (whole, or as a longer or shorter span of the same instruction), the
// plan step's only when it showed every act the coverage record has them
// directing that it asks about, each for its own parties; one more is asked
// about any instruction no answered question showed, but none is added while
// another stage's is open. A disclosure refers
// to a question only when that question showed its instruction. An instruction a
// planned document carries in full for the parties who must take it, left open
// only for a fact its body marks as missing, is a fact gap rather than a
// question of scope: it is disclosed and not asked about. Only the model is
// scripted.

afterEach(cleanupMatters)

const ref = (quote: string) => ({ document: "request", locator: "request", quote })
const requirement = (check: string, quote: string, basis: ChecklistItem["basis"] = "directive"): ChecklistItem =>
  ({ check, basis, source_refs: [ref(quote)], qualifications: [], target: { document: "d1", role: "Supply note", treatment: "operative" } })
const PERIOD = requirement("The note states the delivery period.", "Deliver within thirty days of the order.")
const PAYMENT = requirement("The note states when payment falls due.", "Payment is due on delivery.")
const ADDRESS = requirement("The note names the delivery address.", "Deliver to the buyer's warehouse.")
const ITEMS = [{ id: "c1", item: PERIOD }, { id: "c2", item: PAYMENT }]
const finding = (item_id: string, status: ChecklistFinding["status"], extra: Partial<ChecklistFinding> = {}): ChecklistFinding => ({
  item_id, status, reason: `Focused result ${status}.`, evidence: [], source_refs: [], inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", ...extra,
})
const SCOPE = "Whether this deliverable is meant to carry it"
const left = (item_id: string) => finding(item_id, "open", { open_decisions: [SCOPE] })
const CLOSING = "the open decisions above (supply any missing fact or approval they name). If the package should carry it, say which planned document carries it; otherwise say it is outside this request."

// The checklist's saved output, as the gate reads it.
function checklistMatter(findings: ChecklistFinding[], items: { id: string; item: ChecklistItem }[] = ITEMS) {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items }] }))
  writeFindings(directory, findings)
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  return directory
}

/** A current coverage record of `acts` for the matter's saved plan, as the
 * plan step leaves it, and drafting notes naming the items it gives. */
function recordCoverage(directory: string, acts: CoverageAct[]) {
  const plan = readLeanPlan(directory)!
  writeLeanText(directory, "coverage.json", JSON.stringify({ schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "uncovered", acts, issues: [] }))
  writeLeanText(directory, "drafting-notes.md", coverageItems(directory).map(item => item.id).join("\n"))
  return directory
}

function writeFindings(directory: string, findings: ChecklistFinding[]) {
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: findings.map(row => row.item_id),
    checks: findings.map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))
}

const leftOpen = (gate: ReturnType<typeof finalGate>) => gate.disclosed.filter(item => item.kind === "checklist_open")

test("an instruction left open is asked once, in one package-scope question, and its disclosure refers to it", () => {
  const directory = checklistMatter([left("c1"), finding("c2", "satisfied")])
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate.questions).toHaveLength(1)
  const question = gate.questions![0]!
  expect(question.id.startsWith(PACKAGE_SCOPE_QUESTION)).toBe(true)
  expect(question).toMatchObject({ stage: "gate", status: "open" })
  expect(question).not.toHaveProperty("pid")
  expect(question).not.toHaveProperty("document")
  for (const text of ["\nThe note states the delivery period.", 'request "Deliver within thirty days of the order."', `Open decisions: ${SCOPE}.`, `\nAnswer ${CLOSING}`])
    expect(question.question).toContain(text)
  expect(question.question).not.toMatch(/\bc[12]\b/)
  expect(question.question).not.toContain(PAYMENT.check)
  expect(question.question).not.toContain("For each")
  expect(question.context).toBe("Until a later check settles this instruction, it is disclosed with the review draft as left open; delivery is not held for it. " +
    "Your answer re-runs the stages that use it, including the checklist check; the document plan is kept as saved, and a drafted document changes only by amendment. " +
    'A document the plan does not have is not added: to add one, stop this draft with "Stop draft" if it has not finished, then start a new one, which plans the package from the start.')
  // Still disclosed with its own reason, now referring to the question.
  expect(leftOpen(gate)).toEqual([expect.objectContaining({ status: "disclosed", ref: { checklist_item: "c1", document: "d1", question_id: question.id },
    detail: expect.stringContaining("Checklist requirement left open: The note states the delivery period.") })])
  // Deterministic: the same findings give the same question.
  expect(finalGate(directory).questions).toEqual([question])
  expect(gate).not.toHaveProperty("question")
})

test("several instructions left open share one question, keyed by their verbatim instructions, not by item ids or order", () => {
  const directory = checklistMatter([left("c1"), finding("c2", "open")])
  const [question, ...rest] = finalGate(directory).questions!
  expect(rest).toEqual([])
  expect(question!.question).toContain(`\nThe note states the delivery period. Sources: request "Deliver within thirty days of the order." Open decisions: ${SCOPE}.\n`)
  // A finding that records no open decision asks the lawyer to name what is missing.
  expect(question!.question).toContain('\nThe note states when payment falls due. Sources: request "Payment is due on delivery." No open decision recorded; state what is missing.\n')
  expect(question!.question).toContain(`\nFor each, answer ${CLOSING}`)
  expect(question!.context).toContain("Until a later check settles these instructions, they are disclosed with the review draft as left open; delivery is not held for them.")
  expect(leftOpen(finalGate(directory)).map(item => item.ref.question_id)).toEqual([question!.id, question!.id])
  // Re-extraction renumbers the items, reorders them and rewords the checks;
  // the same instructions, quoted verbatim, ask the same question.
  const renumbered = checklistMatter([finding("s-7", "open"), left("s-3")], [
    { id: "s-7", item: { ...PAYMENT, check: "The note gives the payment date." } },
    { id: "s-3", item: { ...PERIOD, source_refs: [ref("Deliver within  thirty days\nof the order. ")] } },
  ])
  expect(finalGate(renumbered).questions!.map(row => row.id)).toEqual([question!.id])
  // A different set of instructions is a different question.
  expect(finalGate(checklistMatter([left("c1"), finding("c2", "satisfied")])).questions!.map(row => row.id)).not.toEqual([question!.id])
})

test("an answered package-scope question settles only the instructions it showed: a new open instruction is asked about alone; each stays disclosed, referring to an answered question only when it showed that instruction", () => {
  const directory = checklistMatter([left("c1"), finding("c2", "satisfied")])
  const [question] = finalGate(directory).questions!
  replaceStageQuestions(directory, "gate", null, [question!])
  answerQuestion(directory, question!.id, { text: "The supply note carries the delivery period." })
  // A second instruction is now left open as well: the gate asks about it only.
  writeFindings(directory, [left("c1"), left("c2")])
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  const [second, ...rest] = gate.questions!
  expect(rest).toEqual([])
  expect(second!.id).not.toBe(question!.id)
  expect(second!.question).toContain("\nThe note states when payment falls due.")
  expect(second!.question).not.toContain(PERIOD.check)
  expect(second!.question).not.toContain("For each")
  // The answered question asked about c1 only, so only c1 refers to it; c2 refers to the new one.
  expect(leftOpen(gate).map(item => item.ref)).toEqual([{ checklist_item: "c1", document: "d1", question_id: question!.id }, { checklist_item: "c2", document: "d1", question_id: second!.id }])
  replaceStageQuestions(directory, "gate", null, gate.questions!)
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question!.id, status: "answered" }), expect.objectContaining({ id: second!.id, status: "open" })])
  // Answered, both settle their instructions: no third question, and a re-run asks nothing.
  answerQuestion(directory, second!.id, { text: "The supply note carries the payment date." })
  const settled = finalGate(directory)
  expect(settled.questions).toBeUndefined()
  expect(leftOpen(settled).map(item => item.ref.question_id)).toEqual([question!.id, second!.id])
  // A package-scope question another stage asked and the lawyer answered
  // showed the payment instruction for the Buyer, the party of the one act the
  // coverage record has it directing, so only c2 refers to it and the gate asks about c1.
  const payment: CoverageAct = { act: "pay on delivery", source: "request", quote: "Payment is due on delivery.", parties: ["the Buyer"], document: null }
  const planned = recordCoverage(checklistMatter([left("c1"), left("c2")]), [payment])
  writeQuestions(planned, [{ ...coverageScopeQuestion([payment], [])!, id: `${PACKAGE_SCOPE_QUESTION}planned`, status: "answered", answer: "The supply note." }])
  const asked = finalGate(planned)
  expect(asked.questions).toEqual([expect.objectContaining({ stage: "gate", question: expect.stringContaining("\nThe note states the delivery period.") })])
  expect(asked.questions![0]!.question).not.toContain(PAYMENT.check)
  expect(leftOpen(asked).map(item => item.ref.question_id)).toEqual([asked.questions![0]!.id, `${PACKAGE_SCOPE_QUESTION}planned`])
  // A question that showed no instruction is referred to by none, and every instruction is asked.
  writeQuestions(planned, [{ id: `${PACKAGE_SCOPE_QUESTION}planned`, stage: "intake", question: "Which document carries each instruction?", context: "", status: "answered", answer: "The supply note." }])
  const none = finalGate(planned)
  expect(none.questions).toHaveLength(1)
  expect(none.questions![0]!.question).toContain(PERIOD.check)
  expect(none.questions![0]!.question).toContain(PAYMENT.check)
  expect(leftOpen(none).map(item => item.ref.question_id)).toEqual([none.questions![0]!.id, none.questions![0]!.id])
})

test("a reworded span of an instruction an answered question showed is not asked again, whether the new quote is longer or shorter", () => {
  const SENTENCE = "Deliver within thirty days of the order, and invoice on the day of delivery."
  const CLAUSE = "Deliver within thirty days of the order"
  const supplied = (text: string): CoverageAct => ({ act: "deliver the goods", source: "request", quote: text, parties: ["the Supplier"], document: null })
  const answeredOn = (text: string): Question => ({ ...coverageScopeQuestion([supplied(text)], [])!,
    id: `${PACKAGE_SCOPE_QUESTION}planned`, status: "answered", answer: "The supply note." })
  for (const [shown, quoted] of [[SENTENCE, CLAUSE], [CLAUSE, SENTENCE], [SENTENCE, "Deliver within  thirty days of the order, and invoice\non the day of delivery."]]) {
    const directory = recordCoverage(checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note states the delivery period.", quoted!) }]), [supplied(shown!)])
    writeQuestions(directory, [answeredOn(shown!)])
    const gate = finalGate(directory)
    expect(gate.questions).toBeUndefined()
    expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([`${PACKAGE_SCOPE_QUESTION}planned`])
  }
  // A different instruction, or a phrase too short to identify one, is not shown.
  const other = recordCoverage(checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note states the delivery period.", "Deliver within thirty days of the order.") }]),
    [supplied("Deliver to the buyer's warehouse")])
  writeQuestions(other, [answeredOn("Deliver to the buyer's warehouse")])
  expect(finalGate(other).questions).toHaveLength(1)
  const short = recordCoverage(checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note states the delivery period.", "Deliver within thirty days of the order.") }]),
    [supplied("Deliver within")])
  writeQuestions(short, [answeredOn("Deliver within")])
  expect(finalGate(short).questions).toHaveLength(1)
})

test("a scope question that cut a long instruction to an excerpt shows that instruction, and no other", () => {
  const long = `The supplier shall deliver the goods (the "Goods") to the buyer's warehouse ${"within the agreed window ".repeat(10)}after the order.`
  const [question] = finalGate(checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note states the delivery window.", long) }])).questions!
  expect(question!.question).not.toContain(long)
  expect(scopeQuestionShows(question!, [long.replace(/ /g, "\n")])).toBe(true)
  expect(scopeQuestionShows(question!, [long.replace("warehouse", "depot")])).toBe(false)
  // A shorter span the excerpt holds is shown; one the excerpt does not hold is not.
  expect(scopeQuestionShows(question!, [long.slice(0, 100)])).toBe(true)
  expect(scopeQuestionShows(question!, [long.slice(0, 100).replaceAll("the", "a")])).toBe(false)
  // An excerpt at least SHORT_QUOTE long shows the instruction; a shorter
  // fragment shows none, whether cut from its opening or its middle.
  const shows = (cut: string) => scopeQuestionShows({ id: `${PACKAGE_SCOPE_QUESTION}cut`, stage: "gate", context: "", status: "open",
    question: `The note states when payment falls due. Sources: request "${cut}" No open decision recorded; state what is missing.` }, ["Payment is due on delivery."])
  expect(shows("Payment is due on deliv…")).toBe(true)
  expect(shows("Payment is due…")).toBe(false)
  expect(shows("is due on delivery…")).toBe(false)
})

test("while another stage's package-scope question is open the gate adds none; that question stays, and only an instruction it showed refers to it", () => {
  const directory = recordCoverage(checklistMatter([left("c1"), left("c2")]),
    [{ act: "deliver the goods", source: "request", quote: "Deliver within thirty days of the order.", parties: ["the Supplier"], document: null }])
  const PLANNED = `${PACKAGE_SCOPE_QUESTION}planned`
  const open = { id: PLANNED, stage: "draft" as const, document: "d1" as const, question: 'The document plan leaves this instructed act without a document executed by the parties who must take it: "deliver the goods", directed by "Deliver within thirty days of the order." [source: request]; to be taken by "the Supplier"; no planned document carries it.', context: "", status: "open" as const }
  writeQuestions(directory, [open])
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate.questions).toBeUndefined()
  expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([PLANNED, undefined])
  replaceStageQuestions(directory, "gate", null, gate.questions ?? [])
  expect(readQuestions(directory)).toEqual([open])
  // Once the gate's own question is open, the gate still asks it again: its
  // questions replace the ones it asked before.
  const own = checklistMatter([left("c1"), finding("c2", "satisfied")])
  const [first] = finalGate(own).questions!
  replaceStageQuestions(own, "gate", null, [first!])
  expect(finalGate(own).questions).toEqual([first])
})

test("a proposal, observation or open question left open is disclosed and asks nothing; so is a settled instruction with a decision open", () => {
  for (const basis of ["proposal", "observation", "open_question"] as const) {
    const gate = finalGate(checklistMatter([left("c1"), finding("c2", "satisfied")], [{ id: "c1", item: { ...PERIOD, basis } }, { id: "c2", item: PAYMENT }]))
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.questions).toBeUndefined()
    expect(leftOpen(gate)).toEqual([expect.objectContaining({ ref: { checklist_item: "c1", document: "d1" } })])
  }
  const settled = finalGate(checklistMatter([finding("c1", "satisfied", { open_decisions: ["Whether the client approves the period"] }), finding("c2", "satisfied")]))
  expect(settled).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(settled.questions).toBeUndefined()
  expect(leftOpen(settled)).toEqual([expect.objectContaining({ ref: { checklist_item: "c1", document: "d1" } })])
})

test("an instruction left open that a source conflict names is asked only in the conflict question; the rest share the scope question", () => {
  // c1 conflicts with c2's source; c2 and c3 are instructions left open.
  const items = [...ITEMS, { id: "c3", item: ADDRESS }]
  const directory = checklistMatter([finding("c1", "unsatisfied", { conflicting_source_refs: [ref("Payment is due on delivery.")] }), left("c2"), left("c3")], items)
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  const [conflict, scope, ...rest] = gate.questions!
  expect(rest).toEqual([])
  expect(conflict!.id).toMatch(/^checklist-conflict-/)
  expect(conflict!.question).toContain(`First: ${PERIOD.check} Sources: request "Deliver within thirty days of the order." Second: ${PAYMENT.check}`)
  expect(scope!.id.startsWith(PACKAGE_SCOPE_QUESTION)).toBe(true)
  expect(scope!.question).toContain("\nThe note names the delivery address.")
  expect(scope!.question).not.toContain(PAYMENT.check)
  // No item is asked about twice.
  for (const item of [PERIOD, PAYMENT, ADDRESS]) expect(gate.questions!.filter(row => row.question.includes(item.check))).toHaveLength(1)
  expect(gate.disclosed.find(item => item.kind === "checklist_unresolved")!.ref).toMatchObject({ checklist_item: "c1", question_id: conflict!.id })
  expect(leftOpen(gate).map(item => item.ref)).toEqual([{ checklist_item: "c2", document: "d1", question_id: conflict!.id }, { checklist_item: "c3", document: "d1", question_id: scope!.id }])
  // Recorded together, each replaces nothing of the other; answering one
  // leaves the other open and neither is asked again.
  replaceStageQuestions(directory, "gate", null, gate.questions!)
  answerQuestion(directory, scope!.id, { text: "The package carries it in the supply note." })
  replaceStageQuestions(directory, "gate", null, finalGate(directory).questions!)
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: scope!.id, status: "answered" }), expect.objectContaining({ id: conflict!.id, status: "open" })])
  // When the only instruction left open is a conflict's, only the conflict is asked.
  const only = finalGate(checklistMatter([finding("c1", "unsatisfied", { conflicting_source_refs: [ref("Payment is due on delivery.")] }), left("c2"), finding("c3", "satisfied")], items))
  expect(only.questions!.map(row => row.id)).toEqual([conflict!.id])
  expect(leftOpen(only).map(item => item.ref.question_id)).toEqual([conflict!.id])
})

test("an instruction left open against a passage it conflicts with is asked in the scope question, never met: the package stays a review draft", () => {
  const against = finding("c1", "open", { open_decisions: [SCOPE], conflicting_source_refs: [ref("Deliver within twenty days of the order.")] })
  const both = finalGate(checklistMatter([against, left("c2")]))
  expect(both).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(both.questions).toHaveLength(1)
  expect(both.questions![0]!.id.startsWith(PACKAGE_SCOPE_QUESTION)).toBe(true)
  expect(leftOpen(both).map(item => item.ref.question_id)).toEqual([both.questions![0]!.id, both.questions![0]!.id])
  const alone = finalGate(checklistMatter([against, finding("c2", "satisfied")]))
  expect(alone).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(leftOpen(alone)).toEqual([expect.objectContaining({ ref: { checklist_item: "c1", document: "d1", question_id: alone.questions![0]!.id } })])
})

test("no scope question is asked while the checklist check is not done, even with an instruction left open in its output on disk", () => {
  const open = [left("c1"), finding("c2", "satisfied")]
  expect(finalGate(checklistMatter(open)).questions).toHaveLength(1)
  // Output kept from an earlier pass is not this run's verdict until the stage is done.
  for (const status of ["pending", "running", "failed", "blocked"] as const) {
    const directory = checklistMatter(open)
    updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status })
    const gate = finalGate(directory)
    expect(gate.questions).toBeUndefined()
    expect(leftOpen(gate)).toEqual([])
  }
})

test("no scope question is asked when the checklist check's output is missing or incomplete; the gate blocks on that output instead", () => {
  const open = [left("c1"), finding("c2", "satisfied")]
  const missing = ["checklist/extraction.json", "checklist/checks.json", "checklist/report.md"].map(file => {
    const directory = checklistMatter(open)
    rmSync(leanFile(directory, file))
    return directory
  })
  // c2's check is missing from the saved checks.
  const incomplete = checklistMatter(open)
  writeFindings(incomplete, [left("c1")])
  for (const directory of [...missing, incomplete]) {
    const gate = finalGate(directory)
    expect(gate).toMatchObject({ status: "blocked", delivery: null, blocking: [{ stage: "checklist-check", kind: "omission", status: "open" }] })
    expect(gate.questions).toBeUndefined()
    expect(leftOpen(gate)).toEqual([])
  }
})

test("an instruction left open that already refers to the conflict question adds nothing to the scope question's text or id", () => {
  // c1 conflicts with c2's source; c2 and c3 are instructions left open.
  const items = [...ITEMS, { id: "c3", item: ADDRESS }]
  const gate = finalGate(checklistMatter([finding("c1", "unsatisfied", { conflicting_source_refs: [ref("Payment is due on delivery.")] }), left("c2"), left("c3")], items))
  const [conflict, scope, ...rest] = gate.questions!
  expect(rest).toEqual([])
  expect(conflict!.id).toMatch(/^checklist-conflict-/)
  // The scope question is exactly the one c3 alone would raise: c2's quote is
  // neither listed nor hashed into its id.
  const alone = finalGate(checklistMatter([finding("c1", "satisfied"), finding("c2", "satisfied"), left("c3")], items)).questions
  expect([scope]).toStrictEqual(alone!)
  expect(scope!.question).not.toContain("Payment is due on delivery.")
  expect(leftOpen(gate).map(item => [item.ref.checklist_item, item.ref.question_id])).toEqual([["c2", conflict!.id], ["c3", scope!.id]])
})

// Returned while blocked: the scope question asks about a disclosed finding
// that the block neither settles nor depends on, so withholding it would only
// make the lawyer wait for a second stop to answer it; the runner records the
// gate's questions on a blocked result as on a pass.
test("a gate blocked by another open item still returns the scope question, unchanged, so the lawyer can answer it with the block", () => {
  const open = [left("c1"), finding("c2", "satisfied")]
  const unblocked = finalGate(checklistMatter(open)).questions!
  expect(unblocked).toHaveLength(1)
  const directory = checklistMatter(open)
  writeLeanText(directory, "package-review.json", JSON.stringify({ status: "changes-required",
    findings: [{ target: "memo", detail: "The drafting notes omit the delivery timetable." }], limitations: [], delivery: [] }))
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "blocked", delivery: null, blocking: [{ stage: "package-review", kind: "omission", status: "open" }] })
  expect(gate.questions).toStrictEqual(unblocked)
  expect(leftOpen(gate).map(item => item.ref)).toEqual([{ checklist_item: "c1", document: "d1", question_id: unblocked[0]!.id }])
})

test("every open decision recorded on an instruction is put to the lawyer, trimmed, in its recorded order", () => {
  // Recorded order is not alphabetical, and one decision carries stray spaces.
  const decisions = ["Whether a side letter carries the period", "  Which party bears late delivery costs ", "Approval of the period by the client"]
  const gate = finalGate(checklistMatter([finding("c1", "open", { open_decisions: decisions }), finding("c2", "satisfied")]))
  const recorded = "Open decisions: Whether a side letter carries the period; Which party bears late delivery costs; Approval of the period by the client."
  expect(gate.questions!.map(row => row.question)).toEqual([expect.stringContaining(`\nThe note states the delivery period. Sources: request "Deliver within thirty days of the order." ${recorded}\nAnswer the open decisions above`)])
  expect(leftOpen(gate).map(item => item.detail)).toEqual([expect.stringContaining(recorded)])
})

test("the scope question on one instruction, and its disclosure, read exactly as recorded", () => {
  const gate = finalGate(checklistMatter([left("c1"), finding("c2", "satisfied")]))
  expect(gate.questions!.map(row => row.question)).toStrictEqual(["The checklist check left this instruction open, so the package may not carry it:\n" +
    'The note states the delivery period. Sources: request "Deliver within thirty days of the order." Open decisions: Whether this deliverable is meant to carry it.\n' +
    "Answer the open decisions above (supply any missing fact or approval they name). If the package should carry it, say which planned document carries it; otherwise say it is outside this request."])
  expect(leftOpen(gate).map(item => item.detail)).toStrictEqual(["Checklist requirement left open: The note states the delivery period. Finding: Focused result open. " +
    "Limitation: Comments not inspected. Open decisions: Whether this deliverable is meant to carry it."])
})

test("the scope question names no checklist item id or planned-document id, puts each instruction on its own line, and is still read once answered", () => {
  const titled = requirement("d1 states the delivery period.", "Deliver within thirty days of the order.")
  const directory = checklistMatter([finding("source-0a1b2c3d4e5f-1", "open", { open_decisions: ["Whether d1 is meant to carry it"] }), left("source-0a1b2c3d4e5f-2")],
    [{ id: "source-0a1b2c3d4e5f-1", item: titled }, { id: "source-0a1b2c3d4e5f-2", item: PAYMENT }])
  const [question, ...rest] = finalGate(directory).questions!
  expect(rest).toEqual([])
  expect(question!.question).not.toMatch(/source-0a1b2c3d4e5f|\bd1\b/)
  expect(question!.question.split("\n")).toStrictEqual([
    "The checklist check left these instructions open, so the package may not carry them:",
    'The Supply note states the delivery period. Sources: request "Deliver within thirty days of the order." Open decisions: Whether the Supply note is meant to carry it.',
    `The note states when payment falls due. Sources: request "Payment is due on delivery." Open decisions: ${SCOPE}.`,
    `For each, answer ${CLOSING}`,
  ])
  // Answered, it settles both instructions: nothing more is asked, and each disclosure refers to it.
  replaceStageQuestions(directory, "gate", null, [question!])
  answerQuestion(directory, question!.id, { text: "It is outside this request." })
  const gate = finalGate(directory)
  expect(gate.questions).toBeUndefined()
  expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([question!.id, question!.id])
})

// The carried case needs both signals: the saved plan's coverage record has an
// act for every quote the instruction cites in the item's document, executed by
// the parties who must take it, with no act those quotes direct left uncovered,
// and the finding's own body evidence from that document shows an open item.
// Either alone leaves the question asked.
const SUPPLIER = "the Supplier"
const NOTE: Omit<PlannedDocument, "id"> = { title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }], executed_by: [SUPPLIER] }
const SCHEDULE: Omit<PlannedDocument, "id"> = { title: "Delivery schedule", purpose: "Delivery schedule", parts: [{ heading: "Dates", scope: "Delivery dates" }], executed_by: [SUPPLIER] }
const PERIOD_QUOTE = "Deliver within thirty days of the order."
const DELIVER: CoverageAct = { act: "deliver the goods", source: "request", quote: PERIOD_QUOTE, parties: [SUPPLIER], document: "d1" }
const body = (text: string, document = "d1"): ChecklistCurrentEvidence => ({ document, locator: "native PIDs p1", text, kind: "body" })
const MARKED = body("The Supplier shall deliver the goods within thirty days of [insert order date].")
const gap = (item_id: string, evidence = [MARKED]) => finding(item_id, "open", { evidence, open_decisions: ["The order date"] })

/** A checklist matter whose saved plan has a coverage record of that plan. */
function carried(findings: ChecklistFinding[], acts: CoverageAct[] = [DELIVER], options: { items?: typeof ITEMS; documents?: Omit<PlannedDocument, "id">[]; status?: LeanCoverage["status"] } = {}) {
  const directory = checklistMatter(findings, options.items)
  const plan = writeLeanPlan(directory, options.documents ?? [NOTE])
  const coverage: LeanCoverage = { schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: options.status ?? "covered", acts, issues: [] }
  writeLeanText(directory, "coverage.json", JSON.stringify(coverage))
  return directory
}

/** The gate's question ids and the question each instruction left open refers to. */
function asks(directory: string) {
  const gate = finalGate(directory)
  return { questions: (gate.questions ?? []).map(row => row.id), refers: leftOpen(gate).map(item => item.ref.question_id) }
}
const NOT_ASKED = { questions: [], refers: [undefined] }
const asked = (...quotes: string[]) => {
  const id = packageScopeQuestionID(quotes.length ? quotes : [PERIOD_QUOTE])
  return { questions: [id], refers: [id] }
}

test("an instruction its planned document carries for its parties, left open for a fact the body marks as missing, is disclosed and not asked about", () => {
  const gate = finalGate(carried([gap("c1"), finding("c2", "satisfied")]))
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate.questions).toBeUndefined()
  expect(leftOpen(gate)).toStrictEqual([expect.objectContaining({ status: "disclosed", ref: { checklist_item: "c1", document: "d1" },
    detail: "Checklist requirement left open: The note states the delivery period. Finding: Focused result open. Limitation: Comments not inspected. Open decisions: The order date." })])
  // An open-item marker shows the gap as a placeholder does, and the evidence
  // may name the document by its file.
  const marker = body("The Supplier shall deliver within thirty days [open item: confirm the order date].", "supply-note.docx")
  expect(asks(carried([gap("c1", [marker]), finding("c2", "satisfied")]))).toEqual(NOT_ASKED)
})

test("an instruction left open is still asked about when its document's body evidence shows no open item", () => {
  expect(asks(carried([gap("c1"), finding("c2", "satisfied")], [DELIVER], { documents: [NOTE, SCHEDULE] }))).toEqual(NOT_ASKED)
  // An unmarked deferral paragraph, no evidence, a marker in another document's
  // body, or a marker only in a comment.
  for (const evidence of [[body("The delivery period is left to a separate schedule.")], [], [body(MARKED.text, "d2")], [{ ...MARKED, kind: "comment" as const }]])
    expect(asks(carried([gap("c1", evidence), finding("c2", "satisfied")], [DELIVER], { documents: [NOTE, SCHEDULE] }))).toEqual(asked())
})

test("an instruction left open is still asked about unless the saved plan's coverage record carries its act in its document, executed by the parties who must take it", () => {
  const findings = [gap("c1"), finding("c2", "satisfied")]
  expect(asks(carried(findings, [DELIVER], { documents: [NOTE, SCHEDULE] }))).toEqual(NOT_ASKED)
  // No planned document carries the act; another document does; or the
  // document that carries it is not executed by the party who must take it.
  for (const [act, status] of [[{ ...DELIVER, document: null }, "uncovered"], [{ ...DELIVER, document: "d2" }, "covered"], [{ ...DELIVER, parties: ["the Buyer"] }, "uncovered"]] as [CoverageAct, LeanCoverage["status"]][])
    expect(asks(carried(findings, [act], { documents: [NOTE, SCHEDULE], status }))).toEqual(asked())
  // A record of another plan, an unreadable record or none.
  const stale = carried(findings)
  writeLeanText(stale, "coverage.json", JSON.stringify({ ...(readLeanJson(stale, "coverage.json") as LeanCoverage), plan: "0".repeat(64) }))
  const unreadable = carried(findings)
  writeLeanText(unreadable, "coverage.json", "{")
  const missing = carried(findings)
  rmSync(leanFile(missing, "coverage.json"))
  for (const directory of [stale, unreadable, missing]) expect(asks(directory)).toEqual(asked())
})

test("an instruction citing several quotes is still asked about unless its document carries an act for each", () => {
  const ADDRESS_QUOTE = "Deliver to the buyer's warehouse."
  const items = [{ id: "c1", item: { ...PERIOD, source_refs: [ref(PERIOD_QUOTE), ref(ADDRESS_QUOTE)] } }, { id: "c2", item: PAYMENT }]
  const findings = [gap("c1"), finding("c2", "satisfied")]
  // The period is carried; the warehouse instruction matches no carried act.
  expect(asks(carried(findings, [DELIVER], { items }))).toEqual(asked(PERIOD_QUOTE, ADDRESS_QUOTE))
  expect(asks(carried(findings, [DELIVER, { ...DELIVER, act: "deliver to the warehouse", quote: ADDRESS_QUOTE }], { items }))).toEqual(NOT_ASKED)
})

test("an instruction is still asked about when any act its quote directs is left without a document executed by the parties who must take it", () => {
  const findings = [gap("c1"), finding("c2", "satisfied")]
  const consent = { ...DELIVER, act: "consent to the delivery period", parties: ["the Buyer"] }
  // The supply note carries delivery, but the Buyer's consent is in no planned
  // document, or in one the Buyer does not execute.
  for (const act of [{ ...consent, document: null }, consent])
    expect(asks(carried(findings, [DELIVER, act], { status: "uncovered" }))).toEqual(asked())
})

test("an act whose parties the record does not settle carries no instruction, whether or not its document declares who executes it", () => {
  const findings = [gap("c1"), finding("c2", "satisfied")]
  const { executed_by: _, ...UNDECLARED } = NOTE
  for (const documents of [[NOTE], [UNDECLARED]])
    expect(asks(carried(findings, [{ ...DELIVER, parties: [] }], { documents }))).toEqual(asked())
})

test("a bare underscore blank is an execution line, not an open item; a bracketed blank is", () => {
  const findings = (text: string) => [gap("c1", [body(text)]), finding("c2", "satisfied")]
  for (const text of ["The Supplier shall deliver the goods within thirty days of the order. Dated: ________",
    "Signed by the Supplier acting by: ______ Director", "The parties shall meet on ________ to agree the delivery dates."])
    expect(asks(carried(findings(text)))).toEqual(asked())
  for (const text of ["The Supplier shall deliver the goods within thirty days of [____].", "Delivery within thirty days of [_____ order date]. Dated: ________",
    "Delivery within thirty days of [ ______ ]."])
    expect(asks(carried(findings(text)))).toEqual(NOT_ASKED)
})

test("the coverage act joins the instruction by its source and by either quote containing the other, after whitespace and quotation marks are normalized", () => {
  const WAREHOUSE = "Deliver to the buyer's warehouse within thirty days of the order."
  const items = [{ id: "c1", item: requirement("The note names the delivery address.", WAREHOUSE) }]
  const excerpt = { ...DELIVER, quote: "the buyer’s\n  warehouse" }
  expect(asks(carried([gap("c1")], [excerpt], { items }))).toEqual(NOT_ASKED)
  expect(asks(carried([gap("c1")], [{ ...DELIVER, quote: `${WAREHOUSE} Payment is due on delivery.` }], { items }))).toEqual(NOT_ASKED)
  // Another source, or words that do not overlap the instruction's.
  for (const act of [{ ...excerpt, source: "Supply terms.docx" }, { ...DELIVER, quote: "Payment is due on delivery." }])
    expect(asks(carried([gap("c1")], [act], { items }))).toEqual(asked(WAREHOUSE))
})

test("with one instruction carried and another left out, the scope question lists and keys only the one left out", () => {
  const gate = finalGate(carried([gap("c1"), left("c2")]))
  const [question, ...rest] = gate.questions!
  expect(rest).toEqual([])
  expect(question!.id).toBe(packageScopeQuestionID(["Payment is due on delivery."]))
  expect(question!.question).toContain("\nThe note states when payment falls due.")
  for (const text of [PERIOD.check, PERIOD_QUOTE, "For each"]) expect(question!.question).not.toContain(text)
  expect(leftOpen(gate).map(item => item.ref)).toEqual([{ checklist_item: "c1", document: "d1" }, { checklist_item: "c2", document: "d1", question_id: question!.id }])
})

const ANSWER = "The supply note carries the delivery period; plan it there."

// Two instructions from the request; the check leaves the first open because
// whether the planned document carries it is undecided.
function instructionModel() {
  return scriptedModel({
    ...happyPath,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
      const item = (check: string) => ({ check, basis: "directive", source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" } })
      return JSON.stringify({ items: input.source.document === "request" ? [item("The note states the delivery period."), item("The requested supply note is drafted.")] : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      const open = input.item.check.includes("delivery period")
      return JSON.stringify({ item_id: input.item_id, status: open ? "open" : "satisfied", reason: `Focused result for ${input.item_id}.`, evidence: input.current_evidence,
        source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", open_decisions: open ? [SCOPE] : [] })
    },
  })
}

test("an instruction left open completes the run with one open gate question; resume and a gate re-run neither pause nor duplicate it; the answer reaches the accepted draft through its review and is not asked again", async () => {
  const directory = matter()
  const state = await runLean(directory, { sessions: instructionModel().port })
  expect(state.status).toBe("completed")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(readLeanJson(directory, "gate.json")).not.toHaveProperty("questions")
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: expect.stringMatching(/^package-scope-/), stage: "gate", status: "open" })])
  const question = readQuestions(directory)[0]!
  const ids = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items).map((row: any) => row.id) as string[]
  expect(question.question).toContain("\nThe note states the delivery period.")
  for (const text of [...ids, "The requested supply note is drafted."]) expect(question.question).not.toContain(text)
  expect((readLeanJson(directory, "gate.json") as any).disclosed).toContainEqual(expect.objectContaining({ kind: "checklist_open",
    ref: expect.objectContaining({ checklist_item: ids[0], question_id: question.id }) }))
  // The open question keys no step: a resume makes no call and does not stop
  // to wait for it.
  const settled = readLeanState(directory)!
  expect(leanSteps(directory).filter(step => settled.stages[step]?.input_sha !== stageInputSha(directory, step))).toEqual([])
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  // A gate re-run asks the same question once, with no advise call. Like any
  // step that is not done, the re-run gate has no key, so its saved advice is
  // not reused as checked: the advice says the inputs changed.
  expect(question.advice).toEqual({ unavailable: { reason: "The scripted advisor recommends nothing." } })
  updateLeanStage(directory, readLeanState(directory)!, "gate", { status: "pending", input_sha: null, outputs: [], finished_at: null })
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  expect(readQuestions(directory)).toEqual([{ ...question, advice: { unavailable: { reason: expect.stringContaining("supporting inputs changed") } } }])
  // The saved plan is followed and the accepted draft kept: the answer reaches
  // the document through its review, and the checklist check receives it. The
  // instruction still reported open stays disclosed and its answered question
  // is not asked again.
  const plan = readLeanPlan(directory)
  enqueueLeanRequest(directory, { kind: "answer", question_id: question.id, answer: ANSWER })
  const answered = instructionModel()
  expect((await runLean(directory, { sessions: answered.port })).status).toBe("completed")
  expect(answered.count(AGENT.draft)).toBe(0)
  expect(readLeanPlan(directory)).toEqual(plan)
  expect(answered.prompts.filter(row => row.agent === AGENT.review).some(row => row.text.includes(ANSWER))).toBe(true)
  expect(answered.prompts.filter(row => row.agent === CHECKLIST_JOB_AGENTS.check).some(row => row.text.includes(ANSWER))).toBe(true)
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question.id, status: "answered", answer: ANSWER })])
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect((readLeanJson(directory, "gate.json") as any).disclosed).toContainEqual(expect.objectContaining({ kind: "checklist_open",
    ref: expect.objectContaining({ question_id: question.id }) }))
}, 180000)

// Which instructed acts a package-scope question showed the lawyer. Only the
// instruction quotes each asker writes count: the words after "directed by"
// in the plan step's act lines and the source quotes in the gate's rows. A
// party name, an act's description, a document title, a check or an open
// decision is not an instruction quote, even when it quotes one. Each asker
// recognises its own question whatever quotation marks the quote holds or its
// excerpt cut, and case does not matter. An act the plan step asks about is
// shown by an act line only when it listed the act for the same parties with
// the act's own quote, or a longer one holding it, so another party's act in
// the same sentence is still asked, as is an act on the whole sentence when
// the line quoted one clause of it, and by a gate row,
// which names no parties, only when the row's quote holds the act's, not when
// the act quotes the longer sentence around it or the gate cut the row's quote
// to an excerpt. Across askers the parties of a gate row's quote are those of
// every act the coverage record has it directing. Questions written before
// this rule read the same.
describe("scope matching: fix wave 3", () => {
  const DOCS: PlannedDocument[] = [{ id: "d1", title: "Release letter", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["Kestrel Works"] }]
  const actOf = (quote: string, parties: string[], act = "take the instructed step", document: CoverageAct["document"] = null): CoverageAct =>
    ({ act, source: "request", quote, parties, document })
  const answered = (question: Question): Question => ({ ...question, status: "answered", answer: "It is outside this request." })
  /** The plan step's own question on `acts`, answered. */
  const planAnswered = (acts: CoverageAct[], documents = DOCS) => answered(coverageScopeQuestion(acts, documents)!)
  /** Whether the plan step still asks about `act` after the `earlier` answers. */
  const planAsks = (act: CoverageAct, earlier: Question[], documents = DOCS) => coverageScopeQuestion([act], documents, earlier) !== undefined
  /** The gate's question on its open instruction rows, recorded and answered. */
  function gateAnswered(directory: string) {
    const [question] = finalGate(directory).questions!
    replaceStageQuestions(directory, "gate", null, [question!])
    answerQuestion(directory, question!.id, { text: "It is outside this request." })
    return question!
  }

  test("a party name, act description or document title the plan step quoted does not show another act", () => {
    const cases: [CoverageAct, CoverageAct, PlannedDocument[]][] = [
      // The party name "the Orchard Committee" lies in the second act's quote.
      [actOf("the Orchard Committee approves the revised schedule", ["the Orchard Committee"], "approve the revised schedule", "d1"),
        actOf("the Orchard Committee grants the access licence", ["the Orchard Committee"], "grant the access licence"), DOCS],
      // So does the first act's description "approve the merger agreement".
      [actOf("Target must approve the merger agreement before signing.", ["Target"], "approve the merger agreement"),
        actOf("Target must also approve the merger agreement by written consent of its shareholders.", ["Target"], "obtain shareholder consent"), DOCS],
      // And the title of the document the plan has the first act in.
      [actOf("Parent must guarantee the Buyer's obligations.", ["Parent"], "guarantee the obligations", "d1"),
        actOf("Parent must also sign the Deed of Guarantee and Indemnity as security trustee.", ["Parent"], "sign as security trustee"),
        [{ ...DOCS[0]!, title: "Deed of Guarantee and Indemnity", executed_by: ["the Lender"] }]],
    ]
    for (const [first, second, documents] of cases) {
      const earlier = planAnswered([first], documents)
      expect(planAsks(first, [earlier], documents)).toBe(false)
      expect(planAsks(second, [earlier], documents)).toBe(true)
    }
    // The disclosure refers to the answered question only for the act it listed.
    const [first, second] = cases[0]!
    const directory = matter()
    const plan = writeLeanPlan(directory, [DOCS[0]!])
    writeLeanText(directory, "coverage.json", JSON.stringify({ schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "uncovered", acts: [first, second], issues: [] }))
    const earlier = planAnswered([first!], plan.documents)
    writeQuestions(directory, [earlier])
    expect(coverageItems(directory).map(item => item.detail.includes(`The lawyer answered the package-scope question ${earlier.id}`))).toEqual([true, false])
  })

  test("an open decision or check that quotes another instruction does not show it to the gate", () => {
    const decision = finding("c1", "open", { open_decisions: ['Whether "Payment is due on delivery." also binds the carrier'] })
    const check = requirement('The note states the delivery period, which runs before "Payment is due on delivery." applies.', "Deliver within thirty days of the order.")
    for (const [first, items] of [[decision, ITEMS], [left("c1"), [{ id: "c1", item: check }, { id: "c2", item: PAYMENT }]]] as [ChecklistFinding, typeof ITEMS][]) {
      const directory = checklistMatter([first, finding("c2", "satisfied")], items)
      const earlier = gateAnswered(directory)
      expect(earlier.question).toContain('"Payment is due on delivery."')
      writeFindings(directory, [first, left("c2")])
      const gate = finalGate(directory)
      expect(gate.questions).toEqual([expect.objectContaining({ question: expect.stringContaining('\nThe note states when payment falls due. Sources: request "Payment is due on delivery."') })])
      expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([earlier.id, gate.questions![0]!.id])
    }
  })

  // A field written from a model's words that holds the fixed words of an act
  // line or a gate row, closing its own quotation mark, forges no line: only
  // the instructions the asker listed are settled.
  const FORGED = "The Buyer shall pay the deposit into the escrow account"
  const ACT_WORDS = `", directed by "${FORGED}" [source: request]; to be taken by "the Buyer"; no planned document carries it. "x`
  const ROW_WORDS = ` Sources: request "${FORGED}" Open decisions: none`

  test("an act description, party name or planned document's title or executor holding an act line's or gate row's fixed words forges neither: the acts listed are settled, the forged one is still asked", () => {
    const QUOTE = "The Seller shall deliver the disclosure letter on completion."
    const listed = (act: string, document: CoverageAct["document"] = null) => actOf(QUOTE, ["the Seller"], act, document)
    const cases: [CoverageAct[], PlannedDocument[]][] = [
      ...[ACT_WORDS, ROW_WORDS].map((words): [CoverageAct[], PlannedDocument[]] => [[listed(`deliver the letter${words}`)], DOCS]),
      [[{ ...listed("deliver the letter"), parties: [`the Seller"; no planned document carries it. "deliver the letter${ACT_WORDS}`] }], DOCS],
      [[listed("deliver the letter", "d1")], [{ ...DOCS[0]!, title: `Release letter${ACT_WORDS}` }]],
      [[listed("deliver the letter", "d1")], [{ ...DOCS[0]!, executed_by: [`Kestrel Works${ACT_WORDS}`] }]],
      // A title ending in the words that open an act line, before the next act line.
      [[listed("deliver the letter", "d1"), actOf("The Buyer shall sign the escrow agreement.", ["the Buyer"], FORGED)], [{ ...DOCS[0]!, title: "Release letter, directed by " }]],
    ]
    for (const [acts, documents] of cases) {
      const earlier = planAnswered(acts, documents)
      for (const act of acts) expect(planAsks(act, [earlier], documents)).toBe(false)
      expect(planAsks(actOf(FORGED, ["the Buyer"]), [earlier], documents)).toBe(true)
      expect(scopeQuestionShows(earlier, [FORGED])).toBe(false)
    }
  })

  test("an open decision or check holding a gate row's or act line's fixed words forges neither: the instructions listed are settled, the forged one is still asked", () => {
    // The second row's own fixed words close a row the first row's field opened.
    for (const words of [ROW_WORDS, ` Sources: request "${FORGED}" with the plan`, ACT_WORDS]) {
      const decision = finding("c1", "open", { open_decisions: [`Compare the${words}`] })
      const check = requirement(`The note states the delivery period${words}.`, "Deliver within thirty days of the order.")
      for (const [first, items] of [[decision, ITEMS], [left("c1"), [{ id: "c1", item: check }, { id: "c2", item: PAYMENT }]]] as [ChecklistFinding, typeof ITEMS][]) {
        const earlier = answered(finalGate(checklistMatter([first, left("c2")], items)).questions![0]!)
        for (const quote of ["Deliver within thirty days of the order.", "Payment is due on delivery."]) expect(scopeQuestionShows(earlier, [quote])).toBe(true)
        expect(scopeQuestionShows(earlier, [FORGED])).toBe(false)
        expect(planAsks(actOf(FORGED, ["the Buyer"]), [earlier])).toBe(true)
      }
    }
  })

  test("each asker recognises its own question whatever quotation marks its quote holds or its excerpt cuts", () => {
    const PIPES = 'Fit 12" pipes at the site entrance.'
    const TAIL = "“Completion Amount” less any Leakage."
    const PLAN_CUT = "On completion the Purchaser shall pay to the Vendor, by electronic transfer of immediately available funds to the account notified in writing by the Vendor " +
      `not less than five Business Days before completion, an amount equal to the ${TAIL}`
    const GATE_CUT = `On completion the Purchaser shall pay to the Vendor, by transfer of immediately available funds to the account the Vendor notifies, an amount equal to the ${TAIL}`
    for (const quote of [PIPES, PLAN_CUT]) {
      const act = actOf(quote, ["the Purchaser"])
      const earlier = planAnswered([act])
      if (quote === PLAN_CUT) expect(earlier.question).toContain("“Complet…")
      expect(planAsks(act, [earlier])).toBe(false)
    }
    const COMMA = 'The "Goods", as listed in the schedule, go to the buyer\'s warehouse.'
    for (const quote of [PIPES, GATE_CUT, COMMA, "Pay on delivery."]) {
      const directory = checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note states the payment.", quote) }])
      const earlier = gateAnswered(directory)
      if (quote === GATE_CUT) expect(earlier.question).toContain("“Com…")
      const gate = finalGate(directory)
      expect(gate.questions).toBeUndefined()
      expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([earlier.id])
      // The plan step reads the gate's row for an act quoting the same words,
      // but not its excerpt: the words cut away may direct another party's act.
      expect(planAsks(actOf(quote, ["the Purchaser"]), [earlier])).toBe(quote === GATE_CUT)
    }
  })

  test("a quote differing only in case from one an answered question showed is not asked again", () => {
    const SIGN = "The Buyer shall sign the escrow agreement."
    const earlier = planAnswered([actOf(SIGN, ["the Buyer"])])
    expect(planAsks(actOf(SIGN.toLowerCase(), ["the Buyer"]), [earlier])).toBe(false)
    const directory = checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note records the signing.", SIGN) }])
    const asked = gateAnswered(directory)
    writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded",
      items: [{ id: "c1", item: requirement("The note records the signing.", SIGN.toUpperCase()) }] }] }))
    const gate = finalGate(directory)
    expect(gate.questions).toBeUndefined()
    expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([asked.id])
  })

  test("another party's act in the sentence an answered question quoted is still asked, as is the same party's act on the sentence around a clause it quoted; the same party's shorter span is not", () => {
    const SENTENCE = "The Seller shall deliver the disclosure letter and the Buyer shall sign the escrow agreement."
    const SELLER = "The Seller shall deliver the disclosure letter"
    const BUYER = "the Buyer shall sign the escrow agreement"
    // The answered question listed the Buyer's clause; the Seller's act quotes the whole sentence.
    expect(planAsks(actOf(SENTENCE, ["the Seller"]), [planAnswered([actOf(BUYER, ["the Buyer"])])])).toBe(true)
    // It listed the whole sentence for the Buyer; the Seller's act quotes its own clause.
    expect(planAsks(actOf(SELLER, ["the Seller"]), [planAnswered([actOf(SENTENCE, ["the Buyer"])])])).toBe(true)
    // The same party's act on the whole sentence when the question listed one
    // clause of it: that line may have shown another act, so it is asked again.
    expect(planAsks(actOf(SENTENCE, ["Buyer"]), [planAnswered([actOf(BUYER, ["the Buyer"])])])).toBe(true)
    // The same party's act quoted shorter, however its name is written, is not.
    expect(planAsks(actOf(BUYER, ["The buyer"]), [planAnswered([actOf(SENTENCE, ["the Buyer"])])])).toBe(false)
    // An act whose parties are named differently, or not at all, is asked.
    expect(planAsks(actOf(BUYER, ["the Buyer", "the Seller"]), [planAnswered([actOf(BUYER, ["the Buyer"])])])).toBe(true)
    expect(planAsks(actOf(BUYER, []), [planAnswered([actOf(BUYER, ["the Buyer"])])])).toBe(true)
    expect(planAsks(actOf(BUYER, []), [planAnswered([actOf(BUYER, [])])])).toBe(false)
    // A gate row names no parties: the plan step takes from it an act it
    // quoted, whole or in part, but not one quoting the longer sentence around it.
    const gateOn = (quote: string) => gateAnswered(checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note records the signing.", quote) }]))
    expect(planAsks(actOf(SENTENCE, ["the Seller"]), [gateOn(BUYER)])).toBe(true)
    expect(planAsks(actOf(BUYER, ["the Buyer"]), [gateOn(BUYER)])).toBe(false)
    expect(planAsks(actOf(BUYER, ["the Buyer"]), [gateOn(SENTENCE)])).toBe(false)
  })

  test("curly and straight quotation marks in a quote read the same", () => {
    const earlier = planAnswered([actOf("The “Buyer” shall pay the price on completion.", ["the Buyer"])])
    for (const quote of ["The “Buyer” shall pay the price on completion.", 'The "Buyer" shall pay the price on completion.'])
      expect(planAsks(actOf(quote, ["the Buyer"]), [earlier])).toBe(false)
    const shown = (question: string, quote: string) => scopeQuestionShows({ id: `${PACKAGE_SCOPE_QUESTION}x`, stage: "gate", question, context: "", status: "answered" }, [quote])
    expect(shown('The checklist check left this instruction open, so the package may not carry it:\nThe note names the buyer. Sources: request "The “Buyer” is the purchaser." ' +
      "No open decision recorded; state what is missing.\nAnswer the open decisions above.", 'The "Buyer" is the purchaser.')).toBe(true)
  })

  test("package-scope questions as both askers wrote them before this rule are still read", () => {
    const shows = (question: string, quotes: string[]) => scopeQuestionShows({ id: `${PACKAGE_SCOPE_QUESTION}earlier`, stage: "draft", question, context: "", status: "answered", answer: "Outside." }, quotes)
    const PLAN = "The document plan leaves these instructed acts without a document executed by the parties who must take them: " +
      '"approve the revised schedule", directed by "the Orchard Committee approves the revised schedule" [source: request]; to be taken by "the Orchard Committee"; planned in d1 "Release letter", executed by "Kestrel Works". ' +
      '"sign the side letter", directed by "The Agent, the Buyer and the Seller each sign the side letter on completion." [source: request]; to be taken by "the Agent", "the Buyer" and "the Seller"; no planned document carries it. ' +
      "For each, should the package add a document, or change a planned one, so that those parties take it, or is it outside this request? If the package should carry it, say in which document and who executes it."
    const earlier: Question = { id: `${PACKAGE_SCOPE_QUESTION}earlier`, stage: "draft", question: PLAN, context: "", status: "answered", answer: "Outside." }
    expect(planAsks(actOf("the Orchard Committee approves the revised schedule", ["the Orchard Committee"]), [earlier])).toBe(false)
    expect(planAsks(actOf("The Agent, the Buyer and the Seller each sign the side letter on completion.", ["the Seller", "the Agent", "the Buyer"]), [earlier])).toBe(false)
    expect(planAsks(actOf("The Agent, the Buyer and the Seller each sign the side letter on completion.", ["the Agent", "the Buyer"]), [earlier])).toBe(true)
    expect(shows(PLAN, ["the Orchard Committee approves the revised schedule"])).toBe(true)
    expect(shows(PLAN, ["Kestrel Works"])).toBe(false)
    const GATE = "The checklist check left these instructions open, so the package may not carry them: " +
      'c1: The note states the delivery period. Sources: request "Deliver within thirty days of the order." Open decisions: Whether "Deliver to the buyer\'s warehouse." applies. ' +
      'c2: The note states when payment falls due. Sources: request "Payment is due on delivery.", d1 "The price is payable on delivery of the goods." No open decision recorded; state what is missing. ' +
      'c3: The note names the port. Sources: request "Ship to Leith.", d1 "Delivery is at the port of Leith." No open decision recorded; state what is missing. ' +
      "For each, answer the open decisions above (supply any missing fact or approval they name). If the package should carry it, say in which document; otherwise say it is outside this request."
    for (const quote of ["Deliver within thirty days of the order.", "Payment is due on delivery.", "The price is payable on delivery of the goods.", "Ship to Leith."]) expect(shows(GATE, [quote])).toBe(true)
    // Neither words outside the sources nor a phrase shorter than SHORT_QUOTE within one.
    for (const quote of ["Deliver to the buyer's warehouse.", "The note states the delivery period.", "on delivery."]) expect(shows(GATE, [quote])).toBe(false)
    // A question whose text changes is read anew.
    const changed: Question = { ...earlier }
    expect(scopeQuestionShows(changed, ["Kestrel Works then releases the deposit."])).toBe(false)
    changed.question = PLAN.replace("the Orchard Committee approves the revised schedule", "Kestrel Works then releases the deposit.")
    expect(scopeQuestionShows(changed, ["Kestrel Works then releases the deposit."])).toBe(true)
  })

  test("an act line is read at its own fixed words whatever brackets or fixed words its source's name or its quote holds: the act it listed is settled, first or last, and another party's act in its words is still asked and disclosed without the answer", () => {
    const SENTENCE = "The Seller shall deliver the disclosure letter and the Buyer shall pay the deposit into the escrow account."
    const BRACKETED = "The Seller shall deliver the title deeds [see clause 4]; to be taken by hand to the Buyer's solicitors."
    const beside = actOf("the Orchard Committee approves the revised schedule", ["the Buyer"], "approve the revised schedule")
    for (const source of ["request", "Heads of Terms [signed].pdf", "Board pack [2]", "Note [1]; to be taken by hand.pdf"]) for (const quote of [SENTENCE, BRACKETED]) {
      const listed = { ...actOf(quote, ["the Seller"], "deliver the documents"), source }
      const buyers = { ...listed, act: "pay the deposit", parties: ["the Buyer"] }
      for (const acts of [[listed, beside], [beside, listed]]) {
        const earlier = planAnswered(acts)
        expect([listed, beside, buyers].map(act => planAsks(act, [earlier]))).toEqual([false, false, true])
      }
    }
    // The disclosure names the answered question only for the act it listed.
    const listed = { ...actOf(SENTENCE, ["the Seller"], "deliver the documents"), source: "Heads of Terms [signed].pdf" }
    const directory = matter()
    const plan = writeLeanPlan(directory, [DOCS[0]!])
    writeLeanText(directory, "coverage.json", JSON.stringify({ schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "uncovered",
      acts: [listed, { ...listed, act: "pay the deposit", parties: ["the Buyer"] }], issues: [] }))
    const earlier = planAnswered([listed, beside], plan.documents)
    writeQuestions(directory, [earlier])
    expect(coverageItems(directory).map(item => item.detail.includes(`The lawyer answered the package-scope question ${earlier.id}`))).toEqual([true, false])
  })

  test("an act line's parties are read whole whatever quotation marks, commas or \"and\" their names hold", () => {
    const QUOTE = "The Seller shall deliver the disclosure letter on completion."
    const ACME = 'Acme Ltd (the "Seller", "we")'
    const FIRM = 'Smith & Co (the "Agent" and "us")'
    const PIPES = 'Fitters of 12" Pipes Ltd'
    const beside = actOf("the Orchard Committee approves the revised schedule", ["the Buyer"], "approve the revised schedule")
    for (const parties of [[ACME], [ACME, "the Buyer"], [FIRM, ACME, "the Buyer"], [PIPES, "the Buyer", "the Seller"]]) {
      const earlier = planAnswered([actOf(QUOTE, parties), beside])
      expect(planAsks(actOf(QUOTE, parties), [earlier])).toBe(false)
      expect(planAsks(actOf(QUOTE, parties.toReversed()), [earlier])).toBe(false)
      // Other parties' act in the same words is still asked.
      for (const others of [[...parties, "the Escrow Agent"], ["Acme Ltd"], parties.slice(1)])
        if (others.length) expect(planAsks(actOf(QUOTE, others), [earlier])).toBe(true)
    }
  })

  test("a gate row reads a quote it cut to an excerpt from its own start whatever quotation marks its document's name holds", () => {
    const LONG = `The supplier shall deliver the goods (the "Goods") to the buyer's warehouse ${"within the agreed window ".repeat(8)}after the order.`
    expect(LONG.length).toBeGreaterThan(160)
    const source = (document: string, quote: string) => ({ document, locator: document, quote })
    for (const document of ["request", 'Letter re "Orchard".pdf', '"Orchard" side letter'])
      for (const refs of [[source(document, LONG)], [source("request", "Payment is due on delivery."), source(document, LONG)]]) {
        const directory = checklistMatter([left("c1")], [{ id: "c1", item: { ...requirement("The note states the delivery window.", LONG), source_refs: refs } }])
        const earlier = gateAnswered(directory)
        expect(earlier.question).toContain(`${document} "${LONG.slice(0, 159)}…"`)
        expect(scopeQuestionShows(earlier, [LONG])).toBe(true)
        expect(scopeQuestionShows(earlier, [LONG.replace("warehouse", "depot")])).toBe(false)
        // The plan step asks again for an act quoting the same instruction: the
        // words the gate cut away may direct another party's act.
        expect(planAsks(actOf(LONG, ["the Supplier"]), [earlier])).toBe(true)
      }
  })

  // Across askers the parties must agree. A gate row names none, so its quotes
  // stand for the acts the saved plan's coverage record has them directing.
  const SENTENCE = "The Seller shall deliver the disclosure letter and the Buyer shall pay the deposit into the escrow account."
  const SELLER = "The Seller shall deliver the disclosure letter"
  const BUYER = "the Buyer shall pay the deposit into the escrow account"
  const SELLS: PlannedDocument[] = [{ ...DOCS[0]!, executed_by: ["the Seller"] }]
  /** A matter whose checklist leaves `quote` open, with a current coverage record of `acts`. */
  const leftWith = (quote: string, acts: CoverageAct[], check = "The note records the deposit.") =>
    recordCoverage(checklistMatter([left("c1")], [{ id: "c1", item: requirement(check, quote) }]), acts)
  /** Whether the plan step's question on `acts` lists `act` after the `earlier` answers. */
  const lists = (acts: CoverageAct[], act: CoverageAct, earlier: Question[], documents = SELLS) =>
    coverageScopeQuestion(acts, documents, earlier)?.question.includes(`"${act.act}", directed by`) ?? false

  test("a plan-step answer settles a gate instruction only when its act lines showed every act the record has it directing, each for its own parties", () => {
    const sentence = actOf(SENTENCE, ["the Seller"], "deliver the disclosure letter")
    const clause = actOf(SELLER, ["the Seller"], "deliver the disclosure letter")
    const buyer = actOf(BUYER, ["the Buyer"], "pay the deposit")
    // The Seller's act, answered, quoted the sentence or its own clause; the
    // instruction left open also directs the Buyer's payment.
    for (const [shown, quote] of [[sentence, BUYER], [sentence, SENTENCE], [clause, SENTENCE]] as const) {
      const directory = leftWith(quote, [shown, buyer])
      writeQuestions(directory, [planAnswered([shown])])
      const gate = finalGate(directory)
      expect(gate.questions).toHaveLength(1)
      expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([gate.questions![0]!.id])
    }
    // The Seller's own clause directs no act of the Buyer's: the answer settles it.
    const directory = leftWith(SELLER, [sentence, buyer])
    const earlier = planAnswered([sentence])
    writeQuestions(directory, [earlier])
    const gate = finalGate(directory)
    expect(gate.questions).toBeUndefined()
    expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([earlier.id])
  })

  test("a plan-step answer that listed the acts of both parties an instruction directs settles it; one that listed either act alone does not", () => {
    const seller = actOf(SELLER, ["the Seller"], "deliver the disclosure letter")
    const buyer = actOf(BUYER, ["the Buyer"], "pay the deposit")
    // Each act quoted its own clause, or both the whole sentence.
    for (const pair of [[seller, buyer], [{ ...seller, quote: SENTENCE }, { ...buyer, quote: SENTENCE }]]) {
      const directory = leftWith(SENTENCE, pair)
      const both = planAnswered(pair)
      writeQuestions(directory, [both])
      const gate = finalGate(directory)
      expect(gate.questions).toBeUndefined()
      expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([both.id])
      for (const one of pair) {
        writeQuestions(directory, [planAnswered([one])])
        const asked = finalGate(directory)
        expect(asked.questions).toHaveLength(1)
        expect(leftOpen(asked).map(item => item.ref.question_id)).toEqual([asked.questions![0]!.id])
      }
    }
  })

  test("acts whose parties the record does not name are each matched by their own quote: an answer on one settles no gate instruction that also directs another", () => {
    const DEPOSIT = "The deposit shall be paid into the escrow account"
    const KEYS = "the keys shall be handed over on completion"
    const deposit = actOf(DEPOSIT, [], "pay the deposit")
    const keys = actOf(KEYS, [], "hand over the keys")
    const directory = leftWith(`${DEPOSIT} and ${KEYS}.`, [deposit, keys])
    writeQuestions(directory, [planAnswered([deposit])])
    const gate = finalGate(directory)
    expect(gate.questions).toHaveLength(1)
    expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([gate.questions![0]!.id])
    // Listing both settles it.
    const both = planAnswered([deposit, keys])
    writeQuestions(directory, [both])
    const settled = finalGate(directory)
    expect(settled.questions).toBeUndefined()
    expect(leftOpen(settled).map(item => item.ref.question_id)).toEqual([both.id])
  })

  test("an act line shows only its own act's quote, not another act's longer quote that holds it: an answer listing one clause settles no act on the whole sentence, at the plan step or the gate", () => {
    const DEPOSIT = "The deposit shall be paid into the escrow account"
    const WHOLE = `${DEPOSIT} and the keys shall be handed over on completion.`
    // The two acts have the same named parties, or no named party at all.
    for (const parties of [["the Buyer"], []]) {
      const clause = actOf(DEPOSIT, parties, "pay the deposit")
      const whole = actOf(WHOLE, parties, "hand over the keys")
      const earlier = planAnswered([clause])
      expect(earlier.question).not.toContain(`"${whole.act}", directed by`)
      // The plan step still asks about the act on the whole sentence, not the clause's.
      expect([clause, whole].map(act => lists([clause, whole], act, [earlier]))).toEqual([false, true])
      // The gate's row on the whole sentence directs both acts: still asked.
      const directory = leftWith(WHOLE, [clause, whole])
      writeQuestions(directory, [earlier])
      const gate = finalGate(directory)
      expect(gate.questions).toHaveLength(1)
      expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([gate.questions![0]!.id])
      // An answer listing both settles both.
      const both = planAnswered([clause, whole])
      expect([clause, whole].map(act => lists([clause, whole], act, [both]))).toEqual([false, false])
      writeQuestions(directory, [both])
      const settled = finalGate(directory)
      expect(settled.questions).toBeUndefined()
      expect(leftOpen(settled).map(item => item.ref.question_id)).toEqual([both.id])
    }
  })

  // The plan step asks only about acts with an issue, or left out by the
  // lawyer's words; an act the plan carries for its parties needs no showing.
  const TRANSFER = "The Seller shall deliver the transfer and the Seller shall procure the deed of release."
  const transfer = actOf("The Seller shall deliver the transfer", ["the Seller"], "deliver the transfer", "d1")
  const release = actOf("the Seller shall procure the deed of release", ["the Seller"], "procure the deed of release")
  /** leftWith under the plan SELLS, whose d1 the Seller executes. */
  const soldWith = (quote: string, acts: CoverageAct[]) => {
    const directory = checklistMatter([left("c1")], [{ id: "c1", item: requirement("The note records the transfer.", quote) }])
    writeLeanPlan(directory, SELLS)
    return recordCoverage(directory, acts)
  }

  test("a plan-step answer that listed every act it asked about settles a gate instruction that also directs an act the plan carries for its parties", () => {
    // The plan step listed only the deed of release, which no planned document
    // carries, and the lawyer answered that it is outside this request.
    const earlier = planAnswered([transfer, release], SELLS)
    expect(earlier.question).not.toContain(`"${transfer.act}", directed by`)
    // The record then leaves it with no document, or reads the answer as leaving it out.
    for (const deed of [release, { ...release, excluded_by: "It is outside this request." }]) {
      const directory = soldWith(TRANSFER, [transfer, deed])
      writeQuestions(directory, [earlier])
      const gate = finalGate(directory)
      expect(gate.questions).toBeUndefined()
      expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([earlier.id])
    }
    // The same answer leaves the instruction asked when the plan does not carry the transfer either.
    const directory = soldWith(TRANSFER, [{ ...transfer, document: null }, release])
    writeQuestions(directory, [earlier])
    const gate = finalGate(directory)
    expect(gate.questions).toHaveLength(1)
    expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([gate.questions![0]!.id])
  })

  test("a gate instruction directing only an act the plan carries is not settled by a plan-step answer whose act lines did not show that act", () => {
    // One answer listed the deed of release; another, another party's act on the whole sentence.
    for (const earlier of [planAnswered([transfer, release], SELLS), planAnswered([actOf(TRANSFER, ["the Buyer"], "accept the transfer")], SELLS)]) {
      const directory = soldWith(`${transfer.quote}.`, [transfer, release])
      writeQuestions(directory, [earlier])
      const gate = finalGate(directory)
      expect(gate.questions).toHaveLength(1)
      expect(leftOpen(gate).map(item => item.ref.question_id)).toEqual([gate.questions![0]!.id])
    }
  })

  test("a gate answer settles a plan act only when every act the record has its instruction directing is for that act's parties", () => {
    const delivery = actOf(SELLER, ["the Seller"], "deliver the disclosure letter", "d1")
    const buyer = actOf(BUYER, ["the Buyer"], "pay the deposit")
    const whole = actOf(SENTENCE, ["the Buyer"], "pay the deposit")
    const gateOn = (quote: string) => gateAnswered(leftWith(quote, [], "The note records the Seller's delivery of the disclosure letter."))
    // The answer on the sentence that also directs the planned Seller's
    // delivery settles neither the Buyer's clause nor the Buyer's whole sentence.
    for (const act of [buyer, whole]) expect(lists([delivery, act], act, [gateOn(SENTENCE)])).toBe(true)
    // The Buyer's own clause directs no act of the Seller's: the answer settles it.
    expect(lists([delivery, buyer], buyer, [gateOn(BUYER)])).toBe(false)
    // The act's disclosure reads the same record.
    const directory = matter()
    writeLeanPlan(directory, SELLS)
    recordCoverage(directory, [delivery, buyer])
    for (const [quote, cites] of [[SENTENCE, false], [BUYER, true]] as const) {
      const earlier = answered(gateOn(quote))
      writeQuestions(directory, [earlier])
      expect(coverageItems(directory).map(item => item.detail.includes(`The lawyer answered the package-scope question ${earlier.id}`))).toEqual([cites])
    }
  })

  test("a gate answer that cut a long instruction to an excerpt settles no plan act: the words cut away may direct another party's act", () => {
    const LONG = "The Seller shall deliver to the Buyer on or before the completion date the executed disclosure letter together with every schedule, annex and supporting document referred to in it, each certified as true and complete."
    const earlier = gateAnswered(leftWith(LONG, []))
    expect(earlier.question).not.toContain(LONG)
    expect(planAsks(actOf(`${LONG} The Buyer shall then pay the balance of the price into the escrow account.`, ["the Buyer"]), [earlier])).toBe(true)
    // The Seller's act on the sentence itself is asked again too: the plan step's asks are bounded.
    expect(planAsks(actOf(LONG, ["the Seller"]), [earlier])).toBe(true)
  })

  // A gate row's quote holding a quoted term, or several quotes, is read at
  // each ", " too; the plan step judges the row the gate showed, not that part.
  const SIGNS = actOf("The Seller shall sign the deed of release", ["the Seller"], "sign the deed of release")
  const PAYS = actOf("the Buyer shall pay the price into the escrow account", ["the Buyer"], "pay the price")

  test("a gate answer on a quote holding a quoted term settles a plan act only when the whole quote directs no other party's act", () => {
    for (const bank of ['"Bank"', "Bank"]) {
      const earlier = gateAnswered(leftWith(`The Seller shall sign the deed of release in favour of the ${bank}, and the Buyer shall pay the price into the escrow account.`, []))
      expect([SIGNS, PAYS].map(act => lists([SIGNS, PAYS], act, [earlier]))).toEqual([true, true])
    }
    // The Seller's own clause with the quoted term directs no act of the Buyer's: the answer settles it.
    const earlier = gateAnswered(leftWith('The Seller shall sign the deed of release in favour of the "Bank", on completion.', []))
    expect([SIGNS, PAYS].map(act => lists([SIGNS, PAYS], act, [earlier]))).toEqual([false, true])
  })

  test("a gate answer on a row whose first quote it cut to an excerpt settles no plan act, whether or not other quotes follow it", () => {
    const LONG = `The Seller shall sign the deed of release on completion ${"and shall procure the release of every charge over the property ".repeat(2)}` +
      "and the Buyer shall pay the price into the escrow account."
    expect(LONG.length).toBeGreaterThan(160)
    expect(LONG.slice(0, 159)).not.toContain(PAYS.quote)
    const SECOND = ref("Payment is due on delivery of the deed.")
    const gateOn = (refs: ReturnType<typeof ref>[]) =>
      gateAnswered(recordCoverage(checklistMatter([left("c1")], [{ id: "c1", item: { ...requirement("The note records the release.", refs[0]!.quote), source_refs: refs } }]), []))
    for (const refs of [[ref(LONG)], [ref(LONG), SECOND]]) expect(lists([SIGNS, PAYS], SIGNS, [gateOn(refs)])).toBe(true)
    // Uncut, the Seller's clause and the payment term direct no act of the Buyer's: the answer settles it.
    expect(lists([SIGNS, PAYS], SIGNS, [gateOn([ref(`${SIGNS.quote} on completion.`), SECOND])])).toBe(false)
  })

  test("a gate answer settles no plan act when its row also quotes another party's act where quotation marks in a document's name hide the quote's start", () => {
    const source = (document: string, quote: string) => ({ document, locator: document, quote })
    const refs = [source('"Orchard" side letter', `${SIGNS.quote} on completion.`), ref(`${PAYS.quote} on completion.`)]
    const earlier = gateAnswered(recordCoverage(checklistMatter([left("c1")], [{ id: "c1", item: { ...requirement("The note records the release.", refs[0]!.quote), source_refs: refs } }]), []))
    expect(lists([SIGNS, PAYS], PAYS, [earlier])).toBe(true)
  })
})
