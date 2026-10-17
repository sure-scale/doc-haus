import { afterEach, expect, test } from "bun:test"
import { askedQuestions } from "./change-checks"
import type { ChecklistCheckerInput, ChecklistFinding, ChecklistItem, ChecklistProducerInput } from "./checklist-contract"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { checklistProducerInputs } from "./checklist-source-context"
import { finalGate } from "./gate"
import { leanSteps, writeLeanPlan } from "./plan"
import { changesRepairSections } from "./repairs"
import { enqueueLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { answerQuestion, readLeanJson, readLeanState, readQuestions, updateLeanStage, writeLeanText, writeQuestions, type Question } from "./state"
import { cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

// A source conflict among the disclosed checklist findings becomes one batched
// question, asked by the gate after the walk. It pauses nothing: the run
// completes as a review draft with findings and each conflict stays disclosed
// until the lawyer answers. Its answer reaches the checklist check's inputs
// through the ordinary answer path. Only the model is scripted.

afterEach(cleanupMatters)

const ref = (quote: string) => ({ document: "request", locator: "request", quote })
const requirement = (check: string, quote: string): ChecklistItem =>
  ({ check, basis: "directive", source_refs: [ref(quote)], qualifications: [], target: { document: "d1", role: "Supply note", treatment: "operative" } })
const ITEMS = [
  { id: "c1", item: requirement("The note states the delivery period.", "Deliver within thirty days of the order.") },
  { id: "c2", item: requirement("The note states when payment falls due.", "Payment is due on delivery.") },
]
const finding = (item_id: string, status: ChecklistFinding["status"], extra: Partial<ChecklistFinding> = {}): ChecklistFinding => ({
  item_id, status, reason: `Focused result ${status}.`, evidence: [], source_refs: [], inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", ...extra,
})

// The checklist's saved output, as the gate reads it.
function checklistMatter(findings: ChecklistFinding[]) {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items: ITEMS }] }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: findings.map(row => row.item_id),
    checks: findings.map(row => ({ item_id: row.item_id, input_sha: "sha", status: "succeeded", issues: [], finding: row })) }))
  writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
  return directory
}

const conflict = (item_id: string, quote: string) => finding(item_id, "unsatisfied", { conflicting_source_refs: [ref(quote)] })

test("the gate batches every source conflict into one question naming each pair, and links each disclosure to it", () => {
  // Each side reports the other: one pair, one question.
  const directory = checklistMatter([conflict("c1", "Payment is due on delivery."), conflict("c2", "Deliver within thirty days of the order.")])
  const gate = finalGate(directory)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  const question = gate.question!
  // Asymmetric matchers are kept off objects reused below: Bun's
  // toMatchObject writes them into the received value.
  expect(question.id).toMatch(/^checklist-conflict-/)
  expect(question).toMatchObject({ stage: "gate", status: "open" })
  expect(question).not.toHaveProperty("pid")
  expect(question).not.toHaveProperty("document")
  for (const text of ["c1 with c2", "c1: The note states the delivery period.", "c2: The note states when payment falls due.",
    'request "Deliver within thirty days of the order."', 'request "Payment is due on delivery."'])
    expect(question.question).toContain(text)
  expect(question.question.match(/ with /g)).toHaveLength(1)
  const disclosed = gate.disclosed.filter(item => item.kind === "checklist_unresolved")
  expect(disclosed).toHaveLength(2)
  for (const item of disclosed)
    expect(item).toMatchObject({ status: "disclosed", ref: { checklist_class: "source_conflict", question_id: question.id } })
  // Deterministic: the same conflicts give the same question.
  expect(finalGate(directory).question).toEqual(question)
})

test("the changes stages neither see nor re-own the gate's question", () => {
  const directory = checklistMatter([conflict("c1", "Payment is due on delivery."), conflict("c2", "Deliver within thirty days of the order.")])
  const question = finalGate(directory).question!
  const drafting: Question = { id: "q-earlier", stage: "changes", question: "Which delivery address applies?", context: "", status: "open" }
  writeQuestions(directory, [question, drafting])
  const listed = changesRepairSections(directory, "d1").join("\n")
  expect(listed).toContain("q-earlier")
  expect(listed).not.toContain(question.id)
  // A repeat of the gate question's id stays the gate's: it is dropped, so it
  // cannot become a changes-stage question that pauses the run.
  const asked = new Map(readQuestions(directory).map(row => [row.id, row]))
  const repeated = askedQuestions([{ id: question.id, question: question.question }, { id: "q-earlier", question: drafting.question }], asked, "d1")
  expect(repeated[0]).toBeUndefined()
  expect(repeated[1]).toMatchObject({ id: "q-earlier", stage: "changes", status: "open" })
})

test("findings without a source conflict ask nothing and keep their delivery", () => {
  const omission = checklistMatter([finding("c1", "unsatisfied"), finding("c2", "satisfied")])
  const gate = finalGate(omission)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings" })
  expect(gate.question).toBeUndefined()
  expect(gate.disclosed[0]!.ref).not.toHaveProperty("question_id")
  const satisfied = checklistMatter([finding("c1", "satisfied"), finding("c2", "satisfied")])
  expect(finalGate(satisfied)).toMatchObject({ status: "pass", delivery: "execution_ready" })
  expect(finalGate(satisfied).question).toBeUndefined()
})

test("an answer to the conflict question re-keys the checklist check and reaches its source instructions", () => {
  const directory = checklistMatter([conflict("c1", "Payment is due on delivery."), finding("c2", "satisfied")])
  const question = finalGate(directory).question!
  writeLeanText(directory, "questions.json", `${JSON.stringify({ questions: [question] }, null, 2)}\n`)
  const before = stageInputSha(directory, "checklist-check")
  // An open question is not yet an instruction.
  expect(checklistProducerInputs(directory).find(row => row.key === "request")!.input.instructions).toEqual([])
  answerQuestion(directory, question.id, "The payment requirement governs.")
  expect(stageInputSha(directory, "checklist-check")).not.toBe(before)
  expect(checklistProducerInputs(directory).find(row => row.key === "request")!.input.instructions).toContainEqual(
    { document: "request", locator: `.dochaus/lean/questions.json /questions/0/answer (${question.id})`, quote: "The payment requirement governs." })
})

const ANSWER = "Payment on delivery governs; state it in the note."

// Two requirements from the request; the first is reported as contradicted by
// the second's source passage.
function conflictModel() {
  return scriptedModel({
    ...happyPath,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
      const item = (check: string) => ({ check, basis: "directive", source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" } })
      return JSON.stringify({ items: input.source.document === "request" ? [item("The requested supply note is drafted."), item("The note states that payment is due on delivery.")] : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
      const conflicted = input.item.check.startsWith("The requested")
      return JSON.stringify({ item_id: input.item_id, status: conflicted ? "unsatisfied" : "satisfied", reason: `Focused result for ${input.item_id}.`, evidence: input.current_evidence,
        source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.",
        ...(conflicted ? { conflicting_source_refs: input.item.source_refs } : {}) })
    },
  })
}

test("a source conflict completes the run with one open question; resume and a gate re-run neither pause nor duplicate it; the answer reaches the checklist check", async () => {
  const directory = matter()
  const state = await runLean(directory, { sessions: conflictModel().port })
  expect(state.status).toBe("completed")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(readLeanJson(directory, "gate.json")).not.toHaveProperty("question")
  const asked = readQuestions(directory)
  expect(asked).toEqual([expect.objectContaining({ id: expect.stringMatching(/^checklist-conflict-/), stage: "gate", status: "open" })])
  const question = readQuestions(directory)[0]
  const ids = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items).map((row: any) => row.id) as string[]
  expect(ids).toHaveLength(2)
  expect(question!.question).toContain(`${ids[0]} with ${ids[1]}`)
  expect((readLeanJson(directory, "gate.json") as any).disclosed).toContainEqual(expect.objectContaining({ kind: "checklist_unresolved",
    ref: expect.objectContaining({ checklist_item: ids[0], checklist_class: "source_conflict", question_id: question!.id }) }))
  // The open question keys no step: every settled step is current, and a
  // resume makes no call and does not stop to wait for it.
  const settled = readLeanState(directory)!
  expect(leanSteps(directory).filter(step => settled.stages[step]?.input_sha !== stageInputSha(directory, step))).toEqual([])
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  // A gate re-run asks the same question once.
  updateLeanStage(directory, readLeanState(directory)!, "gate", { status: "pending", input_sha: null, outputs: [], finished_at: null })
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  expect(readQuestions(directory)).toEqual([question!])
  // The answer re-runs the stages that consume answers; the checklist check
  // receives it as a settled instruction. The conflict still reported stays
  // disclosed and its answered question is not asked again.
  enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: ANSWER })
  const answered = conflictModel()
  expect((await runLean(directory, { sessions: answered.port })).status).toBe("completed")
  expect(answered.prompts.filter(row => row.agent === CHECKLIST_JOB_AGENTS.check).some(row => row.text.includes(ANSWER))).toBe(true)
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question!.id, status: "answered", answer: ANSWER })])
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft_with_findings" })
  expect((readLeanJson(directory, "gate.json") as any).disclosed).toContainEqual(expect.objectContaining({ ref: expect.objectContaining({ question_id: question!.id }) }))
}, 180000)
