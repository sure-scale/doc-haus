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
import { handedBackItems, LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { answerQuestion, hashStageInputs, readLeanJson, readLeanState, readQuestions, replaceStageQuestions, shortHash, updateLeanStage, writeLeanText, writeQuestions, type Question } from "./state"
import { cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

// A source conflict among the disclosed checklist findings becomes one batched
// question, asked by the gate after the walk. It pauses nothing: the run
// completes as a review draft with findings and each conflict stays disclosed
// until a later check no longer finds it. Its answer reaches the checklist check's inputs
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
function checklistMatter(findings: ChecklistFinding[], items: { id: string; item: ChecklistItem }[] = ITEMS) {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items }] }))
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
  expect(gate.questions).toHaveLength(1)
  const question = gate.questions![0]!
  // Asymmetric matchers are kept off objects reused below: Bun's
  // toMatchObject writes them into the received value.
  expect(question.id).toMatch(/^checklist-conflict-/)
  expect(question).toMatchObject({ stage: "gate", status: "open" })
  expect(question).not.toHaveProperty("pid")
  expect(question).not.toHaveProperty("document")
  expect(question.question).toBe(['Two checklist requirements taken from the sources conflict.',
    'First: The note states the delivery period. Sources: request "Deliver within thirty days of the order."',
    'Second: The note states when payment falls due. Sources: request "Payment is due on delivery."',
    "Which requirement governs, or how should they be reconciled?"].join(" "))
  const disclosed = gate.disclosed.filter(item => item.kind === "checklist_unresolved")
  expect(disclosed).toHaveLength(2)
  for (const item of disclosed)
    expect(item).toMatchObject({ status: "disclosed", ref: { checklist_class: "source_conflict", question_id: question.id } })
  // Deterministic: the same conflicts give the same question.
  expect(finalGate(directory).questions).toEqual([question])
})

test("the changes stages neither see nor re-own the gate's question", () => {
  const directory = checklistMatter([conflict("c1", "Payment is due on delivery."), conflict("c2", "Deliver within thirty days of the order.")])
  const question = finalGate(directory).questions![0]!
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
  expect(gate.questions).toBeUndefined()
  expect(gate.disclosed[0]!.ref).not.toHaveProperty("question_id")
  const satisfied = checklistMatter([finding("c1", "satisfied"), finding("c2", "satisfied")])
  expect(finalGate(satisfied)).toMatchObject({ status: "pass", delivery: "execution_ready" })
  expect(finalGate(satisfied).questions).toBeUndefined()
})

// A question the lawyer handed back to the drafting, the spine's entry made in
// its place citing `evidence` from `source` (the request unless named), and
// drafting notes that disclose the hand-back unless `disclosed` is false.
const HANDED: Question = { id: "q-terms", stage: "intake", question: "Which term governs?", context: "", status: "answered", answer_kind: "decide_for_me", answered_at: "2026-01-01T00:00:00.000Z" }
function handBack(directory: string, evidence: string[], { disclosed = true, kind = "drafting-choice", source = "request" } = {}) {
  writeQuestions(directory, [HANDED])
  writeLeanText(directory, "spine.json", JSON.stringify({ entries: [{ kind, question_id: HANDED.id, detail: "Choice: the note follows the first source.",
    evidence: evidence.map(quote => ({ source_id: source, quote })), limitations: ["The lawyer has not made this choice."] }], limitations: [] }))
  writeLeanText(directory, "drafting-notes.md", disclosed ? handedBackItems(directory).map(item => item.id).join("\n") : "")
  return directory
}
const HINT = "may be a point you asked us to decide earlier; the notes delivered with the draft describe the choice made."
const UNANSWERED = "Leaving this question unanswered keeps the draft as drafted."
const PRICE = "Unit price: 100 per unit"
const EMAIL = "We propose 120 per unit."
const sides = (a: { check: string; quotes: string[] }, b: { check: string; quotes: string[] }) => {
  const item = ({ check, quotes }: { check: string; quotes: string[] }): ChecklistItem => ({ ...requirement(check, quotes[0]!), source_refs: quotes.map(ref) })
  return { items: [{ id: "c1", item: item(a) }, { id: "c2", item: item(b) }],
    findings: [finding("c1", "unsatisfied", { conflicting_source_refs: [ref(b.quotes.at(-1)!)] }), finding("c2", "unsatisfied", { conflicting_source_refs: [ref(a.quotes.at(-1)!)] })] }
}

test("a conflict the lawyer may have handed back is still asked, once and under its own key, with a plain hint; a strict match alone adds the hint", () => {
  const conflicts = () => checklistMatter([conflict("c1", "Payment is due on delivery."), conflict("c2", "Deliver within thirty days of the order.")])
  const plain = finalGate(conflicts()).questions![0]!
  const BOTH = ["Deliver within thirty days of the order.", "Payment is due on delivery."]
  const directory = handBack(conflicts(), BOTH)
  const gate = finalGate(directory)
  // Never met: both stay adverse findings, disclosed with the question, and
  // the hand-back stays disclosed as a choice the lawyer has not made.
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft_with_findings", blocking: [] })
  expect(gate.questions).toHaveLength(1)
  const [question] = gate.questions!
  expect(question!.id).toBe(plain.id)
  expect(question!.question).toBe(plain.question)
  expect(question!.context).toBe(`${plain.context} This conflict ${HINT} ${UNANSWERED}`)
  for (const id of [HANDED.id, "c1", "c2", "checklist-conflict-", "handed-back"]) expect(question!.context).not.toContain(id)
  expect(gate.disclosed.filter(item => item.kind === "checklist_unresolved").map(item => item.ref.question_id)).toEqual([plain.id, plain.id])
  expect(gate.disclosed.filter(item => item.stage === "finalize").map(item => item.ref.question_id)).toEqual([HANDED.id])
  // A choice citing one side only, a passage neither side cites, no passage,
  // the same words in another source, an open item rather than a choice, a
  // side with open decisions of its own, an undisclosed hand-back and a real
  // answer give no hint; the question is asked the same way.
  const hinted = (directory: string) => finalGate(directory).questions![0]!.context.includes(HINT)
  for (const evidence of [[BOTH[0]!], [BOTH[1]!], [...BOTH, "Deliver to the buyer's warehouse."], []]) expect(hinted(handBack(conflicts(), evidence))).toBe(false)
  expect(hinted(handBack(conflicts(), BOTH, { source: "term-sheet" }))).toBe(false)
  expect(hinted(handBack(conflicts(), BOTH, { kind: "open-item" }))).toBe(false)
  // The open decisions an adverse side records are asked in no question, so
  // the hint would read as settling them.
  const undecided = () => checklistMatter([finding("c1", "unsatisfied", { conflicting_source_refs: [ref("Payment is due on delivery.")],
    open_decisions: ["The delivery address is not given in the sources"] }), conflict("c2", "Deliver within thirty days of the order.")])
  const asked = finalGate(undecided()).questions!
  expect(asked.map(row => row.id)).toEqual([plain.id])
  expect(finalGate(handBack(undecided(), BOTH)).questions).toEqual(asked)
  const silent = finalGate(handBack(conflicts(), BOTH, { disclosed: false }))
  expect(silent.status).toBe("blocked")
  expect(silent.questions![0]).toEqual(plain)
  answerQuestion(directory, HANDED.id, "Payment on delivery governs.")
  expect(finalGate(directory).questions).toEqual([plain])
})

test("a choice whose passages do not match the conflict's exactly gives no hint: a wider passage, another term in the same passages, a generic phrase, a shared passage or one side only", () => {
  const SHARED = "Use the term sheet figures for every commercial term."
  const TS = "Unit price: 100 per unit, payable within 30 days of invoice."
  const EM = "We propose 120 per unit, payable within 60 days of invoice."
  const A = "Delivery shall be made in accordance with the order form, within ten days."
  const B = "Delivery shall be made in accordance with the order form, within forty days."
  const cases: [string, ReturnType<typeof sides>, string[]][] = [
    ["payment period, same passages as a price choice", sides({ check: "Payment within 30 days.", quotes: [TS] }, { check: "Payment within 60 days.", quotes: [EM] }), [PRICE, "We propose 120 per unit"]],
    ["generic phrase inside both sides", sides({ check: "Delivery within ten days.", quotes: [A] }, { check: "Delivery within forty days.", quotes: [B] }), ["in accordance with the order form"]],
    ["a shared passage and one side's own", sides({ check: "Price is 100 per unit.", quotes: [SHARED, PRICE] }, { check: "Price is 120 per unit.", quotes: [SHARED, EMAIL] }), [PRICE, SHARED]],
    ["one side only", sides({ check: "Price is 100 per unit.", quotes: [PRICE] }, { check: "Price is 120 per unit.", quotes: [EMAIL] }), [PRICE]],
    ["short passages inside longer ones", sides({ check: "Price is 100 per unit.", quotes: ["100 per unit"] }, { check: "Price is 120 per unit.", quotes: ["120 per unit"] }), [PRICE, EMAIL]],
  ]
  for (const [name, { items, findings }, evidence] of cases) {
    const gate = finalGate(handBack(checklistMatter(findings, items), evidence))
    expect([name, gate.questions?.map(row => row.id.startsWith("checklist-conflict-"))]).toEqual([name, [true]])
    expect([name, gate.questions![0]!.context.includes(HINT)]).toEqual([name, false])
  }
  // The same conflicts, each cited exactly, or as a longer span of the same
  // passage, do get the hint.
  const exact = sides({ check: "Price is 100 per unit.", quotes: [PRICE] }, { check: "Price is 120 per unit.", quotes: ["We propose 120 per unit"] })
  for (const evidence of [[PRICE, "We propose 120 per unit"], [`${PRICE}, payable within 30 days.`, `${EMAIL} Payable in 60 days.`]])
    expect(finalGate(handBack(checklistMatter(exact.findings, exact.items), evidence)).questions![0]!.context).toContain(`This conflict ${HINT}`)
})

test("among several conflicts the hint names only the pair the choice matches, and says nothing of leaving the question unanswered; an answered question is not asked again once the hand-back is disclosed", () => {
  const TEN = "Deliver within ten days of the order.", FORTY = "Deliver within forty days of the order."
  const price = sides({ check: "Price is 100 per unit.", quotes: [PRICE] }, { check: "Price is 120 per unit.", quotes: [EMAIL] })
  const items = [...price.items, { id: "c3", item: requirement("Deliver within ten days.", TEN) }, { id: "c4", item: requirement("Deliver within forty days.", FORTY) }]
  const findings = [...price.findings, conflict("c3", FORTY), conflict("c4", TEN)]
  const directory = handBack(checklistMatter(findings, items), [PRICE, EMAIL], { disclosed: false })
  const first = finalGate(directory).questions![0]!
  expect(first.context).not.toContain(HINT)
  const pair = first.question.indexOf(`"${EMAIL}"`) < first.question.indexOf(`"${TEN}"`) ? 1 : 2
  // The lawyer answers it before the drafting notes disclose the hand-back.
  replaceStageQuestions(directory, "gate", null, [first])
  answerQuestion(directory, first.id, "Ten days; the term sheet price.")
  writeLeanText(directory, "drafting-notes.md", handedBackItems(directory).map(item => item.id).join("\n"))
  const again = finalGate(directory).questions!
  expect(again.map(row => row.id)).toEqual([first.id])
  // Leaving the question unanswered would leave the other pair as drafted
  // too, which no choice settled: that sentence is kept for a question every
  // part of which a choice may have settled.
  expect(again[0]!.context).toBe(`${first.context} Pair ${pair} ${HINT}`)
  replaceStageQuestions(directory, "gate", null, again)
  expect(readQuestions(directory).filter(row => row.status === "open")).toEqual([])
  expect(readQuestions(directory).find(row => row.id === first.id)).toMatchObject({ status: "answered", context: first.context })
})

test("an answer to the conflict question re-keys the checklist check and reaches its source instructions", () => {
  const directory = checklistMatter([conflict("c1", "Payment is due on delivery."), finding("c2", "satisfied")])
  const question = finalGate(directory).questions![0]!
  writeLeanText(directory, "questions.json", `${JSON.stringify({ questions: [question] }, null, 2)}\n`)
  const before = stageInputSha(directory, "checklist-check")
  // An open question is not yet an instruction.
  expect(checklistProducerInputs(directory).find(row => row.key === "request")!.input.instructions).toEqual([])
  answerQuestion(directory, question.id, "The payment requirement governs.")
  expect(stageInputSha(directory, "checklist-check")).not.toBe(before)
  expect(checklistProducerInputs(directory).find(row => row.key === "request")!.input.instructions).toContainEqual(
    { document: "request", locator: `.dochaus/lean/questions.json answer (${question.id})`, quote: "The payment requirement governs." })
})

const ANSWER = "Payment on delivery governs; state it in the note."

// Two requirements from the request; the first is reported as contradicted by
// the second's source passage. With `answerFirst`, once the answer is a
// settled instruction the extractor also extracts it as an item, placed first,
// so the two conflicting items take other ordinal ids.
function conflictModel({ answerFirst = false } = {}) {
  return scriptedModel({
    ...happyPath,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
      const item = (check: string, refs = input.passages) => ({ check, basis: "directive", source_refs: refs, qualifications: [], target: { ...input.targets[0], treatment: "operative" } })
      if (input.source.document !== "request") return JSON.stringify({ items: [] })
      const answer = input.instructions.filter(row => row.quote === ANSWER)
      const first = answerFirst && answer.length ? [item("The note states the payment rule the lawyer chose.", answer)] : []
      return JSON.stringify({ items: [...first, item("The requested supply note is drafted."), item("The note states that payment is due on delivery.")] })
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
  expect(readLeanJson(directory, "gate.json")).not.toHaveProperty("questions")
  const asked = readQuestions(directory)
  expect(asked).toEqual([expect.objectContaining({ id: expect.stringMatching(/^checklist-conflict-/), stage: "gate", status: "open" })])
  const question = readQuestions(directory)[0]
  const ids = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items).map((row: any) => row.id) as string[]
  expect(ids).toHaveLength(2)
  for (const id of ids) expect(question!.question).not.toContain(id)
  expect(question!.question).toContain("First: The note states that payment is due on delivery.")
  expect(question!.question).toContain("Second: The requested supply note is drafted.")
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

test("an answered conflict is not asked again when re-extraction gives its requirements other item ids; its disclosure keeps the answered question", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: conflictModel().port })).status).toBe("completed")
  const [question] = readQuestions(directory)
  const extracted = () => (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items).map((row: any) => row.id) as string[]
  const before = extracted()
  enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: ANSWER })
  expect((await runLean(directory, { sessions: conflictModel({ answerFirst: true }).port })).status).toBe("completed")
  // The answer was extracted first, so the two conflicting items moved.
  const after = extracted()
  expect(after).toHaveLength(3)
  expect(after.slice(1)).not.toEqual(before)
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question!.id, status: "answered", answer: ANSWER })])
  const disclosed = (readLeanJson(directory, "gate.json") as any).disclosed.filter((row: any) => row.ref.checklist_class === "source_conflict")
  expect(disclosed.length).toBeGreaterThan(0)
  for (const row of disclosed) expect(row.ref.question_id).toBe(question!.id)
  for (const id of new Set([...before, ...after])) expect(question!.question).not.toContain(id)
}, 180000)

test("the conflict question names each requirement by its check and sources, never by an internal item id, and its context is true before and after an answer", () => {
  const directory = checklistMatter([conflict("c1", "Payment is due on delivery."), conflict("c2", "Deliver within thirty days of the order.")])
  const [question] = finalGate(directory).questions!
  expect(question!.question).not.toMatch(/\bc[12]\b/)
  for (const text of ["The note states the delivery period.", "The note states when payment falls due.",
    'request "Deliver within thirty days of the order."', 'request "Payment is due on delivery."'])
    expect(question!.question).toContain(text)
  expect(question!.context).toContain("Until a later check no longer finds this conflict, it is disclosed with the review draft")
  expect(question!.context).not.toContain("Until you answer")
  expect(question!.context).toContain("Your answer re-runs the stages that use it, including the checklist check; a drafted document changes only by amendment.")
  expect(question!.context).not.toContain("from the start")
})

test("the same conflicting requirements under other item ids are the same question", () => {
  const first = finalGate(checklistMatter([conflict("c1", "Payment is due on delivery."), conflict("c2", "Deliver within thirty days of the order.")])).questions!
  // Re-extracted with one more item first: c1 and c2 are now c2 and c3.
  const shifted = checklistMatter([finding("c1", "satisfied"), conflict("c2", "Payment is due on delivery."), conflict("c3", "Deliver within thirty days of the order.")])
  writeLeanText(shifted, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded",
    items: [{ id: "c1", item: requirement("The note states the governing law.", "English law governs the note.") }, ...ITEMS.map((row, index) => ({ ...row, id: `c${index + 2}` }))] }] }))
  expect(finalGate(shifted).questions).toEqual(first)
})

// A conflict question answered before its id keyed on the requirements'
// quotes: its id hashed the ordinal item ids, and its text named them.
const LEGACY = (a: string, b: string, payment = "Payment is due on delivery."): Question => ({
  id: `checklist-conflict-${shortHash(JSON.stringify([[a, b]]))}`, stage: "gate", status: "answered", answer: ANSWER, answered_at: "2026-10-01T00:00:00.000Z",
  question: `Checklist requirements taken from the sources conflict: ${a} with ${b}. ${a}: The note states the delivery period. Sources: request "Deliver within thirty days of the order." ${b}: The note states when payment falls due. Sources: request "${payment}" Which requirement governs, or how should they be reconciled?`,
  context: "Until you answer, each conflicting requirement is disclosed with the review draft as not satisfied; delivery is not held for it. Your answer re-runs the drafting from the start, including the checklist check.",
})

test("a conflict answered under an ordinal-id question is not asked again when that question showed the same requirements; any other is asked", () => {
  const shifted = () => {
    const directory = checklistMatter([finding("c1", "satisfied"), conflict("c2", "Payment is due on delivery."), conflict("c3", "Deliver within thirty days of the order.")])
    writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded",
      items: [{ id: "c1", item: requirement("The note states the governing law.", "English law governs the note.") }, ...ITEMS.map((row, index) => ({ ...row, id: `c${index + 2}` }))] }] }))
    return directory
  }
  // Answered when the same two requirements were c1 and c2.
  const same = shifted()
  const legacy = LEGACY("c1", "c2")
  writeQuestions(same, [legacy])
  const gate = finalGate(same)
  expect(gate.questions!.map(row => row.id)).toEqual([legacy.id])
  for (const item of gate.disclosed.filter(item => item.kind === "checklist_unresolved")) expect(item.ref.question_id).toBe(legacy.id)
  replaceStageQuestions(same, "gate", null, gate.questions!)
  expect(readQuestions(same)).toEqual([legacy])
  // One that showed another payment requirement asked about another conflict:
  // this one is asked, and the answered question is left as it was.
  const other = shifted()
  const different = LEGACY("c1", "c2", "Payment is due within sixty days of delivery.")
  writeQuestions(other, [different])
  const asked = finalGate(other).questions!
  expect(asked.map(row => [row.id === different.id, row.status])).toEqual([[false, "open"]])
  replaceStageQuestions(other, "gate", null, asked)
  expect(readQuestions(other)).toEqual([different, expect.objectContaining({ id: asked[0]!.id, status: "open" })])
})

test("an open conflict question in the ordinal-id form is asked in the current form on resume, though the gate's other inputs are unchanged", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: conflictModel().port })).status).toBe("completed")
  const [question] = readQuestions(directory)
  // The matter as an earlier gate left it: its open question in that form,
  // and the gate done from the inputs it keyed then, with no gate contract.
  const items = (readLeanJson(directory, "checklist/extraction.json") as any).sources.flatMap((row: any) => row.items) as { id: string; item: ChecklistItem }[]
  const [a, b] = items.map(row => row.id)
  const quote = items[0]!.item.source_refs[0]!.quote.replace(/\s+/g, " ").trim()
  writeQuestions(directory, [{ id: `checklist-conflict-${shortHash(JSON.stringify([[a, b]]))}`, stage: "gate", status: "open",
    question: `Checklist requirements taken from the sources conflict: ${a} with ${b}. ${a}: ${items[0]!.item.check} Sources: request "${quote}" ${b}: ${items[1]!.item.check} Sources: request "${quote}" Which requirement governs, or how should they be reconciled?`,
    context: "Until you answer, each conflicting requirement is disclosed with the review draft as not satisfied; delivery is not held for it. Your answer re-runs the drafting from the start, including the checklist check." }])
  const earlier = stageInputs(directory, "gate").filter(part => part.name !== "gate-contract")
  updateLeanStage(directory, readLeanState(directory)!, "gate", { input_sha: hashStageInputs("gate", LEAN_PROMPT_VERSION, earlier) })
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  expect(readQuestions(directory)).toEqual([question!])
}, 180000)

test("an answer's source locator survives a later step rewriting questions.json: the extraction stays current, the run completes and an idle resume makes no call", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: conflictModel().port })).status).toBe("completed")
  const [conflict] = readQuestions(directory)
  // The gate asked two questions; the lawyer answers the second. A later
  // step's question rewrite puts answered questions first, so the answer
  // moves within questions.json during the walk.
  writeQuestions(directory, [conflict!, { id: "gate-second", stage: "gate", question: "Should the package carry the instruction?", context: "", status: "open" }])
  enqueueLeanRequest(directory, { kind: "answer", question_id: "gate-second", answer: "It is outside this request." })
  const state = await runLean(directory, { sessions: conflictModel().port })
  expect(state.stop_reason ?? "").not.toContain("Source extraction is incomplete or stale")
  expect(state.status).toBe("completed")
  expect(readQuestions(directory).map(row => row.id)).toEqual(["gate-second", conflict!.id])
  expect(checklistProducerInputs(directory).find(row => row.key === "request")!.input.instructions).toEqual(
    [{ document: "request", locator: ".dochaus/lean/questions.json answer (gate-second)", quote: "It is outside this request." }])
  const settled = readLeanState(directory)!
  expect(leanSteps(directory).filter(step => settled.stages[step]?.input_sha !== stageInputSha(directory, step))).toEqual([])
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
}, 180000)

test("a rewrite that only reorders questions.json re-keys no step: finalize, change-review and the gate stay current and a resume makes no call", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: conflictModel().port })).status).toBe("completed")
  const [conflict] = readQuestions(directory)
  const answered = (id: string): Question => ({ id, stage: "gate", question: `Does ${id} apply?`, context: "", status: "answered", answer: `${id} applies.` })
  writeQuestions(directory, [conflict!, answered("gate-b"), answered("gate-a")])
  expect((await runLean(directory, { sessions: conflictModel().port })).status).toBe("completed")
  const steps = leanSteps(directory).filter(step => step === "finalize" || step === "gate" || step.endsWith(".change-review"))
  expect(steps.length).toBe(3)
  const keys = () => steps.map(step => stageInputSha(directory, step))
  const before = { keys: keys(), questions: readQuestions(directory) }
  writeQuestions(directory, before.questions.toReversed())
  expect(keys()).toEqual(before.keys)
  const settled = readLeanState(directory)!
  expect(leanSteps(directory).filter(step => settled.stages[step]?.input_sha !== stageInputSha(directory, step))).toEqual([])
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toHaveLength(0)
  expect(readQuestions(directory).toSorted((a, b) => a.id.localeCompare(b.id))).toEqual(before.questions.toSorted((a, b) => a.id.localeCompare(b.id)))
}, 180000)
