import { afterEach, expect, test } from "bun:test"
import { checkEvidence, evidencePrompt } from "./evidence"
import { finalGate, type GateResult } from "./gate"
import { intakeQuestions } from "./intake"
import { writeLeanPlan } from "./plan"
import { openItem } from "./quotes"
import { drainLeanRequests, enqueueLeanRequest, queuedAnswerIDs, type LeanRequest, type StoredLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { ADOPTED_ANSWER, answeredQuestionsText, answerQuestion, delegable, LeanQuestionError, readLeanJson, readOpenItems, readQuestions, writeOpenItems, writeQuestions, type Question, type QuestionAdvice } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, NO_FINDINGS, noChanges, scriptedModel, SUPPORT } from "./test-support"

// "Decide for me" adopts the advisor's recommendation as the lawyer's answer.
// An answer is the lawyer's typed words, an offered option picked, or the
// recommended option adopted; each is stored as words, and every stage reads
// the words, an adopted one marked as delegated. Adoption is open only for an
// open judgment, instruction or pure scope question whose advice recommends an
// option (lib/question-delegation.ts). Synthetic matter; no model runs.

afterEach(cleanupMatters)

const RECOMMENDED = "Leave liability uncapped."
const ADVICE: QuestionAdvice = {
  options: [{ id: "o1", text: "Cap liability at the price." }, { id: "o2", text: RECOMMENDED }],
  recommendation: { option: "o2", rationale: "The terms set no cap.", basis: SUPPORT },
}
const UNAVAILABLE: QuestionAdvice = { options: [], unavailable: { reason: "Nothing in the sources supports one option." } }
const OPEN: Question = { id: "q-cap", stage: "intake", question: "Should liability be capped?", context: "The terms are silent.", status: "open", kind: "judgment", advice: ADVICE }
const FILE_NAME: Question = {
  id: "requested-files-d1", stage: "draft", question: "Which file names should the delivered files take?", context: "Named files.", status: "open", kind: "file_name",
  options: [{ id: "o1", text: "As my words ask" }, { id: "o2", text: "As planned" }], advice: { recommendation: { option: "o1", rationale: "The request names the file.", basis: SUPPORT } },
}

function refusal(run: () => unknown) {
  try {
    run()
    return "admitted"
  } catch (error) {
    return error instanceof LeanQuestionError ? error.code : String(error)
  }
}

function outcome(directory: string, request: StoredLeanRequest) {
  return (readLeanJson(directory, `requests/done/${request.id}.json`) as { outcome: { status: string; detail?: string } }).outcome
}

test("an answer is typed words, a picked option or the adopted recommendation, each stored as words", () => {
  const directory = matter()
  const answer = (question: Question, form: Parameters<typeof answerQuestion>[2]) => {
    writeQuestions(directory, [question])
    return answerQuestion(directory, question.id, form)
  }
  const typed = answer(OPEN, { text: ` ${RECOMMENDED}\r\n` })
  expect(typed).toMatchObject({ status: "answered", answer: RECOMMENDED, answered_at: expect.any(String) })
  expect(typed).not.toHaveProperty("answer_kind")
  expect(typed).not.toHaveProperty("option_id")
  expect(answer(OPEN, { option_id: "o1" })).toMatchObject({ status: "answered", answer: "Cap liability at the price.", answer_kind: "option", option_id: "o1" })
  expect(answer(OPEN, { adopt: true })).toMatchObject({ status: "answered", answer: RECOMMENDED, answer_kind: "adopted", option_id: "o2" })
  // A question's own options are picked the same way.
  expect(answer(FILE_NAME, { option_id: "o2" })).toMatchObject({ status: "answered", answer: "As planned", answer_kind: "option", option_id: "o2" })
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: FILE_NAME.id, answer: "As planned" })])
  // The lawyer's later words replace a picked or adopted answer.
  answer(OPEN, { adopt: true })
  const replaced = answerQuestion(directory, OPEN.id, { text: "Cap it at twice the price." })
  expect(replaced).toMatchObject({ status: "answered", answer: "Cap it at twice the price." })
  expect(replaced).not.toHaveProperty("answer_kind")
  expect(replaced).not.toHaveProperty("option_id")
})

test("an option the question does not offer, or empty words, is an invalid answer and records nothing", () => {
  const directory = matter()
  writeQuestions(directory, [OPEN, { ...OPEN, id: "q-plain", kind: "fact", advice: UNAVAILABLE }])
  expect(refusal(() => answerQuestion(directory, "q-cap", { option_id: "o9" }))).toBe("invalid_answer")
  expect(refusal(() => answerQuestion(directory, "q-plain", { option_id: "o1" }))).toBe("invalid_answer")
  expect(refusal(() => answerQuestion(directory, "q-cap", { text: "  " }))).toBe("invalid_answer")
  expect(refusal(() => answerQuestion(directory, "q-missing", { option_id: "o1" }))).toBe("not_found")
  expect(readQuestions(directory).map(question => question.status)).toEqual(["open", "open"])
})

test("Decide for me is offered only for an open judgment, instruction or pure scope question whose advice recommends an option", () => {
  const directory = matter()
  const cases: [string, Question, string][] = [
    ["judgment", OPEN, "admitted"],
    ["pure scope", { ...OPEN, kind: "scope" }, "admitted"],
    ["fact", { ...OPEN, kind: "fact" }, "not_delegable"],
    ["instruction", { ...OPEN, kind: "instruction" }, "admitted"],
    ["conflict", { ...OPEN, kind: "conflict" }, "not_delegable"],
    ["file name", FILE_NAME, "not_delegable"],
    ["mixed scope", { ...OPEN, kind: "scope", mixed: true }, "not_delegable"],
    ["unavailable advice", { ...OPEN, advice: UNAVAILABLE }, "not_delegable"],
    ["advisor failed", { ...OPEN, advice: { unavailable: { reason: "No checked recommendation is available for this question.", reason_code: "advisor_failed" } } }, "not_delegable"],
    ["no advice", { ...OPEN, advice: undefined }, "not_delegable"],
    ["answered", { ...OPEN, status: "answered", answer: "Cap it.", answered_at: "2026-10-19T09:00:00.000Z" }, "already_answered"],
  ]
  for (const [name, question, expected] of cases) {
    writeQuestions(directory, [question])
    expect([name, delegable(question), refusal(() => answerQuestion(directory, question.id, { adopt: true }))]).toEqual([name, expected === "admitted", expected])
  }
  expect(refusal(() => answerQuestion(directory, "q-missing", { adopt: true }))).toBe("not_found")
})

test("the drain applies queued typed, picked and adopted answers in order and refuses what the stored question cannot take", () => {
  const directory = matter()
  writeQuestions(directory, [
    OPEN,
    FILE_NAME,
    { ...OPEN, id: "q-fact", kind: "fact", question: "Who signs for the buyer?" },
    { ...OPEN, id: "q-law", kind: "instruction", question: "Which law governs?" },
  ])
  const requests = [
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-law", answer: "Singapore law." }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: FILE_NAME.id, option_id: "o1" }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-cap", adopt: true }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-fact", adopt: true }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-fact", option_id: "o9" }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-missing", adopt: true }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-fact", answer: "The director.", adopt: true } as unknown as LeanRequest),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-fact", decide_for_me: true } as unknown as LeanRequest),
  ]
  expect(drainLeanRequests(directory)).toEqual({ answered: 3, updated: 0, repairRound: false })
  expect(requests.map(request => outcome(directory, request))).toEqual([
    { status: "answered" },
    { status: "answered" },
    { status: "answered" },
    { status: "rejected", detail: expect.stringMatching(/^not_delegable: /) },
    { status: "rejected", detail: expect.stringMatching(/^invalid_answer: /) },
    { status: "rejected", detail: expect.stringMatching(/^not_found: /) },
    { status: "rejected", detail: "unreadable request" },
    { status: "rejected", detail: "unreadable request" },
  ])
  const stored = new Map(readQuestions(directory).map(question => [question.id, question]))
  expect(stored.get("q-law")).toMatchObject({ status: "answered", answer: "Singapore law." })
  expect(stored.get(FILE_NAME.id)).toMatchObject({ status: "answered", answer: "As my words ask", answer_kind: "option", option_id: "o1" })
  expect(stored.get("q-cap")).toMatchObject({ status: "answered", answer: RECOMMENDED, answer_kind: "adopted", option_id: "o2" })
  expect(stored.get("q-fact")).toMatchObject({ status: "open" })
})

test("an adoption replayed after a restart is already recorded and changes nothing", () => {
  const directory = matter()
  writeQuestions(directory, [OPEN])
  const request = enqueueLeanRequest(directory, { kind: "answer", question_id: "q-cap", adopt: true })
  // The drain applied the adoption, then stopped before moving the request to done.
  const applied = answerQuestion(directory, "q-cap", { adopt: true })
  expect(drainLeanRequests(directory)).toEqual({ answered: 1, updated: 0, repairRound: false })
  expect(outcome(directory, request)).toEqual({ status: "answered" })
  expect(readQuestions(directory)).toEqual([applied])
})

test("the same words typed or picked give the same answers text and keys; adopted, every step reads them as delegated", () => {
  const keys = (form: Parameters<typeof answerQuestion>[2]) => {
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
    writeQuestions(directory, [OPEN, { ...OPEN, id: "q-gate", stage: "gate", question: "Which requirement governs?", kind: "conflict", advice: UNAVAILABLE }])
    answerQuestion(directory, "q-cap", form)
    answerQuestion(directory, "q-gate", { text: "The first." })
    return Object.fromEntries([["answers", answeredQuestionsText(directory)], ...(["intake", "d1.change-plan", "gate", "finalize", "d1.change-review"] as const).map(step => [step, stageInputSha(directory, step)])])
  }
  const typed = keys({ text: RECOMMENDED })
  expect(typed.answers).toContain(`- [q-cap] Q: ${OPEN.question}\n  A: ${RECOMMENDED}`)
  expect(keys({ option_id: "o2" })).toEqual(typed)
  // An adopted recommendation reads as delegated in every step's answers, so each key moves.
  const adopted = keys({ adopt: true })
  expect(adopted.answers).toContain(`- [q-cap] Q: ${OPEN.question}\n  A (${ADOPTED_ANSWER}): ${RECOMMENDED}`)
  for (const step of Object.keys(typed)) expect([step, adopted[step] === typed[step]]).toEqual([step, false])
  // Other words are another answer.
  expect(keys({ option_id: "o1" }).intake).not.toBe(typed.intake)
})

// The spine records settled facts, decisions, recommendations, requirements
// and open items; it records no question and makes no choice for the lawyer.
const QUESTION = "Should late delivery carry a remedy?"
const LAW = "Which law should govern the supply note?"
const spine = (...entries: object[]) => ({ entries, limitations: [] })
const FACT = { kind: "fact", detail: "Delivery is within thirty days.", evidence: SUPPORT, limitations: [] }

function spineMatter(late: "open" | "answered" | "adopted") {
  const directory = matter()
  writeQuestions(directory, [
    { id: "q-law", stage: "intake", question: LAW, context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" },
    { id: "q-late", stage: "intake", question: QUESTION, context: "The terms are silent.", status: "open", kind: "judgment",
      advice: { options: [{ id: "o1", text: "No separate remedy." }, { id: "o2", text: "A price reduction." }], recommendation: { option: "o1", rationale: "The terms are silent.", basis: SUPPORT } } },
  ])
  if (late === "answered") answerQuestion(directory, "q-late", { text: "No separate remedy." })
  if (late === "adopted") answerQuestion(directory, "q-late", { adopt: true })
  return directory
}

test("the spine takes its five entry kinds and refuses a drafting choice or a question_id", () => {
  for (const late of ["open", "answered", "adopted"] as const) {
    const directory = spineMatter(late)
    expect(checkEvidence(directory, "spine", spine())).toEqual([])
    for (const kind of ["decision", "recommendation", "requirement", "fact", "open-item"]) expect(checkEvidence(directory, "spine", spine({ ...FACT, kind }))).toEqual([])
    expect(checkEvidence(directory, "spine", spine({ ...FACT, kind: "drafting-choice" }))).toEqual(["A spine entry kind must be decision, recommendation, requirement, fact or open-item."])
    for (const id of ["q-late", "q-law", "q-other"])
      expect(checkEvidence(directory, "spine", spine({ ...FACT, kind: "decision", question_id: id }))).toEqual(["A spine entry takes no question_id; remove it."])
    const prompt = evidencePrompt(directory, "spine")
    expect(prompt).not.toContain("question_id")
    expect(prompt).not.toContain("drafting-choice")
  }
  // The discrepancy assessment never took one either.
  expect(checkEvidence(spineMatter("adopted"), "discrepancy", { entries: [], gaps: [], limitations: [] })).toEqual([])
})

test("intake told a question is answered is told so however it was answered", () => {
  for (const late of ["answered", "adopted"] as const) {
    const directory = spineMatter(late)
    const reply = JSON.stringify({ questions: [{ question: QUESTION, kind: "judgment", context: "Still silent." }] })
    expect(intakeQuestions(directory, reply)!.issues.join(" ")).toContain("already has an answer")
  }
})

// "Decide for me" through the real coordinator, end to end.
const CHOSEN = "Add no separate late-delivery remedy."
const OTHER = "Reduce the price by one percent for each week of late delivery."
const BASIS = SUPPORT.map(row => ({ ...row, relevance: "The supplied terms fix the period for the supplier's delivery under this note and give no remedy for missing it." }))
// Wording of the deleted hand-back, and the web button's name: neither reaches a model.
const LEFT_TO_DRAFTER = /left (?:\w+ ){0,2}to the draft(?:er|ing)|left_to_draft/i
const BUTTON = /decide[\s_-]for[\s_-]me/i

test("an adopted recommendation runs through the queue as the lawyer's answer, reaches every stage as a delegated decision and settles at the gate", async () => {
  const directory = matter()
  const model = scriptedModel({
    ...happyPath,
    [AGENT.review]: () => NO_FINDINGS,
    [AGENT.changes]: noChanges,
    [AGENT.intake]: (_prompt, index) => index === 0
      ? JSON.stringify({ questions: [
        { question: QUESTION, kind: "judgment", context: "The supplied terms are silent on late delivery." },
        { question: LAW, kind: "instruction", context: "The terms name no governing law." },
      ] })
      : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms", acting_for: { party: "", basis: "none" }, dates: { document_date: { date: "", basis: "today" }, operative_dates: [] } }),
    [AGENT.advise]: prompt => prompt.includes(QUESTION)
      ? JSON.stringify({ options: [{ text: CHOSEN }, { text: OTHER }], recommendation: { option: CHOSEN, rationale: "The terms fix the delivery period and add no remedy.", basis: BASIS } })
      : JSON.stringify({ unavailable: { reason: "The sources name no governing law." } }),
  })
  const run = () => runLean(directory, { sessions: model.port })

  // 1. The run blocks on two advised questions; only the one with a recommendation can be adopted.
  expect((await run()).status).toBe("blocked")
  const asked = (text: string) => readQuestions(directory).find(question => question.question === text)!
  const late = asked(QUESTION), law = asked(LAW)
  expect([delegable(late), delegable(law)]).toEqual([true, false])
  const adopted = enqueueLeanRequest(directory, { kind: "answer", question_id: late.id, adopt: true })
  enqueueLeanRequest(directory, { kind: "answer", question_id: law.id, answer: "Singapore law." })
  expect(queuedAnswerIDs(directory)).toEqual(new Set([late.id, law.id]))

  // 2. The rerun records the adopted option's words as the answer.
  const start = model.prompts.length
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, adopted)).toEqual({ status: "answered" })
  expect(asked(QUESTION)).toMatchObject({ status: "answered", answer: CHOSEN, answer_kind: "adopted", option_id: "o1" })
  const line = `- [${late.id}] Q: ${QUESTION}\n  A (${ADOPTED_ANSWER}): ${CHOSEN}`
  expect(answeredQuestionsText(directory)).toContain(line)
  const prompts = model.prompts.slice(start)
  // The discrepancy record, made before intake asked, is not run again.
  expect(prompts.some(prompt => prompt.agent === AGENT.discrepancy)).toBe(false)
  for (const agent of [AGENT.intake, AGENT.research, AGENT["research-review"], AGENT.acts, AGENT.spine, AGENT.draft,
    AGENT.review, AGENT["review-independent"], AGENT["review-wording"], AGENT["review-client-disadvantage"], AGENT["change-plan"], AGENT.changes, AGENT["package-review"]])
    expect([agent, prompts.some(prompt => prompt.agent === agent && prompt.text.includes(line))]).toEqual([agent, true])
  const notes = prompts.find(prompt => prompt.agent === AGENT.finalize && prompt.text.includes("<section_kind>\ndelegated\n"))!.text
  expect(notes).toContain(`<question>\n${QUESTION}\n`)
  expect(notes).not.toContain(`Answer: ${CHOSEN}`)
  expect(notes).not.toContain(`A: ${CHOSEN}`)
  expect(notes).toContain(`<decision>\nDelegated by the lawyer.\nDecision taken under the delegated judgment: ${CHOSEN}\nReason: The terms fix the delivery period and add no remedy.`)
  expect(notes).not.toContain(late.id)
  for (const prompt of prompts) {
    expect([prompt.agent, LEFT_TO_DRAFTER.exec(prompt.text)?.[0], BUTTON.exec(prompt.text)?.[0]]).toEqual([prompt.agent, undefined, undefined])
    expect([prompt.agent, prompt.text.includes("A: undefined")]).toEqual([prompt.agent, false])
  }

  // 3. The gate settles it as delegated by the lawyer and lists nothing as left to the drafter.
  const gate = readLeanJson(directory, "gate.json") as GateResult
  expect(gate).toMatchObject({ status: "pass", delivery: "execution_ready", blocking: [] })
  expect(gate.items.filter(item => item.ref.question_id === late.id)).toEqual([])
  expect(JSON.stringify(gate)).not.toMatch(LEFT_TO_DRAFTER)
  writeOpenItems(directory, [...readOpenItems(directory), openItem("changes", "question", QUESTION, { question_id: late.id, pid: "p-none" })])
  expect(finalGate(directory).items.find(item => item.kind === "question" && item.ref.question_id === late.id))
    .toMatchObject({ status: "resolved", resolution: "delegated by the lawyer; the recommended option was adopted" })
  writeOpenItems(directory, readOpenItems(directory).filter(item => item.ref.question_id !== late.id))

  // 4. An unchanged Continue, and an adoption replayed into the queue, run no model stage.
  const settled = { prompts: model.prompts.length, sessions: model.sessions.length }
  const resume = enqueueLeanRequest(directory, { kind: "continue" })
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, resume)).toEqual({ status: "resume" })
  const replay = enqueueLeanRequest(directory, { kind: "answer", question_id: late.id, adopt: true })
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, replay)).toEqual({ status: "answered" })
  expect({ prompts: model.prompts.length, sessions: model.sessions.length }).toEqual(settled)

  // 5. The lawyer's own words replace the adopted answer and re-run what reads answers.
  const spines = model.count(AGENT.spine)
  const replaced = enqueueLeanRequest(directory, { kind: "answer", question_id: late.id, answer: OTHER })
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, replaced)).toEqual({ status: "answered" })
  const answer = asked(QUESTION)
  expect(answer).toMatchObject({ status: "answered", answer: OTHER })
  expect(answer).not.toHaveProperty("answer_kind")
  expect(model.count(AGENT.spine)).toBeGreaterThan(spines)
}, 180_000)
