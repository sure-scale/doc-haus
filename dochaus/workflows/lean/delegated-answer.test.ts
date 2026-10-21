import { afterEach, expect, test } from "bun:test"
import { checkAnswerClass, classifyAnswers } from "./answer-class"
import { checkDraftingNotes } from "./checks"
import { finalGate, type GateResult } from "./gate"
import { writeLeanPlan } from "./plan"
import { delegatedDecisionsText } from "./prompts"
import { openItem, quoteFound } from "./quotes"
import { enqueueLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import type { StageSessionPort } from "./stages"
import { answeredQuestionsText, answerQuestion, classifiedAnswer, delegatedQuestions, readLeanJson, readOpenItems, readQuestions, recordAnswerClasses, writeOpenItems, writeQuestions, type Question, type QuestionAdvice } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, NO_FINDINGS, noChanges, scriptedModel, SOURCE, SUPPORT } from "./test-support"

// A lawyer who delegates a question in their own words ("proceed on your own
// judgment") gets what Decide for me gives: the run reads the reply once, in
// a closed set (answered, delegated, declined) a focused job returns and code
// filters, and records a delegating reply to a question whose advice
// recommends an option as the adopted recommendation, keeping the lawyer's
// words. Every stage reads the adopted option's words as settled; the
// drafting notes list each delegated decision with its reason, evidence and
// attribution. Synthetic matter; only the model is scripted.

afterEach(cleanupMatters)

const DELEGATION = "No further information or documents are available. Proceed on your own professional judgment and disclose any unresolved points."
const RECOMMENDED = "Add no separate late-delivery remedy."
const OTHER = "Reduce the price by one percent for each week of late delivery."
const RATIONALE = "The terms fix the delivery period and add no remedy."
const ADVICE: QuestionAdvice = {
  options: [{ id: "o1", text: RECOMMENDED }, { id: "o2", text: OTHER }],
  recommendation: { option: "o1", rationale: RATIONALE, basis: SUPPORT },
}
const QUESTION = "Should late delivery carry a remedy?"
const ASKED: Question = { id: "q-late", stage: "intake", question: QUESTION, context: "The terms are silent.", status: "open", kind: "judgment", advice: ADVICE }
const answered = (question: Question, words = DELEGATION): Question => ({ ...question, status: "answered", answer: words, answered_at: "2026-10-21T09:00:00.000Z" })

test("a delegating reply to a question the recommendation can settle is recorded as the adopted recommendation, keeping the lawyer's words", () => {
  expect(classifiedAnswer(answered(ASKED), "delegated")).toEqual({
    ...answered(ASKED), answer: RECOMMENDED, answer_kind: "adopted", option_id: "o1", answer_class: "delegated", lawyer_words: DELEGATION,
  })
  // Pure scope as Decide for me takes it.
  expect(classifiedAnswer(answered({ ...ASKED, kind: "scope" }), "delegated")).toMatchObject({ answer: RECOMMENDED, answer_kind: "adopted" })
  // Nothing Decide for me could not adopt is adopted: the words stand, with their class.
  for (const question of [{ ...ASKED, kind: "fact" as const }, { ...ASKED, kind: "instruction" as const }, { ...ASKED, kind: "scope" as const, mixed: true as const },
    { ...ASKED, advice: { options: [], unavailable: { reason: "Nothing in the sources supports one option." } } }, { ...ASKED, advice: undefined }])
    expect(classifiedAnswer(answered(question), "delegated")).toEqual({ ...answered(question), answer_class: "delegated" })
  expect(classifiedAnswer(answered(ASKED), "declined")).toEqual({ ...answered(ASKED), answer_class: "declined" })
  expect(classifiedAnswer(answered(ASKED, "Add no remedy."), "answered")).toEqual({ ...answered(ASKED, "Add no remedy."), answer_class: "answered" })
})

test("a class is recorded once, only on the words that were read; later words of the lawyer are read afresh", () => {
  const directory = matter()
  writeQuestions(directory, [answered(ASKED), answered({ ...ASKED, id: "q-law", question: "Which law governs?", kind: "instruction" }, "Singapore law.")])
  recordAnswerClasses(directory, new Map([["q-late", { words: DELEGATION, answer_class: "delegated" as const }], ["q-law", { words: "English law.", answer_class: "answered" as const }]]))
  const [late, law] = readQuestions(directory)
  expect(late).toMatchObject({ answer: RECOMMENDED, answer_kind: "adopted", answer_class: "delegated", lawyer_words: DELEGATION })
  // The lawyer's own words stay quotable as theirs.
  expect(quoteFound(directory, { source_id: "request", quote: "Proceed on your own professional judgment" })).toBe(true)
  // Words that changed since they were read keep no class.
  expect(law).not.toHaveProperty("answer_class")
  // A replayed class changes nothing.
  recordAnswerClasses(directory, new Map([["q-late", { words: RECOMMENDED, answer_class: "declined" as const }]]))
  expect(readQuestions(directory)[0]).toEqual(late)
  // The lawyer's later words replace the decision and are read again.
  const replaced = answerQuestion(directory, "q-late", { text: OTHER })
  expect(replaced).toMatchObject({ status: "answered", answer: OTHER })
  for (const key of ["answer_kind", "option_id", "answer_class", "lawyer_words"]) expect(replaced).not.toHaveProperty(key)
})

test("a delegated decision reaches every stage as Decide for me does; only the notes' input names how it was delegated", () => {
  const keys = (record: (directory: string) => void) => {
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
    writeQuestions(directory, [ASKED])
    record(directory)
    return Object.fromEntries([["answers", answeredQuestionsText(directory)], ...(["discrepancy", "intake", "spine", "d1.change-plan", "d1.change-review", "gate", "finalize"] as const).map(step => [step, stageInputSha(directory, step)])])
  }
  const typed = keys(directory => answerQuestion(directory, "q-late", { text: RECOMMENDED }))
  const adopted = keys(directory => answerQuestion(directory, "q-late", { adopt: true }))
  const delegated = keys(directory => {
    answerQuestion(directory, "q-late", { text: DELEGATION })
    recordAnswerClasses(directory, new Map([["q-late", { words: DELEGATION, answer_class: "delegated" as const }]]))
  })
  expect(delegated.answers).toBe(`- [q-late] Q: ${QUESTION}\n  A: ${RECOMMENDED}`)
  const { finalize: typedNotes, ...typedSteps } = typed
  const { finalize: adoptedNotes, ...adoptedSteps } = adopted
  const { finalize: delegatedNotes, ...delegatedSteps } = delegated
  expect(adoptedSteps).toEqual(typedSteps)
  expect(delegatedSteps).toEqual(typedSteps)
  // The notes disclose a delegated decision, so their key reads it; typed words disclose none.
  expect(new Set([typedNotes, adoptedNotes, delegatedNotes]).size).toBe(3)
})

test("the drafting notes' input gives each delegated question its attribution and the decision's reason, evidence and alternatives, or why none was taken", () => {
  const directory = matter()
  writeQuestions(directory, [
    { ...ASKED, id: "q-a" },
    answered({ ...ASKED, id: "q-b" }),
    answered({ ...ASKED, id: "q-c", kind: "fact", question: "Who signs for the buyer?" }),
    answered({ ...ASKED, id: "q-d", question: "Should title pass on delivery?" }, "I do not know."),
    answered({ ...ASKED, id: "q-e", question: "Which law governs?", kind: "instruction" }, "Singapore law."),
    { ...ASKED, id: "q-f", question: "Should risk pass on delivery?" },
  ])
  answerQuestion(directory, "q-a", { adopt: true })
  recordAnswerClasses(directory, new Map([
    ["q-b", { words: DELEGATION, answer_class: "delegated" as const }],
    ["q-c", { words: DELEGATION, answer_class: "delegated" as const }],
    ["q-d", { words: "I do not know.", answer_class: "declined" as const }],
    ["q-e", { words: "Singapore law.", answer_class: "answered" as const }],
  ]))
  expect(delegatedQuestions(directory).map(question => question.id)).toEqual(["q-a", "q-b", "q-c", "q-d"])
  const decided = (id: string, how: string) => [
    `- [${id}] Q: ${QUESTION}`,
    `  ${how}`,
    `  Decision taken, for the lawyer to confirm: ${RECOMMENDED}`,
    `  Reason: ${RATIONALE}`,
    `  Evidence: "within thirty days of the order date" [source: ${SOURCE}]`,
    `  Alternative not taken: ${OTHER}`,
  ].join("\n")
  expect(delegatedDecisionsText(directory)).toBe([
    decided("q-a", "Delegated by the lawyer."),
    decided("q-b", `Delegated by the lawyer, in their words: "${DELEGATION}"`),
    `- [q-c] Q: Who signs for the buyer?\n  Delegated by the lawyer, in their words: "${DELEGATION}"\n  No decision taken: it asks for a fact or an instruction that only the lawyer can give.`,
    `- [q-d] Q: Should title pass on delivery?\n  Not answered by the lawyer, in their words: "I do not know."\n  No decision taken: the point stays open for the lawyer.`,
  ].join("\n"))
})

test("the drafting notes must list each delegated question before their last attempt", () => {
  const directory = matter()
  writeQuestions(directory, [answered(ASKED)])
  recordAnswerClasses(directory, new Map([["q-late", { words: DELEGATION, answer_class: "delegated" as const }]]))
  const silent = "# Drafting notes\n\nThe draft states the delivery period from the supply terms.\n"
  expect(checkDraftingNotes(directory, silent, []).issues).toEqual([
    'The drafting notes do not list these delegated questions: q-late. For each, say in words what was decided and why, or why the point stays open, ending with its exact id as "[ref: <id>]".',
  ])
  expect(checkDraftingNotes(directory, silent, [], { final: true }).issues).toEqual([])
  expect(checkDraftingNotes(directory, `${silent}\nThe lawyer delegated the late-delivery remedy; none was added, as the terms fix the period. [ref: q-late]\n`, []).issues).toEqual([])
})

function port(reply: (text: string, index: number) => string | Promise<string>) {
  const prompts: { agent: string; text: string }[] = []
  const sessions: StageSessionPort = {
    async createSession() {
      return `ses_${prompts.length}`
    },
    async prompt(input) {
      prompts.push(input)
      return reply(input.text, prompts.length - 1)
    },
  }
  return { sessions, prompts }
}

test("the answer job's reply is one of answered, delegated or declined, or it is retried", () => {
  expect(checkAnswerClass('{"class":"delegated"}')).toEqual({ issues: [], answer_class: "delegated" })
  expect(checkAnswerClass('```json\n{"class":"declined"}\n```')).toEqual({ issues: [], answer_class: "declined" })
  for (const reply of ['{"class":"adopted"}', '{"class":["delegated"]}', '{"answer":"delegated"}', "delegated"])
    expect(checkAnswerClass(reply).issues).toHaveLength(1)
})

test("each unread answer in the lawyer's words is read once by its own job; a reply never passing its check stands as the lawyer's answer, a runtime error records nothing", async () => {
  const directory = matter()
  const state = readLeanJson(directory, "state.json") as Parameters<typeof classifyAnswers>[1]
  writeQuestions(directory, [answered(ASKED), answered({ ...ASKED, id: "q-picked" }, OTHER), { ...ASKED, id: "q-open" }])
  answerQuestion(directory, "q-picked", { option_id: "o2" })
  const job = port(() => '{"class":"delegated"}')
  await classifyAnswers(directory, state, { sessions: job.sessions })
  // Only the typed words are read; a picked option is the lawyer's own choice.
  expect(job.prompts).toHaveLength(1)
  expect(job.prompts[0]).toMatchObject({ agent: AGENT.answer })
  expect(job.prompts[0]!.text).toContain(QUESTION)
  expect(job.prompts[0]!.text).toContain(DELEGATION)
  expect(readQuestions(directory).find(question => question.id === "q-late")).toMatchObject({ answer: RECOMMENDED, answer_kind: "adopted", answer_class: "delegated", lawyer_words: DELEGATION })
  // Read once: nothing left to read asks nothing.
  await classifyAnswers(directory, state, { sessions: job.sessions })
  expect(job.prompts).toHaveLength(1)

  answerQuestion(directory, "q-late", { text: DELEGATION })
  const failing = port(() => { throw new Error("provider unavailable") })
  await classifyAnswers(directory, state, { sessions: failing.sessions })
  expect(readQuestions(directory).find(question => question.id === "q-late")).not.toHaveProperty("answer_class")
  const unreadable = port(() => '{"class":"maybe"}')
  await classifyAnswers(directory, state, { sessions: unreadable.sessions })
  expect(unreadable.prompts).toHaveLength(4)
  expect(readQuestions(directory).find(question => question.id === "q-late")).toMatchObject({ answer: DELEGATION, answer_class: "answered" })
})

test("a question delegated in the lawyer's own words runs through the queue as a recorded decision: every stage follows it, the notes list it and nothing is left open for it", async () => {
  const directory = matter()
  const model = scriptedModel({
    ...happyPath,
    [AGENT.review]: () => NO_FINDINGS,
    [AGENT.changes]: noChanges,
    [AGENT.intake]: (_prompt, index) => index === 0
      ? JSON.stringify({ questions: [{ question: QUESTION, kind: "judgment", context: "The supplied terms are silent on late delivery." }] })
      : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" }),
    [AGENT.advise]: () => JSON.stringify({ options: [{ text: RECOMMENDED }, { text: OTHER }], recommendation: { option: 0, rationale: RATIONALE, basis: SUPPORT } }),
    [AGENT.answer]: prompt => JSON.stringify({ class: prompt.includes(DELEGATION) ? "delegated" : "answered" }),
  })
  const run = () => runLean(directory, { sessions: model.port })

  expect((await run()).status).toBe("blocked")
  const asked = () => readQuestions(directory).find(question => question.question === QUESTION)!
  const id = asked().id
  enqueueLeanRequest(directory, { kind: "answer", question_id: id, answer: DELEGATION })
  const start = model.prompts.length
  expect((await run()).status).toBe("completed")

  // Recorded once, as the adopted recommendation, with the lawyer's words.
  expect(asked()).toMatchObject({ status: "answered", answer: RECOMMENDED, answer_kind: "adopted", option_id: "o1", answer_class: "delegated", lawyer_words: DELEGATION })
  const prompts = model.prompts.slice(start)
  expect(prompts.filter(prompt => prompt.agent === AGENT.answer)).toHaveLength(1)
  // Bound: every stage reads the decision as the lawyer's settled answer, never the delegating words.
  const line = `- [${id}] Q: ${QUESTION}\n  A: ${RECOMMENDED}`
  for (const agent of [AGENT.discrepancy, AGENT.intake, AGENT.research, AGENT.spine, AGENT.draft, AGENT.review, AGENT["change-plan"], AGENT.changes, AGENT["package-review"]])
    expect([agent, prompts.some(prompt => prompt.agent === agent && prompt.text.includes(line))]).toEqual([agent, true])
  for (const prompt of prompts.filter(prompt => prompt.agent !== AGENT.answer && prompt.agent !== AGENT.finalize))
    expect([prompt.agent, prompt.text.includes(DELEGATION)]).toEqual([prompt.agent, false])
  // Disclosed: the notes are given the decision with its attribution, reason and evidence, and list it.
  const notesPrompt = prompts.findLast(prompt => prompt.agent === AGENT.finalize)!.text
  expect(notesPrompt).toContain(`<delegated_decisions>\n- [${id}] Q: ${QUESTION}\n  Delegated by the lawyer, in their words: "${DELEGATION}"\n  Decision taken, for the lawyer to confirm: ${RECOMMENDED}\n  Reason: ${RATIONALE}`)
  expect(readLeanJson(directory, "state.json")).toMatchObject({ status: "completed" })
  const notes = readLeanJson(directory, "gate.json") as GateResult
  expect(notes.items.filter(item => item.ref.question_id === id)).toEqual([])
  expect(readOpenItems(directory).filter(item => item.ref.question_id === id)).toEqual([])
  // A question item about it is settled as delegated by the lawyer, not as the lawyer's own answer.
  writeOpenItems(directory, [...readOpenItems(directory), openItem("changes", "question", QUESTION, { question_id: id, pid: "p-none" })])
  expect(finalGate(directory).items.find(item => item.kind === "question" && item.ref.question_id === id))
    .toMatchObject({ status: "resolved", resolution: "delegated by the lawyer; the recommended option was adopted" })
  writeOpenItems(directory, readOpenItems(directory).filter(item => item.ref.question_id !== id))

  // An unchanged Continue reads nothing again and runs no model stage.
  const settled = model.prompts.length
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await run()).status).toBe("completed")
  expect(model.prompts.length).toBe(settled)
}, 180_000)
