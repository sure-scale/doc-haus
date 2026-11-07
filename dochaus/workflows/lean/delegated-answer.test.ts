import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { checkAnswerClass, classifyAnswers } from "./answer-class"
import { checkDraftingNotes } from "./checks"
import { OutputLimitError } from "./native-stage-result"
import { finalGate, type GateResult } from "./gate"
import { factOpenItems } from "./fact-answers"
import { writeLeanPlan } from "./plan"
import { delegatedDecisionsText, stagePrompt } from "./prompts"
import { derivedRecordNote, openItem, quoteFound, quoteMissIssue } from "./quotes"
import { enqueueLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import type { StageSessionPort } from "./stages"
import { ADOPTED_ANSWER, answeredQuestionsText, answerQuestion, classifiedAnswer, delegatedQuestions, readLeanJson, readOpenItems, readQuestions, recordAnswerClasses, writeLeanText, writeOpenItems, writeQuestions, type Question, type QuestionAdvice } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, NO_FINDINGS, noChanges, scriptedModel, SOURCE, SUPPORT } from "./test-support"

// A lawyer who delegates a question in their own words ("proceed on your own
// judgment") gets what Decide for me gives: the run reads the reply once, in
// a closed set (answered, delegated, declined) a focused job returns and code
// filters, and records a delegating reply to a question whose advice
// recommends an option as the adopted recommendation, keeping the lawyer's
// words. Every stage reads the adopted option's words as settled; the
// drafting notes list each delegated decision with its reason, evidence and
// attribution. Synthetic matter; only the model is scripted.

const workspace = process.env.WORKSPACE_ROOT
const roots: string[] = []
afterEach(() => {
  if (workspace === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = workspace
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
  cleanupMatters()
})

const DELEGATION = "No further information or documents are available. Proceed on your own professional judgment and disclose any unresolved points."
const RECOMMENDED = "Add no separate late-delivery remedy."
const OTHER = "Reduce the price by one percent for each week of late delivery."
const RATIONALE = "The terms fix the delivery period and add no remedy."
const ADVICE: QuestionAdvice = {
  options: [{ id: "o1", text: RECOMMENDED }, { id: "o2", text: OTHER }],
  recommendation: { option: "o1", rationale: RATIONALE, basis: SUPPORT },
}
const QUESTION = "Should late delivery carry a remedy?"
const BASIS = SUPPORT.map(row => ({ ...row, relevance: "The supplied terms fix the period for the supplier's delivery under this note and give no remedy for missing it." }))
const ASKED: Question = { id: "q-late", stage: "intake", question: QUESTION, context: "The terms are silent.", status: "open", kind: "judgment", advice: ADVICE }
const answered = (question: Question, words = DELEGATION): Question => ({ ...question, status: "answered", answer: words, answered_at: "2026-10-21T09:00:00.000Z" })

test("a delegating reply to a question the recommendation can settle is recorded as the adopted recommendation, keeping the lawyer's words", () => {
  expect(classifiedAnswer(answered(ASKED), "delegated")).toEqual({
    ...answered(ASKED), answer: RECOMMENDED, answer_kind: "adopted", option_id: "o1", answer_class: "delegated", lawyer_words: DELEGATION,
  })
  // Pure scope as Decide for me takes it, and an instruction question: words
  // leaving it to the drafter's judgment are the instruction it asks for.
  for (const kind of ["scope", "instruction"] as const)
    expect(classifiedAnswer(answered({ ...ASKED, kind }), "delegated")).toMatchObject({ answer: RECOMMENDED, answer_kind: "adopted", lawyer_words: DELEGATION })
  // A fact, approval or source conflict is not settled by delegating: the words stand, with their class.
  for (const question of [{ ...ASKED, kind: "fact" as const }, { ...ASKED, kind: "conflict" as const }, { ...ASKED, kind: "scope" as const, mixed: true as const },
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
  // The recommendation adopted for them is not: no review may cite it as what the lawyer asked.
  expect(quoteFound(directory, { source_id: "request", quote: RECOMMENDED })).toBe(false)
  expect(quoteFound(directory, { source_id: "answer:q-late", quote: RECOMMENDED })).toBe(false)
  expect(derivedRecordNote(directory, RECOMMENDED)).toContain("recommendation adopted for question q-late")
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

test("words quoted as the lawyer's that are the answers record or the firm's drafting preferences are named as such, without the advice to copy them", () => {
  const directory = matter()
  writeQuestions(directory, [ASKED, { ...ASKED, id: "q-law", question: "Which law governs the note?", kind: "instruction" }])
  answerQuestion(directory, "q-law", { text: "Singapore law." })
  answerQuestion(directory, "q-late", { adopt: true })
  const miss = (quote: string, format: "json" | "prose", options = {}) => quoteMissIssue(directory, "sources[0]", { source_id: "request", quote }, format, options)
  // The record's question is the run's, not the lawyer's: their own answer is pointed to.
  expect(derivedRecordNote(directory, "Which law governs the note?")).toBe(`Those words are the run's record of question q-law, not the lawyer's words; quote the lawyer's own answer under "request".`)
  // An adopted question's record points to the adopted answer only where the check accepts that name.
  expect(derivedRecordNote(directory, QUESTION, { adopted: true })).toBe(`Those words are the run's record of question q-late, not the lawyer's words; quote the adopted answer under "adopted:q-late".`)
  expect(derivedRecordNote(directory, `A (${ADOPTED_ANSWER})`)).toBe("Those words are the run's record of question q-late, not the lawyer's words; quote the lawyer's request or the matter sources that support the point instead.")
  for (const format of ["json", "prose"] as const) {
    const issue = miss("Which law governs the note?", format)
    expect(issue).toBe(`sources[0] quotes "Which law governs the note?" as from "request", but those words are not in the request, user updates or any one of the lawyer's answers. Those words are the run's record of question q-law, not the lawyer's words; quote the lawyer's own answer under "request".`)
    expect(miss(QUESTION, format, { adopted: true })).not.toContain("read-document")
  }
  // The lawyer's own answer is still accepted as theirs.
  expect(quoteFound(directory, { source_id: "request", quote: "Singapore law." })).toBe(true)

  // The firm's saved drafting preferences are not a source: a presentation change cites nothing.
  const root = mkdtempSync(path.join(tmpdir(), "dh-lean-preferences-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences/preferences.json"), JSON.stringify({ houseStyle: "Put each defined term in bold where it is defined." }))
  const preferences = "Those words are the firm's drafting preferences, not a source. A presentation change needs no supporting quote: leave its sources empty and say that it applies the saved preferences, in its limitations or, where it has none, its reason."
  expect(derivedRecordNote(directory, "Put each defined term in bold")).toBe(preferences)
  for (const format of ["json", "prose"] as const) {
    const issue = miss("Put each defined term in bold", format)
    expect(issue).toEndWith(preferences)
    expect(issue).not.toContain("read-document")
  }
  // Words found nowhere keep the generic advice.
  expect(derivedRecordNote(directory, "Deliver the goods by sea.")).toBe("")
  expect(miss("Deliver the goods by sea.", "json")).toContain("read-document")
})

test("a delegated decision reaches every stage as Decide for me does, marked as delegated; only the notes' input names the lawyer's words", () => {
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
  expect(typed.answers).toBe(`- [q-late] Q: ${QUESTION}\n  A: ${RECOMMENDED}`)
  expect(delegated.answers).toBe(`- [q-late] Q: ${QUESTION}\n  A (${ADOPTED_ANSWER}): ${RECOMMENDED}`)
  const { finalize: typedNotes, ...typedSteps } = typed
  const { finalize: adoptedNotes, ...adoptedSteps } = adopted
  const { finalize: delegatedNotes, ...delegatedSteps } = delegated
  expect(delegatedSteps).toEqual(adoptedSteps)
  // Every step from intake on reads the decision as delegated, not as the
  // lawyer's own words; the discrepancy record, before intake, reads no answer.
  for (const step of Object.keys(typedSteps)) expect([step, adoptedSteps[step] === typedSteps[step]]).toEqual([step, step === "discrepancy"])
  // The notes are also given the lawyer's own delegating words.
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
    `- [q-c] Q: Who signs for the buyer?\n  Delegated by the lawyer, in their words: "${DELEGATION}"\n  No decision taken: it needs a fact or an approval that only the lawyer can give.`,
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

test("an answer read whose reply is cut off at its output limit stands as the lawyer's answer, as a reply never passing its check does", async () => {
  const directory = matter()
  const state = readLeanJson(directory, "state.json") as Parameters<typeof classifyAnswers>[1]
  writeQuestions(directory, [answered(ASKED)])
  const cutOff = port(() => { throw new OutputLimitError("The model's reply reached its output limit before giving any answer.") })
  await classifyAnswers(directory, state, { sessions: cutOff.sessions })
  expect(cutOff.prompts).toHaveLength(1)
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
      : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms", acting_for: { party: "", basis: "none" }, dates: { document_date: { date: "", basis: "today" }, operative_dates: [] } }),
    [AGENT.advise]: () => JSON.stringify({ options: [{ text: RECOMMENDED }, { text: OTHER }], recommendation: { option: RECOMMENDED, rationale: RATIONALE, basis: BASIS } }),
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
  // Bound: every stage reads the decision as a delegated one, never the delegating words.
  const line = `- [${id}] Q: ${QUESTION}\n  A (${ADOPTED_ANSWER}): ${RECOMMENDED}`
  expect(prompts.some(prompt => prompt.agent === AGENT.discrepancy)).toBe(false)
  for (const agent of [AGENT.intake, AGENT.research, AGENT.spine, AGENT.draft, AGENT.review, AGENT["change-plan"], AGENT.changes, AGENT["package-review"]])
    expect([agent, prompts.some(prompt => prompt.agent === agent && prompt.text.includes(line))]).toEqual([agent, true])
  for (const prompt of prompts.filter(prompt => prompt.agent !== AGENT.answer && prompt.agent !== AGENT.finalize))
    expect([prompt.agent, prompt.text.includes(DELEGATION)]).toEqual([prompt.agent, false])
  // Disclosed: the notes are given the decision with its attribution, reason and evidence, and list it.
  const notesPrompt = prompts.findLast(prompt => prompt.agent === AGENT.finalize)!.text
  expect(notesPrompt).toContain(`<delegated_decisions>\n- [${id}] Q: ${QUESTION}\n  Delegated by the lawyer, in their words: "${DELEGATION}"\n  Decision taken, for the lawyer to confirm: ${RECOMMENDED}\n  Reason: ${RATIONALE}`)
  // The notes' list of questions points to that section rather than calling it answered.
  expect(notesPrompt).toContain(`- (delegated; see delegated_decisions) ${QUESTION}\n`)
  expect(notesPrompt).not.toContain(`Answer: ${RECOMMENDED}`)
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

test("an explicit factual delegation permits a marked review draft while a declined judgment keeps its existing disclosure", () => {
  const directory = matter()
  const fact = { ...ASKED, id: "q-sign", kind: "fact" as const, question: "Who signs for the buyer?" }
  const unsure = { ...ASKED, id: "q-title", question: "Should title pass on delivery?" }
  const asked = { ...ASKED, id: "q-risk", kind: "fact" as const, question: "When does risk pass?" }
  writeQuestions(directory, [answered(ASKED), answered(fact), answered(unsure, "I do not know."), answered(asked)])
  recordAnswerClasses(directory, new Map([
    ["q-late", { words: DELEGATION, answer_class: "delegated" as const }],
    ["q-sign", { words: DELEGATION, answer_class: "delegated" as const, fact_status: "provisional" as const }],
    ["q-title", { words: "I do not know.", answer_class: "declined" as const }],
    ["q-risk", { words: DELEGATION, answer_class: "delegated" as const, fact_status: "provisional" as const }],
  ]))
  // The fixture's permission says to proceed and disclose unresolved points.
  // It supplies neither the signatory nor the risk-transfer fact.
  const facts = factOpenItems(directory)
  const risk = facts.find(item => item.ref.question_id === "q-risk")!
  const changesItem = openItem("changes", "question", "Question for the lawyer: When does risk pass?", { question_id: "q-risk", placeholder: risk.ref.placeholder })
  writeOpenItems(directory, [changesItem, ...facts.filter(item => item.ref.question_id !== "q-risk")])
  const items = () => finalGate(directory).items.filter(item => item.kind === "question")
  expect(answeredQuestionsText(directory)).toContain("the fact remains unknown; the lawyer permits a review draft")
  expect(answeredQuestionsText(directory)).toContain(`- [q-sign] Q: Who signs for the buyer?`)
  expect(answeredQuestionsText(directory)).toContain(`- [q-title] Q: Should title pass on delivery?\n  A (not answered by the lawyer, so it stays open: every paragraph that depends on it keeps a visible [open item: what the lawyer must supply or confirm] marker, also where its wording is drafted from the sources; never remove that marker or treat the point as settled, and do not ask it again): I do not know.`)

  writeLeanText(directory, "drafting-notes.md", "# Notes\n\nNothing listed.\n")
  expect(items().map(item => [item.ref.question_id, item.status])).toEqual([
    ["q-risk", "open"], ["q-sign", "open"], ["q-title", "open"],
  ])
  expect(finalGate(directory).status).toBe("blocked")
  writeLeanText(directory, "drafting-notes.md", "# Notes\n\nStill needs the signatory [ref: q-sign]. Title [ref: q-title]. Risk [ref: q-risk].\n")
  // Notes alone do not make an unknown fact honest in the document body.
  expect(items().map(item => [item.ref.question_id, item.status])).toEqual([
    ["q-risk", "open"], ["q-sign", "open"], ["q-title", "disclosed"],
  ])
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p-open", sha256: "0", text: facts.map(item => item.ref.placeholder).join("\n") }]))
  expect(items().map(item => [item.ref.question_id, item.status])).toEqual([
    ["q-risk", "disclosed"], ["q-sign", "disclosed"], ["q-title", "disclosed"],
  ])
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
})

test("a provisional factual reply requires its affected paragraph marker and notes until the lawyer supplies the fact", () => {
  const directory = matter()
  const fact = { ...ASKED, id: "q-sign", kind: "fact" as const, question: "Who signs for the buyer?", pid: "p-sign" }
  writeQuestions(directory, [answered(fact)])
  recordAnswerClasses(directory, new Map([["q-sign", { words: DELEGATION, answer_class: "delegated" as const, fact_status: "provisional" as const }]]))
  expect(answeredQuestionsText(directory)).toContain("a visible open-item marker in every dependent paragraph and disclosure in the notes")
  expect(stagePrompt(directory, "d1.change-plan", { contextOnly: true })).toContain("An answer whose label says the point stays open settles nothing: keep its visible open-item marker.")
  writeOpenItems(directory, [openItem("changes", "question", "Question for the lawyer: Who signs for the buyer?", { question_id: "q-sign", pid: "p-sign" })])
  writeLeanText(directory, "drafting-notes.md", "# Notes\n\nStill needs the signatory [ref: q-sign].\n")
  const body = (text: string) => writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p-sign", sha256: "0", text }]))
  const item = () => finalGate(directory).items.find(row => row.kind === "question" && row.ref.question_id === "q-sign")
  body("Signed for the buyer by the director [open item: confirm who signs for the buyer].")
  expect(item()).toMatchObject({ status: "disclosed", resolution: "the fact remains unknown; marked in the body and named in the drafting notes" })
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
  body("Signed for the buyer by the director.")
  expect(item()).toMatchObject({ status: "open", resolution: "the fact remains unknown; its body placeholder is missing" })
  expect(finalGate(directory)).toMatchObject({ status: "blocked" })
  // The lawyer's actual factual answer settles it; the earlier permission did not.
  answerQuestion(directory, "q-sign", { text: "The director signs." })
  recordAnswerClasses(directory, new Map([["q-sign", { words: "The director signs.", answer_class: "answered" as const, fact_status: "supplied" as const }]]))
  expect(item()).toMatchObject({ status: "resolved", resolution: "answered by the lawyer" })
  expect(answeredQuestionsText(directory)).not.toContain("open item")
})
