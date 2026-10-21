import { afterEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { checkEvidence, evidencePrompt } from "./evidence"
import { intakeQuestions } from "./intake"
import { leanSteps, writeLeanPlan } from "./plan"
import { withSourceSnapshot } from "./quotes"
import { drainLeanRequests, enqueueLeanRequest, type LeanRequest, type StoredLeanRequest } from "./requests"
import { stageInputSha } from "./runner"
import { stageInputs } from "./stages"
import { answeredQuestionsText, answerQuestion, delegable, hashStageInputs, LeanQuestionError, readLeanJson, readQuestions, writeLeanText, writeQuestions, type Question, type QuestionAdvice } from "./state"
import { cleanupMatters, matter, SUPPORT } from "./test-support"

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
  expect(adopted.answers).toContain(`- [q-cap] Q: ${OPEN.question}\n  A (delegated by the lawyer; the recommended option was adopted): ${RECOMMENDED}`)
  for (const step of Object.keys(typed)) expect([step, adopted[step] === typed[step]]).toEqual([step, false])
  // Other words are another answer.
  expect(keys({ option_id: "o1" }).intake).not.toBe(typed.intake)
})

// Inputs pinned on the code before hand-backs existed; hashed under a fixed
// label so a prompt version change does not move them. The gate's also keys
// its contract. A draft's keys its outline contract and what it reads of its
// own plan entry, never answers. Change-review, finalize and the gate were
// re-pinned when stage keys stopped carrying when a question was answered (and
// its kind, options and advice), and the draft when the outline gained its
// Form line and when part labels were planned unheaded. The checklist check, change plan, changes and change-review were
// re-pinned when reviews were saved as typed findings (dN/<stage>.json) and the
// change plan began keying the findings earlier documents forward to it. The
// same pins hold for typed or picked words; an adopted recommendation reads as
// delegated, which moves every key that reads the answers and no other.
const ANSWERS = "- [package-scope-d1] Q: Which documents are in scope?\n  A: Only the supply note.\n- [q-law] Q: Which law should govern?\n  A: Singapore law."
const ADOPTED_ANSWERS = ANSWERS.replace("  A: Singapore law.", "  A (delegated by the lawyer; the recommended option was adopted): Singapore law.")
// A draft keys its outline contract and plan entry, never answers.
const UNANSWERED_STEPS = ["d1.draft"]
const INPUTS: Record<string, string> = {
  "discrepancy": "8dac7eefaf1e988e35ffbe6c7f765a951943e9b1299546a9b94dc10e99950a72",
  "intake": "1e491eea042fb2cb31d8340459e58993849d22ce04ec26793d0e495a840061c6",
  "research": "64a0d22c3dd780907955b2e56b3122b94ea72101092d694f899a15e2af0cb13b",
  "research-review": "f30467728cba33fbf89ff10cfba8903cd7aa9647cbd6849d4a3caf293f27a65b",
  "spine": "c038d3823819fc2e8ef176dc1cb855fd7f7a87e1fe077efb1c2aee41f922e076",
  "d1.draft": "68f34f0d87312276c4f5d676b490753a5050865d40c0448aaae368a7d6252376",
  "d1.review": "e80f5bbe084c8ef90d2864358d2caafa09139822584fb0f12cb14d0976e76416",
  "d1.review-independent": "9f1fe22f3e0db01ef5fffe96a09ddf917b5e371a89c58988eb875bcfbbae4aa4",
  "d1.review-wording": "7f86222090d292207f576968b1171dab74bedab8cf5dbc4e3593752d2584b6f5",
  "d1.review-client-disadvantage": "be6786003165e1b19571bc525b2d6dd3ba4afad3a7f707c7f2e58fe4f69ffac1",
  "d1.change-plan": "64bc4586e950cb2aaee6d57179adf09500aee75e27b358e054ae5076f770d7d8",
  "d1.change-prepare": "0e65f05f5cdbac92e874b5150ac1dd21f355a34be5b6a1b918b1be46f580591f",
  "d1.changes": "6c1572e8f947889bf8dcb84f6bb75ba8380fbade2de920ba3a2dcb28c2cf3449",
  "d1.change-review": "0c98552192e0c314b273e3508140684f7d71a97934b771f9e644e5a8f3a0eddc",
  "d1.apply": "3fc66df6877ed50af811c9f7a5bbdf1ccb1c269dc8b176cf392941cd2da83149",
  "finalize": "fcb7ded51c54fad4bb3fd616ba57fbdfe8aa15a2dc1becf5ec5df4e69b537793",
  "checklist-check": "725f51d8b50889f0f431847a3299d55942561e8fabc2bece75410442fd0fa5a7",
  "package-review": "dde115af0ba46b8cb1ecf9bf7eb0c40aff3b78f4d6be9e0f0da3dff2d5d4066f",
  "gate": "0950183c873a0bd663d5367c064a7013025084deb8884311b2f38a54fcb8e229",
}

test("the answers text and every step's inputs are pinned, however the lawyer gave the answer", () => {
  const previous = process.env.WORKSPACE_ROOT
  const workspace = mkdtempSync(path.join(tmpdir(), "lean-dfm-workspace-"))
  process.env.WORKSPACE_ROOT = workspace
  const law: Question = { id: "q-law", stage: "intake", question: "Which law should govern?", context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" }
  const advised: Question = { id: "q-law", stage: "intake", question: "Which law should govern?", context: "Legal basis.", status: "open", kind: "judgment",
    advice: { options: [{ id: "o1", text: "English law." }, { id: "o2", text: "Singapore law." }], recommendation: { option: "o2", rationale: "The terms are Singaporean.", basis: SUPPORT } } }
  try {
    for (const form of [undefined, { option_id: "o2" }, { adopt: true as const }]) {
      const directory = matter()
      writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
      writeQuestions(directory, [
        form ? advised : law,
        { id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", pid: "p-1", document: "d1", status: "open" },
        { id: "package-scope-d1", stage: "intake", question: "Which documents are in scope?", context: "Scope.", status: "answered", answer: "Only the supply note.", answered_at: "2026-10-18T09:05:00.000Z" },
      ])
      if (form) answerQuestion(directory, "q-law", form)
      const adopted = form !== undefined && "adopt" in form
      expect(answeredQuestionsText(directory)).toBe(adopted ? ADOPTED_ANSWERS : ANSWERS)
      const pinned = Object.fromEntries(leanSteps(directory).map(step => [step, withSourceSnapshot(() => hashStageInputs(step, "pin", stageInputs(directory, step)))]))
      if (adopted) expect(Object.keys(pinned).filter(step => pinned[step] === INPUTS[step])).toEqual(UNANSWERED_STEPS)
      else expect(pinned).toEqual(INPUTS)
    }
  } finally {
    if (previous === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previous
    rmSync(workspace, { recursive: true, force: true })
  }
})

// The spine records settled facts, decisions, recommendations, requirements
// and open items; it records no question and makes no choice for the lawyer.
const QUESTION = "Should late delivery carry a remedy?"
const LAW = "Which law should govern the supply note?"
const spine = (...entries: object[]) => ({ entries, limitations: [] })
const FACT = { kind: "fact", detail: "Delivery is within thirty days.", evidence: SUPPORT, limitations: [] }
// Pinned on the code before hand-backs reached the spine.
const SPINE_PROMPT = "aa0471af46a44d8b7a3ccc78fbfb1d1f647d5a746a00e934ca36b8886ea6a283"

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

test("the spine takes its five entry kinds and refuses a drafting choice or a question_id; its prompt is unchanged", () => {
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
    if (late === "open") expect(createHash("sha256").update(prompt).digest("hex")).toBe(SPINE_PROMPT)
  }
  // The discrepancy assessment never took one either.
  expect(checkEvidence(spineMatter("adopted"), "discrepancy", { entries: [], limitations: [] })).toEqual([])
})

test("intake told a question is answered is told so however it was answered", () => {
  for (const late of ["answered", "adopted"] as const) {
    const directory = spineMatter(late)
    const reply = JSON.stringify({ questions: [{ question: QUESTION, kind: "judgment", context: "Still silent." }] })
    expect(intakeQuestions(directory, reply)!.issues.join(" ")).toContain("already has a lawyer's answer")
  }
})

test("a question stored in the old format without answer words, as an old hand-back was, is read as unasked, so its step asks it again", () => {
  const directory = matter()
  const answered = { id: "q-law", stage: "intake", question: LAW, context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" }
  writeLeanText(directory, "questions.json", JSON.stringify({ schema: "dochaus.lean.questions/1", questions: [
    answered,
    { id: "q-late", stage: "intake", question: QUESTION, context: "The terms are silent.", status: "answered", answer_kind: "decide_for_me", answered_at: "2026-10-18T09:05:00.000Z" },
    { id: "q-open", stage: "intake", question: "Who signs?", context: "Execution.", status: "open" },
  ] }))
  expect(readQuestions(directory)).toEqual([answered as Question])
  expect(answeredQuestionsText(directory)).toBe(`- [q-law] Q: ${LAW}\n  A: Singapore law.`)
  // Intake may ask it again, and its key no longer carries it.
  const reply = JSON.stringify({ questions: [{ question: QUESTION, kind: "judgment", context: "Still silent." }] })
  expect(intakeQuestions(directory, reply)).toMatchObject({ issues: [], value: [{ question: QUESTION, status: "open" }] })
})

test("an answer queued for an old-format open question finds no question: the drain records it as rejected not_found, with its words kept", () => {
  const directory = matter()
  writeLeanText(directory, "questions.json", JSON.stringify({ schema: "dochaus.lean.questions/1", questions: [
    { id: "q-open", stage: "intake", question: "Who signs?", context: "Execution.", status: "open" },
  ] }))
  const request = enqueueLeanRequest(directory, { kind: "answer", question_id: "q-open", answer: "The director." })
  expect(drainLeanRequests(directory)).toEqual({ answered: 0, updated: 0, repairRound: false })
  expect(outcome(directory, request)).toEqual({ status: "rejected", detail: "not_found: No such question in this matter's lean drafting run." })
  expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ question_id: "q-open", answer: "The director." })
  expect(readQuestions(directory)).toEqual([])
})
