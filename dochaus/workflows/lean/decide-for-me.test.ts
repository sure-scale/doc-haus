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
import { ADOPTED_ANSWER, answeredQuestionsText, answerQuestion, delegable, hashStageInputs, LeanQuestionError, readLeanJson, readQuestions, writeLeanText, writeQuestions, type Question, type QuestionAdvice } from "./state"
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
  expect(adopted.answers).toContain(`- [q-cap] Q: ${OPEN.question}\n  A (${ADOPTED_ANSWER}): ${RECOMMENDED}`)
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
// spine was re-pinned when it began keying the act record (acts.json), and the
// acts step when its record began stating conflicts (act-record-2) and
// verifying its candidate acts (act-record-4, act-record-5), when its
// checks began placing a reply's JSON break (act-record-6), and when HOW
// began reading the deliverables (act-record-7). The
// same pins hold for typed or picked words; an adopted recommendation reads as
// delegated, which moves every key that reads the answers and no other. The
// instructions step was added, and intake and spine were re-pinned when they
// began keying the instruction list (instructions.md), and the discrepancy
// step when it began reading that list to list gaps and again when it split
// into a differences job and a gaps job (differences-then-gaps-1), and once more
// when the differences prompt was tagged (differences-then-gaps-2), and when the
// jobs ran in parallel with a verify job per gap (differences-gaps-verified-1), and
// when each difference gained a verify job too (differences-gaps-verified-2). Every key that reads
// plan.json was re-pinned when whom the lawyer acts for gained its basis
// (dochaus.lean.plan/2). The draft step was re-pinned when the outline
// contract gained the statement form (statement-form-12), and again when the
// outline began naming every defect in one pass and planning a requested file
// that holds no planned document, and each planned part began to be placed
// from its own side too and to read each entry's limitations
// (outline-one-pass-reverse-18), and when a template's restated title stopped
// counting as a part, a memo named only by a requested file name or only by a
// planned memo's title began to stand, and acts an answered scope question
// settled stopped being sent back for revision (outline-one-pass-reverse-19),
// and when the outline began to be shown the template parts it must account
// for, to read deeper headings as scope, to adopt a lone advice document as
// the memo and never to fold advice (outline-one-pass-reverse-20).
// The spine was re-pinned when its entries
// began to be split into atoms (spine-atoms-1), when a one-atom split began to
// read as atomic (spine-atoms-2), when its checks began asking each figure of
// an entry or atom of its own quotes (spine-atoms-3), and when spine.json began
// keeping the accepted entries beside their atoms, a figure became one with a
// sign, a decimal point, separators or five digits, and its entries' absence
// claims began to be verified (spine-atoms-5); every step after the spine,
// other than the draft, was re-pinned then for keying the atoms. The spine
// alone was re-pinned when an id's digits stopped being read as a number and a
// dotted number after a reference label as a figure (spine-atoms-6), and when
// its absence check began to skip open items, to run on every shaped attempt
// and to record only on the last, and a figure stopped being a bare run of
// digits (spine-atoms-7).
const ANSWERS = "- [package-scope-d1] Q: Which documents are in scope?\n  A: Only the supply note.\n- [q-law] Q: Which law should govern?\n  A: Singapore law."
const ADOPTED_ANSWERS = ANSWERS.replace("  A: Singapore law.", `  A (${ADOPTED_ANSWER}): Singapore law.`)
// A draft keys its outline contract and plan entry, never answers.
const UNANSWERED_STEPS = ["d1.draft"]
const INPUTS: Record<string, string> = {
  "instructions": "7db0182116d00b86390816e69312620dff6c36a8920d47b9190fb725b72906b6",
  "discrepancy": "c4231df63a628e352a61374626dbedcd1051c07c8a8b8f330e20ea2568316eb1",
  "intake": "097361ef1ee9a3958bc6d602d33eac98ff61d441ea7aa050c405152bf241c65d",
  "research": "64a0d22c3dd780907955b2e56b3122b94ea72101092d694f899a15e2af0cb13b",
  "research-review": "f30467728cba33fbf89ff10cfba8903cd7aa9647cbd6849d4a3caf293f27a65b",
  "acts": "ff830f3d37f64416362401fa66710c9fc8e770429b6bc5f2af84baffd3f35a3c",
  "spine": "7719311d661486872e589bf29d615b4ffc5893a39fb4ab7694ff3ae6f16cecef",
  "d1.draft": "91a52c7311ebbc0fcf507f8d921ce27391ecf34ada33f65d60da3eccbeb77482",
  "d1.review": "23af042296cf7f7d3ee97332f0caaf1fac7ab240199ce23d96a33f9c65fc2693",
  "d1.review-independent": "ae78dbf15aa677be04d6d44d0b91d63c34fd2fa03754e51350a6941413143d70",
  "d1.review-wording": "da1716d02cfb2df4b1f9fe71ceec32c3ffe41e66070f508a765ef7aa81e3792a",
  "d1.review-client-disadvantage": "c43e31f535349cf6ad55d3ce52bb9c90b2b7ef81d138cad29091f9f29d626111",
  "d1.change-plan": "20d99921729a43a29acd201db2a6c86a882bae30fd7ee6f07aa3865c04bc0e9c",
  "d1.change-prepare": "5eb87e9b0f4dbadbf0b9916c685b1d24214d3cc4f4d8b408065e0dcbbb7f99db",
  "d1.changes": "132f96d7b820a2985ed5f77066a603023e17b918e497b598473417b55df49396",
  "d1.change-review": "d73733fc9dcdafca5f02603c12d9ad948389a1185887622786bf27fc63ed74b7",
  "d1.apply": "010ef984a0764b469b3062366447829190a21d660fbf1d456b798728b3fb66d5",
  "finalize": "e927aef313b8abe527e12ce7fdac01b0c6f89586bbd3d00431ec013cc074473d",
  "checklist-check": "a44b6aa518efe2b04874468e01db95a9323f20bf98dd2b7805d66e95e3632a20",
  "package-review": "d1c48bcf6dfe13c61403da6a6e73dbb9423d0a0ff5449c2ee5d0eb6bf7480783",
  "gate": "04cf9a0b7d6876b1afe64a4beb8abe0883d0c4a6e26baba2faef68a0220c5fe3",
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
// Pinned on the code before hand-backs reached the spine; re-pinned when the
// spine began carrying unsettled discrepancy gaps as open items.
const SPINE_PROMPT = "495828d2f70b3dc96d2e5f1584315afdd678f845afd3382b77ff3236cf97ea36"

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
  expect(checkEvidence(spineMatter("adopted"), "discrepancy", { entries: [], gaps: [], limitations: [] })).toEqual([])
})

test("intake told a question is answered is told so however it was answered", () => {
  for (const late of ["answered", "adopted"] as const) {
    const directory = spineMatter(late)
    const reply = JSON.stringify({ questions: [{ question: QUESTION, kind: "judgment", context: "Still silent." }] })
    expect(intakeQuestions(directory, reply)!.issues.join(" ")).toContain("already has an answer")
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
