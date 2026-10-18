import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { acceptChangePlan } from "./change-pipeline"
import { checkChanges } from "./change-checks"
import { readParagraphs } from "./checks"
import { focusedChangeReviewInputs } from "./focused-change-review"
import { finalGate, type GateResult } from "./gate"
import { intakeQuestions } from "./intake"
import { leanSteps, readLeanPlan, writeLeanPlan } from "./plan"
import { openItem, withSourceSnapshot } from "./quotes"
import { requestedFilesPrompt } from "./requested-files"
import { drainLeanRequests, enqueueLeanRequest, queuedAnswerIDs, type LeanRequest, type StoredLeanRequest } from "./requests"
import { runLean } from "./runner"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { answeredQuestionsText, answerQuestion, delegateQuestion, handedBackBlock, handedBackQuestionsText, hashStageInputs, readLeanJson, readLeanText, readOpenItems, readQuestions, writeLeanText, writeOpenItems, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, noChanges, scriptedModel, SUPPORT } from "./test-support"

// "Decide for me": the lawyer hands a run question back instead of answering
// it. The hand-back is recorded as a choice the lawyer has not made: the
// drafting notes must disclose it, delivery stays a review draft, and the
// question's own open item is never resolved as answered by the lawyer. Only
// model-asked questions can be handed back. Synthetic matter; only the model is
// scripted.

afterEach(cleanupMatters)

const QUESTION = "Should late delivery carry a remedy?"
const quiet = { ...happyPath, [AGENT.review]: () => "No changes needed.", [AGENT.changes]: noChanges }

/** A completed run whose one intake question the lawyer handed back. */
async function handedBack() {
  const directory = matter()
  const model = scriptedModel({ ...quiet, [AGENT.intake]: (_prompt, index) => index === 0
    ? JSON.stringify({ questions: [{ question: QUESTION, context: "The supplied terms are silent on late delivery." }] })
    : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  const id = readQuestions(directory)[0]!.id
  enqueueLeanRequest(directory, { kind: "answer", question_id: id, decide_for_me: true })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  return { directory, id }
}

function outcome(directory: string, request: StoredLeanRequest) {
  return (readLeanJson(directory, `requests/done/${request.id}.json`) as { outcome: { status: string; detail?: string } }).outcome
}

test("a handed-back question is disclosed as a choice the lawyer has not made and is never delivered execution ready", async () => {
  const { directory, id } = await handedBack()
  const [question] = readQuestions(directory)
  expect(question).toEqual(expect.objectContaining({ id, status: "answered", answer_kind: "decide_for_me", answered_at: expect.any(String) }))
  expect(question).not.toHaveProperty("answer")
  const gate = readLeanJson(directory, "gate.json") as GateResult
  const choice = gate.items.find(item => item.ref.question_id === id)!
  expect(choice).toMatchObject({ stage: "finalize", kind: "omission", status: "disclosed", ref: { repair_target: "memo", question_id: id } })
  expect(choice.detail).toContain(QUESTION)
  expect(choice.detail).not.toContain(id)
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
  // The question's own open item resolves to that disclosure, not to an answer.
  writeOpenItems(directory, [...readOpenItems(directory), openItem("changes", "question", QUESTION, { question_id: id, pid: "p-none" })])
  const own = finalGate(directory).items.find(item => item.kind === "question" && item.ref.question_id === id)!
  expect(own.status).toBe("resolved")
  expect(own.resolution).not.toBe("answered by the lawyer")
  // Notes that do not name the choice block the gate on it.
  const notes = readLeanText(directory, "drafting-notes.md")!
  writeLeanText(directory, "drafting-notes.md", notes.split("\n").filter(line => !line.includes(choice.id)).join("\n"))
  expect(finalGate(directory)).toMatchObject({ status: "blocked", delivery: null, blocking: [expect.objectContaining({ id: choice.id, status: "open" })] })
}, 120_000)

test("a real answer after a hand-back replaces it and its disclosure", async () => {
  const { directory, id } = await handedBack()
  enqueueLeanRequest(directory, { kind: "answer", question_id: id, answer: "Yes: a daily price reduction." })
  expect(drainLeanRequests(directory)).toEqual({ answered: 1, updated: 0, repairRound: false })
  const [question] = readQuestions(directory)
  expect(question).toMatchObject({ id, status: "answered", answer: "Yes: a daily price reduction." })
  expect(question).not.toHaveProperty("answer_kind")
  expect(finalGate(directory).items.some(item => item.ref.question_id === id && item.kind === "omission")).toBe(false)
}, 120_000)

test("the drain applies a hand-back once and refuses it for an answered, unknown or non-model question", () => {
  const directory = matter()
  const questions: Question[] = [
    { id: "q-law", stage: "intake", question: "Which law should govern?", context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" },
    { id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", status: "open" },
    { id: "package-scope-earlier", stage: "draft", question: "Which documents are in scope?", context: "Scope.", status: "open" },
  ]
  writeQuestions(directory, questions)
  const requests = [
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-law", decide_for_me: true }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-missing", decide_for_me: true }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "package-scope-earlier", decide_for_me: true }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-cap", answer: "No cap.", decide_for_me: true } as unknown as LeanRequest),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-cap", decide_for_me: true }),
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-cap", decide_for_me: true }),
  ]
  expect(queuedAnswerIDs(directory)).toEqual(new Set(["q-law", "q-missing", "package-scope-earlier", "q-cap"]))
  // The hand-back counts as an answer, so a settled run restarts on it.
  expect(drainLeanRequests(directory)).toEqual({ answered: 1, updated: 0, repairRound: false })
  expect(requests.map(request => outcome(directory, request))).toEqual([
    { status: "rejected", detail: expect.stringMatching(/^already_answered: /) },
    { status: "rejected", detail: expect.stringMatching(/^not_found: /) },
    { status: "rejected", detail: expect.stringMatching(/^not_delegable: /) },
    { status: "rejected", detail: "unreadable request" },
    { status: "delegated" },
    { status: "rejected", detail: expect.stringMatching(/^already_answered: /) },
  ])
  expect(readQuestions(directory)).toEqual([questions[0]!, {
    id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", status: "answered", answer_kind: "decide_for_me", answered_at: expect.any(String),
  }, questions[2]!])
})

// Inputs pinned on the code before hand-backs existed; hashed under a fixed
// label so a prompt version change does not move them.
const ANSWERS = "- [package-scope-d1] Q: Which documents are in scope?\n  A: Only the supply note.\n- [q-law] Q: Which law should govern?\n  A: Singapore law."
const INPUTS: Record<string, string> = {
  "discrepancy": "8dac7eefaf1e988e35ffbe6c7f765a951943e9b1299546a9b94dc10e99950a72",
  "intake": "1e491eea042fb2cb31d8340459e58993849d22ce04ec26793d0e495a840061c6",
  "research": "64a0d22c3dd780907955b2e56b3122b94ea72101092d694f899a15e2af0cb13b",
  "research-review": "f30467728cba33fbf89ff10cfba8903cd7aa9647cbd6849d4a3caf293f27a65b",
  "spine": "c038d3823819fc2e8ef176dc1cb855fd7f7a87e1fe077efb1c2aee41f922e076",
  "d1.draft": "c9433207902f01fa193b87e4cb2d1bb14fbbd0eed6eeed767579bd51c0247881",
  "d1.review": "e80f5bbe084c8ef90d2864358d2caafa09139822584fb0f12cb14d0976e76416",
  "d1.review-independent": "9f1fe22f3e0db01ef5fffe96a09ddf917b5e371a89c58988eb875bcfbbae4aa4",
  "d1.review-wording": "7f86222090d292207f576968b1171dab74bedab8cf5dbc4e3593752d2584b6f5",
  "d1.review-client-disadvantage": "be6786003165e1b19571bc525b2d6dd3ba4afad3a7f707c7f2e58fe4f69ffac1",
  "d1.change-plan": "27e6854d53eef220e0881676aa0ea75d3dd349ddf1a4f2dadab7d7905790946f",
  "d1.change-prepare": "0e65f05f5cdbac92e874b5150ac1dd21f355a34be5b6a1b918b1be46f580591f",
  "d1.changes": "332ce3588588c3dc6fd37068e9202f5612a4e6a9f21d9632e0e2b15a7bba5271",
  "d1.change-review": "d5ab00d1b03fbfa3aa4104202872c85bd8d96bb8e4acffdb112d70d32ea87924",
  "d1.apply": "3fc66df6877ed50af811c9f7a5bbdf1ccb1c269dc8b176cf392941cd2da83149",
  "finalize": "b38c29092265887c63cf0e8a96bac1cd40201575a9344dca65628f3adedef462",
  "checklist-check": "bf9d2570d825f9efdf7d09b08a635561362fda11c11812346212a16ead317e10",
  "package-review": "dde115af0ba46b8cb1ecf9bf7eb0c40aff3b78f4d6be9e0f0da3dff2d5d4066f",
  "gate": "59610ac871362d90b8662cdd62991b5759a7da08167e298009c3fe59e970e5db",
}

test("without a hand-back the answers text and every step's inputs are unchanged", () => {
  const previous = process.env.WORKSPACE_ROOT
  const workspace = mkdtempSync(path.join(tmpdir(), "lean-dfm-workspace-"))
  process.env.WORKSPACE_ROOT = workspace
  try {
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
    writeQuestions(directory, [
      { id: "q-law", stage: "intake", question: "Which law should govern?", context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" },
      { id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", pid: "p-1", document: "d1", status: "open" },
      { id: "package-scope-d1", stage: "intake", question: "Which documents are in scope?", context: "Scope.", status: "answered", answer: "Only the supply note.", answered_at: "2026-10-18T09:05:00.000Z" },
    ])
    expect(answeredQuestionsText(directory)).toBe(ANSWERS)
    expect([handedBackQuestionsText(directory), handedBackBlock(directory)]).toEqual(["", ""])
    expect(Object.fromEntries(leanSteps(directory).map(step => [step, withSourceSnapshot(() => hashStageInputs(step, "pin", stageInputs(directory, step)))]))).toEqual(INPUTS)
  } finally {
    if (previous === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previous
    rmSync(workspace, { recursive: true, force: true })
  }
})

// The lawyer's answers and the questions handed back are separate projections:
// only answers are authoritative; a hand-back is a choice the drafting makes
// and discloses. Neither shows a timestamp.
const LAW = "Which law should govern the supply note?"

test("the answers text lists only the lawyer's answers; hand-backs are listed apart, without timestamps", () => {
  const directory = matter()
  writeQuestions(directory, [
    { id: "q-law", stage: "intake", question: LAW, context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" },
    { id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", status: "answered", answer_kind: "decide_for_me", answered_at: "2026-10-18T09:05:00.000Z" },
    { id: "q-open", stage: "intake", question: "Who signs?", context: "Execution.", status: "open" },
  ])
  expect(answeredQuestionsText(directory)).toBe(`- [q-law] Q: ${LAW}\n  A: Singapore law.`)
  expect(handedBackQuestionsText(directory)).toBe("- [q-cap] Q: Should liability be capped?")
  const block = handedBackBlock(directory)
  expect(block).toEndWith("\n- [q-cap] Q: Should liability be capped?")
  expect(block).not.toMatch(/2026|controlling|lawyer's answers:/)
  // A hand-back keys every step through "answers", as an answer does.
  expect(stageInputs(directory, "intake").find(part => part.name === "answers")!.bytes).toBe(`${answeredQuestionsText(directory)}\nLeft to the drafting:\n- [q-cap] Q: Should liability be capped?`)
})

test("a hand-back changes the intake key, as an answer does", () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-cap", stage: "intake", question: "Should liability be capped?", context: "The terms are silent.", status: "open" }])
  const key = () => withSourceSnapshot(() => hashStageInputs("intake", LEAN_PROMPT_VERSION, stageInputs(directory, "intake")))
  const open = key()
  delegateQuestion(directory, "q-cap")
  const delegated = key()
  expect(delegated).not.toBe(open)
  expect(stageInputs(directory, "intake").find(part => part.name === "answers")!.bytes).toBe("Left to the drafting:\n- [q-cap] Q: Should liability be capped?")
  answerQuestion(directory, "q-cap", "No cap.")
  expect(key()).not.toBe(delegated)
})

/** A run whose intake asked two questions: the lawyer answered one and handed
 * the other back. Returns the prompts of the rerun that followed. */
async function answeredAndHandedBack() {
  const directory = matter()
  const model = scriptedModel({ ...quiet, [AGENT.intake]: (_prompt, index) => index === 0
    ? JSON.stringify({ questions: [{ question: QUESTION, context: "The supplied terms are silent on late delivery." }, { question: LAW, context: "The terms name no governing law." }] })
    : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  const id = (text: string) => readQuestions(directory).find(question => question.question === text)!.id
  enqueueLeanRequest(directory, { kind: "answer", question_id: id(LAW), answer: "Singapore law." })
  enqueueLeanRequest(directory, { kind: "answer", question_id: id(QUESTION), decide_for_me: true })
  const before = model.prompts.length
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  return { directory, left: id(QUESTION), law: id(LAW), prompts: model.prompts.slice(before) }
}

test("every prompt given the lawyer's answers also lists the hand-back, apart from the answers and their authority", async () => {
  const { directory, left, law, prompts } = await answeredAndHandedBack()
  const block = handedBackBlock(directory)
  const line = `- [${left}] Q: ${QUESTION}`
  expect(block).toContain(line)
  for (const agent of [AGENT.discrepancy, AGENT.intake, AGENT.research, AGENT["research-review"], AGENT.spine, AGENT.draft, AGENT.coverage,
    AGENT.review, AGENT["review-independent"], AGENT["review-wording"], AGENT["review-client-disadvantage"], AGENT["change-plan"], AGENT.changes,
    AGENT["change-review"], AGENT.finalize, AGENT["package-review"]]) {
    const own = prompts.filter(prompt => prompt.agent === agent)
    expect([agent, own.length > 0, own.some(prompt => prompt.text.includes(block))]).toEqual([agent, true, true])
  }
  // The files job runs only for a memo or several documents; its prompt is checked directly.
  expect(requestedFilesPrompt(directory, readLeanPlan(directory)!.documents, "memo")).toContain(block)
  // Every drafting part after the first gets it in its section handoff.
  expect(prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes("Section handoff")).every(prompt => prompt.text.includes(block))).toBe(true)
  for (const prompt of prompts) {
    expect([prompt.agent, prompt.text.includes("A: undefined")]).toEqual([prompt.agent, false])
    // Listed once, in its own block: never among the answers that control.
    expect([prompt.agent, prompt.text.split(line).length - 1]).toEqual([prompt.agent, prompt.text.includes(block) ? 1 : 0])
  }
  const review = prompts.find(prompt => prompt.agent === AGENT.review)!.text
  const questions = prompts.find(prompt => prompt.agent === AGENT["change-review"])!.text.split("<questions>\n")[1]!.split("\n</questions>")[0]!
  expect([questions.includes(law), questions.includes(left)]).toEqual([true, false])
  expect(review).toContain(`their answers remain controlling and must not be re-asked under narrower or different wording:\n- [${law}] Q: ${LAW}\n  A: Singapore law.\n</answers>`)
  // The drafting notes' questions read the hand-back as left to the drafting.
  const notes = prompts.find(prompt => prompt.agent === AGENT.finalize)!.text
  expect(notes).toContain(`- ${left} (left to the drafting): ${QUESTION}`)
  expect(notes).not.toContain(`- ${left} (answered)`)
  expect(notes).toContain(`- ${law} (answered): ${LAW}\n  Answer: Singapore law.`)
}, 120_000)

test("the focused change review lists the hand-back apart from the actual answers", async () => {
  const { directory, left } = await answeredAndHandedBack()
  const delivery = readParagraphs(directory, "d1").find(row => row.text === "The supplier shall deliver the goods.")!
  const plan = { shared_decisions: [], groups: [{ id: "delivery", pids: [delivery.pid], goal: "State the supplied delivery period.", sources: SUPPORT, limitations: [] }], dismissed: [], questions: [], findings: [] }
  expect(acceptChangePlan(directory, "d1", JSON.stringify(plan)).issues).toEqual([])
  const checked = checkChanges(directory, JSON.stringify({ amendments: [{ pid: delivery.pid, updated: "The supplier shall deliver the goods within thirty days of the order date.", reason: "States the source period.", sources: SUPPORT }], questions: [] }), readParagraphs(directory, "d1"), { document: "d1" })
  expect(checked.issues).toEqual([])
  writeLeanText(directory, "d1/changes.json", JSON.stringify(checked.changes))
  const [input] = focusedChangeReviewInputs(directory, "d1")
  expect(input!.text).toContain(handedBackBlock(directory))
  const answers = input!.text.split("<actual_questions_and_answers>\n")[1]!.split("\n</actual_questions_and_answers>")[0]!
  expect(answers).toContain("Singapore law.")
  expect(answers).not.toContain(left)
}, 120_000)

test("intake re-asking a handed-back question is told it was left to the drafting, not answered", () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-cap", stage: "intake", question: "Should liability be capped?", context: "The terms are silent.", status: "open" }])
  delegateQuestion(directory, "q-cap")
  const reply = JSON.stringify({ questions: [{ question: "Should liability be capped?", context: "Still silent." }] })
  const issues = intakeQuestions(directory, reply)!.issues.join(" ")
  expect(issues).toContain("left that question to the drafting to decide")
  expect(issues).not.toContain("lawyer's answer")
})
