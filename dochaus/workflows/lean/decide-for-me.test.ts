import { afterEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { acceptChangePlan } from "./change-pipeline"
import { checkChanges } from "./change-checks"
import { HANDED_BACK_MEANING, handedBackRef } from "./checklist-contract"
import { readParagraphs } from "./checks"
import { checkEvidence, evidencePrompt } from "./evidence"
import { focusedChangeReviewInputs } from "./focused-change-review"
import { finalGate, type GateResult } from "./gate"
import { intakeQuestions } from "./intake"
import { leanSteps, readLeanPlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { openItem, withSourceSnapshot } from "./quotes"
import { requestedFilesPrompt } from "./requested-files"
import { drainLeanRequests, enqueueLeanRequest, queuedAnswerIDs, type LeanRequest, type StoredLeanRequest } from "./requests"
import { researchPrompt } from "./research"
import { runLean } from "./runner"
import { handedBackItems, LEAN_PROMPT_VERSION, researchBasis, stageInputs } from "./stages"
import { answeredQuestionsText, answerQuestion, assertDelegable, delegateQuestion, handedBack as isHandedBack, handedBackBlock, handedBackNotesBlock, handedBackQuestionsText, hashStageInputs, LeanQuestionError, readLeanJson, readLeanText, readOpenItems, readQuestions, writeLeanText, writeOpenItems, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, noChanges, scriptedModel, SUPPORT, type Reply } from "./test-support"

// "Decide for me": the lawyer hands a run question back instead of answering
// it. The hand-back is recorded as a choice the lawyer has not made: the
// drafting notes must disclose it, delivery stays a review draft, and the
// question's own open item is never resolved as answered by the lawyer. Only
// model-asked questions can be handed back. Synthetic matter; only the model is
// scripted.

afterEach(cleanupMatters)

const QUESTION = "Should late delivery carry a remedy?"
const CHOSEN = "Add no separate late-delivery remedy: the terms fix delivery within thirty days and the general law covers late performance."
const quiet = { ...happyPath, [AGENT.review]: () => "No changes needed.", [AGENT.changes]: noChanges }

/** The spine records each question the lawyer handed back as a drafting choice. */
function spineChoosing(directory: string): Reply {
  return () => JSON.stringify({ entries: readQuestions(directory).filter(isHandedBack).map(question => ({
    kind: "drafting-choice", question_id: question.id, detail: CHOSEN, evidence: SUPPORT, limitations: ["The lawyer has not made this choice."] })), limitations: [] })
}

/** A completed run whose one intake question the lawyer handed back. */
async function handedBack() {
  const directory = matter()
  const model = scriptedModel({ ...quiet, [AGENT.spine]: spineChoosing(directory), [AGENT.intake]: (_prompt, index) => index === 0
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
  expect(choice).toMatchObject({ stage: "finalize", kind: "left_to_drafter", status: "disclosed", ref: { repair_target: "memo", question_id: id } })
  expect(choice.detail).toContain(QUESTION)
  expect(choice.detail).not.toContain(id)
  // A hand-back is named as a choice left to the drafter, never called an omission.
  expect(choice.resolution).toBe("left to the drafter; named in the notes")
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
  // The question's own open item resolves to that disclosure, not to an answer.
  writeOpenItems(directory, [...readOpenItems(directory), openItem("changes", "question", QUESTION, { question_id: id, pid: "p-none" })])
  const own = finalGate(directory).items.find(item => item.kind === "question" && item.ref.question_id === id)!
  expect(own.status).toBe("resolved")
  expect(own.resolution).not.toBe("answered by the lawyer")
  // Notes that do not name the choice block the gate on it.
  const notes = readLeanText(directory, "drafting-notes.md")!
  writeLeanText(directory, "drafting-notes.md", notes.split("\n").filter(line => !line.includes(choice.id)).join("\n"))
  expect(finalGate(directory)).toMatchObject({ status: "blocked", delivery: null, blocking: [expect.objectContaining({ id: choice.id, status: "open", resolution: "left to the drafter; the notes do not name it" })] })
  expect(JSON.stringify(finalGate(directory).items.filter(item => item.ref.question_id === id))).not.toMatch(/omission named|name this omission/)
}, 120_000)

test("a real answer after a hand-back replaces it and its disclosure", async () => {
  const { directory, id } = await handedBack()
  enqueueLeanRequest(directory, { kind: "answer", question_id: id, answer: "Yes: a daily price reduction." })
  expect(drainLeanRequests(directory)).toEqual({ answered: 1, updated: 0, repairRound: false })
  const [question] = readQuestions(directory)
  expect(question).toMatchObject({ id, status: "answered", answer: "Yes: a daily price reduction." })
  expect(question).not.toHaveProperty("answer_kind")
  expect(finalGate(directory).items.some(item => item.ref.question_id === id && item.kind === "left_to_drafter")).toBe(false)
}, 120_000)

test("the drain applies a hand-back once, takes a repeat as recorded, and refuses it for an answered, unknown or non-model question", () => {
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
  // The hand-back counts as an answer, so a settled run restarts on it. A
  // repeat finds it recorded and changes nothing (see the replay test below).
  expect(drainLeanRequests(directory)).toEqual({ answered: 2, updated: 0, repairRound: false })
  expect(requests.map(request => outcome(directory, request))).toEqual([
    { status: "rejected", detail: expect.stringMatching(/^already_answered: /) },
    { status: "rejected", detail: expect.stringMatching(/^not_found: /) },
    { status: "rejected", detail: expect.stringMatching(/^not_delegable: /) },
    { status: "rejected", detail: "unreadable request" },
    { status: "delegated" },
    { status: "delegated" },
  ])
  expect(readQuestions(directory)).toEqual([questions[0]!, {
    id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", status: "answered", answer_kind: "decide_for_me", answered_at: expect.any(String),
  }, questions[2]!])
})

test("a hand-back replayed after a restart is recorded as delegated and changes nothing", () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", status: "open" }])
  const request = enqueueLeanRequest(directory, { kind: "answer", question_id: "q-cap", decide_for_me: true })
  // The drain applied the hand-back, then stopped before moving the request to done.
  const applied = delegateQuestion(directory, "q-cap")
  expect(drainLeanRequests(directory)).toEqual({ answered: 1, updated: 0, repairRound: false })
  expect(outcome(directory, request)).toEqual({ status: "delegated" })
  expect(readQuestions(directory)).toEqual([applied])
})

test("assertDelegable admits only an open model-asked question", () => {
  const open: Question = { id: "q-cap", stage: "changes", question: "Should liability be capped?", context: "The terms are silent.", status: "open" }
  const refusal = (target: Question | undefined, id: string) => {
    try {
      assertDelegable(target, id)
      return "admitted"
    } catch (error) {
      return error instanceof LeanQuestionError ? `${error.code}: ${error.message}` : String(error)
    }
  }
  expect(refusal(open, "q-cap")).toBe("admitted")
  expect(refusal(undefined, "q-missing")).toBe("not_found: No such question in this matter's lean drafting run.")
  expect(refusal({ ...open, id: "package-scope-d1" }, "package-scope-d1")).toBe("not_delegable: This question needs the lawyer's own answer; it cannot be left to the drafter to decide.")
  expect(refusal({ ...open, status: "answered", answer: "No cap.", answered_at: "2026-10-18T09:00:00.000Z" }, "q-cap")).toBe("already_answered: This question is already answered.")
  expect(refusal({ ...open, status: "answered", answer_kind: "decide_for_me", answered_at: "2026-10-18T09:00:00.000Z" }, "q-cap")).toBe("already_answered: This question was already left to the drafter to decide.")
})

// Inputs pinned on the code before hand-backs existed; hashed under a fixed
// label so a prompt version change does not move them. The gate's also keys
// its contract, bumped when the hint a disclosed hand-back's choice adds to
// the question that asks its conflict was reworded and narrowed.
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
  "gate": "5a2eea841b0446e2f2581f98994cdf7cb6d0e669da0a5e7d4d28b096f964ef25",
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
const LEFT = "Left to the drafter (not the lawyer's answers):"
// How every stage given a hand-back decides it and applies it in the body.
const DECIDE = "Decide each as the lawyer for the party the lawyer acts for would: follow the instructions and the documents; where they leave the point open or conflict, choose what best serves that party within them, or what is even-handed when the request names no party."
const APPLY = "Apply each choice made as clean operative wording, with no confirmation marker, rejected alternative or precedence clause between the alternatives; the drafting notes give the reason and the alternative."
// The drafting rule for other unsettled choices, and its exception for a hand-back.
const MARKER_RULE = "- A preference is a drafting choice needed to fulfil the requested purpose, within the requested scope. For a new substantive choice not settled by instructions or answers, draft balanced wording and add a visible [open item: confirm ...] marker in the same paragraph saying what was chosen. Do not add confirmation markers to settled decisions or ordinary wording that creates no new substantive choice. A drafted choice is not an invented fact; the marker goes with the drafted wording, never in its place."
const NO_MARKER = "- A point left to the drafter takes no confirmation marker: draft the choice made for it as clean operative wording, without the rejected alternative or a precedence clause between the alternatives."
// The spine, which makes the choice, records it with the wording it takes.
const SPINE_CHOICE = 'kind "drafting-choice" stating the choice made, decided as above, why, and the operative wording it takes, citing sources'

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
  // Downstream steps record and review the choice as a drafting choice, never as the lawyer's.
  expect(block).toContain("None is an agreed term or an answered decision; record and review each as a drafting choice.")
  // What "Decide for me" means is written once, right before the hand-backs.
  expect(block).toEndWith(`${HANDED_BACK_MEANING}\n- [q-cap] Q: Should liability be capped?`)
  expect(block.split(HANDED_BACK_MEANING)).toHaveLength(2)
  // Each is decided for the client within the instructions and documents, and applied cleanly.
  for (const rule of [DECIDE, APPLY]) expect(block).toContain(rule)
  // A hand-back keys every step through "answers", as an answer does.
  expect(stageInputs(directory, "intake").find(part => part.name === "answers")!.bytes).toBe(`${answeredQuestionsText(directory)}\nLeft to the drafting:\n- [q-cap] Q: Should liability be capped?`)
})

test("with no lawyer's answers a hand-back is listed under its own heading, not under the answers heading", () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-cap", stage: "intake", question: "Should liability be capped?", context: "The terms are silent.", status: "answered", answer_kind: "decide_for_me", answered_at: "2026-10-18T09:05:00.000Z" }])
  const block = handedBackBlock(directory)
  for (const prompt of [evidencePrompt(directory, "discrepancy"), evidencePrompt(directory, "spine"), researchPrompt(directory, "intake", researchBasis(directory)), researchPrompt(directory, "research", researchBasis(directory))])
    expect(prompt).toContain(`${LEFT}\n${block}`)
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
  const model = scriptedModel({ ...quiet, [AGENT.spine]: spineChoosing(directory), [AGENT.intake]: (_prompt, index) => index === 0
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
    AGENT["change-review"], AGENT["package-review"]]) {
    const own = prompts.filter(prompt => prompt.agent === agent)
    expect([agent, own.length > 0, own.some(prompt => prompt.text.includes(block))]).toEqual([agent, true, true])
  }
  // The files job runs only for a memo or several documents; its prompt is checked directly.
  expect(requestedFilesPrompt(directory, readLeanPlan(directory)!.documents, "memo")).toContain(block)
  // Every drafting part after the first gets it in its section handoff.
  expect(prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes("Section handoff")).every(prompt => prompt.text.includes(block))).toBe(true)
  // The spine, which makes the choice, and the draft, review and changes
  // stages decide it for the client and apply it as clean operative wording.
  for (const agent of [AGENT.spine, AGENT.draft, AGENT.review, AGENT["review-independent"], AGENT["review-wording"], AGENT["review-client-disadvantage"], AGENT.changes])
    expect([agent, prompts.some(prompt => prompt.agent === agent && prompt.text.includes(DECIDE) && prompt.text.includes(APPLY))]).toEqual([agent, true])
  // Each draft prompt that carries the drafting rules carries the hand-back's exception right after the marker rule.
  const ruled = prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.includes(MARKER_RULE))
  expect([ruled.length > 0, ruled.every(prompt => prompt.text.includes(`${MARKER_RULE}\n${NO_MARKER}`))]).toEqual([true, true])
  // The drafting notes see it in one neutral voice, so nothing echoes as the lawyer's or our words.
  const notesBlock = handedBackNotesBlock(directory)
  expect(notesBlock).toContain(line)
  for (const prompt of prompts) {
    expect([prompt.agent, prompt.text.includes("A: undefined")]).toEqual([prompt.agent, false])
    // Listed once, in its own block: never among the answers that control.
    expect([prompt.agent, prompt.text.split(line).length - 1]).toEqual([prompt.agent, prompt.text.includes(block) || prompt.text.includes(notesBlock) ? 1 : 0])
    // The block has its own heading or tag, never the answers heading.
    expect([prompt.agent, !prompt.text.includes(block) || [`<left_to_drafting>\n${block}`, `${LEFT}\n${block}`].some(text => prompt.text.includes(text))]).toEqual([prompt.agent, true])
  }
  const review = prompts.find(prompt => prompt.agent === AGENT.review)!.text
  const questions = prompts.find(prompt => prompt.agent === AGENT["change-review"])!.text.split("<questions>\n")[1]!.split("\n</questions>")[0]!
  expect([questions.includes(law), questions.includes(left)]).toEqual([true, false])
  expect(review).toContain(`their answers remain controlling and must not be re-asked under narrower or different wording:\n- [${law}] Q: ${LAW}\n  A: Singapore law.\n</answers>`)
  // The drafting notes read the hand-back as the lawyer's request to decide it.
  const notes = prompts.find(prompt => prompt.agent === AGENT.finalize)!.text
  expect(notes).toContain(`<left_to_drafter>\n${notesBlock}\n</left_to_drafter>`)
  expect(notesBlock).toStartWith("The lawyer left these questions to the drafter to decide. Task: ")
  // Nothing the notes' writer is given says a hand-back went unanswered.
  for (const unanswered of ["instead of answering", "not your answer", "or an answer"]) expect(notes).not.toContain(unanswered)
  for (const echo of [block, LEFT, "not the lawyer's answers"]) expect(notes).not.toContain(echo)
  expect(notes).toContain(`- ${left} (left to the drafter): ${QUESTION}`)
  expect(notes).not.toContain(`- ${left} (answered)`)
  expect(notes).toContain(`- ${law} (answered): ${LAW}\n  Answer: Singapore law.`)
  // The spine is told how to record the hand-back, and the notes' open item
  // for it carries what the spine chose, so the memo can say what and why.
  expect(prompts.find(prompt => prompt.agent === AGENT.spine)!.text).toContain(`"question_id"`)
  expect(prompts.find(prompt => prompt.agent === AGENT.spine)!.text).toContain(SPINE_CHOICE)
  const [item] = handedBackItems(directory)
  expect(notes).toContain(`- ${item!.id}: ${item!.detail}\n  Drafting spine (drafting-choice): ${CHOSEN}`)
}, 120_000)

// The lawyer's act is told in plain words: the web button's name never reaches
// a model or the gate's items, so no delivered text can echo it.
const BUTTON = /decide[\s_-]for[\s_-]me/i

test("no prompt, hand-back text or gate item names the Decide for me button", async () => {
  const { directory, left, prompts } = await answeredAndHandedBack()
  for (const prompt of prompts) expect([prompt.agent, BUTTON.exec(prompt.text)?.[0]]).toEqual([prompt.agent, undefined])
  for (const text of [handedBackBlock(directory), handedBackNotesBlock(directory), HANDED_BACK_MEANING, JSON.stringify(handedBackRef(left, QUESTION)), JSON.stringify(readLeanJson(directory, "gate.json"))])
    expect(text).not.toMatch(BUTTON)
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

test("only a hand-back is drafted without a confirmation marker; other unsettled choices keep the marker rule", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
  writeQuestions(directory, [
    { id: "q-law", stage: "intake", question: LAW, context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" },
    { id: "q-late", stage: "intake", question: QUESTION, context: "The terms are silent.", status: "open" },
  ])
  const drafts = () => [stagePrompt(directory, "d1.draft", { outline: true }), stagePrompt(directory, "d1.draft")]
  for (const prompt of drafts()) expect([prompt.includes(MARKER_RULE), prompt.includes(NO_MARKER), prompt.includes(APPLY)]).toEqual([true, false, false])
  delegateQuestion(directory, "q-late")
  for (const prompt of drafts()) expect([prompt.includes(`${MARKER_RULE}\n${NO_MARKER}`), prompt.includes(APPLY)]).toEqual([true, true])
  // The draft agent's own rules keep the marker rule and state the exception.
  const agent = readFileSync(new URL("../../agent/lean-draft.md", import.meta.url), "utf8")
  expect(agent).toContain("- A preference is a drafting choice a competent drafter would make from the request, the documents and ordinary practice. Draft that choice and add a visible [open item: confirm ...] marker in the same paragraph saying what was chosen. A drafted choice is not an invented fact; the marker goes with the drafted wording, never in its place.\n")
  expect(agent).toContain("- A point the prompt lists as left to the drafter takes no confirmation marker: draft the choice made for it as clean operative wording, without the rejected alternative or a precedence clause between the alternatives.\n")
})

test("intake re-asking a handed-back question is told it was left to the drafter, not answered", () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-cap", stage: "intake", question: "Should liability be capped?", context: "The terms are silent.", status: "open" }])
  delegateQuestion(directory, "q-cap")
  const reply = JSON.stringify({ questions: [{ question: "Should liability be capped?", context: "Still silent." }] })
  const issues = intakeQuestions(directory, reply)!.issues.join(" ")
  expect(issues).toContain("left that question to the drafter to decide")
  expect(issues).not.toContain("lawyer's answer")
})

// The spine records each hand-back once, keyed by the stored question id, as a
// drafting choice the lawyer has not made, or as an open item when deciding it
// would need a fact the sources do not give. Never as the lawyer's decision.
const CHOICE = { kind: "drafting-choice", question_id: "q-late", detail: CHOSEN, evidence: SUPPORT, limitations: ["The lawyer has not made this choice."] }
const spine = (...entries: object[]) => ({ entries, limitations: [] })

function spineMatter(late: Question["status"] | "handed back") {
  const directory = matter()
  writeQuestions(directory, [
    { id: "q-law", stage: "intake", question: LAW, context: "Legal basis.", status: "answered", answer: "Singapore law.", answered_at: "2026-10-18T09:00:00.000Z" },
    { id: "q-late", stage: "intake", question: QUESTION, context: "The terms are silent.", ...(late === "handed back" ? { status: "answered", answer_kind: "decide_for_me", answered_at: "2026-10-18T09:05:00.000Z" } : { status: late }) },
  ])
  return directory
}

test("the spine accepts a hand-back recorded once as a drafting choice or an open item", () => {
  const directory = spineMatter("handed back")
  expect(checkEvidence(directory, "spine", spine(CHOICE))).toEqual([])
  expect(checkEvidence(directory, "spine", spine({ ...CHOICE, kind: "open-item", detail: "Whether late delivery carries a remedy needs a commercial term the sources do not give; the body keeps its placeholder." }))).toEqual([])
  expect(evidencePrompt(directory, "spine")).toContain(`"question_id"`)
  for (const rule of [DECIDE, APPLY, SPINE_CHOICE]) expect(evidencePrompt(directory, "spine")).toContain(rule)
  // Its limitation is plain: a choice for the lawyer to confirm, not an answer awaited.
  expect(evidencePrompt(directory, "spine")).toContain("citing sources, with the limitation that it is a drafting choice for the lawyer to confirm;")
  expect(evidencePrompt(directory, "spine")).not.toContain("the lawyer has not made this choice")
})

test("the spine rejects a hand-back left out, recorded as a decision, recorded twice or under an id that is not a hand-back", () => {
  const directory = spineMatter("handed back")
  const issues = (...entries: object[]) => checkEvidence(directory, "spine", spine(...entries))
  const fact = { kind: "fact", detail: "Delivery is within thirty days.", evidence: SUPPORT, limitations: [] }
  const missing = issues(fact)
  expect(missing).toHaveLength(1)
  expect(missing[0]).toContain(`"q-late"`)
  expect(missing[0]).toContain(QUESTION)
  const decision = issues({ ...CHOICE, kind: "decision" })
  expect(decision).toHaveLength(1)
  expect(decision[0]).toContain("decision")
  expect(decision[0]).toContain(`"q-late"`)
  expect(issues(CHOICE, { ...CHOICE, kind: "open-item" })).toEqual([expect.stringContaining("exactly one")])
  // The lawyer answered q-law; q-other was never asked. Neither is a hand-back.
  for (const id of ["q-law", "q-other"]) expect(issues(CHOICE, { ...fact, question_id: id })).toEqual([expect.stringContaining(`"${id}"`)])
  const { question_id: _id, ...unkeyed } = CHOICE
  expect(issues(CHOICE, unkeyed)).toEqual([expect.stringContaining("question_id")])
})

test("the disclosure item tells the lawyer where the choice is and what each outcome needs, and does not depend on the spine", () => {
  const directory = spineMatter("handed back")
  const [item] = handedBackItems(directory)
  // Lawyer-facing, and true whether a choice was made or the point stays open:
  // a plain statement of where the choice is, never an instruction to the notes' writer.
  expect(item!.detail).toBe(`You left this question to the drafter to decide: ${QUESTION} The drafting notes give the choice made, the reason ` +
    "and the alternative not taken, or say why the point stays open. A choice made is for you to confirm; a point left open needs your answer.")
  expect(item!.detail).toContain(QUESTION)
  expect(item!.detail.replace(QUESTION, "")).not.toMatch(/\b(should|must|say what)\b/i)
  expect(item!.ref).toEqual({ repair_target: "memo", question_id: "q-late" })
  expect(item!.kind).toBe("left_to_drafter")
  // The notes' writer gets one neutral task line: each choice once, in plain words.
  const block = handedBackNotesBlock(directory)
  const task = block.split("\n")[0]!
  expect(task).toContain("Task: state each choice once, in plain words, in the statement for the open item that names its question: " +
    "what was chosen, why and the alternative not taken, or that the point stays open and what is needed.")
  expect(task).not.toMatch(/\b(you|your|we|our|us)\b/i)
  expect(block).not.toMatch(/answer/i)
  // A spine that leaves the question open yields the same item, id included.
  writeLeanText(directory, "spine.json", JSON.stringify(spine({ ...CHOICE, kind: "open-item" })))
  expect(handedBackItems(directory)).toEqual([item!])
  // A question without end punctuation still ends its own sentence.
  writeQuestions(directory, [{ id: "q-late", stage: "intake", question: "Deposit of £70,000 or £75,000 ", context: "", status: "answered", answer_kind: "decide_for_me", answered_at: "2026-10-18T09:05:00.000Z" }])
  expect(handedBackItems(directory)[0]!.detail).toStartWith("You left this question to the drafter to decide: Deposit of £70,000 or £75,000. The drafting notes give")
})

// The lawyer leaves a point to the drafter, never "to the drafting": every
// prompt, the draft agent's rules and each retry message say so.
const TO_DRAFTING = /\bleft (?:\w+ ){0,2}to the drafting\b/i
const TO_DRAFTER = /\bleft (?:\w+ ){0,2}to the drafter\b/

test("no prompt, agent rule or retry message says a point was left to the drafting", async () => {
  const { prompts } = await answeredAndHandedBack()
  for (const prompt of prompts) expect([prompt.agent, TO_DRAFTING.exec(prompt.text)?.[0]]).toEqual([prompt.agent, undefined])
  expect(readFileSync(new URL("../../agent/lean-draft.md", import.meta.url), "utf8")).not.toMatch(TO_DRAFTING)
  const directory = spineMatter("handed back")
  expect(evidencePrompt(directory, "spine")).not.toMatch(TO_DRAFTING)
  const fact = { kind: "fact", detail: "Delivery is within thirty days.", evidence: SUPPORT, limitations: [] }
  const { question_id: _id, ...unkeyed } = CHOICE
  const issues = [spine({ ...CHOICE, kind: "other" }), spine(unkeyed), spine(CHOICE, { ...fact, question_id: "q-other" }), spine({ ...CHOICE, kind: "decision" })]
    .flatMap(value => checkEvidence(directory, "spine", value))
  issues.push(...intakeQuestions(directory, JSON.stringify({ questions: [{ question: QUESTION, context: "Still silent." }] }))!.issues)
  expect(issues).toHaveLength(7)
  for (const issue of issues) expect([issue, TO_DRAFTER.test(issue), TO_DRAFTING.test(issue)]).toEqual([issue, true, false])
}, 120_000)

// Pinned on the code before hand-backs reached the spine.
const SPINE_PROMPT = "aa0471af46a44d8b7a3ccc78fbfb1d1f647d5a746a00e934ca36b8886ea6a283"

test("without a hand-back the spine check and prompt are unchanged", () => {
  for (const late of ["open", "answered"] as const) {
    const directory = spineMatter(late)
    if (late === "answered") answerQuestion(directory, "q-late", "No separate remedy.")
    // The new kind is not offered and a stray question_id is ignored, as before.
    expect(checkEvidence(directory, "spine", spine(CHOICE))).toEqual(["A spine entry kind must be decision, recommendation, requirement, fact or open-item."])
    expect(checkEvidence(directory, "spine", spine({ ...CHOICE, kind: "decision", question_id: "q-law" }))).toEqual([])
    expect(checkEvidence(directory, "spine", spine())).toEqual([])
    if (late === "open") expect(createHash("sha256").update(evidencePrompt(directory, "spine")).digest("hex")).toBe(SPINE_PROMPT)
  }
})
