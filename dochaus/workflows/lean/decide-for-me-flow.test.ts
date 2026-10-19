import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import path from "node:path"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { DELIVERED_NOTES } from "./deliverables"
import type { GateResult } from "./gate"
import { readLeanPlan } from "./plan"
import { drainLeanRequests, enqueueLeanRequest, queuedAnswerIDs, type StoredLeanRequest } from "./requests"
import { runLean } from "./runner"
import { handedBackItems } from "./stages"
import { answeredQuestionsText, assertDelegable, handedBack, handedBackBlock, readLeanJson, readLeanText, readQuestions, writeLeanText } from "./state"
import { AGENT, cleanupMatters, drafter, DRAFT, happyPath, matter, noChanges, registering, scriptedModel, SUPPORT, type Reply } from "./test-support"

// "Decide for me" through the real coordinator, end to end: a run blocks on a
// model-asked question, the lawyer hands it back through the request queue the
// ingest answer route uses, the spine and memo checkers refuse a choice
// recorded as approved or left uncovered, the corrected run delivers a review
// draft whose memo discloses the choice in plain words, an unchanged Continue
// does no model work, and a real answer supersedes the hand-back. Synthetic
// matter; only the model is scripted.

const REQUEST = "Draft a short supply note from the supply terms, with a cover memo on the open points."
const MEMO_FILE = "Supply note - cover memo.docx"
const QUESTION = "Should late delivery carry a remedy?"
const CHOSEN = "Add no separate late-delivery remedy: the terms fix delivery within thirty days and the general law covers late performance."
const ANSWER = "Yes: a price reduction of one percent for each week of late delivery."
const DISCLOSED = "The lawyer left this to the drafting, which chose:"

type SpineFault = "approved" | "uncovered" | "stale"
type MemoFault = "uncovered"

/** One-shot faults the test arms before a run: each reply takes the next one,
 * then answers correctly from the matter's current questions. */
const armed = { spine: [] as SpineFault[], finalize: [] as MemoFault[] }

afterEach(() => {
  cleanupMatters()
  armed.spine.length = 0
  armed.finalize.length = 0
})

/** The spine: a correct reply records each standing hand-back as one
 * drafting-choice entry. Faults record it as a decision the client approved,
 * leave it out, or keep a choice for a question the lawyer has since answered. */
function spine(directory: string): Reply {
  return () => {
    const handed = readQuestions(directory).filter(handedBack)
    const fault = armed.spine.shift()
    const entry = (id: string) => ({ kind: "drafting-choice", question_id: id, detail: CHOSEN, evidence: SUPPORT, limitations: ["The lawyer has not made this choice."] })
    const entries = fault === "approved"
      ? handed.map(question => ({ kind: "decision", question_id: question.id, detail: "The client approved adding no separate late-delivery remedy.", evidence: SUPPORT, limitations: [] }))
      : fault === "uncovered" ? []
      : fault === "stale" ? readQuestions(directory).map(question => entry(question.id))
      : handed.map(question => entry(question.id))
    return JSON.stringify({ entries, limitations: [] })
  }
}

/** The memo: one plain-words line per listed open item, each ending in its
 * "[ref: <id>]" marker. An item carrying the spine's drafting choice is
 * disclosed as left to the drafting with what was chosen. A retry arrives in
 * the same session, so the reply keeps reading the last full prompt. The fault
 * leaves out the hand-back's item. */
function memo(): Reply {
  const seen = { prompt: "" }
  return prompt => {
    seen.prompt = prompt.includes("<output_contract>") ? prompt : seen.prompt
    const fault = armed.finalize.shift()
    const blocks = [...seen.prompt.matchAll(/<open_items>\n([\s\S]*?)\n<\/open_items>/g)].map(match => match[1]!)
    const items = blocks.flatMap(block => [...block.matchAll(/^- (\S+): .*(?:\n {2}Drafting spine \(drafting-choice\): (.*))?/gm)].map(match => ({ id: match[1]!, choice: match[2] })))
    const question = (/^- \S+ \(left to the drafter\): (.*)$/m.exec(seen.prompt))?.[1]
    const lines = items.filter(item => !(fault === "uncovered" && item.choice)).map(item => item.choice
      ? `- ${question} ${DISCLOSED} ${item.choice} The lawyer has not made this choice and should confirm it or answer the question. [ref: ${item.id}]`
      : `- The lawyer must resolve this before the draft is used. [ref: ${item.id}]`)
    return ["# Cover memo", "", "The supply terms set the delivery period, and the draft states it.", "", ...lines].join("\n")
  }
}

/** The outline's reply names the requested memo, so the notes are written as it. */
function memoDrafter(): Reply {
  const draft = drafter(DRAFT)
  return (prompt, index) => draft(prompt, index).replace("Acting for: not stated", "Acting for: not stated\nMemo: cover memo")
}

function outcome(directory: string, request: StoredLeanRequest) {
  return (readLeanJson(directory, `requests/done/${request.id}.json`) as { outcome: { status: string; detail?: string } }).outcome
}

const gate = (directory: string) => readLeanJson(directory, "gate.json") as GateResult
const docxText = (directory: string, name: string) => readDocxCommentParagraphs(readFileSync(path.join(directory, name))).map(row => row.paragraph_text).join("\n")

test("a hand-back runs through the queue, its checks and the gate to a review draft, rests on Continue and yields to a real answer", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", REQUEST)
  const deliverables = registering()
  const model = scriptedModel({
    ...happyPath,
    [AGENT.review]: () => "No changes needed.",
    [AGENT.changes]: noChanges,
    [AGENT.draft]: memoDrafter(),
    [AGENT.spine]: spine(directory),
    [AGENT.finalize]: memo(),
    [AGENT.intake]: (_prompt, index) => index === 0
      ? JSON.stringify({ questions: [{ question: QUESTION, context: "The supplied terms are silent on late delivery." }] })
      : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" }),
  })
  const run = () => runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })
  const since = (start: number, agent: string) => model.prompts.slice(start).filter(prompt => prompt.agent === agent).map(prompt => prompt.text)

  // 1. The run blocks on the model's question. The hand-back is admitted as
  //    the ingest answer route admits it (assertDelegable on the stored
  //    question), queued on the same request queue, and drained by the next run.
  const blocked = await run()
  expect(blocked).toMatchObject({ status: "blocked", stop_reason: expect.stringContaining("Waiting for your answers") })
  const [asked] = readQuestions(directory)
  const id = asked!.id
  expect(id).toMatch(/^q-/)
  expect(asked).toMatchObject({ stage: "intake", question: QUESTION, status: "open" })
  expect(model.count(AGENT.spine) + model.count(AGENT.finalize)).toBe(0)
  assertDelegable(asked, id)
  const handBack = enqueueLeanRequest(directory, { kind: "answer", question_id: id, decide_for_me: true })
  expect(queuedAnswerIDs(directory)).toEqual(new Set([id]))

  // 2. The spine first records the choice as one the client approved, then
  //    leaves it out; the memo first leaves its item out. Each is refused by its
  //    checker, and the refusal is fed back as the retry.
  armed.spine.push("approved", "uncovered")
  armed.finalize.push("uncovered")
  const start = model.prompts.length
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, handBack)).toEqual({ status: "delegated" })
  const [delegated] = readQuestions(directory)
  expect(delegated).toEqual(expect.objectContaining({ id, status: "answered", answer_kind: "decide_for_me" }))
  expect(delegated).not.toHaveProperty("answer")
  expect(armed).toEqual({ spine: [], finalize: [] })
  const spines = since(start, AGENT.spine)
  expect(spines).toHaveLength(3)
  expect(spines[1]).toContain("Your reply did not pass the spine checks:")
  expect(spines[1]).toContain(`- The entry with question_id "${id}" is kind "decision", but the lawyer left that question to the drafting and has not decided it. Record it as drafting-choice or open-item, never as a decision.`)
  expect(spines[2]).toContain(`- The lawyer left this question to the drafting to decide: "${id}" "${QUESTION}". Record it in one spine entry with "question_id": "${id}": kind drafting-choice stating the choice made and why, or kind open-item when deciding it would need a fact the sources do not give.`)
  const [choice] = handedBackItems(directory)
  const finals = since(start, AGENT.finalize)
  expect(finals).toHaveLength(2)
  expect(finals[1]).toContain("Your reply did not pass the finalize checks:")
  expect(finals[1]).toContain(`do not mention these open items: ${choice!.id}. Say what each is`)

  // 3. The corrected run: one drafting-choice entry, the memo told what was
  //    chosen, and a review draft whose delivered memo discloses the choice in
  //    plain words, with no internal id.
  expect((readLeanJson(directory, "spine.json") as { entries: { kind: string; question_id?: string }[] }).entries)
    .toEqual([expect.objectContaining({ kind: "drafting-choice", question_id: id, detail: CHOSEN })])
  expect(finals[0]).toContain(`- ${choice!.id}: ${choice!.detail}\n  Drafting spine (drafting-choice): ${CHOSEN}`)
  expect(finals[0]).toContain(`- ${id} (left to the drafter): ${QUESTION}`)
  expect(gate(directory)).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(gate(directory).disclosed).toEqual([expect.objectContaining({ id: choice!.id, status: "disclosed", ref: { repair_target: "memo", question_id: id },
    resolution: "left to the drafter; named in the notes" })])
  expect(readLeanPlan(directory)).toMatchObject({ memo: "cover memo", notes_file: MEMO_FILE })
  expect(deliverables.names).toContain(MEMO_FILE)
  const plain = `${QUESTION} ${DISCLOSED} ${CHOSEN} The lawyer has not made this choice and should confirm it or answer the question.`
  for (const delivered of [readLeanText(directory, DELIVERED_NOTES.markdown)!, docxText(directory, MEMO_FILE)]) {
    expect(delivered).toContain(plain)
    for (const internal of [id, choice!.id, "[ref:", "question_id", "drafting-choice"]) expect(delivered).not.toContain(internal)
  }

  // 4. Continue, and a drain, with nothing changed run no model stage. A second
  //    hand-back is refused at admission; one that reached the queue anyway
  //    drains as already recorded and reruns nothing.
  const settled = { prompts: model.prompts.length, sessions: model.sessions.length }
  const resume = enqueueLeanRequest(directory, { kind: "continue" })
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, resume)).toEqual({ status: "resume" })
  expect(drainLeanRequests(directory)).toEqual({ answered: 0, updated: 0, repairRound: false })
  expect(() => assertDelegable(readQuestions(directory)[0], id)).toThrow("This question was already left to the drafter to decide.")
  const replay = enqueueLeanRequest(directory, { kind: "answer", question_id: id, decide_for_me: true })
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, replay)).toEqual({ status: "delegated" })
  expect({ prompts: model.prompts.length, sessions: model.sessions.length }).toEqual(settled)
  expect(gate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })

  // 5. A real answer supersedes the hand-back: it is recorded as the lawyer's,
  //    the stages that read answers rerun, a spine still carrying the drafting
  //    choice is refused, and nothing is left to disclose for the question.
  const counts = () => [AGENT.intake, AGENT.spine, AGENT.draft, AGENT.finalize].map(agent => model.count(agent))
  const before = counts()
  armed.spine.push("stale")
  const answered = model.prompts.length
  const answer = enqueueLeanRequest(directory, { kind: "answer", question_id: id, answer: ANSWER })
  expect((await run()).status).toBe("completed")
  expect(outcome(directory, answer)).toEqual({ status: "answered" })
  const [question] = readQuestions(directory)
  expect(question).toMatchObject({ id, status: "answered", answer: ANSWER })
  expect(question).not.toHaveProperty("answer_kind")
  expect(answeredQuestionsText(directory)).toBe(`- [${id}] Q: ${QUESTION}\n  A: ${ANSWER}`)
  expect(handedBackBlock(directory)).toBe("")
  expect(handedBackItems(directory)).toEqual([])
  counts().forEach((count, index) => expect(count).toBeGreaterThan(before[index]!))
  const respined = since(answered, AGENT.spine)
  expect(respined).toHaveLength(2)
  expect(respined[0]).toContain(ANSWER)
  expect(respined[0]).not.toContain("left these questions to the drafting")
  expect(respined[1]).toContain("- A spine entry kind must be decision, recommendation, requirement, fact or open-item.")
  expect((readLeanJson(directory, "spine.json") as { entries: { kind: string; question_id?: string }[] }).entries
    .filter(entry => entry.kind === "drafting-choice" || entry.question_id !== undefined)).toEqual([])
  const [refinal] = since(answered, AGENT.finalize)
  expect(refinal).toContain(`- ${id} (answered): ${QUESTION}\n  Answer: ${ANSWER}`)
  expect(refinal).not.toContain("Drafting spine (drafting-choice)")
  expect(refinal).not.toContain("<left_to_drafting>")
  expect(gate(directory).items.filter(item => item.ref.question_id === id)).toEqual([])
  expect(gate(directory)).toMatchObject({ status: "pass", delivery: "execution_ready" })
  expect(readLeanText(directory, DELIVERED_NOTES.markdown)).not.toContain(DISCLOSED)
  expect(docxText(directory, MEMO_FILE)).not.toContain(DISCLOSED)
}, 180_000)
