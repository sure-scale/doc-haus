import { afterEach, expect, test } from "bun:test"
import { readdirSync, readFileSync } from "node:fs"
import { finalGate, type GateResult } from "./gate"
import { openItem } from "./quotes"
import { enqueueLeanRequest, queuedAnswerIDs, type StoredLeanRequest } from "./requests"
import { runLean } from "./runner"
import { answeredQuestionsText, delegable, readLeanJson, readOpenItems, readQuestions, writeOpenItems } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, NO_FINDINGS, noChanges, scriptedModel, SUPPORT } from "./test-support"

// "Decide for me" through the real coordinator, end to end: a run blocks on an
// advised judgment question, the lawyer adopts its recommendation through the
// request queue the ingest answer route uses, and every later stage reads the
// adopted option's words as a decision delegated by the lawyer. The drafting
// notes list it as delegated by the lawyer and the gate settles it so; an unchanged Continue or a replayed adoption does no model work; the
// lawyer's later words replace it. Synthetic matter; only the model is scripted.

afterEach(cleanupMatters)

const QUESTION = "Should late delivery carry a remedy?"
const CHOSEN = "Add no separate late-delivery remedy."
const OTHER = "Reduce the price by one percent for each week of late delivery."
const LAW = "Which law should govern the supply note?"
// Wording of the deleted hand-back, and the web button's name: neither reaches a model.
const LEFT_TO_DRAFTER = /left (?:\w+ ){0,2}to the draft(?:er|ing)|left_to_draft/i
const BUTTON = /decide[\s_-]for[\s_-]me/i

function outcome(directory: string, request: StoredLeanRequest) {
  return (readLeanJson(directory, `requests/done/${request.id}.json`) as { outcome: { status: string; detail?: string } }).outcome
}

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
      : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" }),
    [AGENT.advise]: prompt => prompt.includes(QUESTION)
      ? JSON.stringify({ options: [{ text: CHOSEN }, { text: OTHER }], recommendation: { option: 0, rationale: "The terms fix the delivery period and add no remedy.", basis: SUPPORT } })
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
  const line = `- [${late.id}] Q: ${QUESTION}\n  A (delegated by the lawyer; the recommended option was adopted): ${CHOSEN}`
  expect(answeredQuestionsText(directory)).toContain(line)
  const prompts = model.prompts.slice(start)
  for (const agent of [AGENT.discrepancy, AGENT.intake, AGENT.research, AGENT["research-review"], AGENT.spine, AGENT.draft, AGENT.coverage,
    AGENT.review, AGENT["review-independent"], AGENT["review-wording"], AGENT["review-client-disadvantage"], AGENT["change-plan"], AGENT.changes, AGENT["package-review"]])
    expect([agent, prompts.some(prompt => prompt.agent === agent && prompt.text.includes(line))]).toEqual([agent, true])
  const notes = prompts.find(prompt => prompt.agent === AGENT.finalize)!.text
  expect(notes).toContain(`- (delegated; see delegated_decisions) ${QUESTION}\n`)
  expect(notes).not.toContain(`Answer: ${CHOSEN}`)
  expect(notes).not.toContain(`A: ${CHOSEN}`)
  expect(notes).not.toContain(late.id)
  expect(notes).toContain(`<delegated_decisions>\n- Q: ${QUESTION}\n  Delegated by the lawyer.\n  Decision taken, for the lawyer to confirm: ${CHOSEN}\n  Reason: The terms fix the delivery period and add no remedy.`)
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

test("no agent's rules say a point is left to the drafter", () => {
  const agents = new URL("../../agent/", import.meta.url)
  const names = readdirSync(agents).filter(name => name.endsWith(".md"))
  expect(names.filter(name => name.startsWith("lean-")).length).toBeGreaterThan(10)
  for (const name of names) expect([name, LEFT_TO_DRAFTER.exec(readFileSync(new URL(name, agents), "utf8").replace(/\s+/g, " "))?.[0]]).toEqual([name, undefined])
})
