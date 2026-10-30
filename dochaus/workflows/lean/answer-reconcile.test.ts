import { afterEach, expect, test } from "bun:test"
import { checkConflicts } from "./answer-reconcile"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { answeredQuestionsText, delegable, readLeanJson, readQuestions } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, NO_FINDINGS, noChanges, scriptedModel, SUPPORT, type Reply } from "./test-support"

// Answers given one question at a time, each recommendation adopted in its own
// job, are read together once before any step reads them: answers that cannot
// all be followed become one advised question for the lawyer, and the run
// waits for it. An unchanged set of answers is not read again, and the
// lawyer's answer to that question settles the point for good. Synthetic
// matter; only the model is scripted.

afterEach(cleanupMatters)

const LATE = "Should late delivery carry a remedy?"
const PRICE = "How should a delay in delivery affect the price?"
const NO_REMEDY = "Add no separate late-delivery remedy."
const REDUCTION = "Reduce the price by one percent for each week of late delivery."
const CONFLICT = "One answer adds no remedy for late delivery while the other reduces the price for it."
const BASIS = SUPPORT.map(row => ({ ...row, relevance: "The supplied terms fix the period for the supplier's delivery under this note and give no remedy for missing it." }))
const recommend = (...options: string[]) => JSON.stringify({ options: options.map(text => ({ text })), recommendation: { option: 0, rationale: "The supplied terms support this option.", basis: BASIS } })
const combined = (directory: string) => readQuestions(directory).filter(question => question.id.startsWith("answer-conflict-"))

/** A run that asks two judgment questions, each advised with a
 * recommendation, and reads the answers together with `reconcile`. */
function twoQuestions(reconcile: (ids: [string, string]) => Reply) {
  const directory = matter()
  const ids = (): [string, string] => [LATE, PRICE].map(text => readQuestions(directory).find(question => question.question === text)!.id) as [string, string]
  const model = scriptedModel({
    ...happyPath,
    [AGENT.review]: () => NO_FINDINGS,
    [AGENT.changes]: noChanges,
    [AGENT.intake]: (_prompt, index) => index === 0
      ? JSON.stringify({ questions: [
        { question: LATE, kind: "judgment", context: "The supplied terms are silent on late delivery." },
        { question: PRICE, kind: "judgment", context: "The supplied terms fix the price but not the effect of a delay." },
      ] })
      : JSON.stringify({ ready: true, jurisdiction: "No external legal determination needed for this supplied-source fixture", scope: "Supplied terms" }),
    [AGENT.advise]: prompt => prompt.includes("cannot all be followed together")
      ? recommend(NO_REMEDY, REDUCTION)
      : prompt.includes(`<question>\nKind: judgment\n${LATE}`) ? recommend(NO_REMEDY, REDUCTION) : recommend(REDUCTION, NO_REMEDY),
    [AGENT.reconcile]: (prompt, index, sessionID) => reconcile(ids())(prompt, index, sessionID),
  })
  const run = () => runLean(directory, { sessions: model.port })
  const adoptBoth = () => ids().forEach(question_id => enqueueLeanRequest(directory, { kind: "answer", question_id, adopt: true }))
  return { directory, model, run, ids, adoptBoth }
}

const flags = (ids: string[]) => JSON.stringify({ conflicts: [{ questions: ids, conflict: CONFLICT }] })

test("two adopted recommendations that cannot both be followed become one open, advised question, and the run waits for it", async () => {
  const { directory, model, run, ids, adoptBoth } = twoQuestions(ids => () => flags(ids))
  expect((await run()).status).toBe("blocked")
  adoptBoth()
  const state = await run()
  const [late, price] = ids()
  expect(readQuestions(directory).filter(question => question.id === late || question.id === price).map(question => [question.answer, question.answer_kind])).toEqual([[NO_REMEDY, "adopted"], [REDUCTION, "adopted"]])

  // The job read every answer, each marked as adopted.
  expect(model.count(AGENT.reconcile)).toBe(1)
  const read = model.prompts.find(prompt => prompt.agent === AGENT.reconcile)!.text
  for (const text of [LATE, PRICE, NO_REMEDY, REDUCTION]) expect(read).toContain(text)
  expect(read.match(/the recommended option was adopted/g)).toHaveLength(2)

  // One question, asked at intake and advised; no step ran on the conflicting answers.
  const [asked, ...more] = combined(directory)
  expect(more).toEqual([])
  expect(asked).toMatchObject({ stage: "intake", kind: "judgment", status: "open" })
  expect(asked!.question).toContain(CONFLICT)
  expect(asked!.question).toEndWith(`in place of the answers to ${[late, price].toSorted().map(id => `[${id}]`).join(", ")}.`)
  for (const text of [NO_REMEDY, REDUCTION, late, price]) expect(asked!.context).toContain(text)
  expect(asked!.advice).toHaveProperty("recommendation")
  expect(delegable(asked!)).toBe(true)
  expect(state).toMatchObject({ status: "blocked", stop_reason: `Waiting for your answers: ${asked!.id}` })
  expect(model.count(AGENT.draft)).toBe(0)

  // A Continue without an answer still waits on the same question.
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await run()).status).toBe("blocked")
  expect(combined(directory)).toEqual([asked!])
  expect(model.count(AGENT.reconcile)).toBe(1)
})

test("answers that can all be followed record nothing and the run goes on", async () => {
  const { directory, model, run, adoptBoth } = twoQuestions(() => () => JSON.stringify({ conflicts: [] }))
  await run()
  adoptBoth()
  expect((await run()).status).toBe("completed")
  expect(model.count(AGENT.reconcile)).toBe(1)
  expect(combined(directory)).toEqual([])
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ conflicts: [] })
  expect(model.count(AGENT.draft)).toBeGreaterThan(0)
})

test("a conflict naming an unknown question or a single one is sent back for a retry", async () => {
  const known = new Set(["q-a", "q-b"])
  expect(checkConflicts(JSON.stringify({ conflicts: [{ questions: ["q-a", "q-x"], conflict: CONFLICT }] }), known).issues).toEqual([expect.stringContaining('"q-x"')])
  expect(checkConflicts(JSON.stringify({ conflicts: [{ questions: ["q-a", "q-a"], conflict: CONFLICT }] }), known).issues).toEqual([expect.stringContaining("at least two")])
  expect(checkConflicts(JSON.stringify({ conflicts: [{ questions: ["q-a", "q-b"], conflict: " " }] }), known).issues).toEqual([expect.stringContaining("conflict must say")])
  expect(checkConflicts(JSON.stringify({ conflicts: [{ questions: ["q-b", "q-a"], conflict: "One" }, { questions: ["q-a", "q-b"], conflict: "Two." }] }), known))
    .toEqual({ issues: [], conflicts: [{ questions: ["q-a", "q-b"], conflict: "One. Two." }] })

  // Through the run: the job's own session is told what failed, and its corrected reply stands.
  const { directory, model, run, adoptBoth } = twoQuestions(([late]) => (_prompt, index) => index === 0
    ? JSON.stringify({ conflicts: [{ questions: [late, "q-missing"], conflict: CONFLICT }, { questions: [late], conflict: CONFLICT }] })
    : JSON.stringify({ conflicts: [] }))
  await run()
  adoptBoth()
  expect((await run()).status).toBe("completed")
  const calls = model.prompts.filter(prompt => prompt.agent === AGENT.reconcile)
  expect(calls).toHaveLength(2)
  expect(calls[1]!.sessionID).toBe(calls[0]!.sessionID)
  expect(calls[1]!.text).toContain('"q-missing"')
  expect(calls[1]!.text).toContain("at least two")
  expect(combined(directory)).toEqual([])
})

test("an unchanged set of answers is not read again; a changed answer is", async () => {
  const { directory, model, run, ids, adoptBoth } = twoQuestions(() => () => JSON.stringify({ conflicts: [] }))
  await run()
  adoptBoth()
  expect((await run()).status).toBe("completed")
  expect(model.count(AGENT.reconcile)).toBe(1)
  enqueueLeanRequest(directory, { kind: "continue" })
  await run()
  expect(model.count(AGENT.reconcile)).toBe(1)

  enqueueLeanRequest(directory, { kind: "answer", question_id: ids()[1], answer: "Leave the price unchanged by any delay." })
  await run()
  expect(model.count(AGENT.reconcile)).toBe(2)
})

test("once the lawyer settles a conflict, the same answers flagged again ask nothing more", async () => {
  const { directory, model, run, adoptBoth } = twoQuestions(ids => prompt => {
    const settled = /\[(answer-conflict-[0-9a-f]+)\] Q:/.exec(prompt)?.[1]
    return JSON.stringify({ conflicts: [{ questions: ids, conflict: CONFLICT }, ...(settled ? [{ questions: [settled, ids[0]], conflict: CONFLICT }] : [])] })
  })
  await run()
  adoptBoth()
  expect((await run()).status).toBe("blocked")
  const [asked] = combined(directory)
  enqueueLeanRequest(directory, { kind: "answer", question_id: asked!.id, adopt: true })

  expect((await run()).status).toBe("completed")
  // The settled set was read again, and nothing it flagged was asked.
  expect(model.count(AGENT.reconcile)).toBe(2)
  expect(model.prompts.filter(prompt => prompt.agent === AGENT.reconcile)[1]!.text).toContain(asked!.id)
  expect(combined(directory)).toEqual([expect.objectContaining({ id: asked!.id, status: "answered", answer: NO_REMEDY, answer_kind: "adopted" })])
  // Every step reads the lawyer's answer as settling the point in place of the earlier ones.
  expect(answeredQuestionsText(directory)).toContain(`[${asked!.id}] Q: Some of your answers cannot all be followed together: ${CONFLICT}`)
  expect(answeredQuestionsText(directory)).toContain("This answer settles the point in place of the answers to")
})
