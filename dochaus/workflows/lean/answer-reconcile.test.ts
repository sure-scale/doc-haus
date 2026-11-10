import { afterEach, expect, test } from "bun:test"
import { checkConflicts, conflictQuestion, reconcilePrompt } from "./answer-reconcile"
import { OutputLimitError } from "./native-stage-result"
import { askedAtOrBefore, writeLeanPlan } from "./plan"
import { enqueueLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { LEAN_MAX_RETRIES } from "./stages"
import { answerQuestion, answeredQuestionsText, delegable, readLeanJson, readQuestions, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, NO_FINDINGS, noChanges, scriptedModel, SUPPORT, type Reply } from "./test-support"

// Answers given one question at a time, each recommendation adopted in its own
// job, are read together once before any step reads them, each with the
// document and paragraph it was asked about: answers that cannot all be
// followed become one advised question for the lawyer, and the run waits for
// it. An unchanged set of answers is not read again, and the lawyer's answer to
// that question settles the point for good. A conflict between one document's
// answers is that document's question, so only its steps read the answer.
// Synthetic matter; only the model is scripted.

afterEach(cleanupMatters)

const LATE = "Should late delivery carry a remedy?"
const PRICE = "How should a delay in delivery affect the price?"
const NO_REMEDY = "Add no separate late-delivery remedy."
const REDUCTION = "Reduce the price by one percent for each week of late delivery."
const CONFLICT = "One answer adds no remedy for late delivery while the other reduces the price for it."
const BASIS = SUPPORT.map(row => ({ ...row, relevance: "The supplied terms fix the period for the supplier's delivery under this note and give no remedy for missing it." }))
const recommend = (...options: string[]) => JSON.stringify({ options: options.map(text => ({ text })), recommendation: { option: options[0], rationale: "The supplied terms support this option.", basis: BASIS } })
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
      : JSON.stringify({ ready: true, jurisdiction: "No external legal determination needed for this supplied-source fixture", scope: "Supplied terms", acting_for: { party: "", basis: "none" }, dates: { document_date: { date: "", basis: "today" }, operative_dates: [] } }),
    [AGENT.advise]: prompt => prompt.includes("cannot all be followed together")
      ? recommend(NO_REMEDY, REDUCTION)
      : prompt.includes(`<question>\nKind: judgment\n${LATE}`) ? recommend(NO_REMEDY, REDUCTION) : recommend(REDUCTION, NO_REMEDY),
    [AGENT.reconcile]: (prompt, index, sessionID) => {
      if (prompt.includes("Verify ONE proposed answer conflict")) {
        const rows = JSON.parse(/<directions>\n(.*?)\n<\/directions>/s.exec(prompt)![1]!) as { question_id: string; answer: string }[]
        return JSON.stringify({ status: "conflict", reason: (JSON.parse(/<candidate>\n(.*?)\n<\/candidate>/s.exec(prompt)![1]!) as { explanation: string }).explanation, applicability: "The same late-delivery remedy cannot both be included and excluded.",
          directions: rows.map(row => ({ question_id: row.question_id, quote: row.answer })), basis: [] })
      }
      return reconcile(ids())(prompt, index, sessionID)
    },
  })
  const run = () => runLean(directory, { sessions: model.port })
  const adoptBoth = () => ids().forEach(question_id => enqueueLeanRequest(directory, { kind: "answer", question_id, adopt: true }))
  return { directory, model, run, ids, adoptBoth }
}

// Discovery and each independent candidate review have separate native sessions.
function reconciliationCalls(model: ReturnType<typeof scriptedModel>) {
  const reviews = new Set(model.prompts.filter(row => row.text.includes("Verify ONE proposed answer conflict")).map(row => row.sessionID))
  return model.prompts.filter(row => row.agent === AGENT.reconcile && !reviews.has(row.sessionID))
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
  expect(reconciliationCalls(model).length).toBe(1)
  const read = reconciliationCalls(model)[0]!.text
  for (const text of [LATE, PRICE, NO_REMEDY, REDUCTION]) expect(read).toContain(text)
  expect(read.match(/the recommended option was adopted/g)).toHaveLength(2)
  expect(read.match(/\n {2}Asked about: the whole work\n/g)).toHaveLength(2)

  // One run-level question, asked at intake and advised, quoting the questions
  // it replaces and carrying their ids apart from its words; no step ran on
  // the conflicting answers.
  const [asked, ...more] = combined(directory)
  expect(more).toEqual([])
  expect(asked).toMatchObject({ stage: "intake", kind: "judgment", status: "open", replaces: [late, price].toSorted() })
  expect(asked).not.toHaveProperty("document")
  expect(asked!.question).toContain(CONFLICT)
  expect(asked!.question).toMatch(/in place of the answers to "[^"]+" and "[^"]+"\.$/)
  for (const text of [LATE, PRICE]) expect(asked!.question).toContain(`"${text}"`)
  for (const text of [NO_REMEDY, REDUCTION, LATE, PRICE]) expect(asked!.context).toContain(text)
  for (const text of [late, price]) expect(`${asked!.question}\n${asked!.context}`).not.toContain(text)
  expect(asked!.advice).toHaveProperty("recommendation")
  expect(delegable(asked!)).toBe(true)
  expect(state).toMatchObject({ status: "blocked", stop_reason: `Waiting for your answers: ${asked!.id}` })
  expect(model.count(AGENT.draft)).toBe(0)

  // A Continue without an answer still waits on the same question.
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await run()).status).toBe("blocked")
  expect(combined(directory)).toEqual([asked!])
  expect(reconciliationCalls(model).length).toBe(1)
})

test("answers that can all be followed record nothing and the run goes on", async () => {
  const { directory, model, run, adoptBoth } = twoQuestions(() => () => JSON.stringify({ conflicts: [] }))
  await run()
  adoptBoth()
  expect((await run()).status).toBe("completed")
  expect(reconciliationCalls(model).length).toBe(1)
  expect(combined(directory)).toEqual([])
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ conflicts: [] })
  expect(model.count(AGENT.draft)).toBeGreaterThan(0)
})

test("conflict discovery cut off at its output limit stops drafting without caching agreement", async () => {
  const { directory, model, run, adoptBoth } = twoQuestions(() => () => { throw new OutputLimitError("The model's reply reached its output limit before giving any answer.") })
  await run()
  adoptBoth()
  expect(await run()).toMatchObject({ status: "failed", stop_reason: expect.stringContaining("could not be verified") })
  expect(reconciliationCalls(model).length).toBe(1)
  expect(model.count(AGENT.draft)).toBe(0)
  expect(combined(directory)).toEqual([])
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ verification_failed: true })
  expect(readLeanJson(directory, "answer-reconcile.json")).not.toHaveProperty("conflicts")

  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await run()).status).toBe("failed")
  expect(reconciliationCalls(model).length).toBe(1)
})

test.each([
  ["malformed JSON", "The comparison could not be completed."],
  ["missing conflicts list", JSON.stringify({ compatible: true })],
  ["unknown answer references", JSON.stringify({ conflicts: [{ questions: ["q-absent", "q-missing"], conflict: CONFLICT }] })],
])("exhausted conflict discovery is not agreement: %s", async (_name, reply) => {
  let corrected = false
  const { directory, model, run, adoptBoth, ids } = twoQuestions(() => () => corrected ? JSON.stringify({ conflicts: [] }) : reply)
  expect((await run()).status).toBe("blocked")
  adoptBoth()
  expect(await run()).toMatchObject({ status: "failed", stop_reason: expect.stringContaining("could not be verified") })
  const calls = reconciliationCalls(model)
  expect(calls).toHaveLength(LEAN_MAX_RETRIES + 1)
  expect(new Set(calls.map(row => row.sessionID)).size).toBe(1)
  expect(model.count(AGENT.draft)).toBe(0)
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ verification_failed: true })
  expect(readLeanJson(directory, "answer-reconcile.json")).not.toHaveProperty("conflicts")

  // An unchanged Continue cannot replenish discovery's retry allowance.
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await run()).status).toBe("failed")
  expect(reconciliationCalls(model)).toHaveLength(calls.length)

  // A genuine changed human answer reopens discovery, and only its checked
  // empty result admits drafting and replaces the failed result.
  corrected = true
  enqueueLeanRequest(directory, { kind: "answer", question_id: ids()[1], answer: "Keep the price unchanged if delivery is late." })
  expect((await run()).status).toBe("completed")
  expect(reconciliationCalls(model)).toHaveLength(calls.length + 1)
  expect(model.count(AGENT.draft)).toBeGreaterThan(0)
  expect(readLeanJson(directory, "answer-reconcile.json")).toMatchObject({ conflicts: [] })
  expect(readLeanJson(directory, "answer-reconcile.json")).not.toHaveProperty("verification_failed")
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
  const calls = reconciliationCalls(model)
  expect(calls).toHaveLength(2)
  expect(calls[1]!.sessionID).toBe(calls[0]!.sessionID)
  expect(calls[1]!.text).toContain('"q-missing"')
  expect(calls[1]!.text).toContain("at least two")
  expect(combined(directory)).toEqual([])
})

test("a conflict naming an answer by its id is sent back with how to say it, and on the last attempt is put in words", async () => {
  const named = ([late, price]: [string, string]) => JSON.stringify({ conflicts: [{ questions: [late, price], conflict: `The answer to ${late} adds no remedy while ${price} reduces the price.` }] })
  const { directory, model, run, ids, adoptBoth } = twoQuestions(ids => (_prompt, index) => index === 0 ? named(ids) : flags(ids))
  await run()
  adoptBoth()
  await run()
  const [late, price] = ids()
  const calls = reconciliationCalls(model)
  expect(calls).toHaveLength(2)
  for (const id of [late, price]) expect(calls[1]!.text).toContain(`names the internal id "${id}". The lawyer sees only words: name the question or answer by what it asked.`)
  expect(combined(directory)[0]!.question).toContain(CONFLICT)

  // Never corrected: the last reply's conflict is kept, each id put as its asking.
  const always = twoQuestions(ids => () => named(ids))
  await always.run()
  always.adoptBoth()
  await always.run()
  expect(reconciliationCalls(always.model).length).toBe(LEAN_MAX_RETRIES + 1)
  const [asked] = combined(always.directory)
  for (const id of always.ids()) expect(`${asked!.question}\n${asked!.context}`).not.toContain(id)
  expect(asked!.question).toContain("adds no remedy while")
  expect(asked!.question.split("cannot all be followed together:")[1]!.split("in place of")[0]).toContain(LATE.slice(0, 20))
})

test("an unchanged set of answers is not read again; a changed answer is", async () => {
  const { directory, model, run, ids, adoptBoth } = twoQuestions(() => () => JSON.stringify({ conflicts: [] }))
  await run()
  adoptBoth()
  expect((await run()).status).toBe("completed")
  expect(reconciliationCalls(model).length).toBe(1)
  enqueueLeanRequest(directory, { kind: "continue" })
  await run()
  expect(reconciliationCalls(model).length).toBe(1)

  enqueueLeanRequest(directory, { kind: "answer", question_id: ids()[1], answer: "Leave the price unchanged by any delay." })
  await run()
  expect(reconciliationCalls(model).length).toBe(2)
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
  expect(reconciliationCalls(model).length).toBe(2)
  expect(reconciliationCalls(model)[1]!.text).toContain(asked!.id)
  expect(combined(directory)).toEqual([expect.objectContaining({ id: asked!.id, status: "answered", answer: NO_REMEDY, answer_kind: "adopted" })])
  // Every step reads the lawyer's answer as settling the point in place of the earlier ones.
  expect(answeredQuestionsText(directory)).toContain(`[${asked!.id}] Q: Some of your answers cannot all be followed together: ${CONFLICT}`)
  expect(answeredQuestionsText(directory)).toContain("This answer settles the point in place of the answers to")
})

const PID = "p0123456789ab"
const PLAN = [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }]
/** An answered question, run-level unless `scope` names a document or paragraph. */
const decided = (id: string, question: string, answer: string, scope: Partial<Pick<Question, "document" | "pid" | "stage">> = {}): Question =>
  ({ id, stage: "intake", question, context: "", status: "answered", answer, answer_kind: "option", option_id: "o1", options: [{ id: "o1", text: answer }], kind: "judgment", ...scope })

test("the job is told what each answer was asked about: its document and paragraph, or else the whole work", () => {
  const titles = new Map([["d1", "Supply note"]])
  const prompt = reconcilePrompt([decided("q-a", LATE, NO_REMEDY, { stage: "changes", document: "d1", pid: PID }), decided("q-b", PRICE, REDUCTION, { stage: "changes", document: "d1" }), decided("q-c", LATE, NO_REMEDY)], "Draft the note.", "", titles)
  expect(prompt).toContain(`[q-a] Q: ${LATE}\n  Asked about: Supply note, paragraph [${PID}]\n`)
  expect(prompt).toContain(`[q-b] Q: ${PRICE}\n  Asked about: Supply note\n`)
  expect(prompt).toContain(`[q-c] Q: ${LATE}\n  Asked about: the whole work\n`)
  expect(prompt).toContain("Answers asked about different paragraphs are different points unless following one makes the other impossible.")
})

test("a conflict question belongs to the one document every answer it names belongs to, and is run-level otherwise", () => {
  const a = decided("q-a", LATE, NO_REMEDY, { stage: "changes", document: "d1", pid: PID })
  const b = decided("q-b", PRICE, REDUCTION, { stage: "changes", document: "d1" })
  const c = decided("q-c", "Should the note name a delivery address?", "Name the address given in the terms.")
  const same = conflictQuestion([a, b, c], { questions: ["q-a", "q-b"], conflict: CONFLICT })
  expect(same).toMatchObject({ stage: "intake", document: "d1", replaces: ["q-a", "q-b"] })
  expect(same).not.toHaveProperty("pid")
  expect(conflictQuestion([a, b, c], { questions: ["q-a", "q-c"], conflict: CONFLICT })).not.toHaveProperty("document")
  // The lawyer reads the questions by their words, however many; the ids stay apart.
  const three = conflictQuestion([a, b, c], { questions: ["q-a", "q-b", "q-c"], conflict: CONFLICT })
  expect(three.question).toEndWith(`in place of the answers to "${LATE}", "${PRICE}" and "${c.question}".`)
  expect(three.replaces).toEqual(["q-a", "q-b", "q-c"])
  expect(`${three.question}\n${three.context}`).not.toMatch(/\[q-|\[p0/)
})

test("a conflict between one document's answers is that document's question: answering it keeps the preparation steps done", async () => {
  const directory = matter()
  writeLeanPlan(directory, PLAN)
  writeQuestions(directory, [decided("q-a", LATE, NO_REMEDY, { stage: "changes", document: "d1", pid: PID }), decided("q-b", PRICE, REDUCTION, { stage: "changes", document: "d1", pid: PID })])
  const model = scriptedModel({ ...happyPath, [AGENT.reconcile]: prompt => {
    if (prompt.includes("Verify ONE proposed answer conflict")) {
      const rows = JSON.parse(/<directions>\n(.*?)\n<\/directions>/s.exec(prompt)![1]!) as { question_id: string; answer: string }[]
      return JSON.stringify({ status: "conflict", reason: (JSON.parse(/<candidate>\n(.*?)\n<\/candidate>/s.exec(prompt)![1]!) as { explanation: string }).explanation, applicability: "The same late-delivery remedy cannot both be included and excluded.",
        directions: rows.map(row => ({ question_id: row.question_id, quote: row.answer })), basis: [] })
    }
    return flags(["q-a", "q-b"])
  }, [AGENT.advise]: () => recommend(NO_REMEDY, REDUCTION) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  const [asked, ...more] = combined(directory)
  expect(more).toEqual([])
  expect(asked).toMatchObject({ stage: "intake", document: "d1", status: "open", replaces: ["q-a", "q-b"] })

  const keys = () => Object.fromEntries((["instructions", "spine", "d1.change-plan"] as const).map(step => [step, stageInputSha(directory, step)]))
  const before = keys()
  answerQuestion(directory, asked!.id, { text: NO_REMEDY })
  const after = keys()
  // No preparation step reads it; the document's steps do.
  expect(answeredQuestionsText(directory, { asked: askedAtOrBefore(directory, "spine") })).toBe("")
  expect(answeredQuestionsText(directory, { asked: askedAtOrBefore(directory, "d1.review") })).toContain(`[${asked!.id}] Q:`)
  expect([after.instructions, after.spine]).toEqual([before.instructions, before.spine])
  expect(after["d1.change-plan"]).not.toBe(before["d1.change-plan"])
})
