import { afterEach, expect, test } from "bun:test"
import { intakeQuestions } from "./intake"
import type { ChecklistCheckerInput } from "./checklist-contract"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { finalBodyParagraphs } from "./checks"
import { paragraphShowsOpenItem } from "./gate"
import { provisionalFactMarkerItems } from "./provisional-fact-markers"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { readLeanJson, readLeanText, readOpenItems, readQuestions, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, NO_FINDINGS, cleanupMatters, drafter, happyPath, matter, noChanges, scriptedModel } from "./test-support"

afterEach(cleanupMatters)
const quiet = { ...happyPath, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges }
const ready = JSON.stringify({ ready: true, jurisdiction: "No law-dependent terms", scope: "Supplied terms only", acting_for: { party: "", basis: "none" }, dates: { document_date: { date: "", basis: "today" }, operative_dates: [] } })
const original = (): Question => ({ id: "q-delivery", stage: "intake", kind: "fact", question: "What is the delivery location?", context: "The source omits the location.", status: "answered", answer: "Draft with the delivery location marked as unknown.", fact_status: "provisional", answer_class: "answered" } as Question)

// Only these local pipelines draft the literal missing-fact marker. Check its
// actual current paragraph and exact generated question duty, rather than
// teaching the shared scripted model that every marker check succeeds.
const markedFact = (directory: string, questionID: string) => (prompt: string) => {
  const input: ChecklistCheckerInput = JSON.parse(prompt.split("CHECK_INPUT:\n\n")[1]!)
  const expected = provisionalFactMarkerItems(directory, questionID).find(row => row.id === input.item_id)
  expect(expected).toBeDefined()
  expect(input.item).toEqual(expected!.item)
  const question = readQuestions(directory).find(row => row.id === questionID)!
  expect(question.fact_status).toBe("provisional")
  const current = finalBodyParagraphs(directory, input.item.target.document as "d1")
  const evidence = input.current_evidence.filter(row => row.kind === "body" && row.document === input.item.target.document &&
    row.text.includes(`[open item: ${question.question}]`) && paragraphShowsOpenItem(row.text) && current.some(paragraph => paragraph.text === row.text))
  expect(evidence.length).toBeGreaterThan(0)
  return JSON.stringify({ item_id: input.item_id, status: "satisfied", action_status: "fulfilled", open_decisions: [],
    reason: "The current delivery-location passage visibly preserves this exact missing fact.", evidence, source_refs: input.item.source_refs,
    inspection_scope: "Every location-dependent passage in the synthetic supply note.", limitation: "The location remains unknown; this checks marker handling only." })
}

test("missing-fact intake choices can be performed by the workflow", () => {
  const result = intakeQuestions(matter(), JSON.stringify({ questions: [{ question: "What is the delivery location?", kind: "fact", context: "The source omits the location." }] }))
  expect(result?.issues).toEqual([])
  expect(result?.value?.[0]?.options).toEqual([
    expect.objectContaining({ fact_action: "provisional" }),
    expect.objectContaining({ fact_action: "wait" }),
  ])
})

test("waiting for evidence survives Continue without another status question or drafting", async () => {
  const directory = matter()
  writeQuestions(directory, [{ ...original(), answer: "Wait for me to provide the location.", fact_status: "wait" } as Question])
  const model = scriptedModel(quiet)
  const first = await runLean(directory, { sessions: model.port })
  expect(first.status).toBe("blocked")
  expect(first.stop_reason).toContain("evidence")
  expect(model.prompts).toHaveLength(0)
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  expect(model.prompts).toHaveLength(0)
  expect(readQuestions(directory)).toHaveLength(1)
  enqueueLeanRequest(directory, { kind: "answer", question_id: "q-delivery", answer: "The delivery location is the buyer's depot." })
  const supplied = scriptedModel(quiet)
  expect((await runLean(directory, { sessions: supplied.port })).status).toBe("completed")
}, 120_000)

test("a reworded fact repeat is returned to intake, never shown as a new question", async () => {
  const directory = matter()
  const previous = original()
  writeQuestions(directory, [previous])
  const ask = JSON.stringify({ questions: [{ question: "Has the coordinator obtained the delivery location?", kind: "fact", context: "The location is still not established." }] })
  const model = scriptedModel({ ...quiet, [AGENT.intake]: (_prompt, index) => index === 0 ? ask : ready,
    [CHECKLIST_JOB_AGENTS.check]: markedFact(directory, previous.id),
    [AGENT.draft]: drafter(`# Supply note\n\n## Delivery location\n\nDelivery location: [open item: ${previous.question}].\n`),
    [AGENT.advise]: () => JSON.stringify({ repeat: { question_id: previous.id, question_quote: previous.question, answer_quote: previous.answer, reason: "This requests the same missing location with no new evidence." } }) })
  const result = await runLean(directory, { sessions: model.port })
  expect(readQuestions(directory)).toHaveLength(1)
  expect(model.prompts.filter(row => row.agent === AGENT.intake)[1]?.text).toContain(previous.answer!)
  expect(readOpenItems(directory).some(item => item.ref.question_id === previous.id)).toBe(true)
  expect(readLeanText(directory, "plan.json")).toBeDefined()
  expect({ status: result.status, reason: result.stop_reason }).toMatchObject({ status: "completed" })
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft" })
}, 120_000)

test("a model ignoring repeat feedback cannot redisplay the checked repeat", async () => {
  const directory = matter()
  const previous = original()
  writeQuestions(directory, [previous])
  const ask = JSON.stringify({ questions: [{ question: "Has someone confirmed the delivery location yet?", kind: "fact", context: "The location remains unknown." }] })
  const model = scriptedModel({ [AGENT.intake]: (_prompt, index) => ask.replace("Has someone confirmed", `Status ${index}: has someone confirmed`),
    [AGENT.advise]: () => JSON.stringify({ repeat: { question_id: previous.id, question_quote: previous.question, answer_quote: previous.answer, reason: "Same unresolved location." } }) })
  const result = await runLean(directory, { sessions: model.port })
  expect(result.status).toBe("failed")
  expect(result.stop_reason).toContain("already")
  expect(readQuestions(directory)).toHaveLength(1)
  expect(readLeanText(directory, "plan.json")).toBeUndefined()
  const calls = model.prompts.length
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("failed")
  expect(model.prompts).toHaveLength(calls)
}, 120_000)


test("a provisional fact survives the whole local pipeline into a marked review draft", async () => {
  const directory = matter()
  const previous = original()
  writeQuestions(directory, [previous])
  const model = scriptedModel({ ...quiet,
    [CHECKLIST_JOB_AGENTS.check]: markedFact(directory, previous.id),
    [AGENT.draft]: drafter(`# Supply note\n\n## Delivery location\n\nDelivery location: [open item: ${previous.question}].\n`),
  })
  const result = await runLean(directory, { sessions: model.port })
  expect({ status: result.status, reason: result.stop_reason, gate: readLeanJson(directory, "gate.json") }).toMatchObject({ status: "completed" })
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ items: expect.arrayContaining([expect.objectContaining({ ref: expect.objectContaining({ question_id: previous.id }), status: "disclosed" })]) })
  expect(readQuestions(directory)[0]).toMatchObject({ fact_status: "provisional" })
}, 120_000)


test("a malformed final retry cannot replenish the intake allowance on Continue", async () => {
  const directory = matter()
  const previous = original()
  writeQuestions(directory, [previous])
  const ask = JSON.stringify({ questions: [{ question: "Has the location now been obtained?", kind: "fact", context: "The location is not yet supplied." }] })
  const model = scriptedModel({ [AGENT.intake]: (_prompt, index) => index >= 4 ? "not JSON" : ask,
    [AGENT.advise]: () => JSON.stringify({ repeat: { question_id: previous.id, question_quote: previous.question, answer_quote: previous.answer, reason: "Same missing location." } }) })
  const result = await runLean(directory, { sessions: model.port })
  expect(result.status).toBe("failed")
  expect(result.stages.intake?.feedback.join(" ")).toContain("JSON")
  const calls = model.prompts.length
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("failed")
  expect(model.prompts).toHaveLength(calls)
}, 120_000)


for (const form of ["option", "text"] as const) test(`a queued later update can release a ${form} wait without replacing the answer`, async () => {
  const directory = matter()
  const question = intakeQuestions(directory, JSON.stringify({ questions: [{ question: "What is the delivery location?", kind: "fact", context: "The source omits the location." }] }))!.value![0]!
  writeQuestions(directory, [question])
  const waiting = "Wait for me to supply this fact before drafting further."
  const answer = enqueueLeanRequest(directory, form === "option"
    ? { kind: "answer", question_id: question.id, option_id: "o2" }
    : { kind: "answer", question_id: question.id, answer: waiting })
  const permission = "You may now continue with the delivery location marked as unknown in a review draft."
  const update = enqueueLeanRequest(directory, { kind: "update", text: permission })
  expect(update.id > answer.id).toBe(true)
  // Both requests arrived while work was in flight; processing happens later.
  await new Promise(resolve => setTimeout(resolve, 5))
  const model = scriptedModel({ ...quiet,
    [CHECKLIST_JOB_AGENTS.check]: markedFact(directory, question.id),
    [AGENT.answer]: prompt => prompt.includes('"drafting_permission"')
      ? JSON.stringify({ drafting_permission: "provisional", reason: "The later instruction permits a marked review draft.", basis: [{ source_id: "updated-notes", quote: permission }] })
      : prompt.includes('"evidence_status"')
      ? JSON.stringify({ evidence_status: "missing", reason: "The location is still unknown.", basis: [] })
      : JSON.stringify({ class: "answered", fact_status: "wait", fact_quote: waiting }),
    [AGENT.draft]: drafter(`# Supply note\n\n## Delivery location\n\nDelivery location: [open item: ${question.question}].\n`),
  })
  const result = await runLean(directory, { sessions: model.port })
  expect({ status: result.status, reason: result.stop_reason }).toMatchObject({ status: "completed" })
  expect(readQuestions(directory)[0]).toMatchObject({ fact_status: "provisional", answer_request_id: answer.id, answered_at: answer.received_at,
    answer: form === "option" ? question.options![1]!.text : waiting })
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft" })
}, 120_000)
