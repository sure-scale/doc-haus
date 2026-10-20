import { afterEach, describe, expect, test } from "bun:test"
import { existsSync, readFileSync } from "node:fs"
import { normalize } from "../../lib/quote-verify"
import type { ChecklistCheckerInput, ChecklistProducerInput } from "./checklist-contract"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { PLAN_COVERAGE_REVISIONS, runDraftStage } from "./draft-sections"
import { finalGate } from "./gate"
import { LEAN_COVERAGE_SCHEMA, coverageItems, coveragePlanHash, coverageScopeQuestion, planCoverageCurrent, planCoverageIssues, readLeanCoverage, type CoverageAct } from "./package-coverage"
import { PACKAGE_SCOPE_QUESTION, packageScopeQuestionID } from "./package-scope"
import { readLeanPlan, writeLeanPlan, type PlannedDocument } from "./plan"
import { EXCERPT } from "./quotes"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { outlineBasis } from "./stages"
import { leanFile, readLeanJson, readLeanState, readQuestions, shortHash, updateLeanStage, writeLeanText, writeQuestions, type LeanState, type Question } from "./state"
import { AGENT, cleanupMatters, draftingNotes, matter, noChanges, scriptedModel, type Reply } from "./test-support"

// The plan step's package-scope question: when the outline's coverage check
// still leaves an instructed act without a document executed by the parties
// who must take it after the allowed revisions, the step asks the lawyer one
// package-scope question and stops before saving the plan. The question blocks
// the walk until it is answered; with no plan saved, the outline is drawn up
// again with the answer, and the run asks no second scope
// question about an act that answered question showed, at the plan step or at
// the gate; an act it did not show gets one more question. Re-quoting a shown
// act as a shorter span of the instruction it showed does not count as new;
// quoting the sentence around a clause it showed does, since that clause's
// line may have shown another act. While the gate's scope question is
// open the plan step asks none of its own: the act is disclosed and the open
// question stays the one to answer. Synthetic parties and acts only; only the
// model is scripted, and answers go through the request queue.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare the papers so that the Orchard Committee approves the revised schedule. Kestrel Works then releases the deposit."
const APPROVE = "the Orchard Committee approves the revised schedule"
const MISPLACED = 'must be taken by "the Orchard Committee", but d1 "Release letter" is executed by "Kestrel Works"'
const REVISE = "Your reply did not pass the draft checks:"
const ANSWER = "Keep the package to the release letter; the committee's approval is handled separately."

// One document executed by the party who does not take the act, and the same
// with a companion document executed by the party who does.
const ONE = "Acting for: not stated\n# Release letter\nForm: letter\nExecuted by: Kestrel Works\n## Terms\nThe release."
const TWO = "Acting for: not stated\n# Committee resolution\nForm: other\nExecuted by: the Orchard Committee\n## Terms\nThe approval.\n" +
  "# Release letter\nForm: letter\nExecuted by: Kestrel Works\n## Terms\nThe release."

const act: CoverageAct = { act: "approve the revised schedule", source: "request", quote: APPROVE, parties: ["the Orchard Committee"], document: "d1" }
const ACTS = JSON.stringify({ acts: [act] })
/** The id the plan step gives its question on `acts` alone, through the step
 * itself. It keys on each act's quote and parties, never on its document, so
 * the acts are taken as left in none. */
const planID = (...acts: CoverageAct[]) => coverageScopeQuestion(acts.map(row => ({ ...row, document: null })), [])!.id
/** The plan step's question on `act`. */
const SCOPE_ID = planID(act)
/** The id the gate gives a scope question on the same verbatim instruction: its quotes alone. */
const GATE_ID = `${PACKAGE_SCOPE_QUESTION}${shortHash(JSON.stringify([normalize(APPROVE)]))}`

/** The checklist leaves the request's instruction open, so the gate asks about it. */
const OPEN = "Whether the package is meant to record the committee's approval"
const openChecklist: Record<string, Reply> = {
  [CHECKLIST_JOB_AGENTS.extract]: text => {
    const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
    return JSON.stringify({ items: input.source.document === "request" ? [{ check: "The package records the committee's approval of the revised schedule.", basis: "directive",
      source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" } }] : [] })
  },
  [CHECKLIST_JOB_AGENTS.check]: text => {
    const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
    return JSON.stringify({ item_id: input.item_id, status: "open", reason: `Focused result for ${input.item_id}.`, evidence: input.current_evidence,
      source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", open_decisions: [OPEN] })
  },
}

/** The outline agent: `first` answers the outline prompt, `revised` any
 * revision request, and each part gets one plain paragraph. */
const outliner = (first: string, revised = first): Reply => text =>
  /^Draft part \d+ of \d+ of "/m.test(text) ? "The terms are recorded." : text.startsWith(REVISE) ? revised : first

/** A whole run: the outline keeps one document whatever the coverage job
 * reports, and the later steps find nothing to change. */
const replies = (extra: Record<string, Reply> = {}): Record<string, Reply> => ({
  [AGENT.draft]: outliner(ONE),
  [AGENT.coverage]: () => ACTS,
  [AGENT.review]: () => "No corrections are needed.",
  [AGENT.changes]: noChanges,
  [AGENT.finalize]: draftingNotes,
  ...extra,
})

function prepared() {
  const directory = matter(SOURCE_TEXT)
  writeLeanText(directory, "request.md", REQUEST)
  return directory
}

const runStep = (directory: string, model: ReturnType<typeof scriptedModel>) =>
  runDraftStage(directory, readLeanState(directory)!, { sessions: model.port }, () => {}, "d1", () => {})

const revisions = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))
const partPrompts = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => /^Draft part \d+ of /m.test(prompt.text))
const planSaved = (directory: string) => existsSync(leanFile(directory, "plan.json"))

describe("the plan step's package-scope question", () => {
  test("coverage left uncovered after the revisions blocks the run with one package-scope question at d1.draft; no plan is saved", async () => {
    const directory = prepared()
    const model = scriptedModel(replies())
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("blocked")
    expect(state.stages["d1.draft"]?.status).toBe("blocked")
    expect(state.stop_reason).toContain("d1.draft")
    const questions = readQuestions(directory)
    expect(questions).toEqual([expect.objectContaining({ id: SCOPE_ID, stage: "draft", document: "d1", status: "open" })])
    const [question] = questions
    expect(question).not.toHaveProperty("answer")
    expect(question).not.toHaveProperty("pid")
    for (const text of ['"approve the revised schedule"', `"${APPROVE}" [source: request]`, 'taken by "the Orchard Committee"',
      "add a document, or change a planned one", "outside this request"])
      expect(question!.question).toContain(text)
    expect(question!.context).toContain("Drafting waits for your answer")
    // No plan is saved yet, so the plan is drawn up with the answer.
    expect(question!.context).toContain("The document plan is drawn up with your answer before any part is drafted.")
    // The plan is not saved; the coverage record keeps what was found.
    expect(planSaved(directory)).toBe(false)
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", acts: [act], issues: [expect.stringContaining(MISPLACED)] })
    // Every allowed revision was spent first, and no part was drafted.
    expect(revisions(model)).toHaveLength(PLAN_COVERAGE_REVISIONS)
    expect(model.count(AGENT.coverage)).toBe(PLAN_COVERAGE_REVISIONS + 1)
    expect(partPrompts(model)).toEqual([])
  }, 120_000)

  test("run again without an answer, or continued, the run stays blocked on the same one question and makes no model call", async () => {
    const directory = prepared()
    await runLean(directory, { sessions: scriptedModel(replies()).port })
    const asked = readQuestions(directory)
    expect(asked).toEqual([expect.objectContaining({ id: SCOPE_ID, status: "open" })])

    const idle = scriptedModel({})
    expect((await runLean(directory, { sessions: idle.port })).status).toBe("blocked")
    // Continue resumes the walk; it answers nothing, so the walk stops again at once.
    enqueueLeanRequest(directory, { kind: "continue" })
    const continued = await runLean(directory, { sessions: idle.port })
    expect(continued.status).toBe("blocked")
    expect(continued.stop_reason).toContain(SCOPE_ID)
    expect(idle.sessions).toEqual([])
    expect(idle.prompts).toEqual([])
    expect(readQuestions(directory)).toEqual(asked)
    expect(planSaved(directory)).toBe(false)

    // The same request on a fresh matter asks the same question.
    const again = prepared()
    await runLean(again, { sessions: scriptedModel(replies()).port })
    expect(readQuestions(again).map(row => row.id)).toEqual([SCOPE_ID])
  }, 120_000)

  test("the answer re-runs the outline with it; still uncovered, the plan is saved and no second scope question is asked", async () => {
    const directory = prepared()
    await runLean(directory, { sessions: scriptedModel(replies()).port })
    const [question] = readQuestions(directory)
    expect(question?.id).toBe(SCOPE_ID)
    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: ANSWER })

    const answered = scriptedModel(replies())
    expect((await runLean(directory, { sessions: answered.port })).status).toBe("completed")
    // The outline is asked again, with the answer, and its coverage checked again.
    const outline = answered.prompts.find(prompt => prompt.agent === AGENT.draft)!
    expect(outline.text).toContain(ANSWER)
    expect(revisions(answered)).toHaveLength(PLAN_COVERAGE_REVISIONS)
    expect(readLeanPlan(directory)).toMatchObject({ basis: outlineBasis(directory), documents: [expect.objectContaining({ title: "Release letter" })] })
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", acts: [act] })
    expect(planCoverageCurrent(directory)).toBe(true)
    expect(partPrompts(answered)).toHaveLength(1)
    // One scope question in the run: the answered one, kept as answered.
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: SCOPE_ID, stage: "draft", status: "answered", answer: ANSWER })])
  }, 120_000)

  test("a coverage reply that never passes its check (unchecked) never blocks and the plan is saved; the same outline checked and uncovered blocks", async () => {
    const unchecked = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => "The plan looks complete." })
    const result = await runStep(unchecked, model)
    expect(result.status).toBe("done")
    expect(result.status !== "failed" && (result.questions ?? []).filter(row => row.id.startsWith(PACKAGE_SCOPE_QUESTION))).toEqual([])
    expect(readLeanCoverage(unchecked)).toMatchObject({ status: "unchecked", acts: [] })
    expect(readLeanPlan(unchecked)!.documents.map(row => row.title)).toEqual(["Release letter"])

    const uncovered = prepared()
    const blocked = await runStep(uncovered, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS }))
    expect(blocked).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ id: SCOPE_ID, stage: "draft", status: "open" })] })
    expect(planSaved(uncovered)).toBe(false)
  }, 120_000)

  test("a saved current plan whose coverage record is stale is rechecked, not revised; left uncovered, the step blocks with the question", async () => {
    const directory = prepared()
    const plan = writeLeanPlan(directory, [{ title: "Release letter", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["Kestrel Works"] }], outlineBasis(directory))
    const saved = readFileSync(leanFile(directory, "plan.json"), "utf8")
    expect(planCoverageCurrent(directory)).toBe(false)
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => ACTS })
    const result = await runStep(directory, model)

    expect(result).toMatchObject({ status: "blocked", outputs: ["coverage.json"], openItems: [], reason: expect.stringContaining("d1.draft") })
    expect(result.status === "blocked" && result.questions).toEqual([expect.objectContaining({ id: SCOPE_ID, stage: "draft", status: "open",
      question: expect.stringContaining('taken by "the Orchard Committee"') })])
    // The plan is saved, so the question offers no new document: stopping this draft and starting a new one adds one.
    const [asked] = result.status === "blocked" ? result.questions! : []
    expect(asked!.question).toEndWith("Which planned document, executed by those parties, should carry it, or is it outside this request?")
    expect(asked!.question).not.toContain("add a document")
    expect(asked!.context).toContain('to add a document, stop this draft with "Stop draft", then start a new one, which plans the package from the start.')
    // No outline was asked for and no part drafted: the plan stands as saved.
    expect(model.count(AGENT.draft)).toBe(0)
    expect(model.count(AGENT.coverage)).toBe(1)
    expect(readFileSync(leanFile(directory, "plan.json"), "utf8")).toBe(saved)
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", plan: coveragePlanHash(plan.documents, plan.acting_for), acts: [act] })
  }, 120_000)

  test("once the plan step's scope question is answered, the gate asks none for an instruction left open that the question showed, and the instruction refers to it", async () => {
    const directory = prepared()
    await runLean(directory, { sessions: scriptedModel(replies(openChecklist)).port })
    const [question] = readQuestions(directory)
    expect(question).toMatchObject({ id: SCOPE_ID, stage: "draft", status: "open" })
    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: ANSWER })

    expect((await runLean(directory, { sessions: scriptedModel(replies(openChecklist)).port })).status).toBe("completed")
    const gate = readLeanJson(directory, "gate.json") as { status: string; delivery: string; disclosed: { kind: string; detail: string; ref: Record<string, string> }[] }
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
    expect(gate).not.toHaveProperty("questions")
    // The instruction left open stays disclosed. Its checklist item quotes the
    // request sentence that holds the span the answered question showed, so it
    // refers to that question, as does the uncovered act itself.
    const leftOpen = gate.disclosed.filter(item => item.kind === "checklist_open")
    expect(leftOpen).toHaveLength(1)
    expect(leftOpen[0]!.ref).toMatchObject({ question_id: SCOPE_ID })
    expect(gate.disclosed).toContainEqual(expect.objectContaining({ kind: "omission", detail: expect.stringContaining(`The lawyer answered the package-scope question ${SCOPE_ID};`) }))
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: SCOPE_ID, stage: "draft", status: "answered", answer: ANSWER })])
  }, 180_000)

  test("while the gate's scope question is open, a recheck of an edited plan that leaves an act uncovered asks no second one and keeps the draft: the act is disclosed and answering the open question unblocks the run", async () => {
    const directory = prepared()
    // The plan carries every act at first; the checklist check leaves the instruction open, so the gate asks.
    expect((await runLean(directory, { sessions: scriptedModel(replies({ ...openChecklist, [AGENT.coverage]: () => JSON.stringify({ acts: [] }) })).port })).status).toBe("completed")
    const [asked] = readQuestions(directory)
    expect(asked).toMatchObject({ stage: "gate", status: "open" })
    expect(asked!.id.startsWith(PACKAGE_SCOPE_QUESTION)).toBe(true)
    expect(asked!.id).not.toBe(SCOPE_ID)

    // The lawyer edits the saved plan's executing parties instead of answering; its recheck
    // leaves the act uncovered, and the draft reads nothing the edit changes.
    const plan = readLeanJson(directory, "plan.json") as { documents: PlannedDocument[] }
    writeLeanText(directory, "plan.json", `${JSON.stringify({ ...plan, documents: plan.documents.map(row => ({ ...row, executed_by: ["Kestrel Works", "the Heron Trust"] })) }, null, 2)}\n`)
    const body = readFileSync(leanFile(directory, "d1/body.docx"))
    const updated = scriptedModel(replies(openChecklist))
    const second = await runLean(directory, { sessions: updated.port })
    expect(second.status).toBe("completed")
    expect(updated.count(AGENT.draft)).toBe(0)
    expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(body)
    // One scope question in the run: the open one, still the question to answer.
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: asked!.id, stage: "gate", status: "open" })])
    // The plan is saved with the act recorded as uncovered, and the act is disclosed with the review draft.
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", acts: [act] })
    expect(planCoverageCurrent(directory)).toBe(true)
    const gate = readLeanJson(directory, "gate.json") as { status: string; disclosed: { kind: string; detail: string }[] }
    expect(gate.status).toBe("pass")
    expect(gate.disclosed).toContainEqual(expect.objectContaining({ kind: "omission", detail: expect.stringContaining(`directed by "${APPROVE}" [source: request]`) }))

    // Answering the open question unblocks the run; none other is asked.
    enqueueLeanRequest(directory, { kind: "answer", question_id: asked!.id, answer: ANSWER })
    const third = await runLean(directory, { sessions: scriptedModel(replies(openChecklist)).port })
    expect(third.status).toBe("completed")
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: asked!.id, stage: "gate", status: "answered", answer: ANSWER })])
  }, 300_000)

  // The plan step asks again only about acts no answered scope question showed.
  const shows = (text: string, status: "answered" | "open" = "answered", stage: "draft" | "gate" = "draft") =>
    ({ id: SCOPE_ID, stage, question: `The document plan leaves this instructed act without a document executed by the parties who must take it: "approve the revised schedule", directed by "${text}" [source: request]; to be taken by "the Orchard Committee"; planned in d1 "Release letter", executed by "Kestrel Works".`, context: "", status, ...(status === "answered" ? { answer: ANSWER } : {}) })
  const DEPOSIT = "Kestrel Works then releases the deposit"
  const release: CoverageAct = { act: "release the deposit", source: "request", quote: DEPOSIT, parties: ["the Orchard Committee"], document: "d1" }
  const SENTENCE = "Prepare the papers so that the Orchard Committee approves the revised schedule."

  test("an answered question showed the act and a re-run leaves it uncovered: no new question, and the act is disclosed with the answered-question note; one that showed only another act settles none", async () => {
    const directory = prepared()
    writeQuestions(directory, [shows(APPROVE)])
    const result = await runStep(directory, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS }))
    expect(result.status).toBe("done")
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", acts: [act] })
    expect(coverageItems(directory).map(item => item.detail)).toEqual([expect.stringContaining(`The lawyer answered the package-scope question ${SCOPE_ID};`)])
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: SCOPE_ID, status: "answered" })])
    // An answered question that showed only another act leaves this one to ask.
    const other = prepared()
    writeQuestions(other, [{ ...shows(DEPOSIT), id: `${PACKAGE_SCOPE_QUESTION}deposit` }])
    const asked = await runStep(other, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS }))
    expect(asked).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ id: SCOPE_ID, stage: "draft", status: "open" })] })
    expect(planSaved(other)).toBe(false)
  }, 120_000)

  test("an answered question showed one act and an update adds another uncovered act: one new question names only the new act, with its own id", async () => {
    const directory = prepared()
    writeQuestions(directory, [shows(APPROVE)])
    const result = await runStep(directory, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => JSON.stringify({ acts: [act, release] }) }))
    const questions = result.status === "blocked" ? result.questions : []
    expect(questions).toEqual([expect.objectContaining({ stage: "draft", status: "open", id: planID(release) })])
    expect(questions![0]!.id).not.toBe(SCOPE_ID)
    expect(questions![0]!.question).toContain(`"${DEPOSIT}"`)
    expect(questions![0]!.question).not.toContain(APPROVE)
    expect(planSaved(directory)).toBe(false)
  }, 120_000)

  test("a re-run that quotes a shorter span of the instruction an answered question showed asks nothing; one that quotes the sentence around the clause it showed asks once more", async () => {
    const directory = prepared()
    writeQuestions(directory, [shows(SENTENCE)])
    const result = await runStep(directory, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => JSON.stringify({ acts: [{ ...act, quote: APPROVE }] }) }))
    expect(result.status).toBe("done")
    expect(coverageItems(directory).map(item => item.detail)).toEqual([expect.stringContaining(`The lawyer answered the package-scope question ${SCOPE_ID};`)])
    // The answered line quoted one clause, which may have been another act's:
    // the act on the whole sentence is asked under its own id.
    const longer = prepared()
    writeQuestions(longer, [shows(APPROVE)])
    const whole = { ...act, quote: SENTENCE }
    const asked = await runStep(longer, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => JSON.stringify({ acts: [whole] }) }))
    expect(asked).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ id: planID(whole), stage: "draft", status: "open" })] })
    expect(planSaved(longer)).toBe(false)
  }, 120_000)

  test("while another stage's scope question is open the plan step asks none, even for an act no answered question showed; once that question is answered, the step asks about the act", async () => {
    const directory = prepared()
    const WAREHOUSE = "Deliver to the buyer's warehouse"
    writeQuestions(directory, [shows(WAREHOUSE, "open", "gate")])
    const result = await runStep(directory, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS }))
    expect(result.status).toBe("done")
    expect(planSaved(directory)).toBe(true)
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", acts: [act] })
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: SCOPE_ID, stage: "gate", status: "open" })])
    // Answered, the other stage's question settles only what it showed.
    const answered = prepared()
    writeQuestions(answered, [{ ...shows(WAREHOUSE, "answered", "gate"), id: `${PACKAGE_SCOPE_QUESTION}warehouse` }])
    const asked = await runStep(answered, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS }))
    expect(asked).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ id: SCOPE_ID, stage: "draft", status: "open" })] })
    expect(planSaved(answered)).toBe(false)
  }, 120_000)

  test("an uncovered act's item names an answered scope question only when that question showed the act's quote", () => {
    const directory = prepared()
    const DEPOSIT = "Kestrel Works then releases the deposit"
    const release: CoverageAct = { act: "release the deposit", source: "request", quote: DEPOSIT, parties: ["Kestrel Works"], document: null }
    const plan = writeLeanPlan(directory, [{ title: "Release letter", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["Kestrel Works"] }], outlineBasis(directory))
    writeLeanText(directory, "coverage.json", JSON.stringify({ schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "uncovered",
      acts: [act, release], issues: planCoverageIssues([act, release], plan.documents, directory) }))
    // The plan step asked about the approval alone and the lawyer answered.
    writeQuestions(directory, [{ id: SCOPE_ID, stage: "draft", question: `The document plan leaves this instructed act without a document executed by the parties who must take it: "approve the revised schedule", directed by "${APPROVE}" [source: request]; to be taken by "the Orchard Committee"; planned in d1 "Release letter", executed by "Kestrel Works".`,
      context: "", status: "answered", answer: ANSWER }])
    const [approval, deposit] = coverageItems(directory)
    expect(approval!.detail).toContain(`directed by "${APPROVE}"`)
    expect(approval!.detail).toContain(`The lawyer answered the package-scope question ${SCOPE_ID}; the plan still leaves this act so.`)
    // The deposit's release was not in that question: plain uncovered wording.
    expect(deposit!.detail).toContain(`directed by "${DEPOSIT}"`)
    expect(deposit!.detail).toEndWith("no planned document carries it.")
    expect(deposit!.detail).not.toContain("package-scope")
  })
})

// A package-scope question's id keys on what it asks about: the gate's on the
// instructions' quotes, the plan step's on each act's quote and parties. An
// answered question with the id an asker would give its next question asked
// about those very instructions or acts, even when its text does not show them
// as quoted (scopeQuestionShows). Asked again, that question would be dropped
// as answered and leave the run blocked with nothing to answer, so neither
// asker asks it: the acts or instructions are disclosed, referring to the
// answered question. Each asker also asks at most twice, once and once more:
// the coverage job may quote an answer as a new act, which no question can
// show, so once two of an asker's own package-scope questions are answered what
// it would ask about is disclosed instead. A blocked run always has an open
// question to answer.
describe("scope asker termination: fix wave 3", () => {
  const RELEASE_LETTER: PlannedDocument = { id: "d1", title: "Release letter", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["Kestrel Works"] }
  const MINUTE = "The Orchard Committee signs a minute of its approval."
  const NOTICE = "The Orchard Committee files a notice of the approval."
  const FIRST_ANSWER = `Keep the package to the release letter. ${MINUTE}`
  const SECOND_ANSWER = `That minute is outside this package. ${NOTICE}`
  /** An act the request or an answer directs in `quote`, named by its words. */
  const quoting = (quote: string): CoverageAct => ({ act: quote.replace(/\.$/, ""), source: "request", quote, parties: ["the Orchard Committee"], document: "d1" })
  /** An answered package-scope question with `id`, the id an asker would give
   * its next question; by default its text quotes nothing it asked about. */
  const answeredOn = (id: string, { stage = "draft", answer = ANSWER, question = "Which document carries the instructed act?" }: { stage?: "draft" | "gate"; answer?: string; question?: string } = {}): Question =>
    ({ id, stage, question, context: "", status: "answered", answer })
  /** The plan step's own question on `acts`, as it writes it, answered. */
  const planAnswered = (acts: CoverageAct[], answer = ANSWER): Question =>
    ({ ...coverageScopeQuestion(acts, [RELEASE_LETTER])!, status: "answered", answer })
  /** The gate's own question on one instruction left open, as it writes it, answered. */
  const gateAnswered = (quote: string): Question => ({ ...finalGate(leftOpenMatter([quote])).questions![0]!, status: "answered", answer: ANSWER })
  /** Never blocked with no open question to answer. */
  const answerable = (directory: string, state: LeanState) =>
    state.status !== "blocked" || readQuestions(directory).some(question => question.status === "open")

  // The checklist check left each quoted instruction open; the gate reads its saved output.
  function leftOpenMatter(quotes: string[]) {
    const directory = matter()
    writeLeanPlan(directory, [{ title: "Supply note", purpose: "Supply note", file: "supply-note.docx", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }])
    const items = quotes.map((quote, index) => ({ id: `c${index + 1}`, item: { check: `The note carries instruction ${index + 1}.`, basis: "directive",
      source_refs: [{ document: "request", locator: "request", quote }], qualifications: [], target: { document: "d1", role: "Supply note", treatment: "operative" } } }))
    writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["request"], sources: [{ key: "request", status: "succeeded", items }] }))
    writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: items.map(row => row.id),
      checks: items.map(row => ({ item_id: row.id, input_sha: "sha", status: "succeeded", issues: [], finding: { item_id: row.id, status: "open", reason: "Focused result open.",
        evidence: [], source_refs: [], inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", open_decisions: ["Whether the package carries it"] } })) }))
    writeLeanText(directory, "checklist/report.md", "# Checklist handoff findings\n")
    updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done" })
    return directory
  }
  const leftOpenRefs = (gate: ReturnType<typeof finalGate>) => gate.disclosed.filter(item => item.kind === "checklist_open").map(item => item.ref.question_id)

  test("the plan step never asks a question the lawyer answered: an answered question with the id it would ask settles those acts, though its text shows none, and each act's item refers to it", async () => {
    const answered = answeredOn(SCOPE_ID)
    expect(coverageScopeQuestion([act], [RELEASE_LETTER], [answered])).toBeUndefined()
    // Keyed on the same acts in any order, it settles several acts at once.
    const DEPOSIT = "Kestrel Works then releases the deposit"
    expect(coverageScopeQuestion([act, quoting(DEPOSIT)], [RELEASE_LETTER], [answeredOn(planID(quoting(DEPOSIT), act))])).toBeUndefined()
    // An answered question with another id that shows nothing leaves the act to ask.
    expect(coverageScopeQuestion([act], [RELEASE_LETTER], [{ ...answered, id: `${PACKAGE_SCOPE_QUESTION}other` }])?.id).toBe(SCOPE_ID)
    // So does the gate's id on the same quote, and the act's id leaves another
    // party's act in the same words to ask, with its own id.
    expect(coverageScopeQuestion([act], [RELEASE_LETTER], [answeredOn(GATE_ID, { stage: "gate" })])?.id).toBe(SCOPE_ID)
    const joint = { ...act, parties: [...act.parties, "Kestrel Works"] }
    expect(coverageScopeQuestion([joint], [RELEASE_LETTER], [answered])?.id).toBe(planID(joint))

    const directory = prepared()
    writeQuestions(directory, [answered])
    const result = await runStep(directory, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS }))
    expect(result.status).toBe("done")
    expect(planSaved(directory)).toBe(true)
    expect(coverageItems(directory).map(item => item.detail)).toEqual([expect.stringContaining(`The lawyer answered the package-scope question ${SCOPE_ID};`)])
    expect(readQuestions(directory)).toEqual([answered])
  }, 120_000)

  test("the gate never asks a question the lawyer answered: an answered question with the id it would ask settles those instructions, which refer to it", () => {
    const QUOTE = "Deliver to the buyer's warehouse."
    const directory = leftOpenMatter([QUOTE])
    const answered = answeredOn(packageScopeQuestionID([QUOTE]), { stage: "gate" })
    writeQuestions(directory, [answered])
    const gate = finalGate(directory)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.questions).toBeUndefined()
    expect(leftOpenRefs(gate)).toEqual([answered.id])
    // An answered question with another id that shows nothing leaves the instruction to ask.
    writeQuestions(directory, [{ ...answered, id: `${PACKAGE_SCOPE_QUESTION}other` }])
    expect(finalGate(directory).questions?.map(question => question.id)).toEqual([answered.id])
  })

  test("the plan step keys its question on each act's quote and parties, in any order, however the quote is spaced and the parties written; another party's act in the same words is another question; the gate keys on the quotes alone", () => {
    const minute = quoting(MINUTE)
    expect(planID(act, minute)).toBe(planID(minute, act))
    expect(planID({ ...act, quote: ` ${APPROVE.replaceAll(" ", "\n  ")} ` })).toBe(SCOPE_ID)
    expect(planID({ ...act, parties: ["The Orchard Committee"] })).toBe(SCOPE_ID)
    const both = ["Kestrel Works", "the Orchard Committee"]
    expect(planID({ ...act, parties: both.toReversed() })).toBe(planID({ ...act, parties: both }))
    for (const parties of [both, ["Kestrel Works"], []]) expect(planID({ ...act, parties })).not.toBe(SCOPE_ID)
    // The gate's question on the same instruction keys on its quote alone, as before.
    expect(finalGate(leftOpenMatter([APPROVE])).questions?.map(question => question.id)).toEqual([GATE_ID])
    expect(GATE_ID).not.toBe(SCOPE_ID)
  })

  test("an answered question settles each act it showed, whatever quotation marks its quote holds or wherever its excerpt was cut: only the acts beside it are asked, and once the asks are spent the act's item still names it", async () => {
    const ODD = 'The Orchard Committee fits 12" pipes at the site entrance.'
    const LONG = `The Orchard Committee approves the revised schedule, ${"which sets each delivery date, inspection window and payment milestone for the works, ".repeat(3)}and signs it.`
    const INSPECTOR = "The Orchard Committee appoints an inspector for the works."
    const [minute, notice] = [quoting(MINUTE), quoting(NOTICE)]
    for (const quote of [ODD, LONG]) {
      const shown = { ...quoting(quote), act: "take the scheduled step" }
      const answered = planAnswered([shown])
      // The long quote is shown as an excerpt the step cut.
      if (quote === LONG) expect(answered.question).toContain(`directed by "${quote.slice(0, EXCERPT - 1)}…"`)
      const step = async (questions: Question[]) => {
        const directory = matter(SOURCE_TEXT)
        writeLeanText(directory, "request.md", `${REQUEST} ${quote} ${MINUTE} ${NOTICE}`)
        writeQuestions(directory, questions)
        const coverage = () => JSON.stringify({ acts: [shown, minute, notice] })
        return { directory, result: await runStep(directory, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: coverage })) }
      }
      // The answered question showed the act: only the two beside it are asked, in a question of their own.
      const once = await step([answered])
      const questions = once.result.status === "blocked" ? (once.result.questions ?? []).map(({ id, stage, status, question }) => ({ id, stage, status, question })) : []
      expect(questions.map(({ question: _, ...rest }) => rest)).toEqual([{ id: planID(minute, notice), stage: "draft", status: "open" }])
      const asked = questions[0]!.question
      for (const text of [MINUTE, NOTICE]) expect(asked).toContain(`directed by "${text}"`)
      expect(asked).not.toContain(quote.slice(0, 40))
      // Two of the step's questions answered: none is asked, the plan is saved,
      // and the shown act's item still names the question that showed it.
      const spent = await step([answered, planAnswered([quoting(INSPECTOR)])])
      expect(spent.result.status).toBe("done")
      expect(planSaved(spent.directory)).toBe(true)
      const details = coverageItems(spent.directory).map(item => item.detail)
      expect(details).toHaveLength(3)
      expect(details[0]).toContain(`The lawyer answered the package-scope question ${answered.id};`)
      for (const detail of details.slice(1)) expect(detail).not.toContain("package-scope")
    }
  }, 120_000)

  test("the plan step asks once more, not again: with two of its questions answered, an act quoting the last answer is disclosed, not asked; the gate's answered questions do not count", async () => {
    const coverage = () => JSON.stringify({ acts: [act, quoting(MINUTE), quoting(NOTICE)] })
    const first = planAnswered([act], FIRST_ANSWER)
    const second = planAnswered([quoting(MINUTE)], SECOND_ANSWER)
    // One answered: the act quoting its answer is asked about, once more.
    const once = prepared()
    writeQuestions(once, [first])
    const more = await runStep(once, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => JSON.stringify({ acts: [act, quoting(MINUTE)] }) }))
    expect(more).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ id: planID(quoting(MINUTE)), stage: "draft", status: "open" })] })

    // Two answered: the act quoting the second answer is disclosed, and the plan saved.
    const twice = prepared()
    writeQuestions(twice, [first, second])
    const result = await runStep(twice, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: coverage }))
    expect(result.status).toBe("done")
    expect(planSaved(twice)).toBe(true)
    expect(readQuestions(twice)).toEqual([first, second])
    const details = coverageItems(twice).map(item => item.detail)
    expect(details).toHaveLength(3)
    expect(details[0]).toContain(`The lawyer answered the package-scope question ${first.id};`)
    expect(details[1]).toContain(`The lawyer answered the package-scope question ${second.id};`)
    expect(details[2]).toContain(`directed by "${NOTICE}" [source: request]`)
    expect(details[2]).not.toContain("package-scope")

    // The same two answered by the gate leave the plan step its own question.
    const byGate = prepared()
    writeQuestions(byGate, [{ ...first, stage: "gate" }, { ...second, stage: "gate" }])
    const asked = await runStep(byGate, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: coverage }))
    expect(asked).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ id: planID(quoting(NOTICE)), stage: "draft" })] })
  }, 120_000)

  test("the gate asks once more, not again: with two of its questions answered, a further instruction left open is disclosed without a question; the plan step's answered questions do not count", () => {
    const [PERIOD, PAYMENT, ADDRESS] = ["Deliver within thirty days of the order.", "Payment is due on delivery.", "Deliver to the buyer's warehouse."]
    const directory = leftOpenMatter([PERIOD, PAYMENT, ADDRESS])
    const answered = [gateAnswered(PERIOD), gateAnswered(PAYMENT)]
    writeQuestions(directory, answered)
    const gate = finalGate(directory)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.questions).toBeUndefined()
    expect(leftOpenRefs(gate)).toEqual([answered[0]!.id, answered[1]!.id, undefined])
    // One answered: the gate asks once more.
    writeQuestions(directory, answered.slice(0, 1))
    expect(finalGate(directory).questions?.map(question => question.id)).toEqual([packageScopeQuestionID([PAYMENT, ADDRESS])])
    // Two answered by the plan step: the gate still asks its own.
    writeQuestions(directory, answered.map(question => ({ ...question, stage: "draft" as const })))
    expect(finalGate(directory).questions?.map(question => question.id)).toEqual([packageScopeQuestionID([ADDRESS])])
  })

  test("through the run: an answered question whose text does not show its act's quote ends the next run with the act disclosed, never blocked with no open question", async () => {
    // The quote holds a quotation mark of its own.
    const INCH = 'Fit 12" pipes at the site entrance'
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `The Committee must act: ${INCH}. Kestrel Works then releases the deposit.`)
    const pipes: CoverageAct = { act: "approve the pipe plan", source: "request", quote: INCH, parties: ["the Committee"], document: "d1" }
    const acts = JSON.stringify({ acts: [pipes] })
    const run = () => runLean(directory, { sessions: scriptedModel(replies({ [AGENT.coverage]: () => acts })).port })
    const first = await run()
    expect(first.status).toBe("blocked")
    const [question] = readQuestions(directory)
    expect(question).toMatchObject({ id: planID(pipes), stage: "draft", status: "open" })
    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: ANSWER })

    const second = await run()
    expect(answerable(directory, second)).toBe(true)
    expect(second.status).toBe("completed")
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question!.id, status: "answered", answer: ANSWER })])
    const gate = readLeanJson(directory, "gate.json") as { status: string; disclosed: { kind: string; detail: string }[] }
    expect(gate.status).toBe("pass")
    expect(gate.disclosed).toContainEqual(expect.objectContaining({ kind: "omission",
      detail: expect.stringContaining(`The lawyer answered the package-scope question ${question!.id}; the plan still leaves this act so.`) }))
  }, 180_000)

  test("through the run: answers the coverage job quotes back as new acts bring one more question, then a disclosure; the run ends after two answers", async () => {
    const directory = prepared()
    // Each act the lawyer's answers direct is listed once those answers exist.
    const coverage: Reply = text => JSON.stringify({ acts: [act, ...[MINUTE, NOTICE].filter(quote => text.includes(quote)).map(quoting)] })
    const run = () => runLean(directory, { sessions: scriptedModel(replies({ [AGENT.coverage]: coverage })).port })
    const answer = (text: string) => {
      const open = readQuestions(directory).filter(question => question.status === "open")
      expect(open).toHaveLength(1)
      enqueueLeanRequest(directory, { kind: "answer", question_id: open[0]!.id, answer: text })
      return open[0]!.id
    }
    expect((await run()).status).toBe("blocked")
    expect(answer(FIRST_ANSWER)).toBe(SCOPE_ID)
    expect((await run()).status).toBe("blocked")
    expect(answer(SECOND_ANSWER)).toBe(planID(quoting(MINUTE)))

    const third = await run()
    expect(answerable(directory, third)).toBe(true)
    expect(third.status).toBe("completed")
    expect(readQuestions(directory).map(question => [question.id, question.status])).toEqual([[SCOPE_ID, "answered"], [planID(quoting(MINUTE)), "answered"]])
    const gate = readLeanJson(directory, "gate.json") as { status: string; disclosed: { kind: string; detail: string }[] }
    expect(gate.status).toBe("pass")
    expect(gate.disclosed).toContainEqual(expect.objectContaining({ kind: "omission", detail: expect.stringContaining(`directed by "${NOTICE}" [source: request]`) }))
  }, 300_000)
})
