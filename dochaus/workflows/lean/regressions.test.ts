import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync, writeFileSync } from "node:fs"
import { checkChanges } from "./change-checks"
import { readChangeList, readParagraphs, settleApply } from "./checks"
import { finalGate, type GateResult } from "./gate"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { runApply } from "./apply-stage"
import { stagePrompt } from "./prompts"
import { parsePlan, writeLeanPlan } from "./plan"
import { acceptStageReply } from "./stages"
import { currentOpenItems, leanFile, readLeanJson, readLeanState, readOpenItems, readQuestions, replaceStageOpenItems, updateLeanStage, writeLeanText } from "./state"
import { AGENT, DRAFT, NO_FINDINGS, SOURCE, SUPPORT, changes, cleanupMatters, drafter, happyPath, matter, noChanges, pidOf, sameSession, scriptedModel, type Reply } from "./test-support"

// Regression and causal tests for the lean gate, the repair loop and the
// applier's outcomes. Real coordinator, checkers, applier and gate; only the
// model is scripted.

afterEach(cleanupMatters)

// A change list applied without review can hold overlapping paragraph
// proposals. Exercise that apply/notes boundary directly: runLean batches
// consolidate replacements before change-review, so they never reach it.
async function applyLegacyList(directory: string, model: ReturnType<typeof scriptedModel>, draft = DRAFT) {
  const plan = parsePlan(`Acting for: not stated\n${draft.replace(/^# .+$/m, "$&\nForm: other")}`)
  expect(plan.issues).toEqual([])
  writeLeanPlan(directory, plan.value!.documents)
  const rendered = await acceptStageReply(directory, "d1.draft", draft)
  expect(rendered.issues).toEqual([])
  // Written directly: the changes step refuses two amendments to one paragraph.
  const changes = checkChanges(directory, await model.port.prompt({
    directory, sessionID: "ses_legacy_changes", agent: AGENT.changes, text: stagePrompt(directory, "d1.changes"),
  }), readParagraphs(directory, "d1"), { document: "d1" })
  expect(changes.issues).toEqual([])
  writeLeanText(directory, "d1/changes.json", `${JSON.stringify(changes.changes, null, 2)}\n`)
  replaceStageOpenItems(directory, "changes", "d1", changes.openItems)
  const applied = await runApply(directory, readLeanState(directory)!, "d1")
  expect(applied.status).toBe("done")
  if (applied.status !== "done") throw new Error(applied.reason)
  replaceStageOpenItems(directory, "apply", "d1", applied.openItems)
  updateLeanStage(directory, readLeanState(directory)!, "d1.apply", { status: "done", outputs: applied.outputs })
  const notes = await acceptStageReply(directory, "finalize", await model.port.prompt({
    directory, sessionID: "ses_legacy_notes", agent: AGENT.finalize, text: stagePrompt(directory, "finalize"),
  }))
  expect(notes.issues).toEqual([])
  const gate = finalGate(directory)
  writeLeanText(directory, "gate.json", `${JSON.stringify(gate, null, 2)}\n`)
  expect(gate.status).toBe("pass")
  return applied
}

const DELIVERY = "The supplier shall deliver the goods."
const twoEdits = (second: string): Reply => prompt => JSON.stringify({
  amendments: [
    { pid: pidOf(prompt, DELIVERY), updated: "The supplier shall deliver the goods within thirty days of the order date.", reason: "Period.", sources: SUPPORT },
    { pid: pidOf(prompt, DELIVERY), updated: second, reason: "Second edit.", sources: SUPPORT },
  ],
  questions: [],
})
const gateOf = (directory: string) => readLeanJson(directory, "gate.json") as GateResult
const quiet = { [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges }

describe("applier outcomes are visible open items", () => {
  test("conflicting edits to one paragraph: apply keeps the first, the second is unapplied with the applier's detail and disclosed", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: twoEdits("The supplier shall deliver the goods within ten days.") })
    const state = await applyLegacyList(directory, model)

    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ id: "a1" }], unapplied: [{ id: "a2" }] })
    const item = readOpenItems(directory).find(row => row.kind === "unapplied_amendment")!
    expect(item).toMatchObject({ stage: "apply", ref: { amendment_id: "a2" } })
    expect(item.applier_detail).toMatchObject({ target: "The supplier shall deliver the goods within ten days." })
    expect(model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text).toMatch(/<unapplied_changes>\n- a2/)
    expect(state.status).toBe("done")
    expect(gateOf(directory)).toMatchObject({ status: "pass", delivery: "review_draft", disclosed: [expect.objectContaining({ id: item.id })] })
  }, 120_000)

  test("an overlapping edit is left unapplied; drafting notes that stop naming it block the gate", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: twoEdits("The supplier shall promptly deliver the goods.") })
    expect((await applyLegacyList(directory, model)).status).toBe("done")
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ unapplied: [{ id: "a2", reason: "stale" }] })
    expect(gateOf(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })

    writeFileSync(leanFile(directory, "drafting-notes.md"), "# Drafting notes\n\nNothing to report.\n")
    const gate = finalGate(directory)
    expect(gate).toMatchObject({ status: "blocked", delivery: null })
    expect(gate.blocking).toEqual([expect.objectContaining({ kind: "unapplied_amendment", resolution: "the drafting notes do not name this unapplied change" })])
  }, 120_000)

  test("an applier that throws or reports no outcome makes every affected amendment apply_failed; the body is unchanged and the run completes", async () => {
    const directory = matter()
    expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
    const list = readChangeList(directory, "d1")!

    const missing = settleApply(list, { applied: [], unapplied: [] }, true)
    expect(missing.bodyUnchanged).toBe(false)
    expect(missing.result.unapplied).toEqual([expect.objectContaining({ id: "a1", reason: "apply_failed", detail: expect.objectContaining({ issue: "The applier reported no outcome for this amendment." }) })])
    expect(settleApply(list, undefined, false, "boom")).toMatchObject({ bodyUnchanged: true, result: { applied: [], unapplied: [{ id: "a1", reason: "apply_failed" }] } })

    // A change list the applier refuses to read makes it throw.
    const raw = JSON.parse(readFileSync(leanFile(directory, "d1/changes.json"), "utf8"))
    raw.amendments[0].existing_sha256 = "not-a-hash"
    writeFileSync(leanFile(directory, "d1/changes.json"), JSON.stringify(raw, null, 2))
    const state = await runLean(directory, { sessions: scriptedModel(happyPath).port })
    expect(state.status).toBe("completed")
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [], unapplied: [{ id: "a1", reason: "apply_failed" }] })
    expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)).toContain(DELIVERY)
    const item = readOpenItems(directory).find(row => row.kind === "unapplied_amendment")!
    expect(item.detail).toContain("The applier stopped: invalid_amendments")
    expect(gateOf(directory)).toMatchObject({ status: "pass", delivery: "review_draft", disclosed: [expect.objectContaining({ id: item.id })] })
  }, 120_000)
})

describe("unsupported content in the final body", () => {
  test("an unsupported citation named only in the drafting notes blocks; marked in its paragraph and named, it is disclosed", async () => {
    const cited = "The supplier shall deliver the goods as required by section 4 of the Imaginary Trade Act"
    const blocked = matter()
    await runLean(blocked, { sessions: scriptedModel({ ...happyPath, ...quiet, [AGENT.draft]: drafter(`# Supply note\n\n## Delivery\n\n${cited}.\n\n## Payment\n\nPayment is due on delivery.\n`) }).port })
    const gate = gateOf(blocked)
    expect(gate.status).toBe("blocked")
    expect(gate.blocking).toEqual([expect.objectContaining({ kind: "citation_unsupported", resolution: "named in the notes, but its paragraph has no visible open-item marker" })])

    const marked = matter()
    const state = await runLean(marked, { sessions: scriptedModel({
      ...happyPath, ...quiet, [AGENT.draft]: drafter(`# Supply note\n\n## Delivery\n\n${cited} [open item: confirm this authority].\n\n## Payment\n\nPayment is due on delivery.\n`),
    }).port })
    expect(state.status).toBe("completed")
    expect(gateOf(marked)).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gateOf(marked).disclosed.map(item => item.kind)).toContain("citation_unsupported")
  }, 120_000)

  test("a body placeholder named in the drafting notes is disclosed and the output is a review draft", async () => {
    const directory = matter()
    const state = await runLean(directory, { sessions: scriptedModel({
      ...happyPath, ...quiet, [AGENT.draft]: drafter("# Supply note\n\n## Delivery\n\nThe supplier shall deliver the goods by [open item: confirm the delivery date].\n\n## Payment\n\nPayment is due on delivery.\n"),
    }).port })
    expect(state.status).toBe("completed")
    const gate = gateOf(directory)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.disclosed.map(item => item.kind)).toEqual(["placeholder"])
    expect(readFileSync(leanFile(directory, "drafting-notes-coverage.md"), "utf8")).toContain(gate.disclosed[0]!.id)
  }, 120_000)

  test("a placeholder the draft and the final body both carry is one open item once the body is applied; before that, the draft's finding stands", async () => {
    const marked = drafter("# Supply note\n\n## Delivery\n\nThe supplier shall deliver the goods by [open item: confirm the delivery date].\n\n## Payment\n\nPayment is due on delivery.\n")
    const placeholders = (items: { kind: string; stage: string }[]) => items.filter(item => item.kind === "placeholder").map(item => item.stage)
    const applied = matter()
    expect((await runLean(applied, { sessions: scriptedModel({ ...happyPath, ...quiet, [AGENT.draft]: marked }).port })).status).toBe("completed")
    // The draft's finding is kept for the changes step, but neither the lawyer's
    // list nor the gate counts it next to the final body's.
    expect(placeholders(readOpenItems(applied))).toEqual(["draft", "apply"])
    expect(placeholders(currentOpenItems(applied))).toEqual(["apply"])
    expect(placeholders(gateOf(applied).items)).toEqual(["apply"])

    const stopped = matter()
    const failing = scriptedModel({ ...happyPath, [AGENT.draft]: marked, [AGENT.changes]: () => "no change list" })
    expect((await runLean(stopped, { sessions: failing.port })).status).toBe("failed")
    expect(placeholders(currentOpenItems(stopped))).toEqual(["draft"])
  }, 120_000)
})

describe("misquoted amendments", () => {
  // A retry arrives in the same session and does not repeat the paragraph list,
  // so the scripted redliner keeps the pid it was first shown.
  const redliner = (quote: (index: number) => string): Reply => {
    const seen = { pid: "" }
    return (prompt, index) => {
      seen.pid = index === 0 ? pidOf(prompt, DELIVERY) : seen.pid
      return JSON.stringify({ amendments: [{ pid: seen.pid, updated: "The supplier shall deliver the goods within thirty days of the order date.",
        reason: "Period.", sources: [{ source_id: SOURCE, quote: quote(index) }] }], questions: [] })
    }
  }
  const MISQUOTE = "within 30 days of ordering"

  test("a misquote is retried with the checker's feedback; a retry that quotes exactly is applied", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: redliner(index => index === 0 ? MISQUOTE : SUPPORT[0]!.quote) })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    expect(model.count(AGENT.changes)).toBe(2)
    expect(model.prompts.filter(prompt => prompt.agent === AGENT.changes)[1]!.text)
      .toContain(`quotes "${MISQUOTE}" as from "${SOURCE}", but those words are not in that document.`)
    expect(gateOf(directory)).toMatchObject({ status: "pass", delivery: "execution_ready" })
  }, 120_000)

  test("after the last retry the amendment is withheld: not applied, disclosed, and the run passes as a review draft", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: redliner(() => MISQUOTE) })
    const state = await runLean(directory, { sessions: model.port })

    expect(model.count(AGENT.changes)).toBe(4)
    model.prompts.filter(prompt => prompt.agent === AGENT.changes).slice(1)
      .forEach(prompt => expect(prompt.text).toContain("but those words are not in that document."))
    const list = readChangeList(directory, "d1")!
    expect(list.amendments).toEqual([])
    expect(list.withheld).toEqual([expect.objectContaining({ amendment: expect.objectContaining({ id: "a1" }) })])
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [], unapplied: [] })
    expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)).toContain(DELIVERY)
    const item = readOpenItems(directory).find(row => row.kind === "unapplied_amendment")!
    expect(item).toMatchObject({ stage: "changes" })
    expect(item.detail).toContain(MISQUOTE)
    expect(state.status).toBe("completed")
    expect(gateOf(directory)).toMatchObject({ status: "pass", delivery: "review_draft", disclosed: [expect.objectContaining({ id: item.id })] })
  }, 120_000)
})

describe("questions for the lawyer", () => {
  const QUESTION = "Should late delivery carry a remedy?"

  test("an unanswered question stops application even when its proposed paragraph has a placeholder", async () => {
    const directory = matter()
    const marked: Reply = prompt => JSON.stringify({
      amendments: [{ pid: pidOf(prompt, DELIVERY), updated: "The supplier shall deliver the goods within thirty days of the order date [open item: remedy for late delivery].",
        reason: "Period, with the remedy left open.", sources: SUPPORT }],
      questions: [{ question: QUESTION, kind: "judgment", context: "The terms are silent.", pid: pidOf(prompt, DELIVERY) }],
    })
    const state = await runLean(directory, { sessions: scriptedModel({ ...happyPath, [AGENT.changes]: marked }).port })
    expect(state.status).toBe("blocked")
    expect(gateOf(directory)).toBeUndefined()
    expect(state.stages["d1.apply"]?.status).toBe("pending")
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ status: "open" })])
  }, 120_000)

  const asker = (reply: (pid: string, index: number) => object) => sameSession((prompt, index) => JSON.stringify(reply(pidOf(prompt, DELIVERY), index)))
  const PERIOD = "The supplier shall deliver the goods within thirty days of the order date"
  const asked = (pid: string, marker: string) => ({
    amendments: [{ pid, updated: `${PERIOD}${marker}.`, reason: "Period.", sources: SUPPORT }],
    questions: [{ question: QUESTION, kind: "judgment", context: "The terms are silent.", pid }],
  })
  const unmarked = (pid: string) => asked(pid, "")
  const marked = (pid: string) => asked(pid, " [open item: remedy for late delivery]")
  const retryText = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.changes)[1]!.text

  test("a question whose paragraph will not show it is retried; the accepted question pauses the run", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: asker((pid, index) => index === 0 ? unmarked(pid) : marked(pid)) })
    const state = await runLean(directory, { sessions: model.port })

    const [question] = readQuestions(directory)
    const pid = question!.pid!
    expect(gateOf(directory)).toBeUndefined()
    expect(readOpenItems(directory)).toContainEqual(expect.objectContaining({ kind: "question", ref: expect.objectContaining({ question_id: question!.id, pid }) }))
    expect(state.status).toBe("blocked")
    expect(model.count(AGENT.changes)).toBe(2)
    expect(retryText(model)).toContain(`questions[0] affects paragraph [${pid}], but that paragraph will not show the open item. Add an amendment to [${pid}]`)
  }, 120_000)

  test("the same unmarked question on every attempt: the last attempt is accepted and the run waits for an answer", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: asker(unmarked) })
    const state = await runLean(directory, { sessions: model.port })

    expect(model.count(AGENT.changes)).toBe(4)
    expect(state.stages["d1.changes"]?.status).toBe("done")
    expect(state.status).toBe("blocked")
    expect(gateOf(directory)).toBeUndefined()
    expect(state.stop_reason).toContain("Waiting for your answers")
  }, 120_000)

  test("a question tied to no paragraph is retried with an instruction to tie it to the paragraph it affects", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: asker((pid, index) => index === 0
      ? { ...unmarked(pid), questions: [{ question: QUESTION, kind: "judgment", context: "The terms are silent." }] }
      : marked(pid)) })
    const state = await runLean(directory, { sessions: model.port })

    expect(model.count(AGENT.changes)).toBe(2)
    expect(retryText(model)).toContain('questions[0] is tied to no paragraph of the draft. Set its "pid" to the paragraph whose text depends on the answer')
    expect(state.status).toBe("blocked")
    expect(gateOf(directory)).toBeUndefined()
  }, 120_000)

  test("a question ties to its id: reworded with the same id it stays one question, and once answered it is not asked again", async () => {
    const directory = matter()
    const asking = (wording: string, withID: boolean) => sameSession(prompt => JSON.stringify({ ...JSON.parse(changes(prompt, 0)), questions: [{
      ...(withID ? { id: readQuestions(directory)[0]!.id } : {}), question: wording, kind: "judgment", context: "The terms are silent.",
    }] }))
    await runLean(directory, { sessions: scriptedModel({ ...happyPath, [AGENT.changes]: asking(QUESTION, false) }).port })
    const [question] = readQuestions(directory)

    // Continue cannot bypass the unanswered question or launch another model call.
    enqueueLeanRequest(directory, { kind: "continue" })
    const reworded = scriptedModel({ ...happyPath, [AGENT.changes]: asking("Should a remedy apply to late delivery?", true) })
    expect((await runLean(directory, { sessions: reworded.port })).status).toBe("blocked")
    expect(reworded.prompts).toHaveLength(0)
    expect(readQuestions(directory).map(row => row.id)).toEqual([question!.id])

    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: "No remedy." })
    const answered = scriptedModel({ ...happyPath, [AGENT.changes]: asking("Is a remedy for late delivery wanted?", true) })
    const state = await runLean(directory, { sessions: answered.port })
    // The answer reaches the accepted draft through its review and changes; no redraft.
    expect(answered.counts()).toEqual([0, 1, 1, 1])
    expect(state.status).toBe("completed")
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question!.id, status: "answered", answer: "No remedy." })])
    expect(gateOf(directory)).toMatchObject({ status: "pass", delivery: "execution_ready" })
  }, 120_000)
})
