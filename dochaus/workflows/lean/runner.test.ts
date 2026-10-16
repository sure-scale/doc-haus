import { afterEach, describe, expect, test } from "bun:test"
import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs"
import { runApply } from "./apply-stage"
import { stagePrompt } from "./prompts"
import { readParagraphs } from "./checks"
import type { GateResult } from "./gate"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { acceptStageReply, LEAN_CHILD_PERMISSIONS } from "./stages"
import { leanFile, readLeanJson, readLeanState, readOpenItems, readQuestions } from "./state"
import { AGENT, DRAFT, SOURCE, SUPPORT, changes, cleanupMatters, drafter, happyPath, matter, noChanges, pidOf, sameSession, scriptedModel, type Reply } from "./test-support"

// Real coordinator, checkers, renderer, pointer applier and comparison. Only the
// model is scripted: each agent's reply is a function of the prompt it received.

const NO_HEADING = "Here is the requested document."

afterEach(cleanupMatters)

describe("lean drafting runner", () => {
  test("a failed check is retried in the same session with the checker's issues; the fourth failure stops the run", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.draft]: () => NO_HEADING })
    const state = await runLean(directory, { sessions: model.port })

    expect(model.count(AGENT.draft)).toBe(4)
    expect(new Set(model.prompts.filter(row => row.agent === AGENT.draft).map(prompt => prompt.sessionID)).size).toBe(1)
    expect(model.prompts.filter(row => row.agent === AGENT.draft)[0]!.text).toContain("<output_contract>")
    model.prompts.filter(row => row.agent === AGENT.draft).slice(1).forEach(prompt => {
      expect(prompt.text).toContain("Your reply did not pass the draft checks:")
      expect(prompt.text).toContain("has no \"# \" title line")
    })
    expect(state.status).toBe("failed")
    expect(state.stop_reason).toStartWith("d1.draft: outline: checker failed after 3 retries: ")
    expect(state.stop_reason).toContain("has no \"# \" title line")
    expect(state.stages["d1.draft"]).toMatchObject({ status: "failed", attempts: 4, calls: 4, session_id: "ses_child6" })
    expect(state.stages["d1.review"]?.status).toBe("pending")
    expect(model.sessions.find(row => row.agent === AGENT.draft)).toMatchObject({ parentID: "ses_owner", agent: AGENT.draft, permissions: LEAN_CHILD_PERMISSIONS })
    expect(existsSync(leanFile(directory, "d1/draft.md"))).toBe(false)
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
  }, 120_000)

  test("a retry that passes completes the run through apply, finalize and the gate", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: (prompt, index) => index === 0 ? NO_HEADING : drafter(DRAFT)(prompt, index) })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    // One failed outline, then the outline and one call per part.
    expect(model.count(AGENT.draft)).toBe(4)
    expect(model.prompts.filter(row => row.agent === AGENT.draft)[1]!.text).toContain("has no \"# \" title line")
    expect(model.count(AGENT.merge)).toBe(0)
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ id: "a1", mode: "code" }], unapplied: [] })
    expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text))
      .toContain("The supplier shall deliver the goods within thirty days of the order date.")
    ;["d1/body.docx", "d1/body-applied.docx", "drafting-notes.docx", "d1/redline.docx", "drafting-notes-coverage.md"].forEach(name =>
      expect(existsSync(leanFile(directory, name))).toBe(true))
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "execution_ready", disclosed: [] })
    expect(Object.keys(state.stages)).toEqual(["discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "checklist-check", "package-review", "gate"])
    expect(Object.values(state.stages).every(stage => stage?.status === "done" && stage.input_sha)).toBe(true)
  }, 120_000)

  test("resume skips every stage whose inputs are unchanged and re-runs only downstream of an edit", async () => {
    const directory = matter()
    const first = await runLean(directory, { sessions: scriptedModel(happyPath).port })
    expect(first.status).toBe("completed")

    const idle = scriptedModel({})
    const resumed = await runLean(directory, { sessions: idle.port })
    expect(idle.prompts.length + idle.sessions.length).toBe(0)
    expect(resumed.status).toBe("completed")
    expect(resumed.stages["d1.apply"]?.finished_at).toBe(first.stages["d1.apply"]?.finished_at)

    // This saved planner fixture predates structured findings. Its review history
    // feeds the notes directly: editing review re-runs changes and finalize,
    // while identical amendments leave apply untouched.
    writeFileSync(leanFile(directory, "d1/review.md"), `${readFileSync(leanFile(directory, "d1/review.md"), "utf8")}\nReviewed again by the lawyer.\n`)
    const reviewed = scriptedModel(happyPath)
    const afterReview = await runLean(directory, { sessions: reviewed.port })
    expect(afterReview.status).toBe("completed")
    expect(reviewed.counts()).toEqual([0, 0, 1, 1])
    expect(reviewed.prompts.find(row => row.agent === AGENT.finalize)?.text).toContain("Reviewed again by the lawyer.")
    expect(afterReview.stages.finalize?.input_sha).not.toBe(first.stages.finalize?.input_sha)
    expect(reviewed.count(AGENT["package-review"])).toBe(1)
    expect(afterReview.stages["d1.apply"]?.finished_at).toBe(first.stages["d1.apply"]?.finished_at)

    // An edited change list re-applies from the drafted body and re-runs finalize.
    const list = JSON.parse(readFileSync(leanFile(directory, "d1/changes.json"), "utf8"))
    list.amendments[0].reason = "Edited by the lawyer."
    writeFileSync(leanFile(directory, "d1/changes.json"), JSON.stringify(list, null, 2))
    const edited = scriptedModel(happyPath)
    const rerun = await runLean(directory, { sessions: edited.port })
    expect(rerun.status).toBe("completed")
    expect(edited.counts()).toEqual([0, 0, 0, 1])
    expect(rerun.stages["d1.draft"]?.finished_at).toBe(first.stages["d1.draft"]?.finished_at)
    expect(rerun.stages["d1.apply"]?.finished_at).not.toBe(first.stages["d1.apply"]?.finished_at)
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ id: "a1", mode: "code" }], unapplied: [] })
  }, 120_000)

  test("a review edit with unchanged adjudicated dispositions does not re-run final notes", async () => {
    const directory = matter()
    const replies = { ...happyPath,
      [AGENT["change-plan"]]: () => JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [], findings: [] }),
    }
    const first = await runLean(directory, { sessions: scriptedModel(replies).port })
    expect(first.status).toBe("completed")
    writeFileSync(leanFile(directory, "d1/review.md"), `${readFileSync(leanFile(directory, "d1/review.md"), "utf8")}\nReviewed again by the lawyer.\n`)
    const reviewed = scriptedModel(replies)
    const afterReview = await runLean(directory, { sessions: reviewed.port })
    expect(afterReview.status).toBe("completed")
    expect(reviewed.counts()).toEqual([0, 0, 1, 0])
    expect(afterReview.stages["d1.apply"]?.finished_at).toBe(first.stages["d1.apply"]?.finished_at)
    expect(afterReview.stages.finalize?.input_sha).toBe(first.stages.finalize?.input_sha)
    expect(afterReview.stages.finalize?.finished_at).toBe(first.stages.finalize?.finished_at)
    const idle = scriptedModel({})
    expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
    expect(idle.prompts.length + idle.sessions.length).toBe(0)
  }, 120_000)

  test("a draft misquote goes to the changes stage to repair; left in the body it blocks, and Continue starts one repair round", async () => {
    const directory = matter()
    const draft = `# Supply note\n\n## Delivery\n\nThe supplier promised delivery "within ten days" [source: ${SOURCE}].\n\n## Payment\n\nPayment is due on delivery.\n`
    const repair: Reply = prompt => JSON.stringify({
      amendments: [{ pid: pidOf(prompt, "The supplier promised delivery"), updated: "The supplier promised delivery by [open item: confirm the delivery period].",
        reason: "The quoted period is not in the supply terms.", sources: SUPPORT }],
      questions: [],
    })
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(draft), [AGENT.review]: () => "No corrections are needed.", [AGENT.changes]: noChanges })
    const state = await runLean(directory, { sessions: model.port })

    const draftItem = readOpenItems(directory).find(item => item.stage === "draft")!
    expect(draftItem).toMatchObject({ kind: "quote_miss", ref: { source_id: SOURCE, quote: "within ten days" } })
    const redlinerPrompt = model.prompts.find(prompt => prompt.agent === AGENT.changes)!.text
    expect(redlinerPrompt).toContain("<required_repairs>")
    expect(redlinerPrompt).toContain(draftItem.id)
    expect(model.count(AGENT.draft)).toBe(3)
    const gate = readLeanJson(directory, "gate.json") as GateResult
    expect(state.status).toBe("blocked")
    expect(gate).toMatchObject({ status: "blocked", delivery: null })
    expect(gate.blocking).toEqual([expect.objectContaining({ stage: "apply", kind: "quote_miss", resolution: "named in the notes, but its paragraph has no visible open-item marker" })])
    expect(gate.blocking[0]!.ref.pid).toBe(pidOf(redlinerPrompt, "The supplier promised delivery"))
    expect(readFileSync(leanFile(directory, "drafting-notes.md"), "utf8")).toContain(gate.blocking[0]!.id)
    expect(state.stop_reason).toBe(`gate: blocked by open items ${gate.blocking[0]!.id}`)

    // Without a request, a resume has nothing to change and stays blocked.
    const idle = scriptedModel({})
    expect((await runLean(directory, { sessions: idle.port })).status).toBe("blocked")
    expect(idle.prompts).toHaveLength(0)

    // Continue on the blocked gate re-runs the changes stage with the blocking item.
    const request = enqueueLeanRequest(directory, { kind: "continue" })
    const repaired = scriptedModel({ ...happyPath, [AGENT.changes]: repair })
    const resumed = await runLean(directory, { sessions: repaired.port })
    expect(repaired.counts()).toEqual([0, 0, 1, 1])
    expect(repaired.prompts.find(prompt => prompt.agent === AGENT.changes)!.text).toContain(gate.blocking[0]!.id)
    expect(readLeanJson(directory, "repairs.json")).toMatchObject({ round: 1, items: [{ id: gate.blocking[0]!.id }] })
    expect(readLeanJson(directory, `requests/done/${request.id}.json`)).toMatchObject({ kind: "continue", outcome: { status: "repair_round" } })
    expect(resumed.status).toBe("completed")
    expect(readLeanJson(directory, "gate.json")).toMatchObject({
      status: "pass", delivery: "review_draft", disclosed: [expect.objectContaining({ stage: "apply", kind: "placeholder" })],
    })
  }, 120_000)

  test("legacy overlapping amendments merge through an applier child session; an engine failure stops apply", async () => {
    const twoEdits: Reply = prompt => {
      const pid = pidOf(prompt, "The supplier shall deliver the goods.")
      const source = [{ source_id: SOURCE, quote: "within thirty days of the order date" }]
      return JSON.stringify({
        amendments: [
          { pid, updated: "The supplier shall deliver the goods within thirty days of the order date.", reason: "Delivery period.", sources: source },
          { pid, updated: "The supplier shall promptly deliver the goods.", reason: "Delivery promptness.", sources: source },
        ],
        questions: [],
      })
    }
    const directory = matter()
    const model = scriptedModel({
      ...happyPath,
      [AGENT.changes]: twoEdits,
      [AGENT.merge]: () => "<paragraph>The supplier shall promptly deliver the goods within thirty days of the order date.</paragraph>",
    })
    expect((await acceptStageReply(directory, "d1.draft", DRAFT)).issues).toEqual([])
    expect((await acceptStageReply(directory, "d1.changes", twoEdits(stagePrompt(directory, "d1.changes"), 0))).issues).toEqual([])
    const state = await runApply(directory, readLeanState(directory)!, { sessions: model.port }, "d1")

    expect(state.status).toBe("done")
    expect(readLeanJson(directory, "d1/changes.json")).toMatchObject({ shared_paragraphs: [{ amendment_ids: ["a1", "a2"] }] })
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ id: "a1", mode: "code" }, { id: "a2", mode: "merged" }] })
    expect(model.prompts.find(prompt => prompt.agent === AGENT.merge)!.text).toStartWith("This is a lean drafting run.")
    expect(model.sessions.find(session => session.agent === AGENT.merge)).toMatchObject({ parentID: "ses_owner", permissions: LEAN_CHILD_PERMISSIONS })
    expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text))
      .toContain("The supplier shall promptly deliver the goods within thirty days of the order date.")

    const broken = matter()
    const failing = scriptedModel({ ...happyPath, [AGENT.changes]: twoEdits, [AGENT.merge]: () => { throw new Error("engine unavailable") } })
    expect((await acceptStageReply(broken, "d1.draft", DRAFT)).issues).toEqual([])
    expect((await acceptStageReply(broken, "d1.changes", twoEdits(stagePrompt(broken, "d1.changes"), 0))).issues).toEqual([])
    const stopped = await runApply(broken, readLeanState(broken)!, { sessions: failing.port }, "d1")
    expect(stopped.status).toBe("failed")
    expect(stopped).toMatchObject({ reason: "d1.apply: stopped by a runtime error during merge: engine unavailable" })
    expect(existsSync(leanFile(broken, "d1/apply.json"))).toBe(false)
  }, 120_000)

  test("a question is never answered by the run; an answer repeats drafting and review and is not asked again", async () => {
    const directory = matter()
    const asking = sameSession(prompt => JSON.stringify({ ...JSON.parse(changes(prompt, 0)), questions: [{ question: "Should late delivery carry a remedy?", context: "The terms are silent." }] }))
    const model = scriptedModel({ ...happyPath, [AGENT.changes]: asking })
    const blocked = await runLean(directory, { sessions: model.port })

    const [question] = readQuestions(directory)
    expect(question).toMatchObject({ stage: "changes", document: "d1", status: "open" })
    expect(question!.answer).toBeUndefined()
    expect(blocked.status).toBe("blocked")
    expect(blocked.stop_reason).toContain("Waiting for your answers")
    expect(readLeanJson(directory, "gate.json")).toBeUndefined()
    expect(blocked.stages["d1.apply"]?.status).toBe("pending")

    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: "No remedy; leave the terms silent." })
    const answered = scriptedModel({ ...happyPath, [AGENT.changes]: asking })
    const resumed = await runLean(directory, { sessions: answered.port })
    expect(answered.counts()).toEqual([3, 1, 1, 1])
    expect(answered.prompts.find(prompt => prompt.agent === AGENT.changes)!.text).toContain("No remedy; leave the terms silent.")
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: question!.id, status: "answered" })])
    expect(resumed.status).toBe("completed")
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "execution_ready" })
    expect(readLeanState(directory)!.stop_reason).toBeNull()
    expect(readdirSync(leanFile(directory, "requests")).filter(name => name.endsWith(".json"))).toEqual([])
  }, 120_000)
})
