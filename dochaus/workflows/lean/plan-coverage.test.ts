import { afterEach, describe, expect, test } from "bun:test"
import { existsSync, readFileSync, rmSync } from "node:fs"
import { PLAN_COVERAGE_REVISIONS, runDraftStage } from "./draft-sections"
import { readLeanEvents } from "./events"
import { LEAN_COVERAGE_SCHEMA, coveragePlanHash, planCoverageCurrent, readLeanCoverage, type CoverageAct } from "./package-coverage"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { LEAN_MAX_RETRIES, disclosableItems, outlineBasis } from "./stages"
import { PACKAGE_SCOPE_QUESTION } from "./package-scope"
import { leanFile, readLeanState, readQuestions, updateLeanStage, writeLeanText, writeQuestions, type LeanStageState } from "./state"
import { AGENT, cleanupMatters, matter, scriptedModel } from "./test-support"

// The outline step's coverage check: a coverage job in its own session reads the
// outlined plan, its issues go back to the outline session for a bounded number
// of revisions, and coverage.json records the result before plan.json is saved.
// Synthetic parties and acts only.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare the papers so that the Orchard Committee approves the revised schedule. Kestrel Works then releases the deposit."
const APPROVE = "the Orchard Committee approves the revised schedule"
const MISPLACED = 'must be taken by "the Orchard Committee", but d1 "Release letter" is executed by "Kestrel Works"'
const REVISE = "Your reply did not pass the draft checks:"

// One document executed by the wrong party, then the same with a companion
// document executed by the party that takes the act.
const ONE = "Acting for: not stated\n# Release letter\nExecuted by: Kestrel Works\n## Terms\nThe release."
const TWO = "Acting for: not stated\n# Committee resolution\nExecuted by: the Orchard Committee\n## Terms\nThe approval.\n" +
  "# Release letter\nExecuted by: Kestrel Works\n## Terms\nThe release."

const act: CoverageAct = { act: "approve the revised schedule", source: "request", quote: APPROVE, parties: ["the Orchard Committee"], document: "d1" }
const ACTS = JSON.stringify({ acts: [act] })

/** The outline agent: `first` answers the outline prompt, `revised` any
 * revision request, and each part gets one plain paragraph. */
const outliner = (first: string, revised = first) => (text: string) =>
  /^Draft part \d+ of \d+ of "/m.test(text) ? "The terms are recorded." : text.startsWith(REVISE) ? revised : first

function prepared() {
  const directory = matter(SOURCE_TEXT)
  writeLeanText(directory, "request.md", REQUEST)
  return { directory, state: readLeanState(directory)! }
}

const run = (directory: string, state: ReturnType<typeof prepared>["state"], model: ReturnType<typeof scriptedModel>) =>
  runDraftStage(directory, state, { sessions: model.port }, () => {}, "d1", () => {})

const coverageSessions = (model: ReturnType<typeof scriptedModel>) => model.sessions.filter(session => session.agent === AGENT.coverage)
const revisions = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))

describe("plan coverage at the outline step", () => {
  test("a clean coverage reply is recorded as covered and the plan is saved", async () => {
    const { directory, state } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => ACTS })
    expect((await run(directory, state, model)).status).toBe("done")

    const plan = readLeanPlan(directory)!
    expect(plan.documents.map(row => [row.id, row.title, row.executed_by])).toEqual([
      ["d1", "Committee resolution", ["the Orchard Committee"]], ["d2", "Release letter", ["Kestrel Works"]],
    ])
    expect(readLeanCoverage(directory)).toEqual({
      schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "covered", acts: [act], issues: [],
    })
    expect(planCoverageCurrent(directory)).toBe(true)
    // The coverage job runs in a fresh child session of the run's owner.
    expect(coverageSessions(model)).toEqual([expect.objectContaining({ parentID: "ses_owner", title: expect.stringContaining("draft: coverage") })])
    expect(revisions(model)).toEqual([])
  }, 60_000)

  test("an act whose document the instructions leave out is recorded as covered, with no revision or question, after an outline told to plan no such document", async () => {
    const { directory, state } = prepared()
    const EXCLUDE = "Prepare only documents Kestrel Works executes"
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    const excluded: CoverageAct = { ...act, document: null, excluded_by: EXCLUDE }
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => JSON.stringify({ acts: [excluded] }) })
    expect((await run(directory, state, model)).status).toBe("done")

    const plan = readLeanPlan(directory)!
    expect(plan.documents.map(row => [row.id, row.title, row.executed_by])).toEqual([["d1", "Release letter", ["Kestrel Works"]]])
    expect(readLeanCoverage(directory)).toEqual({
      schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "covered", acts: [excluded], issues: [],
    })
    expect(revisions(model)).toEqual([])
    expect(coverageSessions(model)).toHaveLength(1)
    expect(readQuestions(directory)).toEqual([])
    expect(model.prompts.find(prompt => prompt.agent === AGENT.draft)!.text).toContain("- Under each title write \"Executed by: <party>; <party>\"")
    expect(model.prompts.find(prompt => prompt.agent === AGENT.draft)!.text).toContain("Leave the line out when neither says who executes it. " +
      "Plan no document the instructions leave out of this package. For each legal act the instructions direct, unless they leave out the document of the party who can take it, " +
      "plan the act in a document that party executes, adding a companion document where none is planned.\n")
  }, 60_000)

  test("an act whose document the instructions leave out sends a plan with a document its parties execute back to the outline, and the revised plan without it is covered with no question", async () => {
    const { directory, state } = prepared()
    const EXCLUDE = "Prepare only documents Kestrel Works executes"
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    const excluded: CoverageAct = { ...act, document: null, excluded_by: EXCLUDE }
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO, ONE), [AGENT.coverage]: () => JSON.stringify({ acts: [excluded] }) })
    expect((await run(directory, state, model)).status).toBe("done")

    expect(revisions(model)).toEqual([expect.objectContaining({ text: expect.stringContaining('but d1 "Committee resolution", executed by "the Orchard Committee", is planned and carries no instructed act. ' +
      "If the instructions leave that document out, do not plan it.\n") })])
    expect(readLeanPlan(directory)!.documents.map(row => row.title)).toEqual(["Release letter"])
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", acts: [excluded], issues: [] })
    expect(readQuestions(directory)).toEqual([])
  }, 60_000)

  test("when the revisions keep a document the parties of a left-out act execute, the plan is saved as uncovered and the step asks no question", async () => {
    const { directory, state } = prepared()
    const EXCLUDE = "Prepare only documents Kestrel Works executes"
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    const excluded: CoverageAct = { ...act, document: null, excluded_by: EXCLUDE }
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => JSON.stringify({ acts: [excluded] }) })
    expect((await run(directory, state, model)).status).toBe("done")

    expect(revisions(model)).toHaveLength(PLAN_COVERAGE_REVISIONS)
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", acts: [excluded], issues: [expect.stringContaining('d1 "Committee resolution", executed by "the Orchard Committee", is planned')] })
    expect(readLeanPlan(directory)!.documents.map(row => row.title)).toEqual(["Committee resolution", "Release letter"])
    expect(readQuestions(directory)).toEqual([])
    // The lawyer reads it once, in the act's own item, never as a second item.
    expect(disclosableItems(directory).filter(item => item.stage === "finalize").map(item => item.detail)).toEqual([
      expect.stringMatching(/as leaving out the document for this instructed act, .* The lawyer should also confirm whether d1 "Committee resolution", which is executed by "the Orchard Committee" and carries no instructed act, should be in this package\.$/),
    ])
  }, 60_000)

  test("an act whose document the instructions leave out raises no revision while every planned document its parties execute carries an instructed act", async () => {
    const { directory, state } = prepared()
    const EXCLUDE = "We are not preparing the committee resolution"
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    const excluded: CoverageAct = { ...act, document: null, excluded_by: EXCLUDE }
    const released: CoverageAct = { act: "release the deposit", source: "request", quote: "Kestrel Works then releases the deposit", parties: ["Kestrel Works"], document: "d1" }
    const SHARED = "Acting for: not stated\n# Release letter\nExecuted by: Kestrel Works; the Orchard Committee\n## Terms\nThe release."
    const model = scriptedModel({ [AGENT.draft]: outliner(SHARED), [AGENT.coverage]: () => JSON.stringify({ acts: [released, excluded] }) })
    expect((await run(directory, state, model)).status).toBe("done")

    expect(revisions(model)).toEqual([])
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", acts: [released, excluded], issues: [] })
    expect(readQuestions(directory)).toEqual([])
  }, 60_000)

  test("a coverage issue goes back to the outline session, whose revised plan adds the companion document", async () => {
    const { directory, state } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE, TWO), [AGENT.coverage]: () => ACTS })
    expect((await run(directory, state, model)).status).toBe("done")

    const outline = model.prompts.find(prompt => prompt.agent === AGENT.draft)!
    expect(revisions(model)).toEqual([expect.objectContaining({ sessionID: outline.sessionID, text: expect.stringContaining(MISPLACED) })])
    // Each coverage job has its own session, apart from the outline's.
    expect(coverageSessions(model)).toHaveLength(2)
    const covered = model.prompts.filter(prompt => prompt.agent === AGENT.coverage).map(prompt => prompt.sessionID)
    expect(new Set([outline.sessionID, ...covered]).size).toBe(3)
    expect(readLeanPlan(directory)!.documents.map(row => row.title)).toEqual(["Committee resolution", "Release letter"])
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", acts: [act], issues: [] })
    expect(planCoverageCurrent(directory)).toBe(true)
  }, 60_000)

  test("a coverage reply that never passes its check is recorded as unchecked and the plan is still saved", async () => {
    const { directory, state } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => "The plan looks complete." })
    expect((await run(directory, state, model)).status).toBe("done")

    expect(model.count(AGENT.coverage)).toBe(LEAN_MAX_RETRIES + 1)
    expect(coverageSessions(model)).toHaveLength(1)
    const coverage = readLeanCoverage(directory)!
    expect(coverage).toMatchObject({ status: "unchecked", acts: [] })
    expect(coverage.issues.length).toBeGreaterThan(0)
    expect(revisions(model)).toEqual([])
    expect(readLeanPlan(directory)!.documents.map(row => row.title)).toEqual(["Release letter"])
  }, 60_000)

  test("words that never pass as whole sentences of the lawyer's are dropped on the coverage job's last attempt: the act goes back to the outline and to the lawyer, never unchecked", async () => {
    const { directory, state } = prepared()
    const EXCLUDE = "Prepare only documents Kestrel Works executes"
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => JSON.stringify({ acts: [{ ...act, document: null, excluded_by: "Prepare only documents" }] }) })
    const result = await run(directory, state, model)

    // Each coverage job is retried to its limit, shown the sentence that holds the words.
    expect(model.count(AGENT.coverage)).toBe((LEAN_MAX_RETRIES + 1) * (PLAN_COVERAGE_REVISIONS + 1))
    expect(model.prompts.filter(prompt => prompt.agent === AGENT.coverage && prompt.text.startsWith(REVISE))).toHaveLength(LEAN_MAX_RETRIES * (PLAN_COVERAGE_REVISIONS + 1))
    expect(model.prompts.find(prompt => prompt.agent === AGENT.coverage && prompt.text.startsWith(REVISE))!.text).toContain(`the whole sentence containing them is "${EXCLUDE}."`)
    const uncovered = 'The instructions direct "approve the revised schedule"'
    expect(revisions(model)).toHaveLength(PLAN_COVERAGE_REVISIONS)
    for (const revision of revisions(model)) expect(revision.text).toContain(uncovered)
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", issues: [expect.stringContaining(uncovered)] })
    expect(readLeanCoverage(directory)!.acts).toEqual([{ ...act, document: null }])
    expect(result).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ stage: "draft", status: "open", question: expect.stringContaining(`"${APPROVE}"`) })] })
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
  }, 60_000)

  // Without an answered package-scope question that showed the act the step
  // stops instead (plan-scope-question.test.ts); once one that showed it is
  // answered it saves the plan.
  test("once a package-scope question that showed the act is answered, issues left after the allowed revisions are recorded as uncovered and the last plan is saved", async () => {
    const { directory, state } = prepared()
    writeQuestions(directory, [{ id: `${PACKAGE_SCOPE_QUESTION}earlier`, stage: "draft", question: `The document plan leaves this instructed act without a document executed by the parties who must take it: "approve the revised schedule", directed by "${APPROVE}" [source: request]; to be taken by "the Orchard Committee"; planned in d1 "Release letter", executed by "Kestrel Works".`, context: "", status: "answered", answer: "The release letter only." }])
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS })
    expect((await run(directory, state, model)).status).toBe("done")

    expect(revisions(model)).toHaveLength(PLAN_COVERAGE_REVISIONS)
    expect(coverageSessions(model)).toHaveLength(PLAN_COVERAGE_REVISIONS + 1)
    expect(readLeanCoverage(directory)).toMatchObject({ status: "uncovered", acts: [act], issues: [expect.stringContaining(MISPLACED)] })
    expect(readLeanPlan(directory)!.documents.map(row => row.title)).toEqual(["Release letter"])

    // An answered question that showed no act settles none: the step stops on one about it.
    const other = prepared()
    writeQuestions(other.directory, [{ id: `${PACKAGE_SCOPE_QUESTION}earlier`, stage: "draft", question: "Which document carries each instructed act?", context: "", status: "answered", answer: "The release letter only." }])
    const blocked = await run(other.directory, other.state, scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS }))
    expect(blocked).toMatchObject({ status: "blocked", questions: [expect.objectContaining({ stage: "draft", status: "open", question: expect.stringContaining(`"${APPROVE}"`) })] })
    expect(existsSync(leanFile(other.directory, "plan.json"))).toBe(false)
  }, 60_000)

  test("an outlined plan is current only while its coverage record checked that plan; a hand-written plan needs none", () => {
    const { directory } = prepared()
    const parts = [{ heading: "Terms", scope: "" }]
    const plan = writeLeanPlan(directory, [{ title: "Release letter", purpose: "", parts }], outlineBasis(directory))
    expect(planCoverageCurrent(directory)).toBe(false)
    const record = { schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "covered", acts: [], issues: [] }
    writeLeanText(directory, "coverage.json", JSON.stringify(record))
    expect(planCoverageCurrent(directory)).toBe(true)
    writeLeanText(directory, "coverage.json", JSON.stringify({ ...record, plan: "0".repeat(64) }))
    expect(planCoverageCurrent(directory)).toBe(false)
    writeLeanText(directory, "coverage.json", "{")
    expect(readLeanCoverage(directory)).toBeUndefined()
    writeLeanPlan(directory, [{ title: "Release letter", purpose: "", parts }])
    expect(planCoverageCurrent(directory)).toBe(true)
  })

  test("a plan whose basis is current but whose coverage record is stale is followed as it stands and only rechecked", async () => {
    const { directory, state } = prepared()
    writeLeanPlan(directory, [{ title: "Committee resolution", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["the Orchard Committee"] }], outlineBasis(directory))
    const saved = readFileSync(leanFile(directory, "plan.json"), "utf8")
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => ACTS })
    expect((await run(directory, state, model)).status).toBe("done")

    expect(model.prompts.some(prompt => prompt.text.includes("Reply now with the outline only"))).toBe(false)
    expect(readFileSync(leanFile(directory, "plan.json"), "utf8")).toBe(saved)
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", acts: [act] })
    expect(planCoverageCurrent(directory)).toBe(true)
  }, 60_000)

  test("a refresh that keeps the existing draft creates no draft session, only the coverage session", async () => {
    const { directory, state } = prepared()
    // The step records its accepted draft's key on its row, as the runner keeps it.
    const progress = (patch: Partial<LeanStageState>) => { updateLeanStage(directory, readLeanState(directory)!, "d1.draft", patch) }
    expect((await runDraftStage(directory, state, { sessions: scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => ACTS }).port }, progress, "d1", () => {})).status).toBe("done")
    rmSync(leanFile(directory, "coverage.json"))
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => ACTS })
    const kept = await runDraftStage(directory, readLeanState(directory)!, { sessions: model.port }, progress, "d1", () => {})

    expect(kept.status).toBe("done")
    expect(model.sessions.filter(session => session.agent === AGENT.draft)).toEqual([])
    expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft)).toEqual([])
    const covered = model.prompts.find(prompt => prompt.agent === AGENT.coverage)!.sessionID
    expect(readLeanState(directory)!.stages["d1.draft"]).toMatchObject({ session_id: covered, session_ids: [covered] })
  }, 60_000)

  test("a refresh stopped by the package-scope question creates no draft session", async () => {
    const { directory, state } = prepared()
    writeLeanPlan(directory, [{ title: "Release letter", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["Kestrel Works"] }], outlineBasis(directory))
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE), [AGENT.coverage]: () => ACTS })
    expect((await run(directory, state, model)).status).toBe("blocked")
    expect(model.sessions.filter(session => session.agent === AGENT.draft)).toEqual([])
    expect(coverageSessions(model)).toHaveLength(1)
  }, 60_000)

  test("a pause thrown from a recheck's coverage job propagates and leaves coverage.json and plan.json as they were", async () => {
    const { directory, state } = prepared()
    await run(directory, state, scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => ACTS }))
    // A stale coverage record makes the step recheck the saved plan.
    writeLeanText(directory, "coverage.json", JSON.stringify({ ...readLeanCoverage(directory), plan: "0".repeat(64) }))
    const before = ["coverage.json", "plan.json"].map(file => readFileSync(leanFile(directory, file), "utf8"))
    const stop = new Error("pause requested")
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => { throw stop } })
    await expect(run(directory, state, model)).rejects.toBe(stop)
    expect(coverageSessions(model)).toHaveLength(1)
    expect(["coverage.json", "plan.json"].map(file => readFileSync(leanFile(directory, file), "utf8"))).toEqual(before)
  }, 60_000)

  test("a restart thrown from the coverage job after a revision propagates and writes neither file", async () => {
    const { directory, state } = prepared()
    const stop = new Error("restart requested")
    const model = scriptedModel({
      [AGENT.draft]: outliner(ONE, TWO),
      [AGENT.coverage]: (_, index) => { if (index === 1) throw stop; return ACTS },
    })
    await expect(run(directory, state, model)).rejects.toBe(stop)
    expect(revisions(model)).toHaveLength(1)
    expect(existsSync(leanFile(directory, "coverage.json"))).toBe(false)
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
  }, 60_000)

  test("a pause before any plan exists writes neither file", async () => {
    const { directory, state } = prepared()
    const stop = new Error("pause requested")
    const model = scriptedModel({ [AGENT.draft]: outliner(TWO), [AGENT.coverage]: () => { throw stop } })
    await expect(run(directory, state, model)).rejects.toBe(stop)
    expect(existsSync(leanFile(directory, "coverage.json"))).toBe(false)
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
  })
})

// Observability of the coverage step (record only): the coverage issues that
// drive each outline revision and the recorded outcome reach events.jsonl, and
// the coverage jobs' sessions are reported on the draft step beside the outline's.
describe("plan coverage observability", () => {
  test("an uncovered-then-revised outline logs the driving issues on its revision event and the outcome once", async () => {
    const { directory, state } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE, TWO), [AGENT.coverage]: () => ACTS })
    expect((await run(directory, state, model)).status).toBe("done")

    const events = readLeanEvents(directory).filter(event => event.stage === "draft" && event.document === "d1")
    expect(events.filter(event => event.step?.startsWith("outline revision") && event.event === "model_call")).toEqual([
      expect.objectContaining({ step: `outline revision 1/${PLAN_COVERAGE_REVISIONS}`, issues: [expect.stringContaining(MISPLACED)] }),
    ])
    expect(events.filter(event => event.event === "coverage")).toEqual([
      expect.objectContaining({ step: "coverage", outcome: "covered", issues: [] }),
    ])
  }, 60_000)

  test("the coverage jobs' sessions, then the files job's, appear in the draft step's session_ids after the outline's", async () => {
    const { directory, state } = prepared()
    const model = scriptedModel({ [AGENT.draft]: outliner(ONE, TWO), [AGENT.coverage]: () => ACTS })
    const progress = (patch: Partial<LeanStageState>) => { updateLeanStage(directory, readLeanState(directory)!, "d1.draft", patch) }
    expect((await runDraftStage(directory, state, { sessions: model.port }, progress, "d1", () => {})).status).toBe("done")

    const outline = model.prompts.find(prompt => prompt.agent === AGENT.draft)!.sessionID
    const covered = [...new Set(model.prompts.filter(prompt => prompt.agent === AGENT.coverage).map(prompt => prompt.sessionID))]
    expect(covered).toHaveLength(2)
    // The revised outline plans two documents, so the files job reads it too.
    const filed = [...new Set(model.prompts.filter(prompt => prompt.agent === AGENT.files).map(prompt => prompt.sessionID))]
    expect(filed).toHaveLength(1)
    expect(readLeanState(directory)!.stages["d1.draft"]).toMatchObject({ session_id: outline, session_ids: [outline, ...covered, ...filed] })
  }, 60_000)
})
