import { afterEach, describe, expect, test } from "bun:test"
import { NOT_STATED, parsePlan, planActingFor, readLeanPlan, writeLeanPlan, type StepID } from "./plan"
import { stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { hashStageInputs, writeLeanText } from "./state"
import { AGENT, DRAFT, cleanupMatters, drafter, happyPath, matter, scriptedModel, type Reply } from "./test-support"

// Whom the lawyer acts for: read once by the outline, kept in the plan, shown to
// every model stage and keyed into their inputs. Real coordinator; only the
// model is scripted.

afterEach(cleanupMatters)

const NEUTRAL = "The request does not say whom the lawyer acts for."
const SIDED_REVIEW = "Assess each provision as the reviewing lawyer for the party the lawyer acts for"
const NEUTRAL_REVIEW = "Assess each provision as a reviewing lawyer who acts for neither party"
const actingFor = (party: string): Reply => {
  const draft = drafter(DRAFT)
  return (prompt, index) => draft(prompt, index).replace("Acting for: not stated", `Acting for: ${party}`)
}
const promptsOf = (model: ReturnType<typeof scriptedModel>, agent: string) => model.prompts.filter(prompt => prompt.agent === agent).map(prompt => prompt.text)

describe("whom the lawyer acts for", () => {
  test("when the request does not say, the plan records it and every stage is told to take no side", async () => {
    const directory = matter()
    const model = scriptedModel(happyPath)
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(readLeanPlan(directory)?.acting_for).toEqual(NOT_STATED)
    const [outline, firstPart] = promptsOf(model, AGENT.draft)
    expect(outline).toContain('"Acting for: not stated"')
    expect(firstPart).toContain(NEUTRAL)
    const review = promptsOf(model, AGENT.review)[0]!
    expect(review).toContain(NEUTRAL)
    expect(review).toContain(NEUTRAL_REVIEW)
    expect(review).not.toContain(SIDED_REVIEW)
    expect(promptsOf(model, AGENT.changes)[0]).toContain(NEUTRAL)
    expect(promptsOf(model, AGENT.finalize)[0]).toContain(NEUTRAL)
  }, 120_000)

  test("a party the outline reads from the request is shown to every stage, and the review assesses for that party", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms. We act for the supplier.")
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: actingFor("the supplier") })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(readLeanPlan(directory)?.acting_for).toEqual({ party: "the supplier", stated: true })
    const block = "<acting_for>\nThe lawyer acts for: the supplier (as the request states).\n</acting_for>"
    expect(promptsOf(model, AGENT.draft)[1]).toContain(block)
    const review = promptsOf(model, AGENT.review)[0]!
    expect(review).toContain(block)
    expect(review).toContain(SIDED_REVIEW)
    expect(review).not.toContain(NEUTRAL_REVIEW)
    expect(promptsOf(model, AGENT.changes)[0]).toContain(block)
    expect(promptsOf(model, AGENT.finalize)[0]).toContain(block)
  }, 120_000)

  test("a party the request does not name is retried, and if the outline keeps it the plan records not stated", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: actingFor("the distributor") })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(readLeanPlan(directory)?.acting_for).toEqual(NOT_STATED)
    const outlines = promptsOf(model, AGENT.draft).filter(text => text.includes('"Acting for:" line names "the distributor"'))
    expect(outlines).toHaveLength(LEAN_MAX_RETRIES)
    expect(promptsOf(model, AGENT.review)[0]).toContain(NEUTRAL_REVIEW)
  }, 120_000)

  test("non-answers on the Acting for line are not a party, and a stated party must be named in the request", () => {
    const outline = (line: string) => `Acting for: ${line}\n# Title\n## One`
    const request = "Review the agreement for ABC (UK) Ltd, the supplier."
    for (const line of ["N/A", "n/a", "NA", "Unknown", "none", "None specified", "[not stated]", "(not stated)", "\"not stated\"",
      "'Not specified.'", "not stated (the request does not say)", "Not stated - the request is silent", "unspecified"])
      expect(parsePlan(outline(line), request)).toEqual({ issues: [], value: expect.objectContaining({ acting_for: NOT_STATED }) })
    for (const [line, party] of [["the supplier", "the supplier"], ["ABC (UK) Ltd", "ABC (UK) Ltd"], ["\u201Cthe Supplier\u201D", "the Supplier"],
      ["the supplier (as the request states)", "the supplier"], ["ABC (UK) Ltd \u2014 the supplier", "ABC (UK) Ltd"]])
      expect(parsePlan(outline(line), request).value?.acting_for).toEqual({ party, stated: true })
    const ungrounded = parsePlan(outline("the buyer"), request)
    expect(ungrounded.issues.join(" ")).toContain('names "the buyer", but the request does not name that party')
    expect(ungrounded.value?.acting_for).toEqual(NOT_STATED)
    expect(parsePlan(outline("the buyer")).value?.acting_for).toEqual({ party: "the buyer", stated: true })
  })

  test("a hand-written plan without the slot is not stated, and the lawyer's edit of it re-keys every model step", async () => {
    const directory = matter()
    expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
    const plan = readLeanPlan(directory)!
    // Written by hand: no basis, no acting_for.
    writeLeanPlan(directory, plan.documents.map(row => ({ title: row.title, purpose: row.purpose, parts: row.parts })))
    expect(planActingFor(directory)).toEqual(NOT_STATED)
    expect(stagePrompt(directory, "d1.review")).toContain(NEUTRAL_REVIEW)

    const steps: StepID[] = ["d1.draft", "d1.review", "d1.changes", "finalize", "d1.apply", "gate"]
    const keys = () => steps.map(step => hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)))
    const before = keys()
    writeLeanText(directory, "plan.json", `${JSON.stringify({ ...readLeanPlan(directory)!, acting_for: { party: "the buyer", stated: true } }, null, 2)}\n`)
    expect(planActingFor(directory)).toEqual({ party: "the buyer", stated: true })
    expect(stagePrompt(directory, "d1.review")).toContain(SIDED_REVIEW)
    // Apply keys on the whole plan.json (it measures template coverage from the
    // plan), so it re-keys too; it makes no model call.
    expect(keys().map((key, index) => key !== before[index])).toEqual([true, true, true, true, true, false])
  }, 120_000)
})
