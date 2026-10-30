import { afterEach, describe, expect, test } from "bun:test"
import { readLeanEvents } from "./events"
import { readLeanCoverage, type CoverageAct } from "./package-coverage"
import { leanSteps, readLeanPlan } from "./plan"
import { runLean } from "./runner"
import { readQuestions, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, NO_FINDINGS, cleanupMatters, draftingNotes, listedActs, matter, noChanges, scriptedModel } from "./test-support"

// An answer that moves an instructed act to a separate document keeps that act
// in the package: the coverage jobs list it, blind to the plan, and place it in
// no planned document with no excluded_by, so the plan that leaves it out is
// not taken as covered. Its
// issue goes back to the outline, whose revision plans the separate document,
// and the run drafts it. Synthetic parties and acts only; only the model is
// scripted.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare the papers so that the Orchard Committee approves the revised schedule. Kestrel Works then releases the deposit."
const APPROVE = "the Orchard Committee approves the revised schedule"
const RELEASE = "Kestrel Works then releases the deposit"
const REVISE = "Your reply did not pass the draft checks:"
const ANSWERED: Question = {
  id: "q-papers", stage: "intake", question: "Which acts should the release letter carry?", context: "Needed to plan the package.", status: "answered",
  answer: "Limit the release letter to the release of the deposit. Move the committee's approval to a separate later document.", answer_kind: "adopted",
}

// The release letter alone, and the same with the separate document the answer moves the approval to.
const ONE = "Acting for: not stated\n# Release letter\nForm: letter\nExecuted by: Kestrel Works\n## Terms\nThe release."
const SEPARATE = `${ONE}\n# Committee resolution\nForm: other\nExecuted by: the Orchard Committee\n## Terms\nThe approval.`

const released: CoverageAct = { act: "release the deposit", source: "request", quote: RELEASE, parties: ["Kestrel Works"], document: "d1" }
const approved: CoverageAct = { act: "approve the revised schedule", source: "request", quote: APPROVE, parties: ["the Orchard Committee"], document: null }

/** The listing job: both acts the instructions direct, whichever document is to carry them. */
const listing = () => JSON.stringify({ acts: [released, approved].map(({ document: _, ...act }) => act) })
/** The placement job: the release in the release letter, and the approval in
 * the planned document the committee executes, or in none, never left out. */
const placement = (text: string) => {
  const planned = JSON.parse(/<planned_documents>\n[^\n]*\n([\s\S]*)\n<\/planned_documents>/.exec(text)![1]!) as { id: string; executed_by: string[] | null }[]
  const committee = planned.find(document => document.executed_by?.includes("the Orchard Committee"))
  return JSON.stringify({ placements: listedActs(text).map(row => ({ act: row.index, document: row.act === released.act ? released.document : committee?.id ?? null })) })
}

const PART = /^Draft part \d+ of \d+ of "([^"]*)"/m
/** The outline agent: the release letter alone, the separate document on revision, and one plain paragraph per part. */
const outliner = (text: string) => PART.test(text) ? "The terms are recorded." : text.startsWith(REVISE) ? SEPARATE : ONE

describe("an act an answer moves to a separate document", () => {
  test("listed with no planned document and no excluded_by, it sends the plan back for that document, which the run plans and drafts", async () => {
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", REQUEST)
    writeQuestions(directory, [ANSWERED])
    const model = scriptedModel({
      [AGENT.draft]: outliner, [AGENT.coverageActs]: listing, [AGENT.coverage]: placement, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges, [AGENT.finalize]: draftingNotes,
    })
    const state = await runLean(directory, { sessions: model.port })
    expect(state.status).toBe("completed")

    // The plan without the separate document is not taken as covered: the act goes back to the outline.
    const revisions = readLeanEvents(directory).filter(event => event.stage === "draft" && event.step?.startsWith("outline revision") && event.event === "model_call")
    expect(revisions).toEqual([expect.objectContaining({ issues: [`The instructions direct "approve the revised schedule" ("${APPROVE}" [source: request]), but no planned document carries it. ` +
      'Plan it in a document executed by "the Orchard Committee", adding one when none is planned.'] })])
    expect(readLeanPlan(directory)!.documents.map(row => [row.id, row.title, row.executed_by])).toEqual([
      ["d1", "Release letter", ["Kestrel Works"]], ["d2", "Committee resolution", ["the Orchard Committee"]],
    ])
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", acts: [released, { ...approved, document: "d2" }], issues: [] })
    // The acts are listed once, by two runs, and each plan is placed afresh.
    expect([model.count(AGENT.coverageActs), model.count(AGENT.coverage)]).toEqual([2, 2])
    // The separate document is drafted, and no package-scope question was needed.
    expect(leanSteps(directory)).toContain("d2.draft")
    expect(state.stages["d2.draft"]?.status).toBe("done")
    expect(model.prompts.filter(prompt => PART.exec(prompt.text)?.[1] === "Committee resolution")).toHaveLength(1)
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: ANSWERED.id, status: "answered" })])
  }, 120_000)
})
