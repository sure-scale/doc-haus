import { afterEach, describe, expect, test } from "bun:test"
import { rmSync } from "node:fs"
import { readParagraphs } from "./checks"
import { readLeanDocuments } from "./deliverables"
import type { GateResult } from "./gate"
import { coverageItems, coveragePlanHash, readLeanCoverage } from "./package-coverage"
import { readLeanPlan } from "./plan"
import { openItem } from "./quotes"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES } from "./stages"
import { leanFile, readLeanJson, readLeanText, readOpenItems, readQuestions, writeLeanText, writeOpenItems } from "./state"
import { AGENT, cleanupMatters, draftingNotes, matter, noChanges, registering, scriptedModel, type Reply } from "./test-support"

// Plan coverage through the whole run (runLean): the outline's coverage job,
// its revision requests to the outline session, and what the later steps
// deliver. The step-level behaviour is in plan-coverage.test.ts, the scope
// question and its answer in plan-scope-question.test.ts, and the derived open
// items in coverage-items.test.ts. Synthetic parties and acts only; only the
// model is scripted, and the coverage job answers from the plan it is shown.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare the papers so that the Orchard Committee approves the revised schedule. Kestrel Works then releases the deposit."
const APPROVE = "the Orchard Committee approves the revised schedule"
const RELEASE = "Kestrel Works then releases the deposit"
const REVISE = "Your reply did not pass the draft checks:"

// Outlines. A part headed "Approval" carries the committee's approval and one
// headed "Release" the release of the deposit.
const RELEASE_LETTER = "# Release letter\nExecuted by: Kestrel Works\n## Release\nThe release."
const RESOLUTION = "# Committee resolution\nExecuted by: the Orchard Committee\n## Approval\nThe approval."
const outline = (...documents: string[]) => ["Acting for: not stated", ...documents].join("\n")

type Instructed = { act: string; quote: string; parties: string[]; heading: string }
const approval: Instructed = { act: "approve the revised schedule", quote: APPROVE, parties: ["the Orchard Committee"], heading: "Approval" }
const release: Instructed = { act: "release the deposit", quote: RELEASE, parties: ["Kestrel Works"], heading: "Release" }

/** The planned documents the coverage prompt shows, as JSON after one line of explanation. */
function plannedIn(prompt: string) {
  const json = /<planned_documents>\n[^\n]*\n([\s\S]*?)\n<\/planned_documents>/.exec(prompt)?.[1]
  if (!json) throw new Error("The coverage prompt shows no planned documents.")
  return JSON.parse(json) as { id: string; title: string; executed_by: string[] | null; parts: { heading: string }[] }[]
}

/** The coverage job: each instructed act, carried by the planned document with a
 * part under its heading, or by none. */
const coverageOf = (...acts: Instructed[]): Reply => prompt => {
  const documents = plannedIn(prompt)
  return JSON.stringify({ acts: acts.map(row => ({
    act: row.act, source: "request", quote: row.quote, parties: row.parties,
    document: documents.find(document => document.parts.some(part => part.heading === row.heading))?.id ?? null,
  })) })
}

/** The outline agent: `first` answers the outline prompt, `revised` any
 * revision request, and each part gets one plain paragraph. */
const outliner = (first: string, revised = first): Reply => text =>
  /^Draft part \d+ of \d+ of "/m.test(text) ? "The terms are recorded." : text.startsWith(REVISE) ? revised : first

/** A whole run whose later steps find nothing to change. */
const replies = (extra: Record<string, Reply>): Record<string, Reply> => ({
  [AGENT.review]: () => "No corrections are needed.",
  [AGENT.changes]: noChanges,
  [AGENT.finalize]: draftingNotes,
  ...extra,
})

function prepared(request = REQUEST) {
  const directory = matter(SOURCE_TEXT)
  writeLeanText(directory, "request.md", request)
  return directory
}

type Model = ReturnType<typeof scriptedModel>
const revisions = (model: Model) => model.prompts.filter(prompt => prompt.agent === AGENT.draft && prompt.text.startsWith(REVISE))
const outlinePrompt = (model: Model) => model.prompts.find(prompt => prompt.agent === AGENT.draft)!
const delivered = (directory: string) => readLeanDocuments(directory)!.documents.map(row => `${row.kind}:${row.document}:${row.title}:${row.deliverable}`)
const gateOf = (directory: string) => readLeanJson(directory, "gate.json") as GateResult & { questions?: unknown }

/** Two documents run to completion; then the lawyer adds a party to d2's executing
 * parties only, so the coverage record is out of date while d1's draft step keeps its inputs. */
async function laterPartiesEdited() {
  const directory = prepared()
  const run = () => scriptedModel(replies({ [AGENT.draft]: outliner(outline(RESOLUTION, RELEASE_LETTER)), [AGENT.coverage]: coverageOf(approval, release) }))
  expect((await runLean(directory, { sessions: run().port })).status).toBe("completed")
  const plan = readLeanPlan(directory)!
  const edited = { ...plan, documents: plan.documents.map(row => row.id === "d2" ? { ...row, executed_by: [...row.executed_by!, "the Heron Trust"] } : row) }
  writeLeanText(directory, "plan.json", `${JSON.stringify(edited, null, 2)}\n`)
  return { directory, run, drafted: readLeanText(directory, "d1/draft.md") }
}

describe("plan coverage through the run", () => {
  test("an act no planned document carries brings a companion document executed by its party; both are drafted and delivered and no question is asked", async () => {
    const directory = prepared()
    const deliverables = registering()
    const model = scriptedModel(replies({
      [AGENT.draft]: outliner(outline(RELEASE_LETTER), outline(RELEASE_LETTER, RESOLUTION)),
      [AGENT.coverage]: coverageOf(approval, release),
    }))
    const state = await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })
    expect(state.status).toBe("completed")

    // One revision, in the outline's session, carrying the one coverage issue.
    expect(revisions(model)).toEqual([expect.objectContaining({ sessionID: outlinePrompt(model).sessionID })])
    const issues = revisions(model)[0]!.text.split("\n").filter(line => line.startsWith("- "))
    expect(issues).toEqual([`- The instructions direct "approve the revised schedule" ("${APPROVE}" [source: request]), but no planned document carries it. ` +
      'Plan it in a document executed by "the Orchard Committee", adding one when none is planned.'])
    expect(model.count(AGENT.coverage)).toBe(2)

    expect(readLeanPlan(directory)!.documents.map(row => [row.id, row.title, row.executed_by])).toEqual([
      ["d1", "Release letter", ["Kestrel Works"]], ["d2", "Committee resolution", ["the Orchard Committee"]],
    ])
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", issues: [], acts: [
      expect.objectContaining({ act: approval.act, document: "d2" }), expect.objectContaining({ act: release.act, document: "d1" }),
    ] })
    // Each document is drafted in its own session and delivered.
    for (const title of ["Release letter", "Committee resolution"])
      expect(model.prompts.some(prompt => prompt.agent === AGENT.draft && prompt.text.includes(`Draft part 1 of 1 of "${title}": `))).toBe(true)
    expect((["d1.apply", "d2.apply"] as const).map(step => state.stages[step]?.status)).toEqual(["done", "done"])
    expect(delivered(directory)).toEqual([
      "body:d1:Release letter:Release letter.docx",
      "redline:d1:Release letter:Release letter - redline.docx",
      "body:d2:Committee resolution:Committee resolution.docx",
      "redline:d2:Committee resolution:Committee resolution - redline.docx",
      "drafting_notes:null:null:Release letter - drafting notes.docx",
      "drafting_notes_markdown:null:null:null",
    ])
    expect(deliverables.names).toEqual(expect.arrayContaining(["Release letter.docx", "Committee resolution.docx"]))
    // Covered: nothing about coverage is asked, disclosed or blocking.
    expect(readQuestions(directory)).toEqual([])
    expect(coverageItems(directory)).toEqual([])
    const gate = gateOf(directory)
    expect(gate).toMatchObject({ status: "pass", disclosed: [], blocking: [] })
    expect(gate).not.toHaveProperty("questions")
  }, 120_000)

  // The act sits in a document its party does not execute. The revision either
  // fixes that document's "Executed by:" line or moves the act to a companion
  // document executed by its party.
  for (const [name, first, revised, carrier] of [
    ["the revision adds the party to the carrying document's Executed by line",
      outline("# Schedule approval\nExecuted by: Kestrel Works\n## Approval\nThe approval.", RELEASE_LETTER),
      outline("# Schedule approval\nExecuted by: the Orchard Committee; Kestrel Works\n## Approval\nThe approval.", RELEASE_LETTER),
      { id: "d1", title: "Schedule approval", executed_by: ["the Orchard Committee", "Kestrel Works"] }],
    ["the revision moves the act to a companion document executed by its party",
      outline("# Schedule approval\nExecuted by: Kestrel Works\n## Approval\nThe approval.\n## Release\nThe release."),
      outline(RESOLUTION, RELEASE_LETTER),
      { id: "d1", title: "Committee resolution", executed_by: ["the Orchard Committee"] }],
  ] as const)
    test(`an act planned in a document executed by the wrong party: ${name}, with the issue in the revision prompt`, async () => {
      const directory = prepared()
      const model = scriptedModel(replies({ [AGENT.draft]: outliner(first, revised), [AGENT.coverage]: coverageOf(approval, release) }))
      expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

      expect(revisions(model)).toEqual([expect.objectContaining({ sessionID: outlinePrompt(model).sessionID })])
      expect(revisions(model)[0]!.text.split("\n").filter(line => line.startsWith("- "))).toEqual([
        `- "approve the revised schedule" ("${APPROVE}" [source: request]) must be taken by "the Orchard Committee", but d1 "Schedule approval" is executed by "Kestrel Works". ` +
        'Move the act to a document executed by "the Orchard Committee", or add a companion document executed by "the Orchard Committee".',
      ])
      const plan = readLeanPlan(directory)!
      expect(plan.documents.find(row => row.id === carrier.id)).toMatchObject({ title: carrier.title, executed_by: [...carrier.executed_by] })
      expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", issues: [] })
      expect(readLeanCoverage(directory)!.acts.find(row => row.act === approval.act)?.document).toBe(carrier.id)
      expect(readQuestions(directory)).toEqual([])
      expect(gateOf(directory)).toMatchObject({ status: "pass", disclosed: [], blocking: [] })
    }, 120_000)

  test("a request whose only instructed act is covered on the first pass plans exactly as with no instructed act, without a revision", async () => {
    const request = `Prepare the letter by which ${RELEASE}.`
    const run = async (coverage?: Reply) => {
      const directory = prepared(request)
      const model = scriptedModel(replies({ [AGENT.draft]: outliner(outline(RELEASE_LETTER)), ...(coverage ? { [AGENT.coverage]: coverage } : {}) }))
      expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
      const { basis, ...plan } = readLeanPlan(directory)!
      expect(basis).toEqual(expect.any(String))
      return { directory, model, plan }
    }
    // The coverage step's default reply lists no act.
    const none = await run()
    const one = await run(coverageOf(release))

    expect(one.plan).toEqual(none.plan)
    expect(readLeanCoverage(none.directory)).toMatchObject({ status: "covered", acts: [], issues: [] })
    expect(readLeanCoverage(one.directory)).toMatchObject({ status: "covered", acts: [expect.objectContaining({ act: release.act, document: "d1" })], issues: [] })
    // No revision or checker retry, and the same model calls in the same order.
    for (const { model } of [none, one]) {
      expect(model.prompts.filter(prompt => prompt.text.startsWith(REVISE))).toEqual([])
      expect(model.count(AGENT.coverage)).toBe(1)
    }
    expect(one.model.prompts.map(prompt => prompt.agent)).toEqual(none.model.prompts.map(prompt => prompt.agent))
    expect(readQuestions(one.directory)).toEqual([])
    expect(coverageItems(one.directory)).toEqual([])
  }, 120_000)

  test("an act whose document the request leaves out brings no companion document and no question; the drafting notes and the gate disclose it as outstanding", async () => {
    const EXCLUDE = "Prepare only documents Kestrel Works executes"
    const directory = prepared(`${REQUEST} ${EXCLUDE}.`)
    const excluding: Reply = (prompt, index) => {
      const reply = JSON.parse(coverageOf(approval, release)(prompt, index)) as { acts: { act: string; excluded_by?: string }[] }
      return JSON.stringify({ acts: reply.acts.map(row => row.act === approval.act ? { ...row, excluded_by: EXCLUDE } : row) })
    }
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(outline(RELEASE_LETTER)), [AGENT.coverage]: excluding }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(revisions(model)).toEqual([])
    expect(model.count(AGENT.coverage)).toBe(1)
    expect(readLeanPlan(directory)!.documents.map(row => row.title)).toEqual(["Release letter"])
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", issues: [], acts: [
      expect.objectContaining({ act: approval.act, document: null, excluded_by: EXCLUDE }), expect.objectContaining({ act: release.act, document: "d1" }),
    ] })
    expect(readQuestions(directory)).toEqual([])
    const items = coverageItems(directory)
    expect(items).toEqual([expect.objectContaining({ kind: "omission", detail: expect.stringContaining(`as leaving out the document for this instructed act, so it is listed as outstanding: "${approval.act}"`) })])
    const gate = gateOf(directory)
    expect(gate).toMatchObject({ status: "pass", blocking: [], disclosed: [expect.objectContaining({ id: items[0]!.id, status: "disclosed" })] })
    expect(gate).not.toHaveProperty("questions")
  }, 120_000)

  test("a coverage reply that never passes its check leaves the run going; the drafting notes and the gate disclose the one unchecked item", async () => {
    const directory = prepared()
    const deliverables = registering()
    const model = scriptedModel(replies({ [AGENT.draft]: outliner(outline(RELEASE_LETTER)), [AGENT.coverage]: () => "The plan looks complete." }))
    const state = await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })
    expect(state.status).toBe("completed")

    // One coverage job, retried to its limit, and no outline revision or question.
    expect(model.count(AGENT.coverage)).toBe(LEAN_MAX_RETRIES + 1)
    expect(model.sessions.filter(session => session.agent === AGENT.coverage)).toHaveLength(1)
    expect(revisions(model)).toEqual([])
    expect(readQuestions(directory)).toEqual([])
    expect(readLeanCoverage(directory)).toMatchObject({ status: "unchecked", acts: [] })

    const items = coverageItems(directory)
    expect(items).toEqual([expect.objectContaining({ stage: "finalize", kind: "omission", detail: expect.stringContaining("Plan coverage could not be checked") })])
    const [item] = items
    expect(model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text).toContain(`- ${item!.id}: `)
    expect(readLeanText(directory, "drafting-notes.md")).toContain(`- ${item!.id}: `)
    const gate = gateOf(directory)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.disclosed).toEqual([expect.objectContaining({ id: item!.id, status: "disclosed" })])
    expect(gate).not.toHaveProperty("questions")
    expect(deliverables.names).toContain("Release letter.docx")
  }, 120_000)

  test("a lawyer's edit of a later document's executing parties rechecks coverage before the run goes on; the gate discloses what it leaves open", async () => {
    const directory = prepared()
    const run = () => scriptedModel(replies({ [AGENT.draft]: outliner(outline(RESOLUTION, RELEASE_LETTER)), [AGENT.coverage]: coverageOf(approval, release) }))
    expect((await runLean(directory, { sessions: run().port })).status).toBe("completed")
    expect(readLeanCoverage(directory)).toMatchObject({ status: "covered", issues: [] })

    // The lawyer adds a party the record names nowhere to d2 only; d1's plan entry is unchanged.
    const drafted = readLeanText(directory, "d1/draft.md")
    const plan = readLeanPlan(directory)!
    const edited = { ...plan, documents: plan.documents.map(row => row.id === "d2" ? { ...row, executed_by: [...row.executed_by!, "the Heron Trust"] } : row) }
    writeLeanText(directory, "plan.json", `${JSON.stringify(edited, null, 2)}\n`)
    const model = run()
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    // Checked once more, against the plan as the lawyer left it, which is followed and not outlined
    // again; d1's draft is kept and only d2, whose entry changed, is drafted again (one part).
    expect(model.count(AGENT.coverage)).toBe(1)
    expect(revisions(model)).toEqual([])
    expect(model.count(AGENT.draft)).toBe(1)
    expect(readLeanText(directory, "d1/draft.md")).toBe(drafted)
    expect(readLeanPlan(directory)).toEqual(edited)
    expect(readLeanCoverage(directory)).toMatchObject({
      plan: coveragePlanHash(edited.documents, edited.acting_for), status: "uncovered",
      issues: [expect.stringContaining('d2 "Release letter" declares "Executed by: the Heron Trust"')],
    })
    const items = coverageItems(directory)
    expect(items).toEqual([expect.objectContaining({ detail: expect.stringContaining('declares "Executed by: the Heron Trust"') })])
    expect(gateOf(directory)).toMatchObject({ status: "pass", blocking: [] })
    expect(gateOf(directory).disclosed).toEqual([expect.objectContaining({ id: items[0]!.id, status: "disclosed" })])
  }, 120_000)

  // The recheck keeps d1's draft only while its saved output still passes.
  for (const [name, damage] of [
    ["its paragraph pointers are deleted", (directory: string) => rmSync(leanFile(directory, "d1/paragraphs.json"))],
    ["its text is cleared", (directory: string) => writeLeanText(directory, "d1/draft.md", "")],
  ] as const)
    test(`after that edit, d1's draft is written again, not kept, when ${name}`, async () => {
      const { directory, run, drafted } = await laterPartiesEdited()
      damage(directory)
      const model = run()
      expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

      expect(model.count(AGENT.coverage)).toBe(1)
      expect(model.prompts.filter(prompt => prompt.agent === AGENT.draft).map(prompt => /^Draft part 1 of 1 of "([^"]+)"/m.exec(prompt.text)?.[1]))
        .toEqual(["Committee resolution", "Release letter"])
      expect(readLeanText(directory, "d1/draft.md")).toBe(drafted)
      expect(readParagraphs(directory, "d1")).not.toEqual([])
    }, 120_000)

  test("after that edit, d1's kept draft reports d1's own draft items, not d2's or another step's", async () => {
    const { directory, run, drafted } = await laterPartiesEdited()
    const own = openItem("draft", "omission", "The resolution leaves the meeting date open.", { document: "d1" })
    const reviewed = openItem("review", "omission", "The resolution's quorum is not stated.", { document: "d1" })
    writeOpenItems(directory, [own, openItem("draft", "omission", "The letter leaves the release date open.", { document: "d2" }), reviewed])
    const model = run()
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    // d1's draft is kept with its own item; d2's item went with d2's new draft; d1's review item stays once.
    expect(model.count(AGENT.draft)).toBe(1)
    expect(readLeanText(directory, "d1/draft.md")).toBe(drafted)
    const left = readOpenItems(directory).filter(item => item.stage === "draft" || item.stage === "review")
    expect(left).toHaveLength(2)
    expect(left).toEqual(expect.arrayContaining([own, reviewed]))
  }, 120_000)

  test("an outlined plan whose coverage record is missing is checked again; a plan written by hand without one is followed unchecked", async () => {
    const directory = prepared(`Prepare the letter by which ${RELEASE}.`)
    const run = () => scriptedModel(replies({ [AGENT.draft]: outliner(outline(RELEASE_LETTER)), [AGENT.coverage]: coverageOf(release) }))
    expect((await runLean(directory, { sessions: run().port })).status).toBe("completed")
    const record = readLeanText(directory, "coverage.json")
    expect(record).toBeDefined()

    rmSync(leanFile(directory, "coverage.json"))
    const rechecked = run()
    expect((await runLean(directory, { sessions: rechecked.port })).status).toBe("completed")
    expect([rechecked.count(AGENT.coverage), rechecked.count(AGENT.draft)]).toEqual([1, 0])
    expect(readLeanText(directory, "coverage.json")).toBe(record)

    // Written by hand: no outline basis and no coverage record. Nothing is checked or drafted again.
    const { basis: _, ...handWritten } = readLeanPlan(directory)!
    writeLeanText(directory, "plan.json", `${JSON.stringify(handWritten, null, 2)}\n`)
    rmSync(leanFile(directory, "coverage.json"))
    const followed = run()
    expect((await runLean(directory, { sessions: followed.port })).status).toBe("completed")
    expect([followed.count(AGENT.coverage), followed.count(AGENT.draft)]).toEqual([0, 0])
    expect(coverageItems(directory)).toEqual([])
  }, 120_000)
})

// Not repeated here: exhausting the allowed outline revisions blocks the run with one
// package-scope question, a rerun or Continue makes no model call, and the
// answer re-plans, completes and asks no second question at the plan step or
// the gate (plan-scope-question.test.ts); the answered run's notes and gate
// carry the uncovered act (coverage-items.test.ts, "through the run").
