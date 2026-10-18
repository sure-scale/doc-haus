import { afterEach, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import path from "node:path"
import { checkDraftingNotes } from "./checks"
import { selectedDeliveryPlan } from "./delivery-selection"
import { PLAN_COVERAGE_REVISIONS } from "./draft-sections"
import { finalGate, type GateResult } from "./gate"
import { PACKAGE_SCOPE_QUESTION } from "./package-scope"
import { LEAN_COVERAGE_SCHEMA, coverageCarries, coveragePlanHash, coverageScopeQuestion, planCoverageCurrent, planCoverageIssues, readLeanCoverage, type CoverageAct, type LeanCoverage } from "./package-coverage"
import { readLeanPlan, writeLeanPlan, type LeanPlan, type PlannedDocument } from "./plan"
import { openItem } from "./quotes"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { disclosableItems } from "./stages"
import {
  currentOpenItems, readLeanJson, readLeanState, readOpenItems, readQuestions, replaceStageOpenItems, updateLeanStage, writeLeanText, writeOpenItems, writeQuestions,
  type OpenItem,
} from "./state"
import { AGENT, SOURCE, cleanupMatters, draftingNotes, matter, noChanges, scriptedModel, type Reply } from "./test-support"

// Plan coverage left open after the plan step reaches the lawyer: the open
// items are derived from the current coverage.json, never stored, so the
// drafting notes must name them and a passed gate delivers them as disclosed.
// A record that did not check the saved plan (stale, missing or malformed)
// yields one item saying the saved plan was not checked, never none, unless the
// plan was written by hand without one; it is not the item of a check that ran
// and gave no usable result. Delivery names the package review writes into the
// plan leave the record current. A re-draft that supersedes draft-stage items
// keeps them. Synthetic parties and acts only; only the model is scripted.

afterEach(cleanupMatters)

const SOURCE_TEXT = "Kestrel Works shall release the deposit once the Orchard Committee approves the revised schedule."
const REQUEST = "Prepare the papers so that the Orchard Committee approves the revised schedule. Kestrel Works then releases the deposit."
const APPROVE = "the Orchard Committee approves the revised schedule"
const ANSWER = "Keep the package to the release letter; the committee's approval is handled separately."
const SCOPE_ID = `${PACKAGE_SCOPE_QUESTION}earlier`

const act: CoverageAct = { act: "approve the revised schedule", source: "request", quote: APPROVE, parties: ["the Orchard Committee"], document: "d1" }
const RELEASE: Omit<PlannedDocument, "id"> = { title: "Release letter", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["Kestrel Works"] }
/** A companion document whose declared executing party the record names nowhere. */
const UNNAMED: Omit<PlannedDocument, "id"> = { title: "Trust consent", purpose: "", parts: [{ heading: "Terms", scope: "" }], executed_by: ["the Heron Trust"] }

function prepared() {
  const directory = matter(SOURCE_TEXT)
  writeLeanText(directory, "request.md", REQUEST)
  return directory
}

function recordCoverage(directory: string, plan: LeanPlan, status: LeanCoverage["status"], acts: CoverageAct[], issues: string[]) {
  const coverage: LeanCoverage = { schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status, acts, issues }
  writeLeanText(directory, "coverage.json", `${JSON.stringify(coverage, null, 2)}\n`)
}

const answerScope = (directory: string) => writeQuestions(directory, [{ id: SCOPE_ID, stage: "draft", question: `The document plan leaves this instructed act without a document executed by the parties who must take it: "approve the revised schedule", directed by "${APPROVE}" [source: request]; to be taken by "the Orchard Committee"; planned in d1 "Release letter", executed by "Kestrel Works".`, context: "", status: "answered", answer: ANSWER }])

/** An uncovered plan: the act sits in a document its party does not execute,
 * and a companion document declares a party the record names nowhere. */
function uncovered(directory: string) {
  const plan = writeLeanPlan(directory, [RELEASE, UNNAMED])
  recordCoverage(directory, plan, "uncovered", [act], planCoverageIssues([act], plan.documents, directory))
  answerScope(directory)
  return plan
}

const coverageItems = (directory: string) => disclosableItems(directory).filter(item => item.stage === "finalize")

/** The one item of an unchecked coverage check of `plan`, from a matter of its own. */
function uncheckedItems(plan: LeanPlan) {
  const directory = prepared()
  writeLeanPlan(directory, plan.documents, plan.basis, plan.acting_for)
  recordCoverage(directory, plan, "unchecked", [], ['Reply with the JSON object only: {"acts":[]}.'])
  return coverageItems(directory)
}
/** The one item of a saved plan that no current record checked. */
const notCheckedItems = () => [openItem("finalize", "omission", "Plan coverage was not checked against the saved document plan: no current record of that check exists, " +
  "so it is not known whether each act the instructions direct has a document executed by the parties who must take it. " +
  "The lawyer should confirm that each instructed act is carried by a document executed by the parties who must take it.", { repair_target: "memo" })]
/** Notes saying each item in its detail's words, each planned document by its
 * quoted title rather than its id, and ending with its ref marker. */
const notesNaming = (items: OpenItem[]) => ["# Drafting notes", "", "The release letter records the release.", "",
  ...items.map(item => `- ${item.detail.replace(/\bd\d+ (?=")/g, "")} [ref: ${item.id}]`)].join("\n")

function gate(directory: string, notes: string): GateResult {
  writeLeanText(directory, "drafting-notes.md", `${notes}\n`)
  return finalGate(directory)
}

describe("plan coverage open items", () => {
  test("covered or absent coverage yields no item", () => {
    const directory = prepared()
    const plan = writeLeanPlan(directory, [RELEASE])
    expect(coverageItems(directory)).toEqual([])
    recordCoverage(directory, plan, "covered", [{ ...act, parties: ["Kestrel Works"] }], [])
    expect(coverageItems(directory)).toEqual([])
    expect(finalGate(directory).items).toEqual([])
  })

  test("uncovered after an answered scope question: one item per uncovered act and per other issue, disclosed by the notes and delivered by the gate", () => {
    const directory = prepared()
    uncovered(directory)
    const items = coverageItems(directory)
    expect(items).toHaveLength(2)
    // Ids are content-derived: the same record gives the same ids.
    expect(coverageItems(directory)).toEqual(items)
    expect(items.map(item => item.id)).toEqual(items.map(item => openItem(item.stage, item.kind, item.detail, item.ref).id))
    expect(items.every(item => item.kind === "omission" && item.status === "open" && item.ref.document === undefined && item.ref.repair_target === "memo")).toBe(true)

    const [uncoveredAct, declared] = items
    expect(declared!.detail).toContain('d2 "Trust consent" declares "Executed by: the Heron Trust"')
    // The outline's instruction for fixing the plan is not the lawyer's to read.
    expect(declared!.detail).not.toContain("never supply a party")
    expect(declared!.detail).toEndWith('nor any matter document names "the Heron Trust".')
    expect(uncoveredAct!.detail).toContain(`"approve the revised schedule", directed by "${APPROVE}" [source: request]`)
    expect(uncoveredAct!.detail).toContain('to be taken by "the Orchard Committee"')
    expect(uncoveredAct!.detail).toContain('planned in d1 "Release letter", executed by "Kestrel Works"')
    // The answered question is mentioned, never restated or interpreted.
    expect(uncoveredAct!.detail).toContain(SCOPE_ID)
    expect(items.some(item => item.detail.includes(ANSWER))).toBe(false)
    // Nothing is stored and no question is touched.
    expect(readOpenItems(directory)).toEqual([])
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: SCOPE_ID, status: "answered", answer: ANSWER })])

    // The drafting notes must name each item; copying an item's line passes the quote check.
    expect(checkDraftingNotes(directory, "# Drafting notes\n\nThe release letter records the release.", disclosableItems(directory)).issues)
      .toEqual([expect.stringContaining(items.map(item => item.id).join(", "))])
    expect(checkDraftingNotes(directory, notesNaming(items), disclosableItems(directory)).issues).toEqual([])

    // Unnamed, they block the gate; named, the gate passes and delivers them as disclosed.
    const blocked = gate(directory, "# Drafting notes\n\nThe release letter records the release.")
    expect(blocked.status).toBe("blocked")
    expect(blocked.blocking.map(item => item.id)).toEqual(items.map(item => item.id))
    const passed = gate(directory, notesNaming(items))
    expect(passed).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(passed.disclosed.map(item => [item.id, item.status])).toEqual(items.map(item => [item.id, "disclosed"]))
  })

  test("an act whose document the instructions leave out yields one item, even on a covered record, naming the act, who must take it and the words that leave it out", () => {
    const directory = prepared()
    const EXCLUDE = "Prepare only documents Kestrel Works executes"
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    const plan = writeLeanPlan(directory, [RELEASE])
    const excluded: CoverageAct = { ...act, document: null, excluded_by: EXCLUDE }
    recordCoverage(directory, plan, "covered", [{ ...act, parties: ["Kestrel Works"] }, excluded], [])
    expect(readLeanCoverage(directory)!.acts).toEqual([{ ...act, parties: ["Kestrel Works"] }, excluded])
    // The exclusion is the check's reading of the lawyer's words, for the lawyer to confirm.
    const detail = `Plan coverage: the check of the document plan against the instructions read the lawyer's words "${EXCLUDE}" [source: request] ` +
      "as leaving out the document for this instructed act, so it is listed as outstanding: " +
      `"approve the revised schedule", directed by "${APPROVE}" [source: request]; to be taken by "the Orchard Committee"; no planned document carries it. ` +
      "The lawyer should confirm that this act is to be dealt with outside this package."
    const items = coverageItems(directory)
    expect(items).toEqual([openItem("finalize", "omission", detail, { repair_target: "memo" })])
    expect(coverageItems(directory)).toEqual(items)
    // Unnamed, it blocks the gate; named, the gate passes and delivers it as disclosed.
    expect(gate(directory, "# Drafting notes\n\nThe release letter records the release.").blocking.map(item => item.id)).toEqual([items[0]!.id])
    expect(gate(directory, notesNaming(items))).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [], disclosed: [expect.objectContaining({ id: items[0]!.id, status: "disclosed" })] })
  })

  test("the words that leave an act's document out are quoted as the lawyer's, whatever the act's own source, and pass the notes' quote check; untagged when a matter document is named request", () => {
    const directory = prepared()
    const EXCLUDE = "Prepare only documents Kestrel Works executes"
    writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
    const plan = writeLeanPlan(directory, [RELEASE])
    recordCoverage(directory, plan, "covered", [{ ...act, source: SOURCE, document: null, excluded_by: EXCLUDE }], [])
    const [tagged] = coverageItems(directory)
    expect(tagged!.detail).toContain(`read the lawyer's words "${EXCLUDE}" [source: request] as leaving out`)
    expect(tagged!.detail).toContain(`directed by "${APPROVE}" [source: ${SOURCE}];`)
    expect(checkDraftingNotes(directory, notesNaming([tagged!]), disclosableItems(directory)).issues).toEqual([])
    // A matter document named "request" answers to that name in the quote check, so the lawyer's words go untagged.
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("request", "1", "Minutes of the committee meeting.", 0, 33)
    db.close()
    const [untagged] = coverageItems(directory)
    expect(untagged!.detail).toContain(`read the lawyer's words "${EXCLUDE}" as leaving out`)
    expect(checkDraftingNotes(directory, notesNaming([untagged!]), disclosableItems(directory)).issues).toEqual([])
  })

  describe("a left-out act whose parties all execute a planned document", () => {
    const EXCLUDE = "We are not preparing the committee resolution"
    const excluded: CoverageAct = { ...act, document: null, excluded_by: EXCLUDE }
    /** Another instructed act, carried by the release letter. */
    const released: CoverageAct = { act: "release the deposit", source: "request", quote: "Kestrel Works then releases the deposit", parties: ["Kestrel Works"], document: "d1" }
    /** The release letter, executed by the parties of both acts. */
    const SHARED: Omit<PlannedDocument, "id"> = { ...RELEASE, executed_by: ["Kestrel Works", "the Orchard Committee"] }
    const RESOLUTION: Omit<PlannedDocument, "id"> = { ...RELEASE, title: "Committee resolution", executed_by: ["the Orchard Committee"] }
    /** A second document the act's party executes, with another party. */
    const MINUTES: Omit<PlannedDocument, "id"> = { ...RELEASE, title: "Committee minutes", executed_by: ["the Orchard Committee", "Kestrel Works"] }
    /** An instruction that directs both acts. */
    const BOTH = [{ document: "request", quote: REQUEST }]
    const exclusion = `Plan coverage: the check of the document plan against the instructions read the lawyer's words "${EXCLUDE}" [source: request] ` +
      "as leaving out the document for this instructed act, so it is listed as outstanding: " +
      `"approve the revised schedule", directed by "${APPROVE}" [source: request]; to be taken by "the Orchard Committee"; no planned document carries it. ` +
      "The lawyer should confirm that this act is to be dealt with outside this package."
    /** The outline's revision issue naming the planned `documents` (one, or several). */
    const sentBack = (documents: string, several = false) => `The check read the instructions as leaving out of this package the document for "approve the revised schedule" ("${APPROVE}" [source: request]), ` +
      `which must be taken by "the Orchard Committee", but ${documents}, ` + (several
      ? "are planned and carry no instructed act. Do not plan any of those documents that the instructions leave out."
      : "is planned and carries no instructed act. If the instructions leave that document out, do not plan it.")
    const leftOut = () => {
      const directory = prepared()
      writeLeanText(directory, "request.md", `${REQUEST} ${EXCLUDE}.`)
      return directory
    }

    test("raise nothing when that document carries an instructed act: the act keeps its one unchanged item", () => {
      const directory = leftOut()
      const plan = writeLeanPlan(directory, [SHARED])
      expect(planCoverageIssues([released, excluded], plan.documents, directory)).toEqual([])
      expect(coverageScopeQuestion([released, excluded], plan.documents)).toBeUndefined()
      recordCoverage(directory, plan, "covered", [released, excluded], [])
      expect(coverageItems(directory)).toEqual([openItem("finalize", "omission", exclusion, { repair_target: "memo" })])
      // A document that carries nothing raises nothing either while only some of the act's parties execute it.
      const jointly: CoverageAct = { ...excluded, parties: ["the Orchard Committee", "Kestrel Works"] }
      expect(planCoverageIssues([released, jointly], writeLeanPlan(directory, [SHARED, RESOLUTION]).documents, directory)).toEqual([])
    })

    test("send that document back to the outline when it carries no instructed act; kept planned, it joins the act's one item for the lawyer to confirm, and the gate counts no instruction that directs the act as carried until it is dropped", () => {
      const directory = leftOut()
      const plan = writeLeanPlan(directory, [SHARED, RESOLUTION])
      const issue = sentBack('d2 "Committee resolution", executed by "the Orchard Committee"')
      expect(planCoverageIssues([released, excluded], plan.documents, directory)).toEqual([issue])
      // With no act recorded as carried, the release letter carries none either: both are named.
      expect(planCoverageIssues([excluded], plan.documents, directory)).toEqual([
        sentBack('d1 "Release letter", executed by "Kestrel Works" and "the Orchard Committee", and d2 "Committee resolution", executed by "the Orchard Committee"', true),
      ])
      recordCoverage(directory, plan, "uncovered", [released, excluded], [issue])
      const items = coverageItems(directory)
      expect(items).toEqual([openItem("finalize", "omission", `${exclusion} The lawyer should also confirm whether d2 "Committee resolution", ` +
        'which is executed by "the Orchard Committee" and carries no instructed act, should be in this package.', { repair_target: "memo" })])
      expect(coverageItems(directory)).toEqual(items)
      expect(checkDraftingNotes(directory, notesNaming(items), disclosableItems(directory)).issues).toEqual([])
      expect(gate(directory, notesNaming(items))).toMatchObject({ status: "pass", blocking: [], disclosed: [expect.objectContaining({ id: items[0]!.id, status: "disclosed" })] })
      // While that document stays planned, the release letter does not carry an instruction that also directs the left-out act.
      expect(coverageCarries(directory, BOTH, "d1")).toBe(false)
      // Once the outline drops it, the left-out act raises nothing and the release letter carries the instruction.
      const revised = writeLeanPlan(directory, [SHARED])
      recordCoverage(directory, revised, "covered", [released, excluded], [])
      expect(coverageCarries(directory, BOTH, "d1")).toBe(true)
    })

    test("name every planned document executed by its parties that carries no instructed act, in the one revision issue and in the act's one item", () => {
      const directory = leftOut()
      const plan = writeLeanPlan(directory, [SHARED, RESOLUTION, MINUTES])
      const issue = sentBack('d2 "Committee resolution", executed by "the Orchard Committee", and d3 "Committee minutes", executed by "the Orchard Committee" and "Kestrel Works"', true)
      expect(planCoverageIssues([released, excluded], plan.documents, directory)).toEqual([issue])
      recordCoverage(directory, plan, "uncovered", [released, excluded], [issue])
      const items = coverageItems(directory)
      expect(items).toEqual([openItem("finalize", "omission", `${exclusion} The lawyer should also confirm whether each of d2 "Committee resolution", ` +
        'which is executed by "the Orchard Committee", and d3 "Committee minutes", which is executed by "the Orchard Committee" and "Kestrel Works", ' +
        "should be in this package; none of them carries an instructed act.", { repair_target: "memo" })])
      expect(checkDraftingNotes(directory, notesNaming(items), disclosableItems(directory)).issues).toEqual([])
      expect(gate(directory, notesNaming(items))).toMatchObject({ status: "pass", blocking: [], disclosed: [expect.objectContaining({ id: items[0]!.id, status: "disclosed" })] })
    })
  })

  test("unchecked coverage yields one item asking the lawyer to confirm each instructed act's document", () => {
    const directory = prepared()
    const plan = writeLeanPlan(directory, [RELEASE])
    recordCoverage(directory, plan, "unchecked", [], ['Reply with the JSON object only: {"acts":[]}.'])
    const items = coverageItems(directory)
    expect(items).toEqual([expect.objectContaining({ stage: "finalize", kind: "omission", status: "open" })])
    expect(items[0]!.detail).toBe("Plan coverage could not be checked: the check of the document plan against the instructions gave no usable result, " +
      "so it is not known whether each act the instructions direct has a document executed by the parties who must take it. " +
      "The lawyer should confirm that each instructed act is carried by a document executed by the parties who must take it.")
    expect(coverageItems(directory)).toEqual(items)
    expect(gate(directory, notesNaming(items))).toMatchObject({ status: "pass", delivery: "review_draft", disclosed: [expect.objectContaining({ id: items[0]!.id })] })
  })

  test("a stale record (the plan changed since it was checked) yields the one not-checked item, never none", () => {
    const directory = prepared()
    uncovered(directory)
    expect(coverageItems(directory)).toHaveLength(2)
    const plan = writeLeanPlan(directory, [RELEASE])
    const items = coverageItems(directory)
    expect(items).toEqual(notCheckedItems())
    // Not the item of a check that ran on this plan and gave no usable result.
    expect(items[0]!.id).not.toBe(uncheckedItems(plan)[0]!.id)
    expect(gate(directory, "# Drafting notes\n\nThe release letter records the release.")).toMatchObject({ status: "blocked", blocking: [expect.objectContaining({ id: items[0]!.id })] })
    expect(gate(directory, notesNaming(items))).toMatchObject({ status: "pass", disclosed: [expect.objectContaining({ id: items[0]!.id, status: "disclosed" })] })
  })

  test("a record checked while the lawyer acted for another party is stale: the one not-checked item, and the plan is checked again", () => {
    const directory = prepared()
    const plan = writeLeanPlan(directory, [RELEASE], "0".repeat(64))
    recordCoverage(directory, plan, "covered", [{ ...act, parties: ["Kestrel Works"] }], [])
    expect(planCoverageCurrent(directory)).toBe(true)
    const acting = { party: "Kestrel Works", stated: true }
    expect(coveragePlanHash(plan.documents, acting)).not.toBe(coveragePlanHash(plan.documents, plan.acting_for))
    writeLeanPlan(directory, [RELEASE], plan.basis, acting)
    expect(planCoverageCurrent(directory)).toBe(false)
    expect(coverageItems(directory)).toEqual(notCheckedItems())
  })

  test("an outlined plan without a record yields the one not-checked item; a plan written by hand without one yields none", () => {
    const directory = prepared()
    writeLeanPlan(directory, [RELEASE], "0".repeat(64))
    expect(coverageItems(directory)).toEqual(notCheckedItems())
    writeLeanPlan(directory, [RELEASE])
    expect(coverageItems(directory)).toEqual([])
  })

  test("delivery names the package review selects leave the record current: its items stay", () => {
    const directory = prepared()
    uncovered(directory)
    const items = coverageItems(directory)
    expect(items).toHaveLength(2)
    const named = selectedDeliveryPlan(directory, [
      { document: "d1", role: "redline", file: "release.docx", instruction_quote: "" }, { document: null, role: "drafting_notes", file: "notes.docx", instruction_quote: "" },
    ])!
    writeLeanText(directory, "plan.json", `${JSON.stringify(named, null, 2)}\n`)
    expect(readLeanPlan(directory)!.documents[0]).toMatchObject({ file: "release.docx", deliverable: "redline" })
    expect(coverageItems(directory)).toEqual(items)
  })

  test("a malformed record yields the one not-checked item and never stops the notes check or the gate", () => {
    const directory = prepared()
    const plan = uncovered(directory)
    const notChecked = notCheckedItems()
    const good: LeanCoverage = { schema: LEAN_COVERAGE_SCHEMA, plan: coveragePlanHash(plan.documents, plan.acting_for), status: "uncovered", acts: [act], issues: [] }
    for (const bad of [
      { ...good, schema: "dochaus.lean.coverage/0" },
      { ...good, acts: [{ act: act.act, quote: act.quote }] },
      { ...good, acts: [{ ...act, parties: "the Orchard Committee" }] },
      { ...good, acts: [{ ...act, document: 1 }] },
      { ...good, acts: [{ ...act, document: null, excluded_by: 5 }] },
      { ...good, acts: [{ ...act, excluded_by: "Prepare only documents Kestrel Works executes" }] },
      { ...good, acts: [null] },
      { ...good, issues: [1] },
    ]) {
      writeLeanText(directory, "coverage.json", JSON.stringify(bad))
      expect(readLeanCoverage(directory)).toBeUndefined()
      expect(coverageItems(directory)).toEqual(notChecked)
      expect(checkDraftingNotes(directory, notesNaming(notChecked), disclosableItems(directory)).issues).toEqual([])
      expect(gate(directory, notesNaming(notChecked))).toMatchObject({ status: "pass", disclosed: [expect.objectContaining({ id: notChecked[0]!.id })] })
    }
    writeLeanText(directory, "coverage.json", "{")
    expect(coverageItems(directory)).toEqual(notChecked)
  })

  test("a re-draft that supersedes draft-stage items keeps the coverage items", () => {
    const directory = prepared()
    uncovered(directory)
    const items = coverageItems(directory)
    expect(items).toHaveLength(2)
    const drafted = openItem("draft", "omission", "The draft leaves the release date open.", { document: "d1" })
    writeOpenItems(directory, [drafted])
    const finalized = (result: GateResult) => result.items.filter(item => item.stage === "finalize").map(item => [item.id, item.status])
    // The gate resolves draft-stage items as superseded by the final body; the coverage items stay disclosed.
    const before = gate(directory, notesNaming(items))
    expect(before.status).toBe("pass")
    expect(before.items.find(item => item.id === drafted.id)).toMatchObject({ status: "resolved", resolution: "superseded by the check of the final body" })
    expect(finalized(before)).toEqual(items.map(item => [item.id, "disclosed"]))
    // Once the document is applied its draft-stage items drop out, and a re-draft replaces them; the coverage items remain.
    updateLeanStage(directory, readLeanState(directory)!, "d1.apply", { status: "done" })
    expect(currentOpenItems(directory)).toEqual([])
    replaceStageOpenItems(directory, "draft", "d1", [])
    expect(readOpenItems(directory)).toEqual([])
    expect(coverageItems(directory)).toEqual(items)
    expect(finalized(gate(directory, notesNaming(items)))).toEqual(items.map(item => [item.id, "disclosed"]))
  })
})

// The real run: the plan step asks the scope question, the answer re-plans,
// coverage stays uncovered, and the finalize and gate steps carry the items.
const ONE = "Acting for: not stated\n# Release letter\nExecuted by: Kestrel Works\n## Terms\nThe release."
/** The outline keeps one document, revised or not; each part gets one plain paragraph. */
const outliner: Reply = text => /^Draft part \d+ of \d+ of "/m.test(text) ? "The terms are recorded." : ONE
const replies: Record<string, Reply> = {
  [AGENT.draft]: outliner,
  [AGENT.coverage]: () => JSON.stringify({ acts: [act] }),
  [AGENT.review]: () => "No corrections are needed.",
  [AGENT.changes]: noChanges,
  [AGENT.finalize]: draftingNotes,
}

describe("plan coverage open items through the run", () => {
  test("after the answered scope question the drafting notes are asked to name the act's item and the passed gate delivers it", async () => {
    const directory = prepared()
    await runLean(directory, { sessions: scriptedModel(replies).port })
    const [question] = readQuestions(directory)
    expect(question?.id.startsWith(PACKAGE_SCOPE_QUESTION)).toBe(true)
    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: ANSWER })

    const model = scriptedModel(replies)
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    const items = coverageItems(directory)
    expect(items).toEqual([expect.objectContaining({ detail: expect.stringContaining(question!.id) })])
    const finalize = model.prompts.find(prompt => prompt.agent === AGENT.finalize)!
    expect(finalize.text).toContain(`- ${items[0]!.id}: `)
    const result = readLeanJson(directory, "gate.json") as GateResult
    expect(result).toMatchObject({ status: "pass", delivery: "review_draft" })
    expect(result.disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))
  }, 120_000)

  test("a delivery name the package review writes into the plan keeps the act's item disclosed, without a second check", async () => {
    const naming = "Save the letter as release-letter.docx."
    const directory = matter(SOURCE_TEXT)
    writeLeanText(directory, "request.md", `${REQUEST} ${naming}`)
    const delivery = [{ document: "d1", role: "body", file: "release-letter.docx", instruction_quote: naming }]
    const named = { ...replies, [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery }) }
    await runLean(directory, { sessions: scriptedModel(named).port })
    const [question] = readQuestions(directory)
    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: ANSWER })

    const model = scriptedModel(named)
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    expect(readLeanPlan(directory)!.documents[0]).toMatchObject({ file: "release-letter.docx", deliverable: "body" })
    // The answer re-plans once (an outline and its revisions); the renaming asks for no further check.
    expect(model.count(AGENT.coverage)).toBe(PLAN_COVERAGE_REVISIONS + 1)
    const items = coverageItems(directory)
    expect(items).toEqual([expect.objectContaining({ detail: expect.stringContaining(question!.id) })])
    expect((readLeanJson(directory, "gate.json") as GateResult).disclosed).toContainEqual(expect.objectContaining({ id: items[0]!.id, status: "disclosed" }))
  }, 120_000)
})
