import { afterEach, describe, expect, test } from "bun:test"
import { NOT_STATED, nameKey, namedIn, operativeDocument, parsePlan, planActingFor, readLeanPlan, writeLeanPlan, type StepID } from "./plan"
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
    const outline = (line: string) => `Acting for: ${line}\n# Title\nForm: other\n## One`
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
    // Curly double quotes do not ground a party the request does not name.
    const quoted = parsePlan(outline("\u201CKestrel Works\u201D"), request)
    expect(quoted.issues.join(" ")).toContain('names "Kestrel Works", but the request does not name that party')
    expect(quoted.value?.acting_for).toEqual(NOT_STATED)
  })

  test("a party is named only as whole words, with punctuation and quotation marks folded", () => {
    expect(namedIn("Supply terms with Hannah Ltd.", "Ann")).toBe(false)
    expect(namedIn("Supply terms with Joann Smith.", "Ann")).toBe(false)
    expect(namedIn("Supply terms with Ann, of Hannah Ltd.", "Ann")).toBe(true)
    // A name that folds to nothing is named nowhere.
    expect(namedIn("Supply terms with Ann.", ".")).toBe(false)
    expect(namedIn("Supply terms with Ann.", "")).toBe(false)
    expect(namedIn("Review the agreement with Acme, Inc. dated 1 March.", "Acme Inc.")).toBe(true)
    expect(nameKey("Acme Inc.")).toBe(nameKey("Acme, Inc."))
    expect(nameKey("Acme Inc")).toBe(nameKey("Acme Inc."))
    expect(nameKey("Acme Inc")).not.toBe(nameKey("Acme LLC"))
    expect(nameKey("\u201CKestrel Works\u201D")).toBe(nameKey("Kestrel Works"))
  })

  test("common spellings of one legal form name the same party, while different forms stay different", () => {
    for (const [short, long] of [["Kestrel Works Ltd", "Kestrel Works Limited"], ["Kestrel Works Inc.", "Kestrel Works Incorporated"], ["Kestrel Works & Co.", "Kestrel Works & Company"],
      ["Kestrel Works Corp", "Kestrel Works Corporation"], ["Kestrel Works LLC", "Kestrel Works, L.L.C."], ["Kestrel Works Co., Ltd.", "Kestrel Works Company Limited"],
      ["Kestrel Works LLP", "Kestrel Works L.L.P."], ["Kestrel Works LP", "Kestrel Works L.P."], ["Kestrel Works PLC", "Kestrel Works P.L.C."]] as const) {
      expect(nameKey(short)).toBe(nameKey(long))
      expect(namedIn(`The agreement with ${long} dated 1 March.`, short)).toBe(true)
      expect(namedIn(`The agreement with ${short} dated 1 March.`, long)).toBe(true)
    }
    expect(namedIn("The deposit is Kestrel Works Limited's.", "Kestrel Works Ltd")).toBe(true)
    expect(nameKey("Kestrel Works Ltd")).not.toBe(nameKey("Kestrel Works LLC"))
    expect(namedIn("The agreement with Kestrel Works Limited.", "Kestrel Works LLC")).toBe(false)
    expect(namedIn("The agreement with Kestrel Works.", "Kestrel Works Ltd")).toBe(false)
    // A form word is a legal form only after a name: "the Company" stays "Company".
    expect(nameKey("the Company")).toBe(nameKey("Company"))
    expect(nameKey("the Company")).not.toBe(nameKey("Co"))
    expect(namedIn("The Company shall pay.", "the Company")).toBe(true)
    expect(namedIn("Wren Atelier Co. shall pay.", "the Company")).toBe(false)
  })

  test("a hand-written plan without the slot is not stated, and the lawyer's edit of it re-keys every model step but the accepted draft", async () => {
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
    // The accepted draft is keyed on the outline contract and its own plan entry, not
    // on whom the lawyer acts for. Apply keys on
    // the whole plan.json (it measures template coverage from the plan), so it
    // re-keys too; it makes no model call.
    expect(keys().map((key, index) => key !== before[index])).toEqual([false, true, true, true, true, false])
  }, 120_000)
})

// Who executes each planned document: an optional "Executed by:" line under its
// title, read into the plan for the package-coverage check.
describe("who executes each planned document", () => {
  const outline = (lines: string) => `Acting for: not stated\n# Resolution\nRecords the approval.\nForm: other\n${lines}## Approval\nThe approval.`

  test("an Executed by line names each executing party, or none", () => {
    expect(parsePlan(outline("Executed by: the Orchard Committee; Kestrel Works\n")).value?.documents[0]?.executed_by).toEqual(["the Orchard Committee", "Kestrel Works"])
    expect(parsePlan(outline("- **Executed by:** \u201CKestrel Works\u201D.\n")).value?.documents[0]?.executed_by).toEqual(["Kestrel Works"])
    for (const line of ["Executed by: none", "Executed by: None.", "- Executed by: none (it takes no legal act)"])
      expect(parsePlan(outline(`${line}\n`))).toEqual({ issues: [], value: expect.objectContaining({ documents: [expect.objectContaining({ executed_by: [] })] }) })
    // Lines add up; a non-answer declares nothing; a line under a part is scope text.
    expect(parsePlan(outline("Executed by: Kestrel Works\nExecuted by: Wren Atelier; Kestrel Works\n")).value?.documents[0]?.executed_by).toEqual(["Kestrel Works", "Wren Atelier"])
    expect(parsePlan(outline("Executed by: not stated\n")).value?.documents[0]).not.toHaveProperty("executed_by")
    const late = parsePlan("Acting for: not stated\n# Resolution\nForm: other\n## Approval\nExecuted by: Kestrel Works").value?.documents[0]
    expect(late).not.toHaveProperty("executed_by")
    expect(late?.parts[0]?.scope).toBe("Executed by: Kestrel Works")
  })

  test("a later none or non-answer keeps the parties already named, and none must be the whole answer", () => {
    const executedBy = (lines: string) => parsePlan(outline(lines)).value?.documents[0]?.executed_by
    expect(executedBy("Executed by: Kestrel Works\nExecuted by: none\n")).toEqual(["Kestrel Works"])
    expect(executedBy("Executed by: Kestrel Works\nExecuted by: not stated\n")).toEqual(["Kestrel Works"])
    expect(executedBy("Executed by: None Corp\n")).toEqual(["None Corp"])
    expect(executedBy("Executed by: Kestrel Works\nExecuted by: Kestrel Works.\n")).toEqual(["Kestrel Works"])
  })

  test("a non-answer, however it is phrased, declares no party", () => {
    for (const answer of ["not applicable", "Not applicable (it takes no legal act)", "N.A.", "no party", "No parties.", "none stated", "Not named", "Not identified", "nobody", "TBC", "[to be confirmed]"]) {
      const executed = parsePlan(outline(`Executed by: ${answer}\n`))
      expect(executed.issues).toEqual([])
      expect(executed.value?.documents[0]).not.toHaveProperty("executed_by")
      expect(parsePlan(`Acting for: ${answer}\n# Title\nForm: other\n## One`, "Review the agreement.")).toEqual({ issues: [], value: expect.objectContaining({ acting_for: NOT_STATED }) })
    }
  })

  test("a dash or bracket that continues a name stays part of it, while trailing commentary does not", () => {
    const executedBy = (answer: string) => parsePlan(outline(`Executed by: ${answer}\n`)).value?.documents[0]?.executed_by
    for (const name of ["Heron – Vale LLP", "Heron - Vale LLP", "Heron—Vale LLP", "Wren Atelier Ltd (UK)", "Wren Fund (No. 2)"])
      expect(executedBy(name)).toEqual([name])
    for (const [answer, name] of [["Kestrel Works – as seller", "Kestrel Works"], ["Kestrel Works—the seller", "Kestrel Works"], ["Kestrel Works - the seller.", "Kestrel Works"],
      ["Kestrel Works (the seller)", "Kestrel Works"], ["Kestrel Works (“Seller”)", "Kestrel Works"], ["Heron – Vale LLP – the buyer", "Heron – Vale LLP"],
      ["Wren Atelier Ltd (UK) (as buyer)", "Wren Atelier Ltd (UK)"], ["(Wren Atelier Ltd (UK))", "Wren Atelier Ltd (UK)"], ["“Wren Atelier Ltd (UK)”.", "Wren Atelier Ltd (UK)"]])
      expect(executedBy(answer!)).toEqual([name])
    const request = "We act for Heron – Vale LLP; the other side is Wren Atelier Ltd (UK)."
    expect(parsePlan("Acting for: Heron – Vale LLP\n# Title\nForm: other\n## One", request)).toEqual({ issues: [], value: expect.objectContaining({ acting_for: { party: "Heron – Vale LLP", stated: true } }) })
    expect(parsePlan("Acting for: Wren Atelier Ltd (UK) — the other side\n# Title\nForm: other\n## One", request)).toEqual({ issues: [], value: expect.objectContaining({ acting_for: { party: "Wren Atelier Ltd (UK)", stated: true } }) })
  })

  test("the Executed by label is matched loosely, and only under a title", () => {
    expect(parsePlan(outline("EXECUTED   BY: Kestrel Works\n")).value?.documents[0]?.executed_by).toEqual(["Kestrel Works"])
    const early = parsePlan("Acting for: not stated\nExecuted by: Kestrel Works\n# Resolution\nRecords the approval.\nForm: other\n## Approval\nThe approval.")
    expect(early.issues).toEqual([])
    expect(early.value?.documents[0]).not.toHaveProperty("executed_by")
  })

  test("a missing Executed by line is no issue and leaves the plan as before", () => {
    expect(parsePlan(outline(""), "Approve the schedule.")).toStrictEqual({
      issues: [],
      value: { documents: [{ title: "Resolution", purpose: "Records the approval.", parts: [{ heading: "Approval", scope: "The approval." }], form: "other" }], acting_for: NOT_STATED },
    })
  })

  test("the plan keeps the executing parties, and the drafter reads them", async () => {
    const directory = matter()
    const documents = [{ title: "Resolution", purpose: "", parts: [{ heading: "Approval", scope: "" }], executed_by: ["Kestrel Works"] }, { title: "Advice note", purpose: "", parts: [{ heading: "Advice", scope: "" }], executed_by: [] }]
    writeLeanPlan(directory, documents)
    expect(readLeanPlan(directory)?.documents.map(row => row.executed_by)).toEqual([["Kestrel Works"], []])
    expect(operativeDocument(readLeanPlan(directory)!.documents[0])?.executed_by).toEqual(["Kestrel Works"])
    const plan = readLeanPlan(directory)!
    writeLeanText(directory, "plan.json", JSON.stringify({ ...plan, documents: [{ ...plan.documents[0]!, executed_by: [1] }] }))
    expect(readLeanPlan(directory)).toBeUndefined()
    writeLeanText(directory, "plan.json", JSON.stringify({ ...plan, documents: [{ ...plan.documents[0]!, executed_by: "Kestrel Works" }] }))
    expect(readLeanPlan(directory)).toBeUndefined()
  })
})
