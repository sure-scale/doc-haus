import { afterEach, describe, expect, test } from "bun:test"
import { writeFileSync } from "node:fs"
import path from "node:path"
import { readParagraphs } from "./checks"
import { checkIntakeDates, frozenDates, matterDatesBlock } from "./intake"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { researchPrompt } from "./research"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, researchBasis, stageInputs } from "./stages"
import { hashStageInputs, readLeanJson, readLeanText, readQuestions, writeLeanText, writeQuestions, type OpenItem } from "./state"
import { AGENT, DRAFT, NO_FINDINGS, SOURCE, cleanupMatters, drafter, happyPath, matter, noChanges, pidOf, scriptedModel, type Reply } from "./test-support"

afterEach(cleanupMatters)

const quiet = { ...happyPath, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges }
// The ingest service writes matter.json; a synthetic one stands in for it.
const recordMatter = (directory: string, fields: Record<string, unknown>) =>
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter-fixture", name: "Supply note", ...fields }))
const CLIENT_BLOCK = "<matter_client>\nThe lawyer's client as recorded on the matter: Northwind Traders Ltd\n"

test("a complete brief reaches drafting without forcing an unnecessary questionnaire", async () => {
  const directory = matter()
  const request = "Draft a short supply note under Singapore law. Act for the supplier. Routine contact details may remain placeholders."
  writeLeanText(directory, "request.md", request)
  const model = scriptedModel(quiet)
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.prompts[0]!.text).toContain(request)
  expect(model.prompts.find(row => row.agent === AGENT.intake)!.text).toContain("Do not ask for information already supplied")
  expect(readQuestions(directory)).toEqual([])
}, 120_000)

test("conflicting material instructions can pause before a plan or draft is written", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", "Use Singapore law, but retain the template's England and Wales governing law clause unchanged.")
  const model = scriptedModel({ [AGENT.intake]: () => JSON.stringify({ questions: [
    { question: "Which governing law should the draft use?", kind: "instruction", context: "The requested law conflicts with the instruction to retain the existing clause." },
    { question: "Which governing law should the draft use?", kind: "instruction", context: "The same decision affects the whole document." },
  ] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  expect(model.prompts.map(row => row.agent)).toEqual([AGENT.instructions, AGENT.discrepancy, AGENT.discrepancy, AGENT.intake, AGENT.advise])
  expect(readQuestions(directory)).toHaveLength(1)
  expect(readLeanText(directory, "plan.json")).toBeUndefined()
  expect(readLeanText(directory, "d1/draft.md")).toBeUndefined()
}, 120_000)

test("malformed intake questions fail visibly without drafting through the missing decision", async () => {
  const directory = matter()
  const model = scriptedModel({ [AGENT.intake]: () => '{"questions":[{"question":"Which law?"}]}' })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("failed")
  expect(model.prompts.filter(row => row.agent === AGENT.intake)).toHaveLength(4)
  expect(state.stages["d1.review"]?.status).toBe("pending")
  expect(readLeanText(directory, "d1/draft.md")).toBeUndefined()
}, 120_000)

test("an already answered intake question gets correction feedback instead of being shown again", async () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-existing", stage: "draft", question: "Which law should govern?", context: "Legal basis.", status: "answered", answer: "Singapore law." }])
  const model = scriptedModel({ ...quiet, [AGENT.intake]: (prompt, index) => index === 0
    ? JSON.stringify({ questions: [{ question: "Which law should govern?", kind: "instruction", context: "Legal basis." }] })
    : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms", acting_for: { party: "", basis: "none" }, dates: { document_date: { date: "", basis: "today" }, operative_dates: [] } }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.prompts.filter(row => row.agent === AGENT.intake)[1]!.text).toContain("already has an answer")
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: "q-existing", status: "answered" })])
}, 120_000)

test("an intake question naming an internal id gets feedback on how to say it, and is asked in words on the last attempt", async () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-0a1b2c3d4e", stage: "draft", question: "Which law should govern?", context: "Legal basis.", status: "answered", answer: "Singapore law." }])
  const model = scriptedModel({ ...quiet, [AGENT.intake]: () =>
    JSON.stringify({ questions: [{ question: "Does the answer to q-0a1b2c3d4e also cover disputes?", kind: "instruction", context: "Disputes may need their own forum." }] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  const intake = model.prompts.filter(row => row.agent === AGENT.intake)
  expect(intake[1]!.text).toContain('questions[0].question names the internal id "q-0a1b2c3d4e". The lawyer sees only words: name the question or answer by what it asked.')
  const asked = readQuestions(directory).find(row => row.status === "open")!
  expect(asked.question).not.toContain("q-0a1b2c3d4e")
  expect(asked.question).toContain("Which law should govern?")
}, 120_000)

test("a changed answer during review re-runs the review and reaches the accepted draft as an amendment before delivery", async () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-law", stage: "draft", question: "Which law should govern?", context: "Legal basis.", status: "answered", answer: "England and Wales law." }])
  const DRAFTED = "This note is governed by England and Wales law."
  const AMENDED = "This note is governed by Singapore law."
  let changed = false
  const model = scriptedModel({ ...quiet,
    [AGENT.draft]: drafter(`# Supply note\n\n## Governing law\n\n${DRAFTED}\n`),
    [AGENT.review]: () => {
      if (!changed) {
        changed = true
        enqueueLeanRequest(directory, { kind: "answer", question_id: "q-law", answer: "Singapore law." })
      }
      return NO_FINDINGS
    },
    [AGENT.changes]: (prompt, index) => prompt.includes("Singapore law.") ? JSON.stringify({ amendments: [{ pid: pidOf(prompt, DRAFTED), updated: AMENDED,
      reason: "Follows the lawyer's answer on governing law.", sources: [{ source_id: "request", quote: "Singapore law." }] }], questions: [] }) : noChanges(prompt, index),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.review)).toBe(2)
  // The draft accepted before the answer changed is kept; the answer reaches it through its review and changes.
  const outlines = model.prompts.filter(row => row.agent === AGENT.draft && row.text.includes("Reply now with the outline"))
  expect(outlines).toHaveLength(1)
  expect(readParagraphs(directory, "d1").map(row => row.text)).toContain(DRAFTED)
  expect(readParagraphs(directory, "d1", "paragraphs-applied.json").some(row => row.text === AMENDED)).toBe(true)
}, 120_000)

test("intake alone reads the client recorded on the matter, as data that a party the lawyer names overrides", async () => {
  const directory = matter()
  const intake = () => researchPrompt(directory, "intake", researchBasis(directory))
  expect(intake()).not.toContain("<matter_client>")
  recordMatter(directory, { client: "" })
  expect(intake()).not.toContain("<matter_client>")
  recordMatter(directory, { client: "Northwind Traders Ltd" })
  expect(intake()).toContain(CLIENT_BLOCK)
  expect(intake()).toContain("Do not ask whom the lawyer acts for")
  for (const stage of ["research", "research-review"] as const) expect(researchPrompt(directory, stage, researchBasis(directory))).not.toContain("Northwind Traders Ltd")

  const model = scriptedModel(quiet)
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.prompts.find(row => row.agent === AGENT.intake)!.text).toContain(CLIENT_BLOCK)
}, 120_000)

test("intake is keyed on the matter's client only when one is recorded; no other step is", () => {
  const directory = matter()
  const steps = ["discrepancy", "intake", "research", "research-review", "spine"] as const
  const keys = () => steps.map(step => hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)))
  const before = keys()
  expect(stageInputs(directory, "intake").map(row => row.name)).not.toContain("matter-client")
  for (const fields of [{}, { client: "" }, { client: " \n " }, { client: 42 }]) {
    recordMatter(directory, fields)
    expect(keys()).toEqual(before)
  }
  recordMatter(directory, { client: "Northwind Traders Ltd" })
  expect(stageInputs(directory, "intake")).toContainEqual({ name: "matter-client", bytes: "Northwind Traders Ltd" })
  expect(keys().map((key, index) => key !== before[index])).toEqual([false, true, false, false, false])
})

describe("the matter's dates", () => {
  const TERMS = "The supplier shall deliver the goods within thirty days of the order date."
  const DATED = `The note is to be sent on 1 March 2026. The goods were ordered on 10 February 2026. ${TERMS}`
  const SENT = { date: "2026-03-01", basis: "sources" as const, source_id: SOURCE, quote: "to be sent on 1 March 2026" }
  const ORDERED = { label: "Goods ordered", date: "2026-02-10", source_id: SOURCE, quote: "ordered on 10 February 2026" }
  const DRAFTING_DAY = { document_date: { date: "", basis: "today" as const }, operative_dates: [] }
  const intakeWith = (dates: unknown): Reply => () => JSON.stringify({ ready: true, jurisdiction: "None needed", scope: "Supplied terms", acting_for: { party: "", basis: "none" }, dates })
  const promptsOf = (model: ReturnType<typeof scriptedModel>, agent: string) => model.prompts.filter(prompt => prompt.agent === agent).map(prompt => prompt.text)
  const gateItems = (directory: string) => (readLeanJson(directory, "gate.json") as { items: OpenItem[] }).items

  test("a date the sources fix is frozen with the plan, read by every later stage, and disclosed once", async () => {
    const directory = matter(DATED)
    const model = scriptedModel({ ...happyPath, [AGENT.intake]: intakeWith({ document_date: SENT, operative_dates: [ORDERED] }) })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(readLeanPlan(directory)?.dates).toEqual({ document_date: SENT, operative_dates: [ORDERED] })
    const dated = 'Date the document 2026-03-01, the date the sources fix for it ("to be sent on 1 March 2026"'
    for (const agent of [AGENT.research, AGENT.spine, AGENT.draft, AGENT.review, AGENT.changes, AGENT["package-review"]]) {
      expect([agent, promptsOf(model, agent)[0]]).toEqual([agent, expect.stringContaining(dated)])
    }
    // Every independent explanation has the controlling frozen dates, including
    // operative dates whose periods must not be computed from the system clock.
    const notes = [...promptsOf(model, AGENT.finalize), ...promptsOf(model, AGENT.notesEvidence)]
    expect(notes.length).toBeGreaterThan(1)
    const linking = notes.filter(prompt => prompt.includes("<current_target_lookup>"))
    const prose = notes.filter(prompt => prompt.includes("<section_kind>"))
    expect(linking.length).toBeGreaterThan(0)
    expect(prose.length).toBeGreaterThan(1)
    expect(linking.length + prose.length).toBe(notes.length)
    for (const prompt of linking) {
      expect(prompt).toContain("<assigned_point>")
      expect(prompt).toContain('"targets":[')
    }
    for (const prompt of notes) {
      expect(prompt).toContain(dated)
      expect(prompt).toContain('- Goods ordered: 2026-02-10 ("ordered on 10 February 2026"')
    }
    expect(promptsOf(model, AGENT.review)[0]).toContain('- Goods ordered: 2026-02-10 ("ordered on 10 February 2026"')
    expect(promptsOf(model, AGENT.intake)[0]).not.toContain("<matter_dates>")
    expect(stageInputs(directory, "d1.review").find(part => part.name === "matter-dates")?.bytes).toBe(JSON.stringify(readLeanPlan(directory)?.dates))
    const items = gateItems(directory).filter(item => item.kind === "drafted_default")
    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({ status: "disclosed", detail: expect.stringContaining("dated 2026-03-01, the date the sources fix for them") })
    const defaults = prose.flatMap(prompt => {
      const block = /<open_point>\n([\s\S]*?)\n<\/open_point>/.exec(prompt)?.[1]
      return block ? [JSON.parse(block)] : []
    }).filter(point => point.kind === "drafted_default")
    expect(defaults).toEqual([{ kind: items[0]!.kind, detail: items[0]!.detail, location: items[0]!.ref }])
    expect(readLeanText(directory, "drafting-notes.md")!.split(`[ref: ${items[0]!.id}]`)).toHaveLength(2)
  }, 120_000)

  test("a date whose quotation is not in its source is retried, and on the last attempt the documents bear the drafting day", async () => {
    const directory = matter(DATED)
    const model = scriptedModel({ ...happyPath, [AGENT.intake]: intakeWith({ document_date: { ...SENT, quote: "to be sent on 2 March 2026" }, operative_dates: [ORDERED] }) })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(promptsOf(model, AGENT.intake).filter(text => text.includes('dates.document_date quotes "to be sent on 2 March 2026"'))).toHaveLength(LEAN_MAX_RETRIES)
    expect(readLeanJson(directory, "intake.json")).toMatchObject({ dates: { document_date: { date: "", basis: "today" }, operative_dates: [ORDERED] } })
    expect(readLeanPlan(directory)?.dates?.document_date).toEqual({ date: expect.stringMatching(/^\d{4}-\d{2}-\d{2}$/), basis: "today" })
    expect(gateItems(directory).some(item => item.kind === "drafted_default")).toBe(false)
  }, 120_000)

  test("the readiness record is checked: each date is a calendar date with a quotation found in its source", () => {
    const directory = matter(DATED)
    expect(checkIntakeDates(directory, undefined)).toEqual({ issues: [expect.stringContaining('"dates":{')], value: DRAFTING_DAY })
    expect(checkIntakeDates(directory, DRAFTING_DAY)).toEqual({ issues: [], value: DRAFTING_DAY })
    expect(checkIntakeDates(directory, { document_date: { ...SENT, date: "2026-02-30" }, operative_dates: [] }).issues).toEqual(["dates.document_date needs its date as YYYY-MM-DD."])
    expect(checkIntakeDates(directory, { document_date: { date: "2026-03-01", basis: "request" }, operative_dates: [] }).issues[0]).toContain("without the source_id and exact quote")
    expect(checkIntakeDates(directory, { document_date: { date: "", basis: "later" }, operative_dates: [] }).issues[0]).toContain("needs a basis")
    expect(checkIntakeDates(directory, { document_date: DRAFTING_DAY.document_date, operative_dates: [{ ...ORDERED, label: "" }, ORDERED] }))
      .toEqual({ issues: ["dates.operative_dates[0] needs a label saying what happens on the date."], value: { ...DRAFTING_DAY, operative_dates: [ORDERED] } })
    expect(checkIntakeDates(directory, { document_date: SENT, operative_dates: [ORDERED] })).toEqual({ issues: [], value: { document_date: SENT, operative_dates: [ORDERED] } })
  })

  test("the drafting day is fixed when the plan is made; a run that recorded no dates shows none", () => {
    const directory = matter(DATED)
    writeLeanText(directory, "intake.json", JSON.stringify({ ready: true, dates: DRAFTING_DAY }))
    expect(frozenDates(directory, new Date(2026, 4, 7))).toEqual({ document_date: { date: "2026-05-07", basis: "today" }, operative_dates: [] })
    writeLeanText(directory, "intake.json", JSON.stringify({ ready: true, dates: { document_date: SENT, operative_dates: [] } }))
    expect(frozenDates(directory)).toEqual({ document_date: SENT, operative_dates: [] })

    const undated = matter(DATED)
    writeLeanPlan(undated, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
    expect(matterDatesBlock(undated)).toBe("")
    expect(stagePrompt(undated, "d1.review")).not.toContain("<matter_dates>")
    expect(stageInputs(undated, "d1.review").some(part => part.name === "matter-dates")).toBe(false)
  })
})
