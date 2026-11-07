import { afterEach, describe, expect, test } from "bun:test"
import { checkIntakeDates, frozenDates, matterDatesBlock } from "./intake"
import { readLeanPlan, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { researchPrompt } from "./research"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES, stageInputs } from "./stages"
import { readLeanJson, readLeanText, writeLeanText, type OpenItem } from "./state"
import { AGENT, SOURCE, cleanupMatters, happyPath, matter, scriptedModel, type Reply } from "./test-support"

// The dates the work turns on: fixed by intake from the lawyer's words or the
// sources, frozen with the plan, and shown to every later stage so periods and
// the law in force run from them rather than from the run clock. Real
// coordinator; only the model is scripted.

afterEach(cleanupMatters)

const TERMS = "The supplier shall deliver the goods within thirty days of the order date."
const DATED = `The note is to be sent on 1 March 2026. The goods were ordered on 10 February 2026. ${TERMS}`
const RULE = "Compute every elapsed period, deadline and the law in force from these dates, not from today's date in the system prompt, which is the drafting day."
const SENT = { date: "2026-03-01", basis: "sources" as const, source_id: SOURCE, quote: "to be sent on 1 March 2026" }
const ORDERED = { label: "Goods ordered", date: "2026-02-10", source_id: SOURCE, quote: "ordered on 10 February 2026" }
const DRAFTING_DAY = { document_date: { date: "", basis: "today" as const }, operative_dates: [] }
const intakeWith = (dates: unknown): Reply => () => JSON.stringify({ ready: true, jurisdiction: "None needed", scope: "Supplied terms", acting_for: { party: "", basis: "none" }, dates })
const promptsOf = (model: ReturnType<typeof scriptedModel>, agent: string) => model.prompts.filter(prompt => prompt.agent === agent).map(prompt => prompt.text)
const gateItems = (directory: string) => (readLeanJson(directory, "gate.json") as { items: OpenItem[] }).items

describe("the matter's dates", () => {
  test("a date the sources fix is frozen with the plan, read by every later stage, and disclosed once", async () => {
    const directory = matter(DATED)
    const model = scriptedModel({ ...happyPath, [AGENT.intake]: intakeWith({ document_date: SENT, operative_dates: [ORDERED] }) })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(readLeanPlan(directory)?.dates).toEqual({ document_date: SENT, operative_dates: [ORDERED] })
    const dated = 'Date the document 2026-03-01, the date the sources fix for it ("to be sent on 1 March 2026"'
    for (const agent of [AGENT.research, AGENT.spine, AGENT.draft, AGENT.review, AGENT.changes, AGENT["package-review"]]) {
      expect([agent, promptsOf(model, agent)[0]]).toEqual([agent, expect.stringContaining(dated)])
      expect(promptsOf(model, agent)[0]).toContain(RULE)
    }
    // Every independent explanation has the controlling frozen dates, including
    // operative dates whose periods must not be computed from the system clock.
    const notes = promptsOf(model, AGENT.finalize)
    expect(notes.length).toBeGreaterThan(1)
    for (const prompt of notes) {
      expect(prompt).toContain("<section_kind>")
      expect(prompt).toContain(dated)
      expect(prompt).toContain('- Goods ordered: 2026-02-10 ("ordered on 10 February 2026"')
      expect(prompt).toContain(RULE)
    }
    expect(promptsOf(model, AGENT.review)[0]).toContain('- Goods ordered: 2026-02-10 ("ordered on 10 February 2026"')
    expect(promptsOf(model, AGENT.intake)[0]).not.toContain("<matter_dates>")
    expect(stageInputs(directory, "d1.review").find(part => part.name === "matter-dates")?.bytes).toBe(JSON.stringify(readLeanPlan(directory)?.dates))
    const items = gateItems(directory).filter(item => item.kind === "drafted_default")
    expect(items).toHaveLength(1)
    expect(items[0]).toMatchObject({ status: "disclosed", detail: expect.stringContaining("dated 2026-03-01, the date the sources fix for them") })
    const defaults = notes.flatMap(prompt => {
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
    expect(promptsOf(model, AGENT.review)[0]).toContain(", the drafting day, as neither the lawyer nor the sources give a date.")
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

  test("intake fixes the dates rather than asking only because the request gives none", () => {
    const prompt = researchPrompt(matter(DATED), "intake", "basis")
    expect(prompt).toContain('else one the sources fix for them (basis "sources"), else the drafting day (basis "today", date empty)')
    expect(prompt).toContain("Never ask whom the lawyer acts for, or for a date, only because the request does not say.")
  })
})
