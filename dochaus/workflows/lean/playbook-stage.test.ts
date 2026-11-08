import { afterEach, expect, test } from "bun:test"
import { leanSteps, parsePlan, plannedDocument, reviewGuidance, reviewStages, writeLeanPlan } from "./plan"
import { playbookSelectionIssues } from "./playbook-review"
import { packageReviewPrompt } from "./package-review"
import { stagePrompt } from "./prompts"
import { changePlanPrompt } from "./change-pipeline"
import { runLean, stageInputSha } from "./runner"
import { LEAN_CHILD_PERMISSIONS } from "./stages"
import { readLeanText, writeLeanText } from "./state"
import { AGENT, DRAFT, DRAFTED, SOURCE, SOURCE_TEXT, SUPPORT, cleanupMatters, drafter, happyPath, matter, pidOf, scriptedModel } from "./test-support"

afterEach(cleanupMatters)
const selection = { sources: [{ source_id: SOURCE, quote: SOURCE_TEXT }], reason: "The supplied drafting guidance applies to this supply note." }
const PLAYBOOK_FINDING = "Correction required at the delivery paragraph: the supplied terms require the stated delivery period."
/** The review findings a change plan prompt lists for adjudication. */
const listedFindings = (prompt: string) => JSON.parse(/<review_findings>\n([\s\S]*?)\n<\/review_findings>/.exec(prompt)?.[1] ?? "[]") as Record<string, unknown>[]
const doc = { title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }

test("planning routes the optional review per document and rejects invented evidence", () => {
  const directory = matter()
  expect(leanSteps(directory)).not.toContain("d1.review-playbook")
  const parsed = parsePlan(`Acting for: not stated\n# Supply note\nRecord terms\nForm: other\nPlaybook review: ${JSON.stringify(selection)}\n## Delivery\nDeliver goods\n# Factual inventory\nList sources only\nForm: other\nPlaybook review: {"sources":[],"reason":"A factual inventory imposes no supply terms."}\n## Sources\nList sources`)
  expect(parsed.issues).toEqual([])
  expect(playbookSelectionIssues(directory, parsed.value!.documents)).toEqual([])
  writeLeanPlan(directory, parsed.value!.documents)
  expect(leanSteps(directory)).toContain("d1.review-playbook")
  expect(leanSteps(directory)).not.toContain("d2.review-playbook")
  expect(reviewStages(directory, "d1")).toHaveLength(5)
  expect(reviewStages(directory, "d2")).toHaveLength(4)
  const changed = structuredClone(parsed.value!.documents)
  changed[0]!.playbook_review!.sources[0]!.quote = "No such guidance was supplied."
  expect(playbookSelectionIssues(directory, changed)).toHaveLength(1)
  changed[0]!.playbook_review!.sources[0]!.source_id = "unknown.docx"
  expect(playbookSelectionIssues(directory, changed)).toHaveLength(1)
  expect(parsePlan('Acting for: not stated\n# Note\nPlaybook review: invalid\n## Terms').value).toBeUndefined()
})

const focus = (prompt: string) => /<review_focus>\n([\s\S]*?)\n<\/review_focus>/.exec(prompt)?.[1] ?? ""
const TEMPLATE_POSITIONS = "positions are the provisions its text carries and every note or bracketed direction to the drafter"

test("a document drafted from a template gets the guidance review with the template among its guidance", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...doc, template: SOURCE }, { ...doc, title: "Cover note" }])
  expect(reviewStages(directory, "d1")).toContain("review-playbook")
  expect(leanSteps(directory)).toContain("d1.review-playbook")
  expect(reviewGuidance(directory, "d1")).toEqual([SOURCE])
  // Without a template or a selected playbook, a document keeps the ordinary reviews.
  expect(reviewStages(directory, "d2")).not.toContain("review-playbook")
  expect(leanSteps(directory)).not.toContain("d2.review-playbook")
  const prompt = stagePrompt(directory, "d1.review-playbook")
  expect(prompt).toContain(`<guidance>\nCheck this document against each of these sources, read in full with read-document:\n- ${SOURCE} (the template this document is drafted from)\n</guidance>`)
  expect(focus(prompt)).toContain(TEMPLATE_POSITIONS)
  expect(focus(prompt)).not.toMatch(/restore/i)
  // The other reviews of the templated document keep their focus.
  expect(focus(stagePrompt(directory, "d1.review"))).not.toContain(TEMPLATE_POSITIONS)
})

test("a revision with no selected playbook keeps the ordinary reviews: its original is the baseline, not guidance", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...doc, mode: "revise", template: SOURCE, comparison: SOURCE }, { ...doc, title: "Revised note", mode: "revise", template: SOURCE, playbook_review: selection }])
  expect(reviewGuidance(directory, "d1")).toEqual([])
  expect(reviewStages(directory, "d1")).not.toContain("review-playbook")
  expect(leanSteps(directory)).not.toContain("d1.review-playbook")
  // A selected playbook still schedules it, without the original or the template focus.
  expect(reviewGuidance(directory, "d2")).toEqual([SOURCE])
  expect(focus(stagePrompt(directory, "d2.review-playbook"))).not.toContain(TEMPLATE_POSITIONS)
})

test("the guidance list is the selected playbook sources plus the template, the focus sentence only with a template", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...doc, playbook_review: selection, template: "precedent.md" }, { ...doc, title: "Cover note", playbook_review: selection }])
  expect(reviewGuidance(directory, "d1")).toEqual([SOURCE, "precedent.md"])
  expect(reviewGuidance(directory, "d2")).toEqual([SOURCE])
  const templated = stagePrompt(directory, "d1.review-playbook")
  expect(templated).toContain(`- ${SOURCE} (selected guidance)\n- precedent.md (the template this document is drafted from)`)
  expect(focus(templated)).toContain(TEMPLATE_POSITIONS)
  const selected = stagePrompt(directory, "d2.review-playbook")
  expect(selected).toContain(`- ${SOURCE} (selected guidance)\n</guidance>`)
  expect(focus(selected)).not.toContain(TEMPLATE_POSITIONS)
  expect(packageReviewPrompt(directory)).toContain(`<playbook_review document="d1">\nSelection: ${JSON.stringify(selection)}\nTemplate: precedent.md\n`)
})

test("a changed selection revisits consumers while an unassigned body retains the original reviews", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ ...doc, playbook_review: selection }])
  const steps = ["d1.review-independent", "d1.review-playbook", "d1.change-plan", "d1.change-review", "package-review"] as const
  const before = steps.map(step => stageInputSha(directory, step))
  writeLeanPlan(directory, [{ ...doc, playbook_review: { sources: [], reason: "The lawyer restricted this body to an inventory." } }])
  steps.forEach((step, index) => expect(stageInputSha(directory, step)).not.toBe(before[index]))
  expect(reviewStages(directory, "d1")).not.toContain("review-playbook")
})

test("the native coordinator accepts the specialist before adjudication and final package review", async () => {
  const directory = matter()
  const drafting = drafter(DRAFT)
  const base = scriptedModel({ ...happyPath, [AGENT.draft]: (prompt, index) => {
    const reply = drafting(prompt, index)
    return reply.includes("Acting for:") ? reply.replace("# Supply note", `# Supply note\nPlaybook review: ${JSON.stringify(selection)}`) : reply
  } })
  let checked = false
  const sessions = { ...base.port, async prompt(input: Parameters<typeof base.port.prompt>[0]) {
    if (input.agent === AGENT["review-playbook"]) {
      expect(input.text).toContain(JSON.stringify(selection))
      checked = true
      return JSON.stringify({ coverage: "Playbook coverage: supplied terms require the stated delivery period.", findings: [{
        pid: pidOf(input.text, DRAFTED), origin: "request", balance: "neutral", finding: PLAYBOOK_FINDING, sources: SUPPORT,
      }] })
    }
    if (input.agent === AGENT["change-plan"]) {
      expect(checked).toBe(true)
      // The plan adjudicates the specialist's finding, listed with its review.
      expect(listedFindings(input.text)).toContainEqual(expect.objectContaining({ review: "review-playbook", finding: PLAYBOOK_FINDING }))
    }
    if (input.agent === AGENT["package-review"]) expect(input.text).toContain("Playbook coverage:")
    return base.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions })
  expect(state.status).toBe("completed")
  expect(state.stages["d1.review-playbook"]).toMatchObject({ status: "done", outputs: ["d1/review-playbook.json"] })
  expect(base.sessions.find(session => session.agent === AGENT["review-playbook"])?.permissions).toEqual(LEAN_CHILD_PERMISSIONS)
  expect(plannedDocument(directory, "d1")?.playbook_review).toEqual(selection)
  expect(listedFindings(changePlanPrompt(directory, "d1"))).toContainEqual(expect.objectContaining({ review: "review-playbook", finding: PLAYBOOK_FINDING }))
  expect(packageReviewPrompt(directory)).toContain("Playbook coverage:")
  const before = stageInputSha(directory, "d1.change-plan")
  writeLeanText(directory, "d1/review-playbook.json", JSON.stringify({ coverage: "Revised coverage after a source-backed correction.", findings: [] }))
  expect(stageInputSha(directory, "d1.change-plan")).not.toBe(before)
}, 120_000)

test("an invented planning quotation never enters the best-outline fallback", async () => {
  const directory = matter()
  const invalid = { sources: [{ source_id: SOURCE, quote: "This policy passage was never supplied." }], reason: "Purported applicable guidance" }
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: () => `Acting for: not stated\n# Supply note\nPlaybook review: ${JSON.stringify(invalid)}\n## Delivery\nState delivery terms` })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("failed")
  expect(state.stop_reason).toContain("quote its scope or position exactly")
  expect(readLeanText(directory, "plan.json")).toBeUndefined()
  expect(readLeanText(directory, "d1/draft.md")).toBeUndefined()
  expect(model.count(AGENT["review-playbook"])).toBe(0)
}, 120_000)
