import { afterEach, expect, test } from "bun:test"
import { leanSteps, parsePlan, plannedDocument, reviewStages, writeLeanPlan } from "./plan"
import { playbookSelectionIssues } from "./playbook-review"
import { packageReviewPrompt } from "./package-review"
import { changePlanPrompt } from "./change-pipeline"
import { runLean, stageInputSha } from "./runner"
import { LEAN_CHILD_PERMISSIONS } from "./stages"
import { readLeanText, writeLeanText } from "./state"
import { AGENT, DRAFT, SOURCE, SOURCE_TEXT, cleanupMatters, drafter, happyPath, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)
const selection = { sources: [{ source_id: SOURCE, quote: SOURCE_TEXT }], reason: "The supplied drafting guidance applies to this supply note." }
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
      return "Playbook coverage: supplied terms require the stated delivery period. Correction required at the delivery paragraph."
    }
    if (input.agent === AGENT["change-plan"]) {
      expect(checked).toBe(true)
      expect(input.text).toContain("review-playbook:\nPlaybook coverage:")
    }
    if (input.agent === AGENT["package-review"]) expect(input.text).toContain("Playbook coverage:")
    return base.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions })
  expect(state.status).toBe("completed")
  expect(state.stages["d1.review-playbook"]).toMatchObject({ status: "done", outputs: ["d1/review-playbook.md"] })
  expect(base.sessions.find(session => session.agent === AGENT["review-playbook"])?.permissions).toEqual(LEAN_CHILD_PERMISSIONS)
  expect(plannedDocument(directory, "d1")?.playbook_review).toEqual(selection)
  expect(changePlanPrompt(directory, "d1")).toContain("Playbook coverage:")
  expect(packageReviewPrompt(directory)).toContain("Playbook coverage:")
  const before = stageInputSha(directory, "d1.change-plan")
  writeLeanText(directory, "d1/review-playbook.md", "Revised coverage after a source-backed correction.")
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
