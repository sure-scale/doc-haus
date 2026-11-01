import { afterEach, expect, test } from "bun:test"
import { writeLeanPlan } from "./plan"
import { changePlanPrompt } from "./change-pipeline"
import { DIRECTED_BLANK, stagePrompt } from "./prompts"
import { writeLeanText, writeQuestions } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)

test("factual-record review conditions legal questions on the requested function", () => {
  const directory = matter()
  writeLeanText(directory, "request.md", "Prepare a factual internal approval record, not a new contract or legal opinion.")
  writeLeanPlan(directory, [{ title: "Approval record", purpose: "Record supplied transaction facts", parts: [{ heading: "Facts", scope: "Record supplied facts" }] }])
  for (const step of ["d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage"] as const) {
    const prompt = stagePrompt(directory, step)
    expect(prompt).toContain("factual internal approval record")
    expect(prompt).toContain("Ask for governing law only when an unresolved jurisdiction is necessary")
    expect(prompt).toContain("needs no governing-law question")
    expect(prompt).not.toContain("Where the request and sources point to no governing law, raise that once")
    expect(prompt).toContain("Test essential operative function and legal/source support strictly")
  }
})

test("reviews and revisions preserve answers and do not turn optional refinements into blockers", () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-decided", stage: "intake", question: "Which source controls?", context: "Material conflict", status: "answered", answer: "Use the purchaser's expressly confirmed entity." }])
  for (const step of ["d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.changes"] as const) {
    const prompt = stagePrompt(directory, step)
    expect(prompt).toContain("Use the purchaser's expressly confirmed entity.")
    expect(prompt).toContain("must not become a mandatory interview")
    expect(prompt).toContain("do not reopen the decision, rephrase it as a new question")
    expect(prompt).toContain("Do not ask the human to supply legal research conclusions")
    expect(prompt).toContain("A genuinely indispensable unresolved decision still requires clarification")
  }
  const changes = stagePrompt(directory, "d1.changes")
  expect(changes).toContain("a reviewer's suggested question is not automatically a blocker")
  expect(changes).toContain("not optional refinements or routine details")
  expect(changes).toContain("its updated paragraph must carry a new visible [open item: confirm ...] marker")
})

test("a blank the instructions say to leave is the lawyer's; other blanks are filled from the sources", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Price", scope: "The price" }] }])
  const prompts = {
    ...Object.fromEntries((["d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.changes", "finalize"] as const).map(step => [step, stagePrompt(directory, step)])),
    "d1.change-plan": changePlanPrompt(directory, "d1"),
    "d1.change-prepare": stagePrompt(directory, "d1.change-prepare", { contextOnly: true }),
  }
  for (const [step, prompt] of Object.entries(prompts)) expect([step, prompt.split(DIRECTED_BLANK).length - 1]).toEqual([step, 1])
  expect(DIRECTED_BLANK).toContain("expressly direct be left for the lawyer or client stays as directed until the lawyer's later instruction or answer fills it")
  expect(DIRECTED_BLANK).toContain("offered in the drafting notes as a candidate, never inserted or presented as decided")
  expect(DIRECTED_BLANK).toContain("Fill every other blank or bracket, including a template's own, from the sources as usual.")
})
