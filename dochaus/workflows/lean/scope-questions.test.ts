import { afterEach, expect, test } from "bun:test"
import { writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
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
