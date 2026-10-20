import { afterEach, expect, test } from "bun:test"
import { readParagraphs } from "./checks"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { readLeanText, readQuestions, writeLeanText, writeQuestions } from "./state"
import { AGENT, DRAFT, NO_FINDINGS, cleanupMatters, drafter, happyPath, matter, noChanges, pidOf, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const quiet = { ...happyPath, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges }

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
  expect(model.prompts.map(row => row.agent)).toEqual([AGENT.discrepancy, AGENT.intake, AGENT.advise])
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
    : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.prompts.filter(row => row.agent === AGENT.intake)[1]!.text).toContain("already has a lawyer's answer")
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: "q-existing", status: "answered" })])
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
