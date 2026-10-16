import { afterEach, expect, test } from "bun:test"
import { readParagraphs } from "./checks"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { readLeanText, readQuestions, writeLeanText, writeQuestions } from "./state"
import { AGENT, DRAFT, cleanupMatters, drafter, happyPath, matter, noChanges, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const quiet = { ...happyPath, [AGENT.review]: () => "No changes needed.", [AGENT.changes]: noChanges }

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
    { question: "Which governing law should the draft use?", context: "The requested law conflicts with the instruction to retain the existing clause." },
    { question: "Which governing law should the draft use?", context: "The same decision affects the whole document." },
  ] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  expect(model.prompts.map(row => row.agent)).toEqual([AGENT.discrepancy, AGENT.intake])
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
    ? JSON.stringify({ questions: [{ question: "Which law should govern?", context: "Legal basis." }] })
    : JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms" }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.prompts.filter(row => row.agent === AGENT.intake)[1]!.text).toContain("already has a lawyer's answer")
  expect(readQuestions(directory)).toEqual([expect.objectContaining({ id: "q-existing", status: "answered" })])
}, 120_000)

test("a changed answer during review rebuilds the outline, draft and review before delivery", async () => {
  const directory = matter()
  writeQuestions(directory, [{ id: "q-law", stage: "draft", question: "Which law should govern?", context: "Legal basis.", status: "answered", answer: "England and Wales law." }])
  let changed = false
  let law = "England and Wales"
  const model = scriptedModel({ ...quiet,
    [AGENT.draft]: (prompt, index) => {
      if (prompt.includes("<answers>")) law = prompt.includes("Singapore law.") ? "Singapore" : "England and Wales"
      return drafter(`# Supply note\n\n## Governing law\n\nThis note is governed by ${law} law.\n`)(prompt, index)
    },
    [AGENT.review]: () => {
      if (!changed) {
        changed = true
        enqueueLeanRequest(directory, { kind: "answer", question_id: "q-law", answer: "Singapore law." })
      }
      return "No changes needed."
    },
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.review)).toBe(2)
  const outlines = model.prompts.filter(row => row.agent === AGENT.draft && row.text.includes("Reply now with the outline"))
  expect(outlines).toHaveLength(2)
  expect(outlines[1]!.text).toContain("Singapore law.")
  expect(outlines[1]!.text).not.toContain("England and Wales law.")
  expect(readParagraphs(directory, "d1", "paragraphs-applied.json").some(row => row.text === "This note is governed by Singapore law.")).toBe(true)
}, 120_000)
