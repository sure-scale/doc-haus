import { afterEach, expect, test } from "bun:test"
import { writeLegalCompactionPrompt } from "../../plugin/legal/compaction"
import { checkInstructions, confirmationQuestionID, INSTRUCTIONS_FILE } from "./instructions"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { answeredQuestionsText, readLeanText, readQuestions, writeLeanText, writeQuestions } from "./state"
import { AGENT, NO_FINDINGS, SOURCE, cleanupMatters, happyPath, matter, noChanges, scriptedModel, type Reply } from "./test-support"

afterEach(cleanupMatters)

const quiet = { ...happyPath, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges }
const REQUEST = "Draft a short supply note from the supply terms under Singapore law."
const LAW = { question: "Which law should govern the supply note?", kind: "instruction", context: "The governing law clause depends on it." }
const READY = JSON.stringify({ ready: true, jurisdiction: "Singapore law", scope: "Supplied terms only" })
const DELIVERY = { question: "Which delivery period should the note state?", kind: "judgment", context: "The delivery clause depends on it." }

/** An advisor naming instruction `id` as deciding every question it is asked
 * that offers no options of its own; a code-written question gets no advice.
 * The job that checks that claim confirms it, quoting the lawyer's words,
 * unless `confirmed` is false. */
function deciding(id: string, quote: { source_id: string; quote: string }, confirmed = true): Reply {
  return prompt => prompt.includes("<answer_if_settled>")
    ? JSON.stringify(confirmed ? { decides: true, quote: quote.quote } : { decides: false, reason: "The instruction is only a preference." })
    : prompt.includes("<options>")
    ? JSON.stringify({ unavailable: { reason: "The lawyer decides this." } })
    : JSON.stringify({ options: [{ text: "Follow the instruction" }, { text: "Another approach" }], recommendation: { option: 0, rationale: "The instruction decides it.", basis: [{ ...quote, relevance: "The lawyer's instruction for this work." }], decided_by: id } })
}

test("a question the request's instruction decides is settled, not asked, and intake runs again reading it", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", REQUEST)
  const law = { source_id: "request", quote: "under Singapore law" }
  const model = scriptedModel({
    ...quiet,
    [AGENT.instructions]: () => JSON.stringify({ items: [{ text: "Use Singapore law.", ...law }], documents: [] }),
    [AGENT.intake]: (_prompt, index) => index === 0 ? JSON.stringify({ questions: [LAW] }) : READY,
    [AGENT.advise]: deciding("i1", law),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(readLeanText(directory, INSTRUCTIONS_FILE)).toContain('- [i1] Use Singapore law. (request: "under Singapore law")')
  const [settled] = readQuestions(directory)
  expect(settled).toMatchObject({ status: "answered", answer_kind: "instruction", answer: "Follow the instruction", decided_by: { text: "Use Singapore law.", ...law } })
  // Intake reads the list and, on its second run, the settled answer.
  const intake = model.prompts.filter(row => row.agent === AGENT.intake)
  expect(intake).toHaveLength(2)
  expect(intake[0]!.text).toContain("- [i1] Use Singapore law.")
  expect(intake[1]!.text).toContain(`settled by the lawyer's instruction "Use Singapore law."`)
  // The settled answer is not the lawyer's words, so the list does not read it back.
  expect(answeredQuestionsText(directory, { settled: false })).toBe("")
  expect(model.prompts.filter(row => row.agent === AGENT.instructions)).toHaveLength(1)
}, 120_000)

test("a question the advisor says an instruction decides is asked when the check finds the choice still the lawyer's", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", REQUEST)
  const law = { source_id: "request", quote: "under Singapore law" }
  const model = scriptedModel({
    ...quiet,
    [AGENT.instructions]: () => JSON.stringify({ items: [{ text: "Use Singapore law.", ...law }], documents: [] }),
    [AGENT.intake]: () => JSON.stringify({ questions: [LAW] }),
    [AGENT.advise]: deciding("i1", law, false),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  // Asked with the advisor's recommendation, which no longer claims to settle it.
  const [asked] = readQuestions(directory)
  expect(asked).toMatchObject({ status: "open", advice: { recommendation: { option: "o1" } } })
  expect(asked!.advice).not.toHaveProperty("recommendation.decided_by")
}, 120_000)

test("a fact question is asked even when the advisor names an instruction", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", REQUEST)
  const law = { source_id: "request", quote: "under Singapore law" }
  const model = scriptedModel({
    ...quiet,
    [AGENT.instructions]: () => JSON.stringify({ items: [{ text: "Use Singapore law.", ...law }], documents: [] }),
    [AGENT.intake]: () => JSON.stringify({ questions: [{ ...LAW, kind: "fact" }] }),
    [AGENT.advise]: deciding("i1", law),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  expect(readQuestions(directory).map(row => row.status)).toEqual(["open"])
}, 120_000)

test("a document's directions wait for the lawyer's confirmation, then settle what they decide", async () => {
  const directory = matter()
  const delivery = { source_id: SOURCE, quote: "within thirty days of the order date" }
  const replies = {
    ...quiet,
    [AGENT.instructions]: () => JSON.stringify({ items: [{ text: "State delivery within thirty days of the order date.", ...delivery }], documents: [{ source_id: SOURCE, from: "the client", quote: delivery.quote }] }),
    [AGENT.intake]: ((_prompt, index) => index < 2 ? JSON.stringify({ questions: [DELIVERY] }) : READY) as Reply,
    [AGENT.advise]: deciding("i1", delivery),
  }
  const model = scriptedModel(replies)
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  // Only the confirmation is asked: the point it would decide is left to it.
  const asked = readQuestions(directory)
  expect(asked.map(row => [row.id, row.status, row.options?.map(option => option.text)])).toEqual([[confirmationQuestionID(SOURCE), "open", ["Follow it as my instructions", "Treat it as evidence only"]]])
  expect(readLeanText(directory, INSTRUCTIONS_FILE)).toContain(`## Awaiting the lawyer's confirmation: ${SOURCE}`)
  enqueueLeanRequest(directory, { kind: "answer", question_id: confirmationQuestionID(SOURCE), option_id: "o1" })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(readLeanText(directory, INSTRUCTIONS_FILE)).not.toContain("Awaiting")
  expect(readQuestions(directory).find(row => row.question === DELIVERY.question)).toMatchObject({ status: "answered", answer_kind: "instruction", decided_by: { source_id: SOURCE } })
}, 120_000)

test("the list keeps only instructions resting on verified words from the lawyer or a listed document", () => {
  const directory = matter()
  writeLeanText(directory, "request.md", REQUEST)
  const valid = { items: [{ text: "Use Singapore law.", source_id: "request", quote: "under Singapore law" }], documents: [] }
  expect(checkInstructions(directory, valid).list?.items).toEqual([{ id: "i1", text: "Use Singapore law.", source_id: "request", quote: "under Singapore law" }])
  // Words the lawyer did not write, a document not listed as giving directions,
  // and a designation the lawyer never made are each refused.
  expect(checkInstructions(directory, { items: [{ text: "Use English law.", source_id: "request", quote: "under English law" }], documents: [] }).issues).toHaveLength(1)
  expect(checkInstructions(directory, { items: [{ text: "Deliver in thirty days.", source_id: SOURCE, quote: "within thirty days" }], documents: [] }).issues[0]).toContain("a document listed in documents")
  expect(checkInstructions(directory, { items: [], documents: [{ source_id: SOURCE, from: "the client", quote: "within thirty days", designated_by: "follow the supply terms" }] }).issues).toHaveLength(1)
})

test("an instruction from the lawyer's answer is cited as that answer, verified against its words", () => {
  const directory = matter()
  writeLeanText(directory, "request.md", REQUEST)
  writeQuestions(directory, [{ id: "q-delivery", stage: "intake", question: DELIVERY.question, context: DELIVERY.context, kind: "judgment", status: "answered", answer: "Deliver within thirty days of the order." }])
  const item = (quote: string) => ({ items: [{ text: "Deliver within thirty days.", source_id: "request", quote }], documents: [] })
  expect(checkInstructions(directory, item("within thirty days")).list?.items[0]).toMatchObject({ source_id: "answer:q-delivery", quote: "within thirty days" })
  // Words in the request stay cited as the request; words in no answer are refused.
  expect(checkInstructions(directory, item("under Singapore law")).list?.items[0]?.source_id).toBe("request")
  expect(checkInstructions(directory, item("within forty days")).issues).toHaveLength(1)
})

test("the compaction handoff carries the confirmed instruction list as written", () => {
  const directory = matter()
  writeLeanText(directory, INSTRUCTIONS_FILE, "# The lawyer's instructions\n\n- [i1] Use Singapore law. (request: \"under Singapore law\")\n\n## Awaiting the lawyer's confirmation: notes.docx\n\n- [i2] Cap liability. (notes.docx: \"cap\")\n")
  const output: { context: string[]; prompt?: string } = { context: [] }
  writeLegalCompactionPrompt(directory, output)
  expect(output.prompt).toContain("## Lawyer's Instructions\n- [i1] Use Singapore law. (request: \"under Singapore law\")\n\n## Stage & Gate Status")
  expect(output.prompt).not.toContain("Cap liability")
})
