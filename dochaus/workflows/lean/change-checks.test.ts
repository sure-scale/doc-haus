import { expect, test } from "bun:test"
import { askedQuestions, questionIssues } from "./change-checks"
import type { ParagraphRow } from "./checks"
import type { Question } from "./state"

// The changes contract's "id" names an open or an answered question the reply
// repeats. A reworded repeat of an answered question, a fact's included, is
// not asked again; a repeat of an open one keeps its id.

const rows = new Map<string, ParagraphRow>([["p1", { pid: "p1", sha256: "x", text: "The goods are delivered to [open item: confirm the delivery address]." }]])
const settled: Question = {
  id: "q-address", stage: "changes", question: "What is the delivery address?", context: "", pid: "p1", document: "d1",
  status: "answered", answer: "Not known yet.", kind: "fact",
}
const pending: Question = { id: "q-date", stage: "changes", question: "When is the delivery due?", context: "", pid: "p1", document: "d1", status: "open", kind: "fact" }
const asked = new Map([settled, pending].map(question => [question.id, question]))

test("a reworded repeat naming an answered question passes the contract and is dropped", () => {
  const repeat = { id: "q-address", pid: "p1", question: "Please confirm the address for delivery.", kind: "fact", context: "Needed for the notice." }
  expect(questionIssues(repeat, 0, rows, asked)).toEqual([])
  expect(askedQuestions([repeat], asked, "d1")).toEqual([undefined])
})

test("a reworded repeat naming an open question keeps its id; a new question gets its own", () => {
  const [repeat, fresh] = askedQuestions([
    { id: "q-date", pid: "p1", question: "By which date must delivery happen?", kind: "fact" },
    { pid: "p1", question: "Who signs the delivery note?", kind: "fact" },
  ], asked, "d1")
  expect(repeat).toMatchObject({ id: "q-date", question: "By which date must delivery happen?", status: "open" })
  expect(fresh).toMatchObject({ question: "Who signs the delivery note?", status: "open" })
  expect([settled.id, pending.id]).not.toContain(fresh!.id)
})
