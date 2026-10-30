import { expect, test } from "bun:test"
import { normaliseStarterQuestion, recommendStarter, REVIEW_STARTERS } from "../src/review-starters"

const compare = REVIEW_STARTERS.find((starter) => starter.example)!
const example = compare.example!

const text = (name: string) => example.documents.find((doc) => doc.name === name)?.text

test("every example citation quotes its document exactly at its offsets", () => {
  for (const { document, citation } of example.answers) {
    if (!citation) continue
    const body = text(document)!
    expect(citation.charStart).toBeGreaterThanOrEqual(0)
    expect(citation.charEnd).toBeGreaterThan(citation.charStart)
    expect(citation.charEnd).toBeLessThanOrEqual(body.length)
    expect(body.slice(citation.charStart, citation.charEnd)).toBe(citation.excerpt)
  }
})

test("example has exactly one answer per document and question", () => {
  const cells = example.answers.map((answer) => `${answer.document}#${answer.question}`)
  expect(new Set(cells).size).toBe(cells.length)
  expect(cells.length).toBe(example.documents.length * compare.questions.length)
  for (const answer of example.answers) {
    expect(text(answer.document)).toBeDefined()
    expect(Number.isInteger(answer.question)).toBe(true)
    expect(answer.question).toBeGreaterThanOrEqual(0)
    expect(answer.question).toBeLessThan(compare.questions.length)
  }
})

test("uncited answers read Not addressed with no label, and differs reasons stay short and unquoted", () => {
  for (const answer of example.answers) {
    if (!answer.citation) {
      expect(answer.answer).toBe("Not addressed")
      expect(answer.comparison).toBeUndefined()
    }
    if (answer.comparison === "differs") {
      expect(answer.reason!.length).toBeGreaterThan(0)
      expect(answer.reason!.length).toBeLessThanOrEqual(140)
      expect(answer.reason).not.toMatch(/["“”]/)
    } else expect(answer.reason).toBeUndefined()
  }
})

test("starter ids are unique and each starter's questions stay distinct once normalised", () => {
  expect(new Set(REVIEW_STARTERS.map((starter) => starter.id)).size).toBe(REVIEW_STARTERS.length)
  for (const starter of REVIEW_STARTERS) {
    const keys = starter.questions.map(normaliseStarterQuestion)
    for (const key of keys) expect(key.length).toBeGreaterThan(0)
    expect(new Set(keys).size).toBe(keys.length)
  }
})

test("recommendation counts only source documents", () => {
  expect(recommendStarter([{ role: "source" }, {}])).toEqual({ id: "compare-agreements-v1", reason: "2 documents in this matter" })
  expect(recommendStarter([{ role: "source" }, { role: "deliverable" }])).toEqual({ id: "document-overview-v1", reason: "1 source document in this matter" })
  expect(recommendStarter([{ role: "deliverable" }]).id).toBe("document-overview-v1")
  expect(recommendStarter([]).id).toBe("document-overview-v1")
})
