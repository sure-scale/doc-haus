import { expect, test } from "bun:test"
import { readReviewAnswer } from "../src/api/review-answer"
import type { Part } from "../src/api/session-view"

test("grid uses final answer and its exact earlier source rather than progress or top search hit", () => {
  const messages = [
    { parts: [{ type: "text", text: "Finding the clause." }, { type: "tool", tool: "search-document", state: {
      status: "completed", metadata: { citations: [{ section: "A", excerpt: "Other clause" }, { section: "B", excerpt: "Relevant clause" }] },
    } }] as Part[] },
    { parts: [{ type: "text", text: "The extracted term.\nSource: §B" }] as Part[] },
  ]
  expect(readReviewAnswer(messages)).toEqual({ text: "The extracted term.", citation: { section: "B", excerpt: "Relevant clause" } })
  messages[1]!.parts = [{ type: "text", text: "Another term.\nSource: §C" }] as Part[]
  expect(readReviewAnswer(messages)).toEqual({ text: "Another term.", citation: undefined })
})

test("grid matches displayed structural headings only when the citation identity is unique", () => {
  const citation = { docPath: "/matter/document.docx", section: "12", structuralPath: "Doc(document) › §12 Fees", charStart: 10, charEnd: 40 }
  const parts = [{ type: "tool", tool: "search-document", state: { status: "completed", metadata: { citations: [citation, citation] } } }] as Part[]
  const messages = [{ parts }, { parts: [{ type: "text", text: "The fee term.\nSource: §12 Fees" }] as Part[] }]
  expect(readReviewAnswer(messages).citation).toEqual(citation)
  messages[0]!.parts = [{ type: "tool", tool: "search-document", state: { status: "completed", metadata: { citations: [citation, { ...citation, docPath: "/matter/another.docx" }] } } }] as Part[]
  expect(readReviewAnswer(messages).citation).toBeUndefined()
})
