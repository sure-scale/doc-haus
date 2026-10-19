import { expect, test } from "bun:test"
import type { Citation } from "../src/api/opencode/parts"
import { readReviewAnswer } from "../src/api/review-answer"
import type { Part } from "../src/api/session-view"

test("grid uses final answer and its exact earlier source rather than progress or top search hit", () => {
  const cite = (section: string, excerpt: string): Citation => ({ documentName: "Agreement", docPath: "/matter/agreement.docx", section, excerpt, charStart: 0, charEnd: 0, score: 1 })
  const messages = [
    { parts: [{ type: "text", text: "Finding the clause." }, { type: "tool", tool: "search-document", state: {
      status: "completed", metadata: { citations: [cite("A", "Other clause"), cite("B", "Relevant clause")] },
    } }] as Part[] },
    { parts: [{ type: "text", text: "The extracted term.\nSource: §B" }] as Part[] },
  ]
  expect(readReviewAnswer(messages)).toEqual({ text: "The extracted term.", citation: cite("B", "Relevant clause") })
  messages[1]!.parts = [{ type: "text", text: "Another term.\nSource: §C" }] as Part[]
  expect(readReviewAnswer(messages)).toEqual({ text: "Another term.", citation: undefined })
})

test("grid matches displayed structural headings only when the citation identity is unique", () => {
  const citation: Citation = { documentName: "document.docx", docPath: "/matter/document.docx", section: "12", structuralPath: "Doc(document) › §12 Fees", excerpt: "Fees", charStart: 10, charEnd: 40, score: 1 }
  const parts = [{ type: "tool", id: "tool", sessionID: "session", messageID: "message", tool: "search-document", state: { status: "completed", input: {}, output: "", title: "", metadata: { citations: [citation, citation] } } }] as Part[]
  const messages = [{ parts }, { parts: [{ type: "text", text: "The fee term.\nSource: §12 Fees" }] as Part[] }]
  expect(readReviewAnswer(messages).citation).toEqual(citation)
  messages[0]!.parts = [{ type: "tool", id: "tool", sessionID: "session", messageID: "message", tool: "search-document", state: { status: "completed", input: {}, output: "", title: "", metadata: { citations: [citation, { ...citation, docPath: "/matter/another.docx" }] } } }] as Part[]
  expect(readReviewAnswer(messages).citation).toBeUndefined()
})
