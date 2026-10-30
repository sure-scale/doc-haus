import { expect, test } from "bun:test"
import type { Citation } from "../src/api/opencode/parts"
import { citationLabel, readReviewAnswer } from "../src/api/review-answer"
import type { Part } from "../src/api/session-view"

test("grid uses final answer and its exact earlier source rather than progress or top search hit", () => {
  const cite = (section: string, excerpt: string): Citation => ({ documentName: "Agreement", docPath: "/matter/agreement.docx", section, excerpt, charStart: 0, charEnd: 0, score: 1 })
  const messages = [
    { parts: [{ type: "text", text: "Finding the clause." }, { type: "tool", tool: "search-document", state: {
      status: "completed", metadata: { citations: [cite("A", "Other clause"), cite("B", "Relevant clause")] },
    } }] as Part[] },
    { parts: [{ type: "text", text: "The extracted term.\nSource: §B" }] as Part[] },
  ]
  expect(readReviewAnswer(messages, { name: "Agreement" })).toEqual({ text: "The extracted term.", citation: cite("B", "Relevant clause") })
  messages[1]!.parts = [{ type: "text", text: "Another term.\nSource: §C" }] as Part[]
  expect(readReviewAnswer(messages, { name: "Agreement" })).toEqual({ text: "Another term.", citation: undefined })
})

test("grid matches displayed structural headings only when the citation identity is unique", () => {
  const citation: Citation = { documentName: "document.docx", docPath: "/matter/document.docx", section: "12", structuralPath: "Doc(document) › §12 Fees", excerpt: "Fees", charStart: 10, charEnd: 40, score: 1 }
  const parts = [{ type: "tool", id: "tool", sessionID: "session", messageID: "message", tool: "search-document", state: { status: "completed", input: {}, output: "", title: "", metadata: { citations: [citation, citation] } } }] as Part[]
  const messages = [{ parts }, { parts: [{ type: "text", text: "The fee term.\nSource: §12 Fees" }] as Part[] }]
  expect(readReviewAnswer(messages, { name: "document.docx" }).citation).toEqual(citation)
  messages[0]!.parts = [{ type: "tool", id: "tool", sessionID: "session", messageID: "message", tool: "search-document", state: { status: "completed", input: {}, output: "", title: "", metadata: { citations: [citation, { ...citation, docPath: "/matter/another.docx" }] } } }] as Part[]
  expect(readReviewAnswer(messages, { name: "document.docx" }).citation).toBeUndefined()
})

// Two matter documents both have a Termination clause; the row is MSA.docx.
const termination = (documentName: string): Citation => ({ documentName, docPath: `/matter/${documentName}`, section: "Termination", excerpt: `${documentName} termination.`, charStart: 100, charEnd: 130, score: 1 })
const searched = (...citations: Citation[]) => ({ parts: [{ type: "tool", id: "tool", sessionID: "session", messageID: "message", tool: "search-document",
  state: { status: "completed", input: {}, output: "", title: "", metadata: { citations } } }] as Part[] })
const answered = { parts: [{ type: "text", text: "Either party on 30 days' notice.\nSource: §Termination" }] as Part[] }
const msa = { name: "MSA.docx", docPath: "/matter/MSA.docx" }

test("grid never cites another document's passage that shares the row's section label", () => {
  expect(readReviewAnswer([searched(termination("NDA.docx")), answered], msa).citation).toBeUndefined()
})

test("grid cites the row's own passage when another document shares its section label", () => {
  expect(readReviewAnswer([searched(termination("NDA.docx"), termination("MSA.docx")), answered], msa).citation).toEqual(termination("MSA.docx"))
})

test("a source label names the cited heading, with § only for numbered sections", () => {
  const label = (section: string, structuralPath?: string) => citationLabel({ section, structuralPath })
  expect(label("12", "Doc(Supply terms.docx) › §12 Governing law")).toBe("§12 Governing law")
  expect(label("7.2", "Doc(MSA.docx) › §7 Termination › §7.2 Notice")).toBe("§7.2 Notice")
  expect(label("LOGISTICS SERVICES AGREEMENT", "Doc(LSA.docx) › LOGISTICS SERVICES AGREEMENT")).toBe("LOGISTICS SERVICES AGREEMENT")
  expect(label("Schedule 1", "Doc(LSA.docx) › Schedule 1 › Rates")).toBe("Rates")
  expect(label("3", "Doc(LSA.docx)")).toBe("§ 3")
  expect(label("4")).toBe("§ 4")
  expect(label("§4")).toBe("§ 4")
  expect(label("Preamble")).toBe("Preamble")
  expect(label("Termination", "")).toBe("Termination")
})

test("a clause-sentence heading is cut at a word and keeps its number", () => {
  const heading = "§5 The Supplier shall deliver the Goods to the Delivery Location on the Delivery Date and shall bear the risk"
  const shown = citationLabel({ section: "5", structuralPath: `Doc(Supply.docx) › ${heading}` })
  expect(shown).toBe("§5 The Supplier shall deliver the Goods to the…")
  expect(heading.startsWith(shown.slice(0, -1))).toBe(true)
})
