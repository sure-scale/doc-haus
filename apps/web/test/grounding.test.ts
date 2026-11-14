import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import type { Citation } from "../src/api/opencode/parts"
import type { Part } from "../src/api/session-view"
import { answerGrounding } from "../src/components/chat/grounding"
import { readTurn, toTurns } from "../src/components/chat/turn-model"
import GroundingBadge from "../src/components/GroundingBadge"

// The grounding state is structural: model-authored `[Document § section]`
// markers bind a statement to a quotation the legal plugin verified. These cases
// pin the states, the refusal to promote retrieval hits, and the honest limits.

const EXCERPT = "The Supplier shall deliver the synthetic cartons within thirty days of the order."
const CITATION: Citation = { documentName: "MSA.docx", docPath: "/matter/MSA.docx", section: "7.2", excerpt: EXCERPT, charStart: 0, charEnd: EXCERPT.length }
const HIT: Citation = { ...CITATION, section: "9.1", excerpt: "Either party may terminate for material breach.", score: 0.82 }

function tool(name: string, metadata: Record<string, unknown> = {}, status: "completed" | "error" = "completed"): Part {
  return {
    type: "tool", id: `tool-${name}`, sessionID: "ses", messageID: "msg", tool: name,
    state: status === "completed"
      ? { status, input: {}, output: "ok", title: name, metadata }
      : { status, input: {}, error: "failed", title: name, metadata },
  }
}

function text(value: string): Part {
  return { type: "text", id: "answer", sessionID: "ses", messageID: "msg", text: value }
}

const grounding = (value: string, citations: Citation[] = [], parts: Part[] = []) =>
  answerGrounding({ text: value, citations, parts: [...parts] })

test("a greeting with no evidence work is chat and shows no badge", () => {
  const result = grounding("Hello. How can I help with this matter?")
  expect(result.state).toBe("chat")
  expect(renderToStaticMarkup(createElement(GroundingBadge, { grounding: result }))).toBe("")
})

test("every substantive statement linked to a verified quotation is grounded", () => {
  const answer = [
    "The supplier must deliver the cartons within thirty days of each order [MSA § 7.2].",
    "Delivery timing is therefore measured from the order date, not from signature [MSA § 7.2].",
  ].join("\n\n")
  const result = grounding(answer, [CITATION], [tool("cite", { citations: [CITATION] })])
  expect(result.state).toBe("grounded")
  expect(result.supported).toBe(2)
  expect(result.unlinked).toEqual([])
  const html = renderToStaticMarkup(createElement(GroundingBadge, { grounding: result }))
  expect(html).toContain("Grounded")
  expect(html).toContain("2 statements linked")
})

test("a retrieval hit is not span-verified support", () => {
  const answer = "Either party may terminate for material breach [MSA § 9.1]."
  const result = grounding(answer, [HIT], [tool("search-document", { citations: [HIT] })])
  expect(result.state).toBe("abstain")
  expect(result.citations).toBe(0)
  expect(renderToStaticMarkup(createElement(GroundingBadge, { grounding: result }))).toContain("No verified support")
})

test("a substantive block with no marker makes the answer limited", () => {
  const answer = [
    "The supplier must deliver the cartons within thirty days of each order [MSA § 7.2].",
    "In practice this gives the customer very little time to inspect a delivery before the next order is placed.",
  ].join("\n\n")
  const result = grounding(answer, [CITATION], [tool("cite", { citations: [CITATION] })])
  expect(result.state).toBe("limited")
  expect(result.supported).toBe(1)
  expect(result.unlinked).toHaveLength(1)
  expect(result.unlinked[0]).toStartWith("In practice this gives")
})

test("verified citations with no claim mapping cannot be grounded", () => {
  const answer = "The agreement contains a thirty day delivery window and a material breach termination right."
  const result = grounding(answer, [CITATION, HIT], [tool("cite", { citations: [CITATION] })])
  expect(result.state).toBe("limited")
  expect(result.supported).toBe(0)
  expect(result.citations).toBe(1)
  const html = renderToStaticMarkup(createElement(GroundingBadge, { grounding: result }))
  expect(html).toContain("does not link them to its statements")
})

test("a short uncited line stating an amount is still material", () => {
  const answer = [
    "Delivery is due within thirty days of each order [MSA § 7.2].",
    "The agreed figure is $42,500.",
  ].join("\n\n")
  const result = grounding(answer, [CITATION], [tool("cite", { citations: [CITATION] })])
  expect(result.state).toBe("limited")
  expect(result.unlinked).toEqual(["The agreed figure is $42,500."])
})

test("evidence was consulted but no verified quotation supports the answer: abstain", () => {
  const answer = "The agreement has no cap on liability for the supplier."
  const result = grounding(answer, [], [tool("search-document")])
  expect(result.state).toBe("abstain")
})

test("a refused quotation is never upgraded to grounded", () => {
  const answer = "The supplier must deliver the cartons within thirty days of each order [MSA § 7.2]."
  const result = grounding(answer, [CITATION], [tool("cite", { citations: [CITATION], rejectedCitations: [{ documentName: "MSA.docx", section: "" }] })])
  expect(result.state).toBe("limited")
  expect(result.rejected).toBe(1)
})

test("a quoted passage with no verified quotation behind it makes the answer limited", () => {
  const answer = "The clause reads: “The Supplier shall deliver the synthetic cartons within thirty days of the order.” [MSA § 7.2]"
  const unsupported = "“Either party may terminate for material breach on written notice to the other party.”"
  const withUnverified = `${answer}\n\n${unsupported}`
  const result = grounding(withUnverified, [CITATION], [tool("cite", { citations: [CITATION] })])
  expect(result.state).toBe("limited")
  expect(result.unverifiedQuotes).toBe(1)
})

test("a marker naming the wrong section is unbound and cannot ground its block", () => {
  const answer = "The supplier must deliver the cartons within thirty days of each order [MSA § 12.4]."
  const result = grounding(answer, [CITATION], [tool("cite", { citations: [CITATION] })])
  expect(result.state).toBe("abstain")
  expect(result.unbound).toBe(1)
})

test("a correct passage quoted from the wrong instrument does not bind", () => {
  const answer = "The supplier must deliver the cartons within thirty days of each order [Amendment.docx § 7.2]."
  const result = grounding(answer, [CITATION], [tool("cite", { citations: [CITATION] })])
  expect(result.state).toBe("abstain")
  expect(result.unbound).toBe(1)
})

test("live and stored turns compute the same grounding state", () => {
  const parts = [
    text("The supplier must deliver the cartons within thirty days of each order [MSA § 7.2]."),
    tool("cite", { citations: [CITATION] }),
  ]
  const live = readTurn(new Map(parts.map(part => [part.id, part])), new Map([["msg", "assistant"]]), "", "ses")
  const stored = toTurns([{ info: { id: "msg", role: "assistant" }, parts: JSON.parse(JSON.stringify(parts)) }], "")[0]!
  expect(stored.grounding).toEqual(live.grounding)
  expect(stored.grounding?.state).toBe("grounded")
})
