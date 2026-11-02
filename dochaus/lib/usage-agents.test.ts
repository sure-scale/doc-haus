import { expect, test } from "bun:test"
import { usageAgent, usageAgentKey, WORD_SUGGESTIONS_AGENT } from "./usage-agents"

test("background agents carry their Settings names and group", () => {
  expect(usageAgent("chronology-assess")).toEqual({ key: "chronology-assess", label: "Suggest a chronology", group: "background" })
  expect(usageAgent("review-compare").group).toBe("background")
  expect(usageAgent("narrator").label).toBe("Narrate drafting progress")
})

test("Word suggestions are told apart from Q&A by session metadata", () => {
  expect(usageAgentKey("qa", { metadata: { kind: "suggestions" } })).toBe(WORD_SUGGESTIONS_AGENT)
  expect(usageAgent("qa", { metadata: { kind: "suggestions" } })).toEqual({ key: WORD_SUGGESTIONS_AGENT, label: "Suggest ideas in Word", group: "background" })
  expect(usageAgent("qa", { metadata: { kind: "chat" } })).toEqual({ key: "qa", label: "Questions", group: "work" })
})

test("other agents are work with a readable name", () => {
  expect(usageAgent("lean-advise").label).toBe("Question check")
  expect(usageAgent("lean-draft-section")).toEqual({ key: "lean-draft-section", label: "Draft section", group: "work" })
  expect(usageAgent(undefined).key).toBe("unknown")
})
