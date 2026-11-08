import { expect, test } from "bun:test"
import { readdirSync } from "node:fs"
import path from "node:path"
import { OTHER_WORK, usageAgent, usageAgentKey, WORD_SUGGESTIONS_AGENT } from "./usage-agents"

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

test("a drafting run's steps group into a few plain tasks", () => {
  expect(usageAgent("lean-placement")).toEqual({ key: "lean-placement", label: "Planning", group: "work" })
  expect(usageAgent("lean-spine-level").label).toBe("Planning")
  expect(usageAgent("lean-research-review").label).toBe("Research")
  expect(usageAgent("draft").label).toBe("Drafting")
  expect(usageAgent("lean-draft").label).toBe("Drafting")
  expect(usageAgent("lean-change-review").label).toBe("Review")
  expect(usageAgent("lean-checklist-check").label).toBe("Checks")
})

test("a name with no plain label reads as other work, never its internal id", () => {
  expect(usageAgent("lean-synthetic-step")).toEqual({ key: "lean-synthetic-step", label: OTHER_WORK, group: "work" })
  expect(usageAgent(undefined)).toEqual({ key: "unknown", label: OTHER_WORK, group: "work" })
})

test("every shipped agent has a plain name, and every drafting step a task", () => {
  const shipped = ["agent", "mode"].flatMap((dir) =>
    readdirSync(path.join(import.meta.dir, "..", dir)).filter((file) => file.endsWith(".md")).map((file) => file.slice(0, -3)))
  expect(shipped.filter((name) => name.startsWith("lean-")).length).toBeGreaterThan(0)
  for (const name of shipped) {
    const { label } = usageAgent(name)
    expect({ name, label }).not.toEqual({ name, label: OTHER_WORK })
    if (name.startsWith("lean-")) expect({ name, label }).toEqual({ name, label: expect.stringMatching(/^(Planning|Research|Drafting|Review|Checks)$/) })
  }
})
