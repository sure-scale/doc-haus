import { expect, test } from "bun:test"
import { readdirSync } from "node:fs"
import path from "node:path"
import { OTHER_WORK, usageAgent, usageAgentKey, WORD_SUGGESTIONS_AGENT } from "./usage-agents"

test("background agents resolve to the background group with non-empty labels", () => {
  for (const key of ["chronology-assess", "review-compare", "narrator"]) {
    const resolved = usageAgent(key)
    expect(resolved.key).toBe(key)
    expect(resolved.group).toBe("background")
    expect(resolved.label).not.toBe(OTHER_WORK)
    expect(resolved.label).not.toBe(key)
  }
})

test("Word suggestions are told apart from Q&A by session metadata", () => {
  expect(usageAgentKey("qa", { metadata: { kind: "suggestions" } })).toBe(WORD_SUGGESTIONS_AGENT)
  expect(usageAgent("qa", { metadata: { kind: "suggestions" } })).toMatchObject({ key: WORD_SUGGESTIONS_AGENT, group: "background" })
  expect(usageAgent("qa", { metadata: { kind: "chat" } })).toMatchObject({ key: "qa", group: "work" })
  expect(usageAgent("qa", { metadata: { kind: "suggestions" } }).label).not.toBe(usageAgent("qa", { metadata: { kind: "chat" } }).label)
})

test("a drafting run's steps group into a few plain tasks", () => {
  expect(usageAgent("lean-placement")).toMatchObject({ key: "lean-placement", group: "work" })
  expect(usageAgent("lean-placement").label).toBe(usageAgent("lean-spine-level").label)
  expect(usageAgent("draft").label).toBe(usageAgent("lean-draft").label)
  expect(new Set([
    usageAgent("lean-placement").label,
    usageAgent("lean-research-review").label,
    usageAgent("lean-draft").label,
    usageAgent("lean-change-review").label,
    usageAgent("lean-checklist-check").label,
  ]).size).toBe(5)
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
