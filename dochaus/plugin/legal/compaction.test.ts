import { afterEach, expect, test } from "bun:test"
import { writeLegalCompactionPrompt } from "./compaction"
import { readLeanState, writeLeanState, writeLeanText } from "../../workflows/lean/state"
import { cleanupMatters, matter } from "../../workflows/lean/test-support"

afterEach(cleanupMatters)

test("lean compaction references native durable state and preserves the repair count and proposal posture", () => {
  const directory = matter()
  const state = readLeanState(directory)!
  writeLeanState(directory, { ...state, status: "blocked", stop_reason: "Waiting for a material client decision.", stages: { ...state.stages, "d1.change-plan": { ...state.stages["d1.change-plan"]!, status: "blocked" } } })
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 4, substantive_attempts: 4, items: [], requested_at: new Date().toISOString() }))
  const output = { context: [] as string[], prompt: "" }
  writeLegalCompactionPrompt(directory, output)
  expect(output.prompt).toContain(".dochaus/lean/state.json")
  expect(output.prompt).toContain('"substantive_attempts":4')
  expect(output.prompt).toContain("d1.change-plan")
  expect(output.prompt).toContain("Waiting for a material client decision.")
  expect(output.prompt).not.toContain("not a drafting matter")
})

test("completed workflow handoff preserves the current operation and names only existing durable records", () => {
  const directory = matter()
  const state = readLeanState(directory)!
  writeLeanState(directory, { ...state, status: "completed", stop_reason: null })
  writeLeanText(directory, "questions.json", JSON.stringify({ questions: [] }))
  const output = { context: [] as string[], prompt: "" }
  writeLegalCompactionPrompt(directory, output)
  const template = output.prompt.split("\n<template>\n")[1]!.split("\n</template>")[0]!
  expect(template).toContain(".dochaus/lean/state.json: current workflow state")
  expect(template).toContain(".dochaus/lean/questions.json: saved questions and human answers")
  expect(template).not.toContain(".dochaus/lean/repairs.json:")
})
