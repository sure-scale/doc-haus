import { afterEach, expect, test } from "bun:test"
import { readLeanDocuments } from "./deliverables"
import { readLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { writeLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

test("one planned substantive body supplies a requested body, tracked redline and accompanying notes", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", "Prepare a revised supply note with its tracked redline and accompanying drafting notes.")
  const prompt = stagePrompt(directory, "d1.draft", { outline: true })
  expect(prompt).toContain("workflow renders the final Word document")
  expect(prompt).toContain("generates its tracked redline")
  expect(prompt).toContain("one final drafting-notes memo")
  const model = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(readLeanPlan(directory)?.documents).toHaveLength(1)
  expect(readLeanDocuments(directory)?.documents.map(row => row.kind)).toEqual(["body", "redline", "drafting_notes", "drafting_notes_markdown"])
  expect(model.sessions.filter(row => row.agent === AGENT.finalize)).toHaveLength(1)
  expect(model.prompts.some(row => row.agent === AGENT.draft && /Draft part .* of "Tracked/.test(row.text))).toBe(false)
}, 120_000)
