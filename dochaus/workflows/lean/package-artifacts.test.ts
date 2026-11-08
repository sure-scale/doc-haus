import { afterEach, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { readLeanDocuments } from "./deliverables"
import { focusedNotesInputs } from "./focused-notes"
import { readLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { writeLeanText } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

test("one planned substantive body supplies a requested body, tracked redline and accompanying notes", async () => {
  const directory = matter()
  writeLeanText(directory, "request.md", "Prepare a revised supply note with its tracked redline and accompanying drafting notes.")
  const prompt = stagePrompt(directory, "d1.draft", { outline: true })
  expect(prompt).toContain("workflow renders the final Word document")
  expect(prompt).toContain("generates its tracked redline")
  expect(prompt).toContain("one final drafting-notes memo")
  const model = scriptedModel(happyPath)
  const deliverables = registering()
  expect((await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })).status).toBe("completed")
  expect(readLeanPlan(directory)?.documents).toHaveLength(1)
  const documents = readLeanDocuments(directory)!.documents
  expect(documents.map(row => row.kind)).toEqual(["body", "redline", "drafting_notes", "drafting_notes_markdown"])
  documents.forEach(row => expect(existsSync(path.join(directory, row.path))).toBe(true))
  expect(documents.filter(row => row.deliverable !== null)).toHaveLength(3)
  documents.filter(row => row.deliverable !== null).forEach(row => expect(existsSync(path.join(directory, row.deliverable!))).toBe(true))
  expect(model.sessions.filter(row => row.agent === AGENT.finalize)).toHaveLength(focusedNotesInputs(directory).length)
  expect(model.prompts.some(row => row.agent === AGENT.draft && /Draft part .* of "Tracked/.test(row.text))).toBe(false)
}, 120_000)
