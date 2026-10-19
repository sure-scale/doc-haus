import { afterAll, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { LegalPlugin } from "../plugin/legal/policy"
import { recordSessionAgent } from "../plugin/legal/session"
import { stagePrompt } from "../workflows/lean/prompts"
import { cleanupMatters, matter } from "../workflows/lean/test-support"

const workspace = mkdtempSync(path.join(tmpdir(), "drafting-presentation-"))
const originalWorkspace = process.env.WORKSPACE_ROOT
process.env.WORKSPACE_ROOT = workspace
mkdirSync(path.join(workspace, ".preferences"))
const preferences = path.join(workspace, ".preferences", "preferences.json")
const initial = { attorney: "A. Counsel", firm: "Counsel Practice", posture: "conservative", formality: "formal", detail: "detailed", dateFormat: "iso", numberStyle: "numerals", houseStyle: "Use active voice. Draft an action plan before revising documents.", process: "plan-first" }

afterAll(() => {
  cleanupMatters()
  rmSync(workspace, { recursive: true, force: true })
  if (originalWorkspace === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = originalWorkspace
})

async function systemFor(agent: string) {
  const hooks = await LegalPlugin({ directory: workspace })
  const sessionID = `presentation-${agent}`
  recordSessionAgent(sessionID, agent)
  const output = { system: [] as string[] }
  await hooks["experimental.chat.system.transform"]!({ sessionID } as never, output)
  return output.system.join("\n")
}

test("narration receives current presentation preferences without drafting authority or plan phases", async () => {
  writeFileSync(preferences, JSON.stringify(initial))
  const narrator = await systemFor("narrator")
  expect(narrator).toContain("Use formal legal drafting style")
  expect(narrator).toContain("Explain your reasoning in detail")
  expect(narrator).toContain("2026-06-11")
  expect(narrator).toContain("plain numerals")
  expect(narrator).toContain("Use active voice")
  expect(narrator).toContain("presentation only")
  expect(narrator).not.toContain("Documents are prepared by")
  expect(narrator).not.toContain("Draft conservatively")
  expect(narrator).not.toContain("<plan_first_drafting>")
  expect(narrator).not.toContain("Inventory. Review")

  const draft = await systemFor("lean-drafting-notes")
  expect(draft).toContain("Documents are prepared by A. Counsel of Counsel Practice")
  expect(draft).toContain("Draft conservatively")
  expect(draft).toContain("Use active voice")
  for (const agent of ["draft", "router"]) {
    const system = await systemFor(agent)
    expect(system).not.toContain("<drafting_preferences>")
    expect(system).not.toContain("Use active voice")
  }

  writeFileSync(preferences, JSON.stringify({ ...initial, formality: "plain", detail: "concise", houseStyle: "Use descriptive labels." }))
  const next = await systemFor("narrator")
  expect(next).toContain("Prefer plain-language drafting")
  expect(next).toContain("Keep explanations concise")
  expect(next).toContain("Use descriptive labels")
  expect(next).not.toContain("Explain your reasoning in detail")
  expect(next).not.toContain("Use active voice")
})

test("human-facing notes and narration defer style and detail to saved preferences", () => {
  const notes = stagePrompt(matter(), "finalize") + readFileSync(new URL("../agent/lean-drafting-notes.md", import.meta.url), "utf8")
  expect(notes).toContain("Follow the lawyer's explanation-detail, formality and house-style preferences")
  expect(notes).not.toContain("Write concise notes")
  expect(notes).not.toContain("Explain briefly what each drafted document does")
  const narrator = readFileSync(new URL("../agent/narrator.md", import.meta.url), "utf8")
  expect(narrator).toContain("explanation-detail, formality and house-style preferences")
  expect(narrator).not.toContain("one or two short, plain sentences")
  expect(narrator).not.toContain("calm, specific and brief")
})

// Answers and hand-backs restart a run waiting on questions (the answer route
// launches the run; it stays blocked while any question is open), so the
// narrator keeps Continue drafting for a run stopped with nothing waiting.
test("narration sends the lawyer to Continue drafting only for a run with no question waiting", () => {
  const narrator = readFileSync(new URL("../agent/narrator.md", import.meta.url), "utf8").replace(/\s+/g, " ")
  expect(narrator).toContain("Continue drafting is for a stopped run with no question waiting; a run waiting only on questions resumes by itself once every question is answered or left to the drafting.")
  expect(narrator).not.toContain("resume a stopped run with Continue drafting")
})
