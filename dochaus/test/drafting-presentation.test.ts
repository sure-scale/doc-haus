import { afterAll, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { LegalPlugin } from "../plugin/legal/policy"
import { recordSessionAgent } from "../plugin/legal/session"

const workspace = mkdtempSync(path.join(tmpdir(), "drafting-presentation-"))
const originalWorkspace = process.env.WORKSPACE_ROOT
process.env.WORKSPACE_ROOT = workspace
mkdirSync(path.join(workspace, ".preferences"))
const preferences = path.join(workspace, ".preferences", "preferences.json")
const initial = { attorney: "A. Counsel", firm: "Counsel Practice", posture: "conservative", formality: "formal", detail: "detailed", dateFormat: "iso", numberStyle: "numerals", houseStyle: "Use active voice. Draft an action plan before revising documents.", process: "plan-first" }

afterAll(() => {
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
