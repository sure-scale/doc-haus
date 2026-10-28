import { expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { LegalPlugin } from "../plugin/legal/policy"
import { recordSessionAgent } from "../plugin/legal/session"
import { stageDraftingPreferences } from "../workflows/lean/stages"

test("firm practice areas reach substantive and library agents and update lean inputs", async () => {
  const previous = process.env.WORKSPACE_ROOT
  const workspace = mkdtempSync(path.join(tmpdir(), "practice-areas-"))
  const matter = path.join(workspace, "matter")
  const library = path.join(workspace, "library")
  mkdirSync(matter)
  mkdirSync(library)
  mkdirSync(path.join(workspace, ".preferences"))
  const preferences = path.join(workspace, ".preferences", "preferences.json")
  process.env.WORKSPACE_ROOT = workspace
  const save = (practiceAreas: string[]) => writeFileSync(preferences, JSON.stringify({ practiceAreas }))
  const system = async (agent: string, directory = matter) => {
    const hooks = await LegalPlugin({ directory } as Parameters<typeof LegalPlugin>[0])
    const sessionID = `practice-areas-${agent}`
    recordSessionAgent(sessionID, agent)
    const output = { system: [] as string[] }
    await hooks["experimental.chat.system.transform"]!({ sessionID } as never, output)
    return output.system.join("\n")
  }
  try {
    save(["Employment", "Corporate"])
    for (const agent of ["lean-review", "legal-review"]) {
      const context = await system(agent)
      expect(context).toContain('["Employment","Corporate"]')
      expect(context).toContain("Do not assume any area is the subject of this matter")
    }
    for (const builder of ["playbook-importer", "template-builder", "skill-builder", "agent-builder", "workflow-builder"]) {
      expect(await system(builder, library)).toContain('["Employment","Corporate"]')
      expect(await system(builder)).toContain('["Employment","Corporate"]')
    }
    expect(await system("router")).not.toContain("<practice_areas>")
    for (const agent of ["lean-checklist-extract", "lean-checklist-check"]) {
      expect(await system(agent)).not.toContain("<practice_areas>")
    }
    const before = stageDraftingPreferences(matter)
    save(["Corporate"])
    expect(stageDraftingPreferences(matter)).not.toBe(before)
    expect(await system("lean-review")).toContain('["Corporate"]')
    save([])
    expect(await system("legal-review")).not.toContain("<practice_areas>")
  } finally {
    if (previous === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previous
    rmSync(workspace, { recursive: true, force: true })
  }
})
