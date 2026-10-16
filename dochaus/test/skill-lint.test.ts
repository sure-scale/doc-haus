import { readFileSync } from "node:fs"
import path from "node:path"
import { describe, expect, test } from "bun:test"

const scriptPath = path.resolve(import.meta.dir, "..", "script", "lint-legal-runtime.ts")

describe("lint-legal-runtime skill wiring", () => {
  test("script imports lintSkill from skill-authoring-rules", () => {
    const source = readFileSync(scriptPath, "utf8")
    expect(source).toContain("from \"../lib/skill-authoring-rules\"")
    expect(source).toContain("lintSkill")
  })

  test("script invokes lintSkill on skill files in the per-file loop", () => {
    const source = readFileSync(scriptPath, "utf8")
    expect(source).toMatch(/lintSkill\(/)
    expect(source).toContain("isSkillFile")
    expect(source).toContain("skillInputFor")
  })
})
