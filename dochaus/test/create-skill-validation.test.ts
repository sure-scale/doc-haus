import { describe, expect, test } from "bun:test"
import { lintSkill } from "../lib/skill-authoring-rules"
import type { SkillInput } from "../lib/skill-authoring-rules"

// The create-skill and update-skill tools call lintSkill on their args before
// POSTing to ingest. We test the policy by exercising lintSkill directly with
// the same input shape the tools build, plus a regression check that the tool
// source imports and calls lintSkill before the fetch.

import { readFileSync } from "node:fs"
import path from "node:path"

describe("create-skill / update-skill validation policy", () => {
  test("lintSkill rejects the inputs the tools must reject", () => {
    const badName: SkillInput = {
      name: "Bad_Name",
      path: "",
      frontmatter: "",
      description: "",
      body: "",
      builtins: new Set(),
    }
    const findings = lintSkill(badName)
    const errors = findings.filter((f) => f.severity === "error")
    expect(errors.length).toBeGreaterThan(0)
  })

  test("create-skill.ts imports lintSkill and calls it before fetch", () => {
    const source = readFileSync(
      path.resolve(import.meta.dir, "..", "tool", "create-skill.ts"),
      "utf8",
    )
    expect(source).toContain("skill-authoring-rules")
    expect(source).toMatch(/lintSkill\(/)
    const fetchIdx = source.indexOf("fetch(")
    const lintIdx = source.indexOf("lintSkill(")
    expect(fetchIdx).toBeGreaterThan(-1)
    expect(lintIdx).toBeGreaterThan(-1)
    expect(lintIdx).toBeLessThan(fetchIdx)
  })

  test("update-skill.ts imports lintSkill and calls it before fetch", () => {
    const source = readFileSync(
      path.resolve(import.meta.dir, "..", "tool", "update-skill.ts"),
      "utf8",
    )
    expect(source).toContain("skill-authoring-rules")
    expect(source).toMatch(/lintSkill\(/)
    const fetchIdx = source.indexOf("fetch(")
    const lintIdx = source.indexOf("lintSkill(")
    expect(fetchIdx).toBeGreaterThan(-1)
    expect(lintIdx).toBeGreaterThan(-1)
    expect(lintIdx).toBeLessThan(fetchIdx)
  })
})
