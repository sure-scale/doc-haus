import { describe, expect, test } from "bun:test"
import { lintSkill, parseSkillFile } from "../lib/skill-authoring-rules"
import type { SkillInput } from "../lib/skill-authoring-rules"

function makeInput(overrides: Partial<SkillInput> = {}): SkillInput {
  return {
    name: "example-skill",
    path: "dochaus/skill/example-skill/SKILL.md",
    frontmatter: 'description: Example skill. Use when testing the lint rules.',
    description: "Example skill. Use when testing the lint rules.",
    body: ["# Example", "", "Short body."].join("\n"),
    references: [],
    builtins: new Set<string>(),
    ...overrides,
  }
}

describe("parseSkillFile", () => {
  test("extracts frontmatter, description, body", () => {
    const raw = [
      "---",
      "name: foo",
      "description: Foo skill. Use when fooing.",
      "---",
      "",
      "# Body",
      "Content.",
      "",
    ].join("\n")
    const parsed = parseSkillFile(raw)
    expect(parsed.frontmatter).toContain("name: foo")
    expect(parsed.description).toBe("Foo skill. Use when fooing.")
    expect(parsed.body).toContain("# Body")
  })

  test("returns empty description when frontmatter is absent", () => {
    const parsed = parseSkillFile("# No frontmatter\nBody.")
    expect(parsed.description).toBe("")
    expect(parsed.body).toContain("No frontmatter")
  })

  test("extracts bare-scalar description containing an apostrophe", () => {
    const raw = [
      "---",
      "name: foo",
      "description: Handles the matter's documents. Use when reviewing.",
      "---",
      "",
      "# Body",
      "",
    ].join("\n")
    expect(parseSkillFile(raw).description).toBe("Handles the matter's documents. Use when reviewing.")
  })

  test("extracts bare-scalar description containing embedded quotes", () => {
    const raw = [
      "---",
      "name: foo",
      'description: Answers "what is market" for common clauses. Use when judging.',
      "---",
      "",
      "# Body",
      "",
    ].join("\n")
    expect(parseSkillFile(raw).description).toBe('Answers "what is market" for common clauses. Use when judging.')
  })
})

describe("lintSkill", () => {
  test("valid skill returns no findings", () => {
    const findings = lintSkill(makeInput())
    expect(findings).toEqual([])
  })

  test("frontmatter-description-missing when description empty", () => {
    const findings = lintSkill(makeInput({ description: "", frontmatter: "" }))
    const finding = findings.find((f) => f.rule === "frontmatter-description-missing")
    expect(finding).toBeDefined()
    expect(finding?.severity).toBe("error")
  })

  test("description-trigger-missing when no Use when/whenever/before/after/while", () => {
    const findings = lintSkill(
      makeInput({ description: "A skill with no trigger phrase." }),
    )
    const trigger = findings.find((f) => f.rule === "description-trigger-missing")
    expect(trigger).toBeDefined()
    expect(trigger?.severity).toBe("error")
  })

  test("description passes with 'Use before filing'", () => {
    const findings = lintSkill(
      makeInput({ description: "Pre-flight skill. Use before filing a brief." }),
    )
    expect(findings.find((f) => f.rule === "description-trigger-missing")).toBeUndefined()
  })

  test("name-pattern rejects uppercase and underscores", () => {
    const findings = lintSkill(makeInput({ name: "Bad_Name" }))
    expect(findings.find((f) => f.rule === "name-pattern" && f.severity === "error")).toBeDefined()
  })

  test("name-pattern accepts lowercase-hyphenated", () => {
    const findings = lintSkill(makeInput({ name: "good-name-2" }))
    expect(findings.find((f) => f.rule === "name-pattern")).toBeUndefined()
  })

  test("name-collision-builtin when name in builtins", () => {
    const findings = lintSkill(
      makeInput({ name: "cite-check", builtins: new Set(["cite-check"]) }),
    )
    expect(findings.find((f) => f.rule === "name-collision-builtin")).toBeDefined()
  })

  test("name-playbook-reserved when name starts with playbook-", () => {
    const findings = lintSkill(makeInput({ name: "playbook-thing" }))
    expect(findings.find((f) => f.rule === "name-playbook-reserved")).toBeDefined()
  })

  test("references-nesting-too-deep when references/foo/bar/baz.md present", () => {
    const findings = lintSkill(
      makeInput({ references: ["references/foo/bar/baz.md"] }),
    )
    expect(findings.find((f) => f.rule === "references-nesting-too-deep")).toBeDefined()
  })

  test("references one level deep passes", () => {
    const findings = lintSkill(
      makeInput({ references: ["references/foo.md", "references/sub/bar.md"] }),
    )
    expect(findings.find((f) => f.rule === "references-nesting-too-deep")).toBeUndefined()
  })

  test("body-over-budget-hard fires over 1000 lines", () => {
    const longBody = "# Title\n\n" + "line\n".repeat(1001)
    const findings = lintSkill(makeInput({ body: longBody }))
    expect(findings.find((f) => f.rule === "body-over-budget-hard" && f.severity === "error")).toBeDefined()
  })

  test("body-over-budget-soft warns between 500 and 1000 lines", () => {
    const mediumBody = "# Title\n\n" + "line\n".repeat(600)
    const findings = lintSkill(makeInput({ body: mediumBody }))
    const soft = findings.find((f) => f.rule === "body-over-budget-soft")
    expect(soft).toBeDefined()
    expect(soft?.severity).toBe("warn")
    const hard = findings.find((f) => f.rule === "body-over-budget-hard")
    expect(hard).toBeUndefined()
  })

  test("duplicate-section-heading warns on repeated heading text", () => {
    const body = ["# Title", "", "## Same", "content", "", "## Same", "more", ""].join("\n")
    const findings = lintSkill(makeInput({ body }))
    const dup = findings.find((f) => f.rule === "duplicate-section-heading")
    expect(dup).toBeDefined()
    expect(dup?.severity).toBe("warn")
  })

  test("procedure-missing-completion warns when <procedure> lacks terminal state", () => {
    const body = [
      "# Skill",
      "",
      "Intro.",
      "",
      "<procedure>",
      "1. Do thing one.",
      "2. Do thing two.",
      "</procedure>",
      "",
      "Some closing line with no terminal vocabulary.",
      "",
    ].join("\n")
    const findings = lintSkill(makeInput({ body }))
    const missing = findings.find((f) => f.rule === "procedure-missing-completion")
    expect(missing).toBeDefined()
    expect(missing?.severity).toBe("warn")
  })

  test("procedure passes when terminal state word appears in last 30%", () => {
    const body = [
      "# Skill",
      "",
      "<procedure>",
      "1. Do thing.",
      "</procedure>",
      "",
      "The verdict is the deliverable is safe to file only when all checks pass.",
      "",
    ].join("\n")
    const findings = lintSkill(makeInput({ body }))
    expect(findings.find((f) => f.rule === "procedure-missing-completion")).toBeUndefined()
  })

  test("multiple findings returned in one call", () => {
    const findings = lintSkill(
      makeInput({
        name: "Bad_Name",
        description: "",
        frontmatter: "",
      }),
    )
    expect(findings.length).toBeGreaterThanOrEqual(2)
    const rules = new Set(findings.map((f) => f.rule))
    expect(rules.has("name-pattern")).toBe(true)
    expect(rules.has("frontmatter-description-missing")).toBe(true)
  })
})

describe("lintProseFrequency", () => {
  test("returns no findings for plain high-frequency prose", () => {
    const findings = lintSkill(makeInput({ body: "# Title\n\nUse the tool to find the answer before you start.\n" }))
    expect(findings.find((f) => f.rule === "prose-frequency-plainness")).toBeUndefined()
  })

  test("flags 'utilize' as lower-frequency than 'use'", () => {
    const findings = lintSkill(makeInput({ body: "# Title\n\nUtilize the tool to read the file.\n" }))
    const hit = findings.find((f) => f.rule === "prose-frequency-plainness")
    expect(hit).toBeDefined()
    expect(hit?.severity).toBe("warn")
    expect(hit?.message.toLowerCase()).toContain("utilize")
  })

  test("flags the wordy phrase 'in order to'", () => {
    const findings = lintSkill(makeInput({ body: "In order to start, read the file.\n" }))
    expect(findings.find((f) => f.rule === "prose-frequency-plainness")).toBeDefined()
  })

  test("one finding per distinct rare phrase even if repeated", () => {
    const body = "Utilize the tool. Then utilize it again to utilize more.\n"
    const hits = lintSkill(makeInput({ body })).filter((f) => f.rule === "prose-frequency-plainness")
    expect(hits.length).toBe(1)
    expect(hits[0].message).toContain("3")
  })

  test("multiple distinct rare phrases each produce a finding", () => {
    const body = "Utilize the tool. Endeavor to ascertain the answer subsequent to reading.\n"
    const hits = lintSkill(makeInput({ body })).filter((f) => f.rule === "prose-frequency-plainness")
    expect(hits.length).toBeGreaterThanOrEqual(4)
  })

  test("ignores rare phrases inside fenced code blocks", () => {
    const body = [
      "# Title",
      "",
      "Use the tool.",
      "",
      "```",
      "Utilize the deprecated API in order to preserve behavior.",
      "```",
      "",
    ].join("\n")
    const findings = lintSkill(makeInput({ body }))
    expect(findings.find((f) => f.rule === "prose-frequency-plainness")).toBeUndefined()
  })

  test("ignores rare phrases inside inline code spans", () => {
    const body = "Call `utilize_legacy()` to keep the old path.\n"
    const findings = lintSkill(makeInput({ body }))
    expect(findings.find((f) => f.rule === "prose-frequency-plainness")).toBeUndefined()
  })

  test("does not flag legal terms of art that are not in the prose list", () => {
    const body = "Notwithstanding the foregoing, the indemnification clause survives termination.\n"
    const findings = lintSkill(makeInput({ body }))
    expect(findings.find((f) => f.rule === "prose-frequency-plainness")).toBeUndefined()
  })

  test("imported lintProseFrequency is callable directly", async () => {
    const mod = await import("../lib/skill-authoring-rules")
    expect(typeof mod.lintProseFrequency).toBe("function")
    const hits = mod.lintProseFrequency("Utilize the tool.\n")
    expect(hits.length).toBe(1)
    expect(hits[0].rule).toBe("prose-frequency-plainness")
  })
})
