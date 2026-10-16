import { tool } from "./legacy-definition"
import { existsSync, readdirSync } from "node:fs"
import path from "node:path"
import { lintSkill } from "../lib/skill-authoring-rules"

// doc.haus update-skill tool. Rewrites a custom firm skill in
// WORKSPACE_ROOT/.skills. Built-in skills are read-only.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Update an existing custom firm skill: replaces its description and full markdown body. Built-in skills cannot be updated. Use list-skills first to read the current body — supply the complete revised body, not a diff.",
  args: {
    name: tool.schema.string().describe("Exact name of the custom skill to update"),
    description: tool.schema.string().describe("Revised one-line summary including when to use it"),
    content: tool.schema.string().describe("The complete revised skill body as markdown"),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "update-skill", patterns: [args.name], metadata: { skill: args.name } })
    const findings = lintSkill({
      name: args.name,
      path: `dochaus/skill/${args.name}/SKILL.md`,
      frontmatter: `description: ${args.description}`,
      description: args.description,
      body: args.content,
      references: [],
      builtins: builtinSkillNames(),
    })
    const errors = findings.filter((f) => f.severity === "error")
    if (errors.length) {
      return `Skill rejected — fix these shape errors before updating:\n${errors.map((e) => `- [${e.rule}] ${e.message}`).join("\n")}`
    }
    const warns = findings.filter((f) => f.severity === "warn")
    const res = await fetch(`${ingestUrl}/skills/${encodeURIComponent(args.name)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description: args.description, content: args.content }),
    })
    if (!res.ok) return `Update failed (${res.status}): ${((await res.json()) as { error: string }).error}`
    const warnNote = warns.length
      ? `\n\nWarnings (skill updated; consider addressing):\n${warns.map((w) => `- [${w.rule}] ${w.message}`).join("\n")}`
      : ""
    return {
      title: `Updated skill ${args.name}`,
      output: `Updated "${args.name}".${warnNote}`,
      metadata: { skill: args.name },
    }
  },
})

function builtinSkillNames(): Set<string> {
  const root = path.resolve(import.meta.dir, "..", "skill")
  if (!existsSync(root)) return new Set()
  return new Set(readdirSync(root).filter((entry) => existsSync(path.join(root, entry, "SKILL.md"))))
}
