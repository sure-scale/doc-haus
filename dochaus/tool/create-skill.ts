import { tool } from "./legacy-definition"
import { existsSync, readdirSync } from "node:fs"
import path from "node:path"
import { lintSkill } from "../lib/skill-authoring-rules"

// doc.haus create-skill tool. Adds a firm skill to WORKSPACE_ROOT/.skills, owned
// by the ingest service (single writer over WORKSPACE_ROOT). The frontmatter
// (name + quoted description) is added by the library; supply only the body.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    'Create a new firm skill in the skill library. A skill is reference knowledge (clause standards, checklists, drafting guidance) the specialist agents load on demand when its description matches the task at hand. Use list-skills first — names must not collide with existing skills. Names are lowercase-hyphenated and must not start with "playbook-" (playbooks have their own importer). The description decides when agents load the skill, so state the trigger conditions in it.',
  args: {
    name: tool.schema.string().describe('Skill name, lowercase-hyphenated, e.g. "indemnity-standards"'),
    description: tool.schema
      .string()
      .describe(
        'One-line summary including when to use it, e.g. "Firm positions on indemnification. Use when judging indemnity clauses."',
      ),
    content: tool.schema.string().describe("The skill body as markdown — the frontmatter is added by the library"),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "create-skill", patterns: [args.name], metadata: { skill: args.name } })
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
      return `Skill rejected — fix these shape errors before creating:\n${errors.map((e) => `- [${e.rule}] ${e.message}`).join("\n")}`
    }
    const warns = findings.filter((f) => f.severity === "warn")
    const res = await fetch(`${ingestUrl}/skills`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(args),
    })
    if (!res.ok) return `Create failed (${res.status}): ${((await res.json()) as { error: string }).error}`
    const warnNote = warns.length
      ? `\n\nWarnings (skill created; consider addressing):\n${warns.map((w) => `- [${w.rule}] ${w.message}`).join("\n")}`
      : ""
    return {
      title: `Created skill ${args.name}`,
      output: `Created "${args.name}" in the firm skill library. Agents load it automatically when its description matches their task.${warnNote}`,
      metadata: { skill: args.name },
    }
  },
})

function builtinSkillNames(): Set<string> {
  const root = path.resolve(import.meta.dir, "..", "skill")
  if (!existsSync(root)) return new Set()
  return new Set(readdirSync(root).filter((entry) => existsSync(path.join(root, entry, "SKILL.md"))))
}
