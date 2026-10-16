import { tool } from "./legacy-definition"

// doc.haus list-skills tool. Returns the firm's skill library — repo-shipped
// reference skills (read-only) plus the custom ones in WORKSPACE_ROOT/.skills —
// with full bodies, so the skill builder can avoid duplicates and revise
// existing skills. Playbooks are managed separately and never appear here.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "List the firm's skills (name, description, full markdown body, and whether each is built-in). Always call this first before creating, updating, or deleting a skill — built-in skills are read-only, and names must not collide.",
  args: {},
  async execute() {
    const res = await fetch(`${ingestUrl}/skills`)
    const { skills } = (await res.json()) as {
      skills: { name: string; description: string; content: string; builtin: boolean }[]
    }
    const lines =
      skills.length === 0
        ? "No skills yet."
        : skills
            .map(
              (s) =>
                `[${s.name}]${s.builtin ? " (built-in, read-only)" : ""} ${s.description}\n---\n${s.content}\n---`,
            )
            .join("\n\n")
    return {
      title: `${skills.length} skill(s)`,
      output: lines,
      metadata: { skills: skills.map((s) => s.name) },
    }
  },
})
