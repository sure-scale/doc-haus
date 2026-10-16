import { tool } from "./legacy-definition"

// doc.haus list-agents tool. Returns the firm's specialist subagents — the
// repo-shipped ones (read-only) plus the custom ones in dochaus/agents.json —
// with the custom agents' instructions, so the agent builder can avoid
// duplicates and revise existing agents.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "List the firm's specialist subagents (name, label, description, and for custom ones their task instructions). Always call this first before creating, updating, or deleting an agent — built-in specialists are read-only, and names must not collide.",
  args: {},
  async execute() {
    const res = await fetch(`${ingestUrl}/agents`)
    const { agents } = (await res.json()) as {
      agents: { name: string; label: string; description: string; instructions: string; builtin: boolean }[]
    }
    const lines =
      agents.length === 0
        ? "No agents yet."
        : agents
            .map((a) =>
              a.builtin
                ? `[${a.name}] (built-in, read-only) ${a.description}`
                : `[${a.name}] ${a.label} — ${a.description}\n  Instructions:\n${a.instructions
                    .split("\n")
                    .map((l) => `  ${l}`)
                    .join("\n")}`,
            )
            .join("\n\n")
    return {
      title: `${agents.length} agent(s)`,
      output: lines,
      metadata: { agents: agents.map((a) => a.name) },
    }
  },
})
