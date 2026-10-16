import { tool } from "./legacy-definition"

// doc.haus update-agent tool. Rewrites a custom specialist subagent's label,
// description, and instructions. The name stays immutable.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Update an existing custom specialist subagent: replaces its label, description, and task instructions. Built-in specialists cannot be updated, and the name is immutable. Use list-agents first to read the current instructions — supply the complete revised instructions, not a diff.",
  args: {
    name: tool.schema.string().describe("Exact name of the custom agent to update"),
    label: tool.schema.string().describe("Human-readable agent name"),
    description: tool.schema.string().describe("Revised one-line summary of what this specialist checks"),
    instructions: tool.schema.string().describe("The complete revised task instructions as markdown"),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "update-agent", patterns: [args.name], metadata: { agent: args.name } })
    const res = await fetch(`${ingestUrl}/agents/${encodeURIComponent(args.name)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ label: args.label, description: args.description, instructions: args.instructions }),
    })
    if (!res.ok) return `Update failed (${res.status}): ${((await res.json()) as { error: string }).error}`
    return {
      title: `Updated agent ${args.name}`,
      output: `Updated "${args.name}".`,
      metadata: { agent: args.name },
    }
  },
})
