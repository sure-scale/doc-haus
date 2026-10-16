import { tool } from "./legacy-definition"

// doc.haus delete-agent tool. Removes a custom specialist subagent. The registry
// refuses while any workflow still references the agent as a step.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Delete a custom specialist subagent. Built-in specialists cannot be deleted, and deletion is refused while any workflow uses the agent as a step (the error names the workflows). Confirm with the user before calling this — deletion is permanent.",
  args: {
    name: tool.schema.string().describe("Exact name of the custom agent to delete"),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "delete-agent", patterns: [args.name], metadata: { agent: args.name } })
    const res = await fetch(`${ingestUrl}/agents/${encodeURIComponent(args.name)}`, { method: "DELETE" })
    if (!res.ok) return `Delete failed (${res.status}): ${((await res.json()) as { error: string }).error}`
    return {
      title: `Deleted agent ${args.name}`,
      output: `Deleted "${args.name}" from the firm's specialist agents.`,
      metadata: { agent: args.name },
    }
  },
})
