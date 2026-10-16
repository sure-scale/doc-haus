import { tool } from "./legacy-definition"

// doc.haus delete-skill tool. Removes a custom firm skill from
// WORKSPACE_ROOT/.skills. Built-in skills cannot be deleted.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Delete a custom firm skill from the skill library. Built-in skills cannot be deleted. Confirm with the user before calling this — deletion is permanent.",
  args: {
    name: tool.schema.string().describe("Exact name of the custom skill to delete"),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "delete-skill", patterns: [args.name], metadata: { skill: args.name } })
    const res = await fetch(`${ingestUrl}/skills/${encodeURIComponent(args.name)}`, { method: "DELETE" })
    if (!res.ok) return `Delete failed (${res.status}): ${((await res.json()) as { error: string }).error}`
    return {
      title: `Deleted skill ${args.name}`,
      output: `Deleted "${args.name}" from the firm skill library.`,
      metadata: { skill: args.name },
    }
  },
})
