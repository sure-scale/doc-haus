import { tool } from "./legacy-definition"

// doc.haus delete-workflow tool. Removes a custom workflow from the registry and
// deletes its generated orchestrator agent .md. Built-in workflows (shipped in
// the repo) cannot be deleted — only custom workflows registered via
// create-workflow are in scope.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Delete a custom workflow from the firm's workflow registry. Removes the workflow record and its generated agent definition. Built-in workflows cannot be deleted — only custom workflows created via create-workflow. Use list-workflows first to get the exact name.",
  args: {
    name: tool.schema
      .string()
      .describe("Exact workflow name (slug) from list-workflows — immutable registry key"),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "delete-workflow", patterns: [args.name], metadata: { workflow: args.name } })

    const res = await fetch(`${ingestUrl}/workflows/${encodeURIComponent(args.name)}`, { method: "DELETE" })

    if (!res.ok) {
      const { error } = (await res.json()) as { error: string }
      return error
    }

    return {
      title: `Deleted workflow ${args.name}`,
      output: `Workflow "${args.name}" has been removed from the registry and its agent definition deleted.`,
      metadata: { workflow: args.name },
    }
  },
})
