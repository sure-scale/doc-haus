import { tool } from "./legacy-definition"

// doc.haus update-workflow tool. Replaces the definition of an existing custom
// workflow — its label, description, scope, prompt, and steps. The name is the
// immutable slug used as the registry key and agent filename; it cannot be
// changed here. Only custom workflows in the registry can be updated.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Update an existing custom workflow's label, description, scope, prompt, or steps. Pass the immutable workflow name from list-workflows as the first argument. Only custom workflows in the registry can be updated — built-in workflows cannot be modified. Use list-workflows first to get the exact name.",
  args: {
    name: tool.schema
      .string()
      .describe("Exact workflow name (slug) from list-workflows — immutable registry key"),
    label: tool.schema.string().describe("Updated human-readable workflow name"),
    description: tool.schema.string().describe("Updated one-line summary of what this workflow does"),
    scope: tool.schema
      .enum(["matter", "document"])
      .describe(
        "matter = opens its own new chat across the whole matter; document = runs in the current document thread",
      ),
    prompt: tool.schema
      .string()
      .optional()
      .describe("The message sent when a user launches the workflow. Omit to keep the current value."),
    steps: tool.schema
      .array(
        tool.schema.object({
          agent: tool.schema
            .string()
            .describe("Exact subagent name from list-workflows — never invent a name"),
          instructions: tool.schema
            .string()
            .optional()
            .describe("Per-step focus for this subagent"),
        }),
      )
      .describe("Updated ordered list of pipeline steps. Replaces the existing steps entirely."),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "update-workflow", patterns: [args.name], metadata: { workflow: args.name } })

    // The server PUT replaces the whole record, so an omitted prompt must be
    // backfilled from the current registry entry to honor "omit to keep".
    const { workflows } = (await (await fetch(`${ingestUrl}/workflows`)).json()) as {
      workflows: { name: string; prompt: string }[]
    }
    const current = workflows.find((w) => w.name === args.name)

    const body = {
      label: args.label,
      description: args.description,
      scope: args.scope,
      prompt: args.prompt ?? current?.prompt ?? `Run the "${args.label}" workflow on this matter.`,
      steps: args.steps.map((s) => ({ agent: s.agent, instructions: s.instructions ?? "" })),
    }

    const res = await fetch(`${ingestUrl}/workflows/${encodeURIComponent(args.name)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    })

    if (!res.ok) {
      const { error } = (await res.json()) as { error: string }
      return error
    }

    const workflow = (await res.json()) as { name: string; label: string; steps: { agent: string; instructions?: string }[] }
    const stepSummary = workflow.steps.map((s, i) => `  ${i + 1}. ${s.agent}${s.instructions ? ` — ${s.instructions}` : ""}`).join("\n")

    return {
      title: `Updated workflow ${workflow.name}`,
      output: `Updated "${workflow.label}" (workflow "${workflow.name}").\nSteps:\n${stepSummary}`,
      metadata: { workflow: workflow.name },
    }
  },
})
