import { tool } from "./legacy-definition"

// doc.haus create-workflow tool. Registers a new custom workflow in the firm's
// workflow registry (dochaus/workflows.json) and generates its orchestrator agent
// .md in dochaus/agent/. The ingest service owns both writes.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Create a new custom workflow in the firm's workflow registry. A workflow is an ordered pipeline of specialist subagents launched as one repeatable process on a matter. Use list-workflows first to discover available subagent names — only exact names from that list are valid. The workflow's name is derived from the label (slugified) and is immutable once created.",
  args: {
    label: tool.schema.string().describe("Human-readable workflow name, e.g. \"Full Contract Review\""),
    description: tool.schema
      .string()
      .describe("One-line summary of what this workflow does, shown in the Workflows launcher"),
    scope: tool.schema
      .enum(["matter", "document"])
      .describe(
        "matter = opens its own new chat across the whole matter; document = runs in the current document thread",
      ),
    prompt: tool.schema
      .string()
      .optional()
      .describe(
        'The message sent when a user launches the workflow. Defaults to "Run the \\"<label>\\" workflow on this matter." when omitted.',
      ),
    steps: tool.schema
      .array(
        tool.schema.object({
          agent: tool.schema
            .string()
            .describe("Exact subagent name from list-workflows — never invent a name"),
          instructions: tool.schema
            .string()
            .optional()
            .describe("Per-step focus for this subagent, e.g. \"focus on IP ownership clauses\""),
        }),
      )
      .describe(
        "Ordered list of pipeline steps. Each step's output is passed to the next. Order matters: a summarizer goes last; a challenger attacks findings from steps before it.",
      ),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "create-workflow", patterns: [args.label], metadata: { workflow: args.label } })

    const body = {
      label: args.label,
      description: args.description,
      scope: args.scope,
      prompt: args.prompt ?? `Run the "${args.label}" workflow on this matter.`,
      steps: args.steps.map((s) => ({ agent: s.agent, instructions: s.instructions ?? "" })),
    }

    const res = await fetch(`${ingestUrl}/workflows`, {
      method: "POST",
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
      title: `Created workflow ${workflow.name}`,
      output: `Registered "${workflow.label}" as workflow "${workflow.name}".\nSteps:\n${stepSummary}`,
      metadata: { workflow: workflow.name },
    }
  },
})
