import { tool } from "./legacy-definition"

// doc.haus create-agent tool. Registers a new custom specialist subagent in the
// firm's agent registry (dochaus/agents.json) and generates its .md in
// dochaus/agent/. The ingest service owns both writes and wraps the instructions
// with the standard citation and output discipline; supply only the specialist's
// task instructions.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Create a new custom specialist subagent. A specialist analyzes a matter's documents for one focus (e.g. IP ownership, data privacy, change-of-control) and reports cited findings; once created it is immediately available as a workflow pipeline step. The agent's name is derived from the label (slugified) and is immutable once created. Use list-agents first — names must not collide. Supply only the task instructions; the citation and output rules are added automatically.",
  args: {
    label: tool.schema.string().describe('Human-readable agent name, e.g. "IP Ownership Specialist"'),
    description: tool.schema
      .string()
      .describe("One-line summary of what this specialist checks, shown in the agent library and workflow builder"),
    instructions: tool.schema
      .string()
      .describe(
        "The specialist's task instructions as markdown: what to review, what to flag, and what counts as a finding. Do not include citation or output formatting rules — those are added automatically.",
      ),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "create-agent", patterns: [args.label], metadata: { agent: args.label } })
    const res = await fetch(`${ingestUrl}/agents`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(args),
    })
    if (!res.ok) return `Create failed (${res.status}): ${((await res.json()) as { error: string }).error}`
    const agent = (await res.json()) as { name: string; label: string }
    return {
      title: `Created agent ${agent.name}`,
      output: `Registered "${agent.label}" as specialist agent "${agent.name}". It is now available as a workflow step — compose it into a workflow from the Workflows page.`,
      metadata: { agent: agent.name },
    }
  },
})
