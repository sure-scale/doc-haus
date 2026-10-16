import { tool } from "./legacy-definition"

// doc.haus update-playbook tool. Revises an imported playbook in
// WORKSPACE_ROOT/.playbooks: its description, its full body, or both. The server
// implements true keep-semantics — an omitted field retains its current value,
// and the version/last-reviewed frontmatter survives every update — so a
// description-only edit can never wipe the clause sections. Repo-shipped
// playbooks are read-only.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    'Update an existing imported playbook: pass "name" (exact playbook name from list-playbooks) plus "description" (revised one-line summary), "content" (the complete revised playbook body as markdown — one "##" section per clause type with Preferred/Fallbacks/Unacceptable/Rationale and ```approved fences holding clause text copied verbatim, never paraphrased), or both. An omitted field keeps its current value. Use list-playbooks first to read the current body — supply the complete revised body, not a diff. Repo-shipped playbooks cannot be updated.',
  args: {
    name: tool.schema.string().describe('Exact name of the playbook to update, e.g. "playbook-saas-msa"'),
    description: tool.schema
      .string()
      .optional()
      .describe("Revised one-line summary of what this playbook covers. Omit to keep the current description."),
    content: tool.schema
      .string()
      .optional()
      .describe("The complete revised playbook body as markdown, with approved clause text inside ```approved fences copied verbatim from the source. Omit to keep the current body."),
  },
  async execute(args, ctx) {
    if (args.description === undefined && args.content === undefined)
      return 'Nothing to update — supply "description", "content", or both.'

    await ctx.ask({ permission: "update-playbook", patterns: [args.name], metadata: { playbook: args.name } })

    const res = await fetch(`${ingestUrl}/playbooks/${encodeURIComponent(args.name)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ description: args.description, content: args.content }),
    })
    if (!res.ok) return `Update failed (${res.status}): ${((await res.json()) as { error: string }).error}`

    const playbook = (await res.json()) as { name: string; matters: string[] }
    return {
      title: `Updated playbook ${playbook.name}`,
      output: [
        `Updated "${playbook.name}".`,
        playbook.matters.length
          ? `It is bound to: ${playbook.matters.join(", ")} — reviews on those matters use the revised positions from now on.`
          : "",
      ]
        .filter(Boolean)
        .join(" "),
      metadata: { playbook: playbook.name },
    }
  },
})
