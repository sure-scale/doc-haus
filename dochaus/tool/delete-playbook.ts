import { tool } from "./legacy-definition"

// doc.haus delete-playbook tool. Removes an imported playbook from
// WORKSPACE_ROOT/.playbooks. The server refuses while any matter is still bound
// to the playbook — deleting it would leave that matter's review pipeline
// pointing at nothing — so bindings must be cleared first. Repo-shipped
// playbooks cannot be deleted.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Delete an imported playbook from the firm playbook library. Refused while any matter is still bound to the playbook — the error names the bound matters; unbind each (via the matter's playbook setting) before retrying. Repo-shipped playbooks cannot be deleted. Deletion is permanent — confirm with the user before calling this. Use list-playbooks first to get the exact name and see current bindings.",
  args: {
    name: tool.schema.string().describe('Exact name of the playbook to delete, e.g. "playbook-saas-msa"'),
  },
  async execute(args, ctx) {
    await ctx.ask({ permission: "delete-playbook", patterns: [args.name], metadata: { playbook: args.name } })
    const res = await fetch(`${ingestUrl}/playbooks/${encodeURIComponent(args.name)}`, { method: "DELETE" })
    if (!res.ok) return `Delete failed (${res.status}): ${((await res.json()) as { error: string }).error}`
    return {
      title: `Deleted playbook ${args.name}`,
      output: `Deleted "${args.name}" from the firm playbook library.`,
      metadata: { playbook: args.name },
    }
  },
})
