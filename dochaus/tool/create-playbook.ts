import { tool } from "./legacy-definition"

// doc.haus create-playbook tool. Imports a firm playbook skill into the
// per-deployment library (WORKSPACE_ROOT/.playbooks), owned by the ingest
// service. The body is the playbook skill markdown — one `##` section per clause
// type with Preferred/Fallbacks/Unacceptable/Rationale and ```approved fences
// carrying the firm-approved clause text copied byte-exact from the source
// document. The name must start "playbook-" so it qualifies as a playbook skill
// (the discovery scan and the matter playbook binding both require that prefix).

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    'Import a firm playbook into the playbook library. Pass "name" (the skill name, must start with "playbook-", e.g. "playbook-saas-msa"), "description" (a one-line summary of what the playbook covers — surfaced when a matter is bound to a playbook, so make it specific), and "content" (the playbook body as markdown: one "##" section per clause type, each with **Preferred:**, optional **Fallbacks:**, **Unacceptable:**, **Rationale:**, an ```approved fence holding the firm-approved replacement text, and optional ```approved-fallback fences). Approved clause text inside the fences must be copied verbatim from the source document, never paraphrased. The frontmatter is added by the library; supply only the body. Returns the skill name created.',
  args: {
    name: tool.schema
      .string()
      .describe('Skill name for the new playbook, must start with "playbook-", e.g. "playbook-saas-msa"'),
    description: tool.schema
      .string()
      .describe('A one-line summary of what this playbook covers, e.g. "Firm playbook for SaaS master service agreements"'),
    content: tool.schema
      .string()
      .describe(
        'The playbook body as markdown — one "##" section per clause type with Preferred/Fallbacks/Unacceptable/Rationale and ```approved fences holding clause text copied verbatim from the source document',
      ),
  },
  async execute(args, ctx) {
    if (!args.name.startsWith("playbook-")) return `Playbook name must start with "playbook-". Got "${args.name}".`

    const existing = (await (await fetch(`${ingestUrl}/playbooks`)).json()) as { name: string }[]
    if (existing.some((p) => p.name === args.name)) return `A playbook named ${args.name} already exists. Pick a different name.`

    // Adding a playbook to the firm's shared library is gated on the user's
    // approval. Asked before the upload so a rejection costs nothing.
    await ctx.ask({ permission: "create-playbook", patterns: [args.name], metadata: { playbook: args.name } })

    const res = await fetch(`${ingestUrl}/playbooks`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: args.name, description: args.description, content: args.content }),
    })
    if (!res.ok) return `Playbook composed but upload failed (${res.status}): ${await res.text()}`

    return {
      title: `Created playbook ${args.name}`,
      output: `Imported ${args.name} into the firm playbook library. Bind it to a matter via the matter's playbook setting to review against it.`,
      metadata: { playbook: args.name },
    }
  },
})
