import { tool } from "./legacy-definition"
import path from "node:path"

// doc.haus delete-template tool. Removes a template from the firm's global
// library (WORKSPACE_ROOT/.templates, owned by the ingest service). Documents
// already drafted from the template are standalone copies and are unaffected;
// only future drafting loses the base.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "Delete a template from the firm template library. Deletion is permanent — confirm with the user before calling this. Documents already drafted from the template are unaffected; it only stops being available as a drafting base. Use list-templates first to get the exact name.",
  args: {
    name: tool.schema.string().describe('Exact file name of the template to delete, e.g. "consulting-agreement.docx"'),
  },
  async execute(args, ctx) {
    const name = path.basename(args.name.endsWith(".docx") ? args.name : `${args.name}.docx`)

    const { templates } = (await (await fetch(`${ingestUrl}/templates`)).json()) as { templates: { name: string }[] }
    if (!templates.some((t) => t.name === name)) return `No template named ${name}. Use list-templates to see the library.`

    await ctx.ask({ permission: "delete-template", patterns: [name], metadata: { template: name } })

    const res = await fetch(`${ingestUrl}/templates?name=${encodeURIComponent(name)}`, { method: "DELETE" })
    if (!res.ok) return `Delete failed (${res.status}): ${await res.text()}`

    return {
      title: `Deleted template ${name}`,
      output: `Deleted ${name} from the firm template library.`,
      metadata: { template: name },
    }
  },
})
