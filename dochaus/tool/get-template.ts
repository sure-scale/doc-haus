import { tool } from "./legacy-definition"
import path from "node:path"

// doc.haus get-template tool. Returns the full plain text of one template in the
// firm library, plus its description and placeholders. list-templates returns
// metadata only; this fetches the body, which template-builder needs before
// revising a template with update-template — you cannot fix a clause you have
// not read. Read-only, so it carries no approval gate.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    'Read the full text of a template in the firm template library, with its description and placeholders. Use before update-template to load the current body — supply update-template the complete revised body, not a diff. Use list-templates first to find the exact template name.',
  args: {
    name: tool.schema.string().describe('Exact template file name from list-templates, e.g. "consulting-agreement.docx"'),
  },
  async execute(args) {
    const name = path.basename(args.name.endsWith(".docx") ? args.name : `${args.name}.docx`)

    const { templates } = (await (await fetch(`${ingestUrl}/templates`)).json()) as {
      templates: { name: string; description: string; placeholders: { text: string }[] }[]
    }
    const template = templates.find((t) => t.name === name)
    if (!template) return `No template named ${name}. Use list-templates to see the library.`

    const res = await fetch(`${ingestUrl}/templates/text?name=${encodeURIComponent(name)}`)
    if (!res.ok) return `Could not read ${name} (${res.status}): ${await res.text()}`
    const { text } = (await res.json()) as { text: string }

    return {
      title: `Read template ${name}`,
      output: [
        `Description: ${template.description || "(none)"}`,
        `Placeholders: ${template.placeholders.length ? template.placeholders.map((p) => p.text).join(", ") : "(none)"}`,
        "---",
        text,
      ].join("\n"),
      metadata: { template: name, placeholders: template.placeholders.map((p) => p.text) },
    }
  },
})
