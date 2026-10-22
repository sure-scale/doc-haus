import { tool } from "./legacy-definition"
import { fileURLToPath } from "node:url"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { withParagraphPointers } from "../lib/drafting-amendment"
import { isTemplateDraftDirectory, TEMPLATE_DRAFT_REVIEW_MESSAGE } from "../lib/template-draft-client"

// doc.haus update-template tool. Revises an existing template in the firm's
// global library: a new body (recomposed onto the styled blank seed exactly like
// create-template), a new description, or both. Omitted fields keep their
// current values — a description-only edit never touches the document bytes.
// The same no-client-data rule as create-template applies: every variable term
// stays a unique descriptive [insert ...] placeholder.

const templatesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "templates")
const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    'Update an existing template in the firm template library. Pass "name" (exact file name from list-templates) plus "content" (the complete revised template body as markdown — replaces the whole document; use get-template first to read the current body, then supply the full revision, not a diff), "description" (a revised one-line summary), or both. An omitted field keeps its current value. The template must still contain NO real client data: every party name, date, monetary amount, address, and reference number stays a UNIQUE descriptive placeholder like "[insert consultant name]". Returns the placeholders the revised template exposes.',
  args: {
    name: tool.schema
      .string()
      .describe('Exact file name of the template to update, e.g. "consulting-agreement.docx"'),
    content: tool.schema
      .string()
      .optional()
      .describe("The complete revised template body as markdown, with every variable term a unique [insert ...] placeholder. Omit to keep the current document body."),
    description: tool.schema
      .string()
      .optional()
      .describe("Revised one-line summary of what this template is for. Omit to keep the current description."),
  },
  async execute(args, ctx) {
    if (isTemplateDraftDirectory(ctx.directory)) return TEMPLATE_DRAFT_REVIEW_MESSAGE
    const name = path.basename(args.name.endsWith(".docx") ? args.name : `${args.name}.docx`)
    if (args.content === undefined && args.description === undefined)
      return 'Nothing to update — supply "content", "description", or both.'

    const { templates } = (await (await fetch(`${ingestUrl}/templates`)).json()) as {
      templates: { name: string; description: string }[]
    }
    const current = templates.find((t) => t.name === name)
    if (!current) return `No template named ${name}. Use list-templates to see the library.`

    // Changing the firm's shared library is gated on the user's approval. Asked
    // before any Docxodus work so a rejection costs nothing.
    await ctx.ask({ permission: "update-template", patterns: [name], metadata: { template: name } })

    if (args.content === undefined) {
      const res = await fetch(`${ingestUrl}/templates?name=${encodeURIComponent(name)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ description: args.description }),
      })
      if (!res.ok) return `Description update failed (${res.status}): ${await res.text()}`
      return {
        title: `Updated template ${name}`,
        output: `Updated the description of ${name}. The document body is unchanged.`,
        metadata: { template: name },
      }
    }

    const dx = await docxodus()
    const session = dx.openDocxSession(await Bun.file(path.join(templatesDir, "_base.docx")).bytes(), {})

    // Same composition as create-template: the seed's single empty paragraph is
    // the insertion anchor for the whole markdown body, then the seed paragraph
    // is dropped.
    const anchor = Object.keys(session.project().anchorIndex).find((id) => id.startsWith("p:"))!
    const inserted = session.insertParagraph(anchor, "after", args.content)
    if (!inserted.success) {
      session.close()
      return `Template build failed: ${inserted.error?.message ?? JSON.stringify(inserted.error)}`
    }
    session.deleteBlock(anchor)

    const placeholders = session.findPlaceholders().map((p) => p.match.text)
    const bytes = withParagraphPointers(session.save())
    session.close()

    // The library POST overwrites the existing .docx under the same name and sets
    // the description from the form, so the current one is re-sent when omitted.
    const form = new FormData()
    form.append("file", new File([new Uint8Array(bytes)], name))
    form.append("description", args.description ?? current.description)
    const res = await fetch(`${ingestUrl}/templates`, { method: "POST", body: form })
    if (!res.ok) return `Template composed but upload failed (${res.status}): ${await res.text()}`

    return {
      title: `Updated template ${name}`,
      output: [
        `Replaced ${name} in the firm template library.`,
        placeholders.length
          ? `Placeholders: ${placeholders.join(", ")}.`
          : "It exposes no placeholders — confirm every variable term is a [insert ...] placeholder.",
      ].join(" "),
      metadata: { template: name, placeholders },
    }
  },
})
