import { tool } from "./legacy-definition"

// doc.haus list-templates tool. Enumerates the firm's drafting templates and, for
// each, the placeholders Docxodus finds in it — so the model knows exactly which
// fills to gather before calling draft-document. The templates live in the global
// library owned by the ingest service (WORKSPACE_ROOT/.templates); we read them
// over HTTP so ingest stays the single writer over WORKSPACE_ROOT.

const ingestUrl = process.env.INGEST_URL ?? "http://127.0.0.1:4500"

export default tool({
  description:
    "List the document templates available for drafting, with a description of what each is for, the placeholders each one needs filled, and its optional clauses (keep-or-omit; pass declined ones to draft-document's omit). Use before draft-document to pick a template by its description and gather its fills.",
  args: {},
  async execute() {
    const { templates } = (await (await fetch(`${ingestUrl}/templates`)).json()) as {
      templates: {
        name: string
        description: string
        placeholders: { text: string; kind: string; hint?: string }[]
      }[]
    }
    return {
      title: `${templates.length} template(s)`,
      output: JSON.stringify(
        templates.map((t) => ({
          template: t.name,
          description: t.description,
          placeholders: t.placeholders.filter((p) => !p.text.startsWith("[optional:")),
          optionalClauses: t.placeholders
            .filter((p) => p.text.startsWith("[optional:"))
            .map((p) => p.text.slice("[optional:".length, -1).trim()),
        })),
        null,
        2,
      ),
      metadata: { templates: templates.map((t) => t.name) },
    }
  },
})
