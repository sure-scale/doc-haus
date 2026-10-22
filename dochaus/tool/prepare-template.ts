import { tool } from "./legacy-definition"
import { attachedTemplateDraft, templateDraftRequest, type TemplateDraftInfo } from "../lib/template-draft-client"

export default tool({
  description:
    "Prepare a reusable template by editing only this conversation's attached Word copy. Read it with read-template-source first, establish intended use/party position/jurisdiction and ask focused questions for material unknowns. Each edit targets a paragraph_handle and a zero-based occurrence of an exact before_quote within that paragraph. Preserve legal substance and Word structure; replace client details with descriptive [insert ...] placeholders. Do not replace legal constants merely because they are names, dates or numbers. Batches are atomic. This prepares a preview only: it never saves to the library. The user reviews the actual Word proposal and changes, then uses Save to library. After each successful batch, re-read for new handles before another edit.",
  args: {
    revision: tool.schema.number().int().min(0).describe("Current revision returned by read-template-source."),
    name: tool.schema.string().min(1).describe("A general, client-free template filename ending in .docx."),
    description: tool.schema.string().min(1).describe("One-line purpose including material scope or position limits; no client details."),
    edits: tool.schema.array(tool.schema.object({
      paragraph_handle: tool.schema.string().min(1),
      occurrence: tool.schema.number().int().min(0).describe("Zero-based occurrence of the exact quote in that paragraph."),
      before_quote: tool.schema.string().min(1),
      replacement: tool.schema.string(),
    })).max(80).describe("Focused exact-text replacements. Empty only when the copy needs no further textual changes."),
    notes: tool.schema.array(tool.schema.string()).optional().describe("Material qualifications and deliberate optional-clause choices for the user's review; notes do not resolve missing instructions."),
  },
  async execute(args, ctx) {
    const draft = await attachedTemplateDraft(ctx.directory, ctx.abort)
    if (draft.status === "published") throw new Error("This template has already been saved. Start a new conversation to prepare another copy.")
    const prepared = await templateDraftRequest<TemplateDraftInfo>(draft.id, "/edits", {
      method: "POST", signal: ctx.abort, headers: { "content-type": "application/json" },
      body: JSON.stringify({ ...args, edits: args.edits.map((edit, index) => ({ ...edit, id: `edit-${index + 1}` })) }),
    })
    return {
      title: `Prepared template ${prepared.name}`,
      output: `Prepared ${prepared.name} for review with ${prepared.changes.length} recorded change(s). ` +
        "The source document is unchanged and the proposal has not been saved to the library. Ask the user to open Review proposed template to inspect the Word document and changes, edit the name and description if needed, and choose Save to library. " +
        `Current revision: ${prepared.revision}.` + (prepared.notes.length ? ` Review notes: ${JSON.stringify(prepared.notes)}` : ""),
      metadata: { templateDraft: prepared.id, revision: prepared.revision, changes: prepared.changes.length },
    }
  },
})
