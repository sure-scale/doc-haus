import { tool } from "./legacy-definition"
import { attachedTemplateDraft, templateDraftRequest } from "../lib/template-draft-client"
import type { DocxRevisionPacket } from "../lib/docx-amendments"
import { UNTRUSTED_DOCUMENT_NOTICE } from "../lib/untrusted"
import { detectInjection } from "../lib/untrusted-text"

export default tool({
  description:
    "Read a bounded page of this conversation's attached document working copy, including tables, headers and footers. No matter or path argument is accepted. Read all pages before saying all client details have been removed. Start at offset=0 and follow next_offset and next_text_offset until null. The returned revision and paragraph_handle locate exact edits for prepare-template; re-read after preparing edits because handles change. This is a drafting base, not evidence that its negotiated provisions are firm policy.",
  args: {
    offset: tool.schema.number().int().min(0).optional().describe("Zero-based paragraph index, default 0."),
    text_offset: tool.schema.number().int().min(0).optional().describe("Unicode character offset within the first paragraph, default 0; use the returned continuation."),
    limit: tool.schema.number().int().min(1).max(4000).optional().describe("Maximum document text characters in this page, default 4000."),
    count: tool.schema.number().int().min(1).max(20).optional().describe("Maximum paragraphs in this page, default 8."),
  },
  async execute(args, ctx) {
    const draft = await attachedTemplateDraft(ctx.directory, ctx.abort)
    const projection = await templateDraftRequest<DocxRevisionPacket>(draft.id, "/text", { signal: ctx.abort })
    const offset = args.offset ?? 0
    const textOffset = args.text_offset ?? 0
    const total = projection.paragraphs.length
    if (offset > total || offset === total && textOffset > 0)
      throw new Error("The paragraph offset is beyond this document. Use the previous page's continuation or start at offset=0.")
    const page = [] as { index: number; paragraph_handle: string; scope: string; text: string; text_offset: number;
      total_characters: number; table: { table: number; row: number; column: number } | null }[]
    let remaining = args.limit ?? 4000
    let index = offset
    let nextTextOffset = textOffset
    while (index < total && page.length < (args.count ?? 8) && remaining > 0) {
      const paragraph = projection.paragraphs[index]
      const characters = Array.from(paragraph.text)
      if (nextTextOffset > characters.length) throw new Error("The text offset is beyond its paragraph. Use the previous page's continuation.")
      const text = characters.slice(nextTextOffset, nextTextOffset + remaining).join("")
      const returned = Array.from(text).length
      page.push({ index, paragraph_handle: paragraph.paragraph_handle, scope: paragraph.scope, text,
        text_offset: nextTextOffset, total_characters: characters.length,
        table: paragraph.table ? { table: paragraph.table.table, row: paragraph.table.row, column: paragraph.table.column } : null })
      remaining -= returned
      nextTextOffset += returned
      if (nextTextOffset < characters.length) break
      index++
      nextTextOffset = 0
    }
    const next = index < total ? index : null
    const warning = [...new Set(detectInjection(page.map(paragraph => paragraph.text).join("\n")).map(finding => finding.detail))]
    const cursor = { offset, text_offset: textOffset, total_paragraphs: total, next_offset: next, next_text_offset: next === null ? null : nextTextOffset }
    return {
      title: "Read attached template document",
      output: `<untrusted-document>\n${JSON.stringify(page)}\n</untrusted-document>\n\n${UNTRUSTED_DOCUMENT_NOTICE}\n` +
        (warning.length ? `Injection warning: ${warning.join("; ")}\n` : "") +
        `Revision: ${draft.revision}. ${JSON.stringify(cursor)}\n` +
        (next === null ? "End of attached document (EOF)." : `Continue with offset=${next}, text_offset=${nextTextOffset}.`) +
        (draft.notes.length ? `\nPreparation notes: ${JSON.stringify(draft.notes)}` : ""),
      metadata: { templateDraft: draft.id, revision: draft.revision, page: cursor, truncated: next !== null },
    }
  },
})
