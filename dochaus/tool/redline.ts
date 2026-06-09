import { tool } from "@opencode-ai/plugin"
import { existsSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"

// doc.haus redline tool. Rewrites a whole clause — the paragraph a
// search-document citation points at — as a native Word tracked change, so a
// reviewer accepts or rejects the new wording in Microsoft Word.
//
// This is the citation-driven, clause-level counterpart to the surgical
// find/replace tools (word-integration, tracked-changes): `clause` locates the
// paragraph and the whole paragraph is rewritten to `replacement`. We match on
// the block's flat text, whitespace-tolerant, because citation excerpts come
// from a different text extraction (mammoth) than Docxodus' own projection and
// their spacing/char offsets don't align — so we only ever LOCATE the block, we
// never offset-index into it.
//
// Session edits are always plain; native w:ins/w:del come only from
// compareDocuments(original, edited, { authorName }). So we rewrite an in-memory
// copy, then redline it against the on-disk original and write that back.

export default tool({
  description:
    "Rewrite a whole clause as a tracked change, using a passage retrieved from search-document. Locates the paragraph containing `clause` and replaces its entire text with `replacement`, recorded as a native Word revision attributed to an author (a reviewer accepts or rejects it in Word). Use this to redline a clause you found as a citation; use tracked-changes for a surgical word/phrase swap.",
  args: {
    document: tool.schema.string().describe("Document file name within the matter (the docPath from a citation)"),
    clause: tool.schema
      .string()
      .describe("Text from the clause to rewrite — the citation excerpt, or a sentence within it. Used to locate the paragraph."),
    replacement: tool.schema
      .string()
      .describe("The new clause text. Replaces the whole located paragraph; markdown is supported."),
    author: tool.schema.string().optional().describe("Name to attribute the tracked change to (default: doc.haus)"),
  },
  async execute(args, ctx) {
    const file = path.isAbsolute(args.document) ? args.document : path.join(ctx.directory, args.document)
    if (!existsSync(file)) return `Document not found in this matter: ${args.document}`

    const dx = await docxodus()
    const original = await Bun.file(file).bytes()
    const session = dx.openDocxSession(original, {})

    const target = session.findByText(args.clause, { ignoreWhitespace: true })
    if (!target) {
      session.close()
      return `Clause not found in ${path.basename(file)}: ${JSON.stringify(args.clause)}`
    }

    const result = session.replaceText(target.id, args.replacement)
    if (!result.success) {
      session.close()
      return `Redline failed: ${result.error?.message ?? JSON.stringify(result.error)}`
    }

    const edited = session.save()
    session.close()

    const author = args.author ?? "doc.haus"
    const redline = await dx.compareDocuments(original, edited, { authorName: author })
    await Bun.write(file, redline)

    return {
      title: `Redlined a clause in ${path.basename(file)}`,
      output: `Rewrote the clause matching ${JSON.stringify(args.clause)} in ${path.basename(file)}, recorded as a tracked change attributed to ${author}. Reviewer can accept or reject in Word.`,
      metadata: { document: file, clause: args.clause, replacement: args.replacement, anchor: target.id, author },
    }
  },
})
