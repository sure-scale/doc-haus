import { tool } from "@opencode-ai/plugin"
import { existsSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"

// doc.haus word-integration tool. Reads and edits the matter's canonical Word
// (.docx) document in place via Docxodus, so edits round-trip through Microsoft
// Word. `services/ingest` writes that canonical .docx into the matter directory;
// this tool operates on the same file. Edits here are plain (non-tracked);
// tracked changes are a separate tool.

export default tool({
  description:
    'Read or edit the matter\'s canonical Word (.docx) document in place. Use action "read" to get the document text, or "replace" to change wording. Edits are plain (not tracked changes) and re-open correctly in Microsoft Word.',
  args: {
    document: tool.schema.string().describe("Document file name within the matter (the docPath from a citation)"),
    action: tool.schema.enum(["read", "replace"]).describe('"read" the document text, or "replace" text in place'),
    find: tool.schema.string().optional().describe("For replace: the exact text to find"),
    replace: tool.schema.string().optional().describe("For replace: the replacement text"),
  },
  async execute(args, ctx) {
    const file = path.isAbsolute(args.document) ? args.document : path.join(ctx.directory, args.document)
    if (!existsSync(file)) return `Document not found in this matter: ${args.document}`

    const dx = await docxodus()
    const session = dx.openDocxSession(await Bun.file(file).bytes(), {})

    if (args.action === "read") {
      const markdown = session.project().markdown
      session.close()
      return { title: `Read ${path.basename(file)}`, output: markdown }
    }

    if (!args.find || args.replace === undefined) {
      session.close()
      return 'Replace requires both "find" and "replace".'
    }

    const targets = session.findAllByText(args.find)
    if (!targets.length) {
      session.close()
      return `Text not found in ${path.basename(file)}: ${JSON.stringify(args.find)}`
    }

    const results = targets.flatMap((t) => session.replaceTextRange(t.id, args.find!, args.replace!))
    const failed = results.find((r) => !r.success)
    if (failed) {
      session.close()
      return `Edit failed: ${failed.error?.message ?? JSON.stringify(failed.error)}`
    }

    await Bun.write(file, session.save())
    session.close()
    return {
      title: `Edited ${path.basename(file)}`,
      output: `Replaced ${results.length} occurrence(s) of ${JSON.stringify(args.find)} with ${JSON.stringify(args.replace)} in ${path.basename(file)}.`,
      metadata: { document: file, find: args.find, replace: args.replace, edits: results.length },
    }
  },
})
