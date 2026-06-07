import { tool } from "@opencode-ai/plugin"
import { existsSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"

// doc.haus tracked-changes tool. Applies an edit to the matter's canonical Word
// (.docx) document as a native Word tracked change (w:ins/w:del), attributed to
// an author, so a reviewer can accept or reject it in Microsoft Word.
//
// Docxodus session edits are always plain; native tracked changes come from
// compareDocuments(original, edited, { authorName }), which diffs the two whole
// documents into a single revision-marked output. So we edit a throwaway copy in
// memory, then redline it against the on-disk original and write the redline back.

export default tool({
  description:
    "Edit the matter's canonical Word (.docx) document as a tracked change. Finds exact text and replaces it, recording the change as a native Word insertion/deletion attributed to an author, which a reviewer can accept or reject. Use this for negotiation redlines; use word-integration for silent (non-tracked) edits.",
  args: {
    document: tool.schema.string().describe("Document file name within the matter (the docPath from a citation)"),
    find: tool.schema.string().describe("The exact text to find"),
    replace: tool.schema.string().describe("The replacement text"),
    author: tool.schema.string().optional().describe("Name to attribute the tracked change to (default: doc.haus)"),
  },
  async execute(args, ctx) {
    const file = path.isAbsolute(args.document) ? args.document : path.join(ctx.directory, args.document)
    if (!existsSync(file)) return `Document not found in this matter: ${args.document}`

    const dx = await docxodus()
    const original = await Bun.file(file).bytes()
    const session = dx.openDocxSession(original, {})

    const targets = session.findAllByText(args.find)
    if (!targets.length) {
      session.close()
      return `Text not found in ${path.basename(file)}: ${JSON.stringify(args.find)}`
    }

    const results = targets.flatMap((t) => session.replaceTextRange(t.id, args.find, args.replace))
    const failed = results.find((r) => !r.success)
    if (failed) {
      session.close()
      return `Edit failed: ${failed.error?.message ?? JSON.stringify(failed.error)}`
    }

    const edited = session.save()
    session.close()

    const author = args.author ?? "doc.haus"
    const redline = await dx.compareDocuments(original, edited, { authorName: author })
    await Bun.write(file, redline)

    return {
      title: `Tracked change in ${path.basename(file)}`,
      output: `Recorded ${results.length} tracked change(s) replacing ${JSON.stringify(args.find)} with ${JSON.stringify(args.replace)} in ${path.basename(file)}, attributed to ${author}. Reviewer can accept or reject in Word.`,
      metadata: { document: file, find: args.find, replace: args.replace, edits: results.length, author },
    }
  },
})
