import { tool } from "./legacy-definition"
import { existsSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { assertDirectWorkingDocxWriteAllowed } from "../lib/working-docx"
import { withArtifactMutationLocksAsync } from "../capabilities/artifacts/mutation-lock"

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

    // Replace rewrites the canonical .docx in place with no review step, so it
    // is gated on the matter owner's approval (permission "word-integration" in
    // opencode.json). ctx.ask blocks until they reply and throws on reject —
    // asked before the Docxodus session opens so a rejection leaks nothing.
    // "Always" approves future edits to this document only.
    if (args.action === "replace") {
      if (!args.find || args.replace === undefined) return 'Replace requires both "find" and "replace".'
      await ctx.ask({
        permission: "word-integration",
        patterns: [file],
        always: [file],
        metadata: { document: path.basename(file), find: args.find, replace: args.replace },
      })
    }

    const operate = async () => {
      if (args.action !== "read") assertDirectWorkingDocxWriteAllowed(ctx.directory,file)
      const dx = await docxodus()
      const session = dx.openDocxSession(await Bun.file(file).bytes(),{})
      try {
        if (args.action === "read") return {title:`Read ${path.basename(file)}`,output:session.project().markdown}
        const targets = session.findAllByText(args.find!)
        if (!targets.length) return `Text not found in ${path.basename(file)}: ${JSON.stringify(args.find)}`
        const results = targets.flatMap((target) => session.replaceTextRange(target.id,args.find!,args.replace!))
        const failed = results.find((result) => !result.success)
        if (failed) return `Edit failed: ${failed.error?.message ?? JSON.stringify(failed.error)}`
        await Bun.write(file,session.save())
        return {title:`Edited ${path.basename(file)}`,
          output:`Replaced ${results.length} occurrence(s) of ${JSON.stringify(args.find)} with ${JSON.stringify(args.replace)} in ${path.basename(file)}.`,
          metadata:{document:file,find:args.find,replace:args.replace,edits:results.length}}
      } finally {session.close()}
    }
    return args.action === "read" ? operate() : withArtifactMutationLocksAsync(ctx.directory,[file],operate)
  },
})
