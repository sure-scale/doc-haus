import { tool } from "./legacy-definition"
import { existsSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { recordRedlineProposal, withRedlineDocumentLock } from "../lib/redlines"
import { draftingAttribution } from "../lib/docx-attribution"

// doc.haus tracked-changes tool. Proposes a surgical find/replace edit to the
// matter's canonical Word (.docx) document as a tracked change, attributed to an
// author, that a reviewer accepts or rejects in the doc.haus app.
//
// The proposal is recorded against the matter's index; the canonical .docx stays
// clean (the accepted state) until a reviewer accepts, when ingest bakes the edit
// in. So here we only confirm the text exists — we never modify the file.

export default tool({
  description:
    "Propose a surgical edit to the matter's canonical Word (.docx) document as a tracked change. Finds exact text and proposes replacing it, attributed to an author. The change is recorded as a pending redline the user reviews and accepts or rejects in the doc.haus app — the document is not modified until they accept. Use this for a word/phrase swap; use redline for a whole-clause rewrite, and word-integration for silent (non-tracked) edits.",
  args: {
    document: tool.schema.string().describe("Document file name within the matter (the docPath from a citation)"),
    find: tool.schema.string().describe("The exact text to find"),
    replace: tool.schema.string().describe("The replacement text"),
    comment: tool.schema
      .string()
      .min(1)
      .max(2000)
      .describe("Lawyer-facing review comment explaining the deficiency or risk, the reason for the change, and its source or cross-reference. Stored with the pending change, never inserted into operative text."),
    author: tool.schema.string().optional().describe("Name to attribute the tracked change to (default: saved attorney, then firm, then doc.haus)"),
  },
  async execute(args, ctx) {
    const file = path.isAbsolute(args.document) ? args.document : path.join(ctx.directory, args.document)
    // doc.haus's own records (kept versions, the registry) are not matter documents.
    if (insideDochausRecords(ctx.directory, file) || !existsSync(file)) return `Document not found in this matter: ${args.document}`

    const dx = await docxodus()
    const session = dx.openDocxSession(await Bun.file(file).bytes(), {})

    const targets = session.findAllByText(args.find)
    if (!targets.length) {
      session.close()
      return `Text not found in ${path.basename(file)}: ${JSON.stringify(args.find)}`
    }
    session.close()

    const author = args.author ?? draftingAttribution().author
    // Recording the proposal is gated on the matter owner's approval (permission
    // "tracked-changes" in opencode.json) — the redline review queue is itself a
    // work product, so the assistant must not stack proposals into it unasked.
    // ctx.ask blocks until they reply and throws on reject. "Always" approves
    // future proposals against this document only.
    await ctx.ask({
      permission: "tracked-changes",
      patterns: [file],
      always: [file],
      metadata: { document: path.basename(file), find: args.find, replace: args.replace, comment: args.comment, author },
    })

    const proposal = await withRedlineDocumentLock(ctx.directory, file, async () => {
      if (!existsSync(file)) return { error: `Document not found in this matter: ${args.document}` }
      const current = dx.openDocxSession(await Bun.file(file).bytes(), {})
      const currentTargets = current.findAllByText(args.find)
      if (!currentTargets.length) {
        current.close()
        return { error: `Text changed before approval and is no longer present in ${path.basename(file)}.` }
      }
      const anchor = currentTargets[0].id
      current.close()
      const recorded = recordRedlineProposal(ctx.directory, {
        docPath: file,
        docName: path.basename(file),
        scope: "phrase",
        findText: args.find,
        oldText: args.find,
        newText: args.replace,
        author,
        reviewComment: args.comment,
        anchorId: anchor,
      })
      return { ...recorded, anchor, matches: currentTargets.length }
    })
    if ("error" in proposal) return proposal.error

    const superseded = proposal.superseded.length
      ? ` Supersedes pending redline${proposal.superseded.length === 1 ? "" : "s"} ${proposal.superseded.map((candidate) => `#${candidate.id}`).join(", ")} on the same passage — only this latest edit stays pending.`
      : ""
    return {
      title: `Proposed tracked change in ${path.basename(file)}`,
      output: `Proposed replacing ${JSON.stringify(args.find)} with ${JSON.stringify(args.replace)} in ${path.basename(file)} (${proposal.matches} occurrence(s)), attributed to ${author}. Recorded as pending redline #${proposal.id} with its lawyer-facing review comment — the user reviews and accepts or rejects it in the doc.haus app.${superseded}`,
      metadata: { document: file, find: args.find, replace: args.replace, comment: args.comment, matches: proposal.matches, author, redline: proposal.id, superseded: proposal.superseded.map((candidate) => candidate.id) },
    }
  },
})

function insideDochausRecords(directory: string, file: string) {
  return path.relative(path.resolve(directory), path.resolve(file)).split(path.sep)[0]?.toLowerCase() === ".dochaus"
}
