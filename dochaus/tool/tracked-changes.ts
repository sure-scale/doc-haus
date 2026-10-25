import { tool } from "./legacy-definition"
import { existsSync, lstatSync, realpathSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { recordRedlineProposal, withRedlineDocumentLock } from "../lib/redlines"
import { draftingAttribution } from "../lib/docx-attribution"
import { findPastListLabel } from "../lib/list-label"
import { pendingRevisionAnchors, pendingRevisionRefusal } from "../lib/docx-amendments"

// doc.haus tracked-changes tool. Proposes a surgical find/replace edit to the
// matter's canonical Word (.docx) document as a tracked change, attributed to an
// author, that a reviewer accepts or rejects in the doc.haus app.
//
// The proposal is recorded against the matter's index; the canonical .docx stays
// clean (the accepted state) until a reviewer accepts, when ingest bakes the edit
// in. So here we only confirm the text exists — we never modify the file.

export default tool({
  description:
    "Propose a surgical edit to the matter's canonical Word (.docx) document as a tracked change. Finds exact text and proposes replacing it, attributed to an author. The change is recorded as a pending redline the user reviews and accepts or rejects in the doc.haus app — the document is not modified until they accept. Use this for a word/phrase swap; use redline for a whole-clause rewrite, and word-integration for silent (non-tracked) edits. A list number Word prints before a paragraph (\"2.1\") is not text; leave it out of find and replace.",
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
    if (!matterDocument(ctx.directory, file)) return `Document not found in this matter: ${args.document}`

    const dx = await docxodus()
    const session = dx.openDocxSession(await Bun.file(file).bytes(), {})

    const located = locate(session, args.find, args.replace)
    session.close()
    if (!located.targets.length) return `Text not found in ${path.basename(file)}: ${JSON.stringify(args.find)}`
    if (located.error) return located.error
    const { find, replace } = located

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
      metadata: { document: path.basename(file), find, replace, comment: args.comment, author },
    })

    const proposal = await withRedlineDocumentLock(ctx.directory, file, async () => {
      if (!existsSync(file)) return { error: `Document not found in this matter: ${args.document}` }
      const current = dx.openDocxSession(await Bun.file(file).bytes(), {})
      const again = locate(current, args.find, args.replace)
      // Accepting the change replaces every occurrence of the recorded text.
      const matches = again.find === find ? current.findAllByText(find).length : 0
      current.close()
      if (again.error) return { error: again.error }
      if (!again.targets.length || !matches) return { error: `Text changed before approval and is no longer present in ${path.basename(file)}.` }
      const anchor = again.targets[0]!.id
      const recorded = recordRedlineProposal(ctx.directory, {
        docPath: file,
        docName: path.basename(file),
        scope: "phrase",
        findText: find,
        oldText: find,
        newText: replace,
        author,
        reviewComment: args.comment,
        anchorId: anchor,
      })
      return { ...recorded, anchor, matches }
    })
    if ("error" in proposal) return proposal.error

    const superseded = proposal.superseded.length
      ? ` Supersedes pending redline${proposal.superseded.length === 1 ? "" : "s"} ${proposal.superseded.map((candidate) => `#${candidate.id}`).join(", ")} on the same passage — only this latest edit stays pending.`
      : ""
    return {
      title: `Proposed tracked change in ${path.basename(file)}`,
      output: `Proposed replacing ${JSON.stringify(find)} with ${JSON.stringify(replace)} in ${path.basename(file)} (${proposal.matches} occurrence(s)), attributed to ${author}. Recorded as pending redline #${proposal.id} with its lawyer-facing review comment — the user reviews and accepts or rejects it in the doc.haus app.${superseded}`,
      metadata: { document: file, find, replace, comment: args.comment, matches: proposal.matches, author, redline: proposal.id, superseded: proposal.superseded.map((candidate) => candidate.id) },
    }
  },
})

// Ingest registers documents by name in the matter folder itself: only a file
// saved directly there, by its real path and not through a link, is a matter
// document. doc.haus's own records (.dochaus/: kept versions, the registry)
// never are, under any spelling of their path.
function matterDocument(directory: string, file: string) {
  try {
    return !lstatSync(file).isSymbolicLink() && path.dirname(realpathSync(file)) === realpathSync(directory)
  } catch {
    return false
  }
}

/** Where `find` is in the document, past a copied list number (findPastListLabel).
 * Accept replaces the recorded text wherever it occurs, untracked, so an
 * occurrence in a paragraph with a pending revision is refused (the touch rule). */
function locate(session: ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>, find: string, replace: string) {
  const located = findPastListLabel(find, replace, text => session.findAllByText(text), id => session.getAnchorInfo(id)?.visibleText, { recorded: true })
  const refused = located.targets.length && !located.error ? pendingRevisionRefusal(pendingRevisionAnchors(session), located.targets.map(target => target.id), located.find, "proposed") : undefined
  return refused ? { ...located, error: refused } : located
}
