import { tool } from "./legacy-definition"
import { existsSync, lstatSync, realpathSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { recordRedlineProposal, withRedlineDocumentLock } from "../lib/redlines"
import { draftingAttribution } from "../lib/docx-attribution"

// doc.haus redline tool. Proposes rewriting a whole clause — the paragraph a
// search-document citation points at — as a tracked change a reviewer accepts or
// rejects in the doc.haus app.
//
// This is the citation-driven, clause-level counterpart to the surgical
// find/replace tools (word-integration, tracked-changes): `clause` locates the
// paragraph and the whole paragraph is rewritten to `replacement`. We match on
// the block's flat text, whitespace-tolerant, because citation excerpts come
// from a different text extraction (mammoth) than Docxodus' own projection and
// their spacing/char offsets don't align — so we only ever LOCATE the block, we
// never offset-index into it.
//
// The proposal is recorded against the matter's index; the canonical .docx stays
// clean (the accepted state) until a reviewer accepts, when ingest bakes the edit
// in. So here we only locate the clause and capture its current text — we never
// modify the file.

export default tool({
  description:
    "Propose rewriting a whole clause as a tracked change, using a passage retrieved from search-document. Locates the paragraph containing `clause` and proposes replacing its entire text with `replacement`, attributed to an author. The change is recorded as a pending redline the user reviews and accepts or rejects in the doc.haus app — the document is not modified until they accept. Use this to redline a clause you found as a citation; use tracked-changes for a surgical word/phrase swap.",
  args: {
    document: tool.schema.string().describe("Document file name within the matter (the docPath from a citation)"),
    clause: tool.schema
      .string()
      .describe("Text from the clause to rewrite — the citation excerpt, or a sentence within it. Used to locate the paragraph."),
    replacement: tool.schema
      .string()
      .describe("The new clause text. Replaces the whole located paragraph; markdown is supported."),
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

    const target = session.findByText(args.clause, { ignoreWhitespace: true })
    if (!target) {
      session.close()
      // Search and read show a Word file's pending revisions as labels; changes
      // to such a file wait until they are resolved, as at accept.
      if (/\[Tracked (insertion|deletion|move)/.test(args.clause))
        return `${path.basename(file)} contains tracked changes. Accept or reject them in Word before proposing changes to it.`
      return `Clause not found in ${path.basename(file)}: ${JSON.stringify(args.clause)}`
    }
    session.close()

    const author = args.author ?? draftingAttribution().author
    // Recording the proposal is gated on the matter owner's approval (permission
    // "redline" in opencode.json) — the redline review queue is itself a work
    // product, so the assistant must not stack proposals into it unasked.
    // ctx.ask blocks until they reply and throws on reject. "Always" approves
    // future proposals against this document only.
    await ctx.ask({
      permission: "redline",
      patterns: [file],
      always: [file],
      metadata: { document: path.basename(file), clause: args.clause, replacement: args.replacement, comment: args.comment, author },
    })

    const proposal = await withRedlineDocumentLock(ctx.directory, file, async () => {
      if (!existsSync(file)) return { error: `Document not found in this matter: ${args.document}` }
      const current = dx.openDocxSession(await Bun.file(file).bytes(), {})
      const anchored = current.findByText(args.clause, { ignoreWhitespace: true })
      if (!anchored) {
        current.close()
        return { error: `Clause changed before approval and is no longer present in ${path.basename(file)}.` }
      }
      // projectAnchor emits the block's addressing token (`{#p:body:UNID}`)
      // ahead of its text. Strip those internal markers from reviewer-facing
      // old text while the document lock still protects the anchor bytes.
      const oldText = current.projectAnchor(anchored.id).markdown.replace(/\{#[^}]*\}/g, "").trim()
      current.close()
      const recorded = recordRedlineProposal(ctx.directory, {
        docPath: file,
        docName: path.basename(file),
        scope: "clause",
        findText: args.clause,
        oldText,
        newText: args.replacement,
        author,
        reviewComment: args.comment,
        anchorId: anchored.id,
      })
      return { ...recorded, oldText, anchor: anchored.id }
    })
    if ("error" in proposal) return proposal.error

    const superseded = proposal.superseded.length
      ? ` Supersedes pending redline${proposal.superseded.length === 1 ? "" : "s"} ${proposal.superseded.map((candidate) => `#${candidate.id}`).join(", ")} on the same clause — only this latest rewrite stays pending.`
      : ""
    return {
      title: `Proposed redline in ${path.basename(file)}`,
      output: `Proposed rewriting the clause matching ${JSON.stringify(args.clause)} in ${path.basename(file)}, attributed to ${author}. Recorded as pending redline #${proposal.id} with its lawyer-facing review comment — the user reviews and accepts or rejects it in the doc.haus app.${superseded}`,
      metadata: { document: file, clause: args.clause, oldText: proposal.oldText, replacement: args.replacement, comment: args.comment, anchor: proposal.anchor, author, redline: proposal.id, superseded: proposal.superseded.map((candidate) => candidate.id) },
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
