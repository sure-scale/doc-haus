import { createHash } from "node:crypto"
import { tool } from "./legacy-definition"
import { lstatSync, realpathSync } from "node:fs"
import path from "node:path"
import { docxodus } from "../lib/docxodus"
import { recordRedlineProposal, withRedlineDocumentLock } from "../lib/redlines"
import { draftingAttribution, ourAuthors } from "../lib/docx-attribution"
import { findPastListLabel, withoutListLabel } from "../lib/list-label"
import { applyLegacyClause, LegacyPhraseRefusal, hasTrackedRevisions, pendingRevisionAnchors, pendingRevisionRefusal, trackedRevisionsReason } from "../lib/docx-amendments"
import { validWordChatFormatting } from "../lib/word-chat-amendment-contract"
import { registeredWordRole } from "../lib/legal-db"

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
    "Propose rewriting a whole clause as a tracked change, using a passage retrieved from search-document. Locates the paragraph containing `clause` and proposes replacing its entire text with `replacement`, attributed to an author. The change is recorded as a pending redline the user reviews and accepts or rejects in the doc.haus app — the document is not modified until they accept. Use this to redline a clause you found as a citation; use tracked-changes for a surgical word/phrase swap. Optional formatting changes the complete existing paragraph together with the wording; provide the exact full paragraph as clause and plain replacement text. A list number Word prints before a paragraph (\"2.1\") is not text; leave it out of clause and replacement.",
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
    formatting: tool.schema.object({
      style: tool.schema.enum(["Title", "Heading1", "Heading2", "Heading3", "Normal"]).optional(),
      spacingBefore: tool.schema.number().optional(),
      spacingAfter: tool.schema.number().optional(),
      lineSpacing: tool.schema.number().optional(),
    }).strict().optional().describe("Optional paragraph formatting for this complete existing paragraph. Spacing is in points; lineSpacing is a multiple of single spacing. Structural changes are not supported."),
    author: tool.schema.string().optional().describe("Name to attribute the tracked change to (default: saved attorney, then firm, then doc.haus)"),
  },
  async execute(args, ctx) {
    const file = path.isAbsolute(args.document) ? args.document : path.join(ctx.directory, args.document)
    if (!matterDocument(ctx.directory, file)) return `Document not found in this matter: ${args.document}`

    const author = args.author ?? draftingAttribution().author
    // Our side's names, as Accept counts them (ingest's redline.ts).
    const own = [...ourAuthors(), author]
    // As Request edits: a Word file other than an uploaded source is refused
    // while it has tracked changes, beside which an untracked edit would not show.
    const bytes = await Bun.file(file).bytes()
    const role = registeredWordRole(ctx.directory, file)
    if (role !== "source" && hasTrackedRevisions(bytes)) return trackedRevisionsReason
    const dx = await docxodus()
    const session = dx.openDocxSession(bytes, {})

    const located = locate(session, args.clause, args.replacement, own)
    if (args.formatting) {
      const targetText = located.targets.length === 1 ? paragraphText(session, located.targets[0]!) : undefined
      const target = { id: "paragraph", text: targetText ?? "", paragraphs: [{ id: "paragraph", text: targetText ?? "" }], startOffset: 0, endOffset: targetText?.length ?? 0, capturedAt: "", saved: true }
      if (!targetText || args.clause !== targetText || located.replace !== args.replacement || /[\r\n\t\v]/.test(args.replacement) || !args.replacement ||
          !validWordChatFormatting(target, [{ paragraph: 0, ...args.formatting }], args.replacement)) {
        session.close()
        return "Paragraph formatting requires one uniquely matched complete paragraph and plain replacement text without structural changes."
      }
    }
    session.close()
    if (!located.targets.length) {
      // Search and read show a Word file's pending revisions as labels; changes
      // to such a file wait until they are resolved, as at accept.
      if (/\[Tracked (insertion|deletion|move)/.test(args.clause))
        return `${path.basename(file)} contains tracked changes. Accept or reject them in Word before proposing changes to it.`
      return `Clause not found in ${path.basename(file)}: ${JSON.stringify(args.clause)}`
    }
    if (located.error) return located.error
    const { find: clause, replace: replacement } = located
    const options = { pendingRevisions: "layer" as const, ownAuthors: own, revisionAuthor: author }
    let guarded: Awaited<ReturnType<typeof applyLegacyClause>> | undefined
    if (!args.formatting) {
      try { guarded = await applyLegacyClause(bytes, clause, replacement, options, "proposed") }
      catch (error) { return clauseRefusal(error) }
    }

    // Recording the proposal is gated on the matter owner's approval (permission
    // "redline" in opencode.json) — the redline review queue is itself a work
    // product, so the assistant must not stack proposals into it unasked.
    // ctx.ask blocks until they reply and throws on reject. "Always" approves
    // future proposals against this document only.
    await ctx.ask({
      permission: "redline",
      patterns: [file],
      always: [file],
      metadata: { document: path.basename(file), clause: guarded?.before ?? clause, replacement, comment: args.comment, author, formatting: args.formatting },
    })

    const proposal = await withRedlineDocumentLock(ctx.directory, file, async () => {
      if (!matterDocument(ctx.directory, file)) return { error: `Document not found in this matter: ${args.document}` }
      const fresh = await Bun.file(file).bytes()
      if (registeredWordRole(ctx.directory, file) !== role || (role !== "source" && hasTrackedRevisions(fresh))) return { error: trackedRevisionsReason }
      if (guarded && createHash("sha256").update(fresh).digest("hex") !== guarded.source_sha256) return { error: "The Word document changed while approval was pending. Prepare a fresh change." }
      const current = dx.openDocxSession(fresh, {})
      const again = locate(current, args.clause, args.replacement, own)
      const anchored = again.find === clause && !again.error ? again.targets[0] : undefined
      if (!anchored || (args.formatting && (again.targets.length !== 1 || paragraphText(current, anchored) !== args.clause))) {
        current.close()
        return { error: again.error ?? `Clause changed before approval and is no longer present in ${path.basename(file)}.` }
      }
      // projectAnchor emits the block's addressing token (`{#p:body:UNID}`)
      // ahead of its text. Strip those internal markers from reviewer-facing
      // old text while the document lock still protects the anchor bytes.
      const oldText = current.projectAnchor(anchored.id).markdown.replace(/\{#[^}]*\}/g, "").trim()
      current.close()
      if (guarded) {
        try { await applyLegacyClause(fresh, clause, replacement, options, "proposed") }
        catch (error) { return { error: clauseRefusal(error) } }
      }
      const recorded = recordRedlineProposal(ctx.directory, {
        docPath: file,
        docName: path.basename(file),
        scope: "clause",
        findText: guarded?.before ?? clause,
        oldText,
        newText: replacement,
        author,
        reviewComment: args.comment,
        anchorId: anchored.id,
        ...(guarded ? { targetJson: JSON.stringify({ version: 1, kind: "legacy_clause", source_sha256: guarded.source_sha256, target_handle: guarded.target_handle, amendment_id: "legacy-clause", prepared_role: role, prepared_path: realpathSync(file) }) } : {}),
        formatting: args.formatting,
      })
      return { ...recorded, oldText, anchor: anchored.id }
    })
    if ("error" in proposal) return proposal.error

    const superseded = proposal.superseded.length
      ? ` Supersedes pending redline${proposal.superseded.length === 1 ? "" : "s"} ${proposal.superseded.map((candidate) => `#${candidate.id}`).join(", ")} on the same clause — only this latest rewrite stays pending.`
      : ""
    return {
      title: `Proposed redline in ${path.basename(file)}`,
      output: `Proposed rewriting the clause matching ${JSON.stringify(clause)} in ${path.basename(file)}, attributed to ${author}. Recorded as pending redline #${proposal.id} with its lawyer-facing review comment — the user reviews and accepts or rejects it in the doc.haus app.${superseded}`,
      metadata: { document: file, clause, oldText: proposal.oldText, replacement, comment: args.comment, formatting: args.formatting, anchor: proposal.anchor, author, redline: proposal.id, superseded: proposal.superseded.map((candidate) => candidate.id) },
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

/** The paragraph `clause` locates, past a copied list number (findPastListLabel).
 * Accept rewrites the whole of the first paragraph holding the recorded
 * clause text, so that paragraph is refused while it has a pending revision:
 * someone else's, or one of `own` the rewrite would take with it (the touch rule). */
function locate(session: ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>, clause: string, replacement: string, own: readonly string[]) {
  const located = findPastListLabel(clause, replacement, text => session.findAllByText(text, { ignoreWhitespace: true }), id => session.getAnchorInfo(id)?.visibleText, { recorded: true, whole: true })
  const refused = located.targets.length && !located.error ? pendingRevisionRefusal(session, pendingRevisionAnchors(session), [located.targets[0]!.id], located.find, "proposed", own, true) : undefined
  return refused ? { ...located, error: refused } : located
}

// Docxodus displays Word's automatic list label in visibleText, but that label
// is not paragraph content and must not become a replay target.
function paragraphText(session: ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>, target: { id: string; autoNumberPrefix?: string }) {
  const shown = session.getAnchorInfo(target.id)?.visibleText
  return shown === undefined ? undefined : withoutListLabel(shown, target.autoNumberPrefix?.trim())
}

function clauseRefusal(error: unknown) {
  if (error instanceof LegacyPhraseRefusal) return error.message
  if (error instanceof Error && /^docx_field_replacement(?::|$)/.test(error.message)) return "This change touches a Word field that could not be safely preserved. Make the change in Word, then prepare a fresh change."
  if (error instanceof Error && error.message === "unsupported_legacy_clause_markup") return "This clause uses formatting or structure that cannot be safely reproduced. Make the change in Word, or request plain text."
  return "This clause change cannot be safely prepared. Review the passage in Word and prepare a fresh change."
}
