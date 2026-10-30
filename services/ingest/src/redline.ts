import { docxodus } from "./docxodus"
import { applyDocxParagraphFormatting } from "../../../dochaus/lib/docx-paragraph-formatting"
import type { WordChatParagraphFormat } from "../../../dochaus/lib/word-chat-amendment-contract"
import { compareDocxWithPointers, paragraphTargets, resolveParagraphTarget, withStableRevisionDates, type ParagraphTarget } from "../../../dochaus/lib/drafting-amendment"
import type { RedlineRow } from "./db"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { createDocxLocalization, applyDocxParagraphEdits, applyLegacyPhrase, applyLegacyClause, LegacyPhraseRefusal, DocxOverlappingAmendmentsError, DocxTrackedRevisionsError, hasTrackedRevisions, openRevision, othersPending, pendingRevisionAnchors, pendingRevisionRefusal, quoted, type DocxRevisionOptions } from "../../../dochaus/lib/docx-amendments"
import { createHash } from "node:crypto"
import { draftingAttribution, ourAuthors } from "../../../dochaus/lib/docx-attribution"
import { normalizeUntrustedText } from "../../../dochaus/lib/untrusted-text"
import { listNames } from "./document-versions"
import { exportDocxComments, preserveDocxComments, readDocxCommentParagraphPositions, readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"
import { assertLayeredRoundTrip, assertRedlineRoundTrip } from "../../../dochaus/lib/redline-round-trip"
import { discloseSourceRevisions } from "../../../dochaus/lib/docx-revision-views"
import { docxProtectionMessage } from "./docx-protection-message"
import { paragraphInsertionProposalMarker } from "../../../dochaus/lib/paragraph-insertion-proposal"
import { authorParagraphInsertionReview, resolveParagraphInsertionReview } from "../../../dochaus/lib/docx-paragraph-insertion-review"
import { checkedParagraphInsertionOperation, paragraphInsertionBinding, ParagraphInsertionProposalError, insertionChanged, insertionSourceOnly, insertionUnsupported } from "./paragraph-insertion-proposals"
import { assertExclusiveRedlineSingleton, exclusiveRedlineProposalMarker, RedlineProposalConflictError, tableRowProposalMarker, legacyClauseProposalMarker } from "../../../dochaus/lib/table-row-proposal"
import { paragraphSpacingProposalMarker } from "../../../dochaus/lib/paragraph-spacing-proposal"
import { authorExactParagraphSpacing, resolveExactParagraphSpacing } from "../../../dochaus/lib/docx-exact-spacing"
import { authorExactParagraphSpacingReview } from "../../../dochaus/lib/docx-exact-spacing-review"
import { checkedParagraphSpacingOperation, paragraphSpacingBinding, ParagraphSpacingProposalError, spacingChanged, spacingSourceOnly, spacingUnsupported } from "./paragraph-spacing-proposals"
import { authorDocxTableRowInsertion, resolveDocxTableRowInsertion } from "../../../dochaus/lib/docx-table-row-insertion"
import { authorDocxTableRowReview } from "../../../dochaus/lib/docx-table-row-review"
import { authorDocxTableRowDeletion, resolveDocxTableRowDeletion } from "../../../dochaus/lib/docx-table-row-deletion"
import { authorDocxTableRowDeletionReview } from "../../../dochaus/lib/docx-table-row-deletion-review"
import { assertTableRowSingleton, checkedTableRowOperation, rowChanged, rowSourceOnly, tableRowBinding, tableRowSourceHasRevisions, TableRowProposalError } from "./table-row-proposals"

// Replay and bake pending redline proposals against a .docx.
//
// The canonical document on disk is the accepted state: clean, except that an
// uploaded source keeps the other side's pending revisions and takes accepted
// edits beside them as tracked changes (bake). A proposal records how to reproduce one edit — a surgical find/replace ('phrase')
// or a whole-paragraph rewrite ('clause') — located the same way the dochaus tools
// located it originally. We never store offsets: text is re-resolved against the
// live document every time, so edits stay valid as the document changes underneath.

type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
export class RedlinePreviewConflictError extends Error {}
/** A batch Accept refused for a reason the lawyer can act on; its message says which proposals and what to do. */
export class RedlineAcceptConflictError extends Error {}
/** An older proposal its text replay could not write. The message is the
 * replay's reason; `refusal` is Accept all's, which names the proposal. */
export class RedlineReplayError extends Error {
  readonly refusal: string
  constructor(row: RedlineRow, reason: string, actionable = false, readonly skippable = false) {
    super(reason)
    // The touch rule's refusal already names the change and what to do.
    this.refusal = actionable || reason.startsWith("The change to ") ? reason
      : `The change to ${quoted(row.find_text)} cannot be written with the rest: its words have changed, or another proposal in this list changes them. Reject it, accept the rest, then request that edit again.`
  }
}
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

// Ingest's own install of the same Docxodus package declares the session type
// again; the touch rule only calls the session's methods.
const shared = (session: Session) => session as unknown as Parameters<typeof pendingRevisionAnchors>[0]
// A replay session's paragraphs with a pending revision, read once: its
// untracked edits add none.
const pendingRevisions = new WeakMap<Session, Map<string, string[]>>()
const pendingIn = (session: Session) => pendingRevisions.get(session)
  ?? pendingRevisions.set(session, pendingRevisionAnchors(shared(session))).get(session)!

// Reproduce one proposal's edit on an open session. Returns an error message
// (rather than throwing) when the anchor text can no longer be found — e.g. an
// earlier proposal in the same replay already rewrote the passage this one
// anchors to. Callers decide whether that is fatal: the redlined view skips the
// row and renders the rest; accept must fail loudly rather than bake a partial set.
// It is refused in a paragraph with someone else's pending revision (the touch
// rule): untracked, doc.haus's words would become that revision's author's. Our
// own side's (ours) hold no paragraph, as in Request edits, but an edit inside
// one, a whole-clause rewrite of its paragraph included, is refused.
function applyProposal(session: Session, row: RedlineRow): { ok: true } | { ok: false; error: string } {
  if (row.scope === "clause") {
    const target = session.findByText(row.find_text, { ignoreWhitespace: true })
    if (!target) return { ok: false, error: `Clause no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` }
    const refused = pendingRevisionRefusal(shared(session), pendingIn(session), [target.id], row.find_text, "applied", ours([rowAuthor(row)]), true)
    if (refused) return { ok: false, error: refused }
    const result = session.replaceText(target.id, row.new_text)
    if (!result.success) return { ok: false, error: `Redline #${row.id} failed: ${result.error?.message ?? "unknown error"}` }
    return { ok: true }
  }

  const targets = session.findAllByText(row.find_text)
  if (!targets.length) return { ok: false, error: `Text no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` }
  const refused = pendingRevisionRefusal(shared(session), pendingIn(session), targets.map((t) => t.id), row.find_text, "applied", ours([rowAuthor(row)]))
  if (refused) return { ok: false, error: refused }
  const results = targets.flatMap((t) => session.replaceTextRange(t.id, row.find_text, row.new_text))
  const failed = results.find((r) => !r.success)
  if (failed) return { ok: false, error: `Redline #${row.id} failed: ${failed.error?.message ?? "unknown error"}` }
  return { ok: true }
}

// The redlined view the viewer renders and the download with proposals: the
// document with every pending proposal as native w:ins/w:del the browser paints
// green/red. With no pending rows it is the document unchanged.
//
// A list with revision-bound proposals is the file with them written in place
// as tracked changes, as Accept writes them (trackedExport), older proposals in
// it included, on a plain file and an uploaded markup alike; on a plain file,
// a list with an older clause rewrite is what Accept writes compared with the
// file. Older proposals alone are replayed and compared with the document;
// without review notes they retain the tolerant preview replay. Revision-bound
// proposals and annotated proposals fail closed: exporting only part of their
// tracked changes or silently losing a stored note is misleading.
// `state` is the request's redlineState for these bytes, when already read.
// `options` is how the document is read.
export async function buildRedlined(original: Uint8Array, rows: RedlineRow[], state?: RedlineState, options: DocxRevisionOptions = {}): Promise<Uint8Array> {
  if (!rows.length) return original
  if (rows.some(paragraphInsertionProposalMarker)) {
    try {const row=await checkedInsertion(original,rows,state),operation=await checkedParagraphInsertionOperation(original,row);return (await authorParagraphInsertionReview(original,operation.request,row.review_comment)).bytes}
    catch(error){throw new RedlinePreviewConflictError(insertionFailure(error))}
  }
  if (rows.some(paragraphSpacingProposalMarker)) {
    try {
      const row = await checkedSpacing(original, rows, state), operation = await checkedParagraphSpacingOperation(original, row)
      return (await authorExactParagraphSpacingReview(original, operation.request, row.review_comment)).bytes
    } catch (error) { throw new RedlinePreviewConflictError(spacingFailure(error)) }
  }
  if (rows.some(tableRowProposalMarker)) {
    try {
      const row = await checkedTableRow(original, rows, state)
      const operation = await checkedTableRowOperation(original,row)
      return operation.kind === "table_row_delete"
        ? (await authorDocxTableRowDeletionReview(original,operation.request,row.review_comment)).bytes
        : (await authorDocxTableRowReview(original,operation.request,row.review_comment)).bytes
    } catch (error) { throw new RedlinePreviewConflictError(rowFailure(error)) }
  }
  const annotated = rows.some(row => row.review_comment.trim())
  // Paragraph comments cannot safely follow a structural replay by ordinal.
  if (annotated && rows.some(row => /[\r\n\t]/.test(row.find_text + row.new_text)))
    throw new RedlinePreviewConflictError("Word review comments require paragraph-preserving proposals. Prepare a fresh change list without structural replacements.")
  if (rows.some((row) => row.formatting_json != null)) {
    try { return await trackedExport(original, rows, state, { ...options, pendingRevisions: "layer" }) }
    catch (error) {
      if (error instanceof RedlinePreviewConflictError) throw error
      if (error instanceof RedlineReplayError) throw new RedlinePreviewConflictError(error.refusal)
      const protection = docxProtectionMessage(error)
      if (protection) throw new RedlinePreviewConflictError(protection)
      throw new RedlinePreviewConflictError("A formatting proposal cannot be safely previewed. Prepare a fresh change list against the current document.")
    }
  }
  if (rows.some((row) => row.target_json != null)) {
    try {
      // An older clause rewrite replaces its whole paragraph, which a tracked
      // write refuses where the paragraph holds a bookmark, as every deliverable
      // paragraph holds its pointer's. A file without tracked changes takes
      // Accept's edits untracked, so the comparison shows them.
      if (rows.some((row) => row.target_json == null && row.scope === "clause") && !(await (await docxodus()).getRevisions(original)).length) {
        const written = await writeProposals(original, rows, state, options, { locate: annotated })
        const paragraphs = readDocxCommentParagraphPositions(written.bytes)
        return commentRedline(original, written.bytes, rows, written.positions.map(at => paragraphs.findIndex(row => at !== undefined && row.xml_paragraph_indexes.includes(at))))
      }
      // The tracked write reads its own changes back beside any the file has.
      return await trackedExport(original, rows, state, { ...options, pendingRevisions: "layer" })
    } catch (error) {
      if (error instanceof RedlinePreviewConflictError) throw error
      if (error instanceof RedlineReplayError) throw new RedlinePreviewConflictError(error.refusal)
      const protection = docxProtectionMessage(error)
      if (protection) throw new RedlinePreviewConflictError(protection)
      throw new RedlinePreviewConflictError("A proposal no longer resolves against this Word revision. Prepare a fresh change list before exporting its tracked changes and review comments.")
    }
  }
  try {
    if (!annotated) {
      let bytes = original
      const applied: RedlineRow[] = []
      for (const row of rows) {
        try { bytes = (await bakeOlder(bytes, [row], false, false, options)).bytes; applied.push(row) }
        catch (error) {
          if (!(error instanceof RedlineReplayError) || !error.skippable) throw error
          console.warn("[redline] skipping in redlined view", diagnosticErrorMetadata(error.message))
        }
      }
      return applied.length ? await commentRedline(original, bytes, applied, applied.map(() => -1)) : original
    }
    const written = await writeProposals(original, rows, state, options, { locate: true })
    const paragraphs = readDocxCommentParagraphPositions(written.bytes)
    return await commentRedline(original, written.bytes, rows, written.positions.map(at => paragraphs.findIndex(row => at !== undefined && row.xml_paragraph_indexes.includes(at))))
  } catch (error) {
    if (error instanceof RedlineReplayError) throw new RedlinePreviewConflictError(error.refusal)
    const protection = docxProtectionMessage(error)
    throw new RedlinePreviewConflictError(protection ?? "The proposals cannot be safely previewed. Prepare a fresh change list.")
  }
}

async function commentRedline(original: Uint8Array, modified: Uint8Array, rows: RedlineRow[], targets: number[]) {
  // Docxodus dates every compared revision 2000-01-01; an export is dated when made.
  const compared = withStableRevisionDates(await compareDocxWithPointers(original, modified, { authorName: previewAuthor(rows) }), new Date().toISOString())
  const redline = (() => {
    try { return preserveDocxComments(original, compared) }
    catch { throw new RedlinePreviewConflictError("An existing Word comment could not be preserved at its original unchanged anchor or with its original links and formatting dependencies. Review comments on edited ranges, complex fields, objects, threaded comments and attachments separately before exporting.") }
  })()
  const finish = async (bytes: Uint8Array) => {
    try {
      const output = await discloseSourceRevisions(bytes, original, draftingAttribution().author)
      await assertRedlineRoundTrip(output.bytes, original, modified)
      return output.bytes
    }
    catch (error) { throw new RedlinePreviewConflictError(error instanceof Error ? error.message : "The Word redline could not be checked against its source and intended text.") }
  }
  if (!rows.some(row => row.review_comment.trim())) return finish(redline)
  const paragraphs = readDocxCommentParagraphs(modified)
  const exported = readDocxCommentParagraphs(redline)
  if (JSON.stringify(paragraphs.map(row => row.paragraph_text)) !== JSON.stringify(exported.map(row => row.paragraph_text)))
    throw new RedlinePreviewConflictError("The Word comparison changed paragraph boundaries; review comments cannot be attached safely. Prepare a fresh change list.")
  return finish(withReviewComments(redline, rows, targets.map(target => paragraphs[target])))
}

// Each annotated row's review comment, under its stored author, on its
// accepted paragraph of `bytes` (paragraphs[i] for rows[i]).
function withReviewComments(bytes: Uint8Array, rows: RedlineRow[], paragraphs: ({ paragraph_index: number; paragraph_text: string } | undefined)[]) {
  const comments = rows.flatMap((row, index) => {
    if (!row.review_comment.trim()) return []
    const paragraph = paragraphs[index]
    if (!paragraph || !row.author.trim())
      throw new RedlinePreviewConflictError(`Review comment for proposal #${row.id} has no supported body paragraph or stored author. Prepare a fresh change list before exporting.`)
    return [{ id: String(row.id), text: row.review_comment, author: row.author,
      anchor: { paragraph_index: paragraph.paragraph_index, paragraph_text: paragraph.paragraph_text } }]
  })
  try {
    const output = exportDocxComments(bytes, comments)
    if (output.unresolved.length)
      throw new RedlinePreviewConflictError(`Word review comments could not be attached: ${output.unresolved.map(row => `proposal #${row.id}: ${row.reason}`).join(" ")}`)
    return output.bytes
  } catch (error) {
    if (error instanceof RedlinePreviewConflictError) throw error
    throw new RedlinePreviewConflictError("The stored Word review comments could not be exported safely. Check their text, authors and document anchors before exporting again.")
  }
}

/** Whether these proposals export layered on the document's own pending
 * revisions: revision-bound proposals read in layer mode (an uploaded source).
 * Nothing of theirs is accepted, so such an export has no disclosure. */
export function layeredRedline(rows: RedlineRow[], options: DocxRevisionOptions) {
  return options.pendingRevisions === "layer" && rows.some(row => row.target_json != null)
}

// The file with the proposals written into it in place, as native tracked
// changes each under its proposal's author, and their review comments:
// accepting those changes gives what Accept writes. On the other side's markup
// their revisions and comments are kept as they are; comparison would accept
// them. Track Changes is left on, as a returned redline travels.
async function trackedExport(original: Uint8Array, rows: RedlineRow[], state: RedlineState | undefined, options: DocxRevisionOptions) {
  const ownAuthors = ours(rows.map(rowAuthor))
  const annotated = rows.some(row => row.review_comment.trim())
  const baked = await writeProposals(original, rows, state, options, { ownAuthors })
  const tracked = await writeProposals(original, rows, state, options, { ownAuthors, tracked: true, locate: annotated })
  let output = tracked.bytes
  if (annotated) {
    // The edited paragraph's position finds its accepted row, which a deleted
    // mark may join to another.
    const accepted = readDocxCommentParagraphPositions(output)
    output = withReviewComments(output, rows, tracked.positions.map(at => at === undefined ? undefined : accepted.find(row => row.xml_paragraph_indexes.includes(at))))
  }
  try { await assertLayeredRoundTrip(output, original, baked.bytes) }
  catch (error) { throw new RedlinePreviewConflictError(error instanceof Error ? error.message : "The Word redline could not be checked against its source and intended text.") }
  return output
}

// The position among the main body's w:p elements of each body paragraph of
// this revision, by paragraph_handle.
async function bodyParagraphPositions(bytes: Uint8Array, options: DocxRevisionOptions) {
  const opened = await openRevision(new Uint8Array(bytes), options)
  try {
    const positions = new Map(Array.from(opened.documents["word/document.xml"]!.getElementsByTagNameNS(W, "p")).map((node, index) => [node, index]))
    return new Map(opened.paragraphs.flatMap((paragraph, ordinal) => paragraph.part === "word/document.xml"
      ? [[opened.revision.paragraphs[ordinal]!.paragraph_handle, positions.get(paragraph.node)!] as const] : []))
  } finally { opened.session.close() }
}

// Our own side's names on a document: every name doc.haus signs with and the
// proposals' recorded authors. Their pending changes hold no paragraph.
const ours = (authors: readonly string[]) => [...ourAuthors(), ...authors]
const rowAuthor = (row: RedlineRow) => row.author || draftingAttribution().author
// Each proposal's edit tracked under its own recorded author.
const trackedAs = (rows: RedlineRow[]) => ({ revisionAuthor: rowAuthor(rows[0]!),
  amendmentAuthors: Object.fromEntries(rows.map((row) => [String(row.id), rowAuthor(row)])) })

function previewAuthor(rows: RedlineRow[]) {
  const authors = [...new Set(rows.map(rowAuthor))]
  // Docxodus comparisons accept only one author for every new native revision.
  // Refuse a combined export rather than silently overwrite stored identities.
  if (authors.length > 1)
    throw new RedlinePreviewConflictError("These proposals have different authors and cannot share one Word redline preview. Review proposals from one author at a time or prepare a fresh change list.")
  return authors[0]!
}

// Bake the given proposals into the document, returning its new canonical bytes:
// the new accepted state. Accept fails loudly: a proposal that no longer
// resolves throws rather than silently baking a partial set into the canonical
// document. `state` is the request's redlineState for these bytes, when already
// read for the stale check. A clean document takes the edits untracked. In layer
// mode (an uploaded source) a document with pending revisions keeps them
// pending and takes the edits beside them as native tracked changes, each under
// its proposal's author (an attorney named between two Prepares leaves two), so
// the accepted file is itself the redline that goes back; the round trip proves
// rejecting only those gives the document as it was, and accepting them the
// same edits untracked.
export async function bake(original: Uint8Array, rows: RedlineRow[], state?: RedlineState, options: DocxRevisionOptions = {}): Promise<Uint8Array> {
  if (rows.some(paragraphInsertionProposalMarker)) {
    try {const row=await checkedInsertion(original,rows,state),operation=await checkedParagraphInsertionOperation(original,row),authored=await authorParagraphInsertionReview(original,operation.request,row.review_comment);return await resolveParagraphInsertionReview(original,authored.bytes,operation.request,row.review_comment,"accept")}
    catch(error){throw new RedlineAcceptConflictError(insertionFailure(error))}
  }
  if (rows.some(paragraphSpacingProposalMarker)) {
    try {
      const row = await checkedSpacing(original, rows, state), operation = await checkedParagraphSpacingOperation(original, row)
      const authored = await authorExactParagraphSpacing(original, operation.request)
      return await resolveExactParagraphSpacing(original, authored.trackedBytes, operation.request, "accept")
    } catch (error) { throw new RedlineAcceptConflictError(spacingFailure(error)) }
  }
  if (rows.some(tableRowProposalMarker)) {
    try {
      const row = await checkedTableRow(original, rows, state), operation = await checkedTableRowOperation(original,row)
      if (operation.kind === "table_row_delete") {
        const authored = await authorDocxTableRowDeletion(original,operation.request)
        return await resolveDocxTableRowDeletion(original,authored.bytes,operation.request,"accept")
      }
      const authored = await authorDocxTableRowInsertion(original, operation.request)
      return await resolveDocxTableRowInsertion(original, authored.bytes, operation.request, "accept")
    } catch (error) { throw new RedlineAcceptConflictError(rowFailure(error)) }
  }
  try {
    if (options.pendingRevisions !== "layer" || !(await (await docxodus()).getRevisions(original)).length)
      return (await writeProposals(original,rows,state,options)).bytes
    // Every proposal's author is our side, whichever name it was prepared under.
    const ownAuthors = ours(rows.map(rowAuthor))
    const plain = (await writeProposals(original,rows,state,options,{ownAuthors})).bytes
    const tracked = (await writeProposals(original,rows,state,options,{ownAuthors,tracked:true})).bytes
    await assertLayeredRoundTrip(tracked,original,plain)
    return tracked
  } catch (error) {
    if (!(error instanceof DocxOverlappingAmendmentsError)) throw error
    const [first,second] = rows.filter((row) => error.amendments.includes(String(row.id)))
    throw new RedlineAcceptConflictError(`The changes to ${quoted(first!.find_text)} and ${quoted(second!.find_text)} overlap, so Accept all cannot write both. Accept the one you want on its own, then reject the other and request that edit again.`)
  }
}

// The proposals written as Accept writes them, for bake and the preview alike:
// revision-bound ones against these bytes, then older ones replayed by their
// text on the result, as accepting each in turn would. `tracked` writes each
// as native tracked changes under its proposal's author. With `locate`, also
// each annotated row's edited paragraph, by its position among the body's w:p
// elements, for its review comment: an edit keeps every w:p in place.
async function writeProposals(original: Uint8Array, rows: RedlineRow[], state: RedlineState | undefined, options: DocxRevisionOptions,
  how: { ownAuthors?: readonly string[]; tracked?: boolean; locate?: boolean } = {}) {
  assertExclusiveRedlineSingleton(rows)
  const bound = rows.filter((row) => row.target_json != null)
  const older = rows.filter((row) => row.target_json == null)
  const written = bound.length ? await bakeBound(original, bound, state,
    { ...options, ...how.ownAuthors ? { ownAuthors: how.ownAuthors } : {}, ...how.tracked ? trackedAs(bound) : {} }) : undefined
  const replayed = older.length ? await bakeOlder(written?.bytes ?? original, older, how.tracked, how.locate, options) : undefined
  const bytes = replayed?.bytes ?? written?.bytes ?? original
  if (!how.locate) return { bytes, positions: [] }
  const handles = written && await bodyParagraphPositions(original, options)
  return { bytes, positions: rows.map((row) => row.target_json != null
    ? handles!.get(written!.bindings[bound.indexOf(row)]!.paragraph_handle)
    : replayed!.positions[older.indexOf(row)]) }
}

// Older proposals, replayed by their text; `tracked`, as native tracked changes
// under each proposal's author. With `locate`, an annotated row's position is
// that of the one accepted paragraph its replay changes.
async function bakeOlder(original: Uint8Array, rows: RedlineRow[], tracked = false, locate = false, options: DocxRevisionOptions = {}) {
  const dx = await docxodus()
  let bytes = original
  const positions: Array<number | undefined> = []
  for (const row of rows) {
    if (row.formatting_json != null) {
      if (row.scope !== "clause" || row.target_json != null || /[\r\n\t]/.test(row.find_text + row.new_text))
        throw new RedlineReplayError(row, "Formatting requires one complete existing paragraph.")
      let format: Omit<WordChatParagraphFormat, "paragraph">
      try { format = JSON.parse(row.formatting_json) } catch { throw new RedlineReplayError(row, "The formatting proposal is invalid.") }
      const opened = await openRevision(bytes, { ...options, pendingRevisions: "layer", ownAuthors: ours([rowAuthor(row)]),
        ...(tracked ? { revisionAuthor: rowAuthor(row) } : {}) })
      let position: number, handle: string
      try {
        const matches = opened.paragraphs.filter(p => p.part === "word/document.xml" && p.match.text === row.find_text)
        if (matches.length !== 1) throw new RedlineReplayError(row, "The complete formatting target is missing or ambiguous.")
        const selected = matches[0]!
        handle = opened.revision.paragraphs[opened.paragraphs.indexOf(selected)]!.paragraph_handle
        position = Array.from(opened.documents["word/document.xml"]!.getElementsByTagNameNS(W, "p")).indexOf(selected.node)
        if (position < 0) throw new RedlineReplayError(row, "The complete formatting target is unavailable.")
        const refused = pendingRevisionRefusal(opened.session, pendingRevisionAnchors(opened.session), [selected.match.enclosingAnchor.id], row.find_text, "applied", ours([rowAuthor(row)]), true)
        if (refused) throw new RedlineReplayError(row, refused)
      } finally { opened.session.close() }
      if (row.new_text !== row.find_text) {
        try {
          bytes = (await applyDocxParagraphEdits(bytes, [{ id: String(row.id), paragraph_handle: handle, occurrence: 0,
            before_quote: row.find_text, replacement: row.new_text }], undefined,
            { ...options, pendingRevisions: "layer", ownAuthors: ours([rowAuthor(row)]), ...(tracked ? { revisionAuthor: rowAuthor(row) } : {}) })).bytes
        } catch (error) {
          const protection = docxProtectionMessage(error)
          throw new RedlineReplayError(row, protection ?? "The paragraph wording could not be changed safely.", Boolean(protection))
        }
      }
      try {
        bytes = await applyDocxParagraphFormatting(bytes, { paragraphText: row.new_text, format, author: rowAuthor(row), tracked, allowNoop: row.find_text !== row.new_text,
          options: { ...options, pendingRevisions: "layer", ownAuthors: ours([rowAuthor(row)]) } })
      } catch (error) {
        const protection = docxProtectionMessage(error)
        throw new RedlineReplayError(row, protection ?? "The formatting target changed or cannot be formatted safely.", Boolean(protection))
      }
      positions.push(locate && row.review_comment.trim() ? position : undefined)
      continue
    }
    if (row.scope === "phrase") {
      const before = locate && row.review_comment.trim() ? readDocxCommentParagraphPositions(bytes) : undefined
      try {
        bytes = (await applyLegacyPhrase(bytes, row.find_text, row.new_text,
          { ...options, pendingRevisions: "layer", ownAuthors: ours([rowAuthor(row)]), ...(tracked ? { revisionAuthor: rowAuthor(row) } : {}) })).bytes
      } catch (error) {
        const protection = docxProtectionMessage(error)
        if (error instanceof LegacyPhraseRefusal) throw new RedlineReplayError(row, error.kind === "missing" ? `Text no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` : error.message, error.kind === "pending", true)
        throw new RedlineReplayError(row, protection ?? "The wording cannot be safely changed at every occurrence.", Boolean(protection))
      }
      const after = before && readDocxCommentParagraphPositions(bytes)
      const changed = before && after?.filter((paragraph, index) => paragraph.paragraph_text !== before[index]?.paragraph_text)
      positions.push(before && after && after.length === before.length && changed?.length === 1 ? changed[0]!.xml_paragraph_indexes[0] : undefined)
      continue
    }
    const before = locate && row.review_comment.trim() ? readDocxCommentParagraphPositions(bytes) : undefined
    try {
      bytes = (await applyLegacyClause(bytes, row.find_text, row.new_text,
        { ...options, pendingRevisions: "layer", ownAuthors: ours([rowAuthor(row)]), ...(tracked ? { revisionAuthor: rowAuthor(row) } : {}) })).bytes
    } catch (error) {
      const protection = docxProtectionMessage(error)
      if (error instanceof LegacyPhraseRefusal) throw new RedlineReplayError(row, error.kind === "missing" ? `Clause no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` : error.message, error.kind === "pending", true)
      throw new RedlineReplayError(row, protection ?? "The clause cannot be safely changed. Review it in Word and prepare a fresh change.", true)
    }
    const after = before && readDocxCommentParagraphPositions(bytes)
    const changed = before && after?.filter((paragraph, index) => paragraph.paragraph_text !== before[index]?.paragraph_text)
    positions.push(before && after && after.length === before.length && changed?.length === 1 ? changed[0]!.xml_paragraph_indexes[0] : undefined)
  }
  return { bytes, positions }
}

// Version 1 binds a proposal to the whole Word revision it was prepared against.
// Version 2 binds it to one body paragraph's exact text and the selected
// occurrence of its quote inside that paragraph; base_sha256 is diagnostic only.
// The paragraph is the one its doc.haus pointer names (pid) or, on an uploaded
// source where it has none, the one Word's own unique w14:paraId names (para_id).
type RevisionBinding = { version:1; source_sha256:string; amendment_id:string; target_handle:string; kind?: "legacy_clause"; prepared_role?: "source" | "deliverable" | null; prepared_path?: string }
type ParagraphBinding = { version:2; amendment_id:string; paragraph_sha256:string; occurrence:number; base_sha256:string } & ({ pid:string } | { para_id:string })
export function binding(row:RedlineRow):RevisionBinding|ParagraphBinding {
  if (paragraphSpacingProposalMarker(row)) throw new Error("A spacing proposal requires its exact formatting writer.")
  if (tableRowProposalMarker(row) || paragraphInsertionProposalMarker(row)) throw new Error("A structural proposal requires its structural writer.")
  if (!row.target_json) throw new Error("The proposal has no exact Word revision binding.")
  // Bound proposals use the wording writer; it cannot apply stored formatting.
  if (row.formatting_json != null) throw new Error("Formatting cannot be combined with a Word revision binding. Prepare a fresh change list.")
  const parsed:unknown = JSON.parse(row.target_json)
  const value = parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string,unknown> : {}
  const hex = (text:unknown) => typeof text === "string" && /^[a-f0-9]{64}$/.test(text)
  const named = (text:unknown) => typeof text === "string" && text.length > 0
  if (value.kind === "legacy_clause") {
    if (row.scope !== "clause" || value.version !== 1 || !hex(value.source_sha256) || !named(value.target_handle) || !named(value.amendment_id) ||
      !(value.prepared_role === null || value.prepared_role === "source" || value.prepared_role === "deliverable") ||
      typeof value.prepared_path !== "string" || value.prepared_path !== row.doc_path ||
      Object.keys(value).some(key => !["version", "kind", "source_sha256", "target_handle", "amendment_id", "prepared_role", "prepared_path"].includes(key)))
      throw new Error("The proposal's Word revision binding is invalid.")
    return value as RevisionBinding
  }
  if (value.kind !== undefined) throw new Error("The proposal's Word revision binding is invalid.")
  if (value.version === 1 && hex(value.source_sha256) && named(value.target_handle) && named(value.amendment_id))
    return value as RevisionBinding
  if (value.version === 2 && named(value.amendment_id) && (named(value.pid) ? !("para_id" in value) : named(value.para_id) && !("pid" in value)) &&
    hex(value.paragraph_sha256) && hex(value.base_sha256) &&
    typeof value.occurrence === "number" && Number.isSafeInteger(value.occurrence) && value.occurrence >= 0)
    return value as ParagraphBinding
  throw new Error("The proposal's Word revision binding is invalid.")
}

const paragraphChanged = "The paragraph this change targets has changed since it was prepared. Prepare a fresh change list."
const trackedChanges = "This Word file contains tracked changes. Accept or reject them in Word before changes can be applied."
const unreadableRevision = "This Word revision cannot be read paragraph by paragraph. Prepare a fresh change list."
export type RedlineState = { sha256:string; targets:ParagraphTarget[]|null; tracked?:true; registeredRole?: "source" | "deliverable" | null; documentPath?: string; invalidClauseIds?: number[] }
// Read once per request from the bytes being reviewed. Paragraph targets are
// read only when a paragraph-bound proposal needs them, or to explain why a
// whole-revision proposal is stale; an unreadable revision leaves those
// proposals stale rather than failing the whole change list.
export async function redlineState(bytes:Uint8Array,rows:RedlineRow[],options:DocxRevisionOptions = {},registeredRole?: "source" | "deliverable" | null, documentPath?: string):Promise<RedlineState> {
  const sha256 = createHash("sha256").update(bytes).digest("hex")
  const invalidClauseIds: number[] = []
  for (const row of rows.filter(legacyClauseProposalMarker)) {
    try {
      const value = binding(row)
      if (value.version === 1 && value.source_sha256 === sha256) {
        const packet = await createDocxLocalization(bytes, [{ id: String(row.id), before_quote: row.find_text, replacement: row.new_text }], {}, options)
        requireCompleteClause(packet, row, value.target_handle)
      }
    } catch { invalidClauseIds.push(row.id) }
  }
  if (rows.some(exclusiveRedlineProposalMarker)) return { sha256, targets: [], registeredRole, documentPath, invalidClauseIds, ...(await tableRowSourceHasRevisions(bytes) ? { tracked: true } : {}) }
  // An older proposal is baked untracked: beside tracked changes, only on an
  // uploaded source; any other Word file with them refuses it, as Prepare does.
  if (options.pendingRevisions !== "layer" && rows.some((row) => row.target_json == null) && hasTrackedRevisions(bytes))
    return {sha256,targets:null,tracked:true,registeredRole,documentPath,invalidClauseIds}
  const bindings = rows.flatMap((row) => {try {return [binding(row)]} catch {return []}})
  if (!bindings.some((value) => value.version === 2 || value.source_sha256 !== sha256)) return {sha256,targets:[],registeredRole,documentPath,invalidClauseIds}
  try {return {sha256,targets:await paragraphTargets(bytes,options),registeredRole,documentPath,invalidClauseIds}}
  catch (error) {
    if (error instanceof DocxTrackedRevisionsError) return {sha256,targets:null,tracked:true,registeredRole,documentPath,invalidClauseIds}
    console.warn("[redline] paragraph targets unavailable",diagnosticErrorMetadata(error))
    return {sha256,targets:null,registeredRole,documentPath,invalidClauseIds}
  }
}
// A target's para_id is null unless no other body paragraph repeats it.
function boundParagraph(value:ParagraphBinding,targets:ParagraphTarget[]) {
  const target = "pid" in value ? resolveParagraphTarget(targets,value.pid) : targets.find((row) => row.para_id === value.para_id)
  return target?.sha256 === value.paragraph_sha256 ? target : undefined
}
export function redlineStaleReason(row:RedlineRow,state:RedlineState,protectedWord:boolean) {
  if(paragraphInsertionProposalMarker(row)) {
    try {const value=paragraphInsertionBinding(row);return state.registeredRole!==value.prepared_role?insertionChanged:state.tracked?insertionSourceOnly:value.source_sha256!==state.sha256?insertionChanged:undefined}
    catch(error){return insertionFailure(error)}
  }
  if (paragraphSpacingProposalMarker(row)) {
    try {
      const value = paragraphSpacingBinding(row)
      return state.registeredRole !== value.prepared_role ? spacingChanged : state.tracked ? spacingSourceOnly : value.source_sha256 !== state.sha256 ? spacingChanged : undefined
    } catch (error) { return spacingFailure(error) }
  }
  if (tableRowProposalMarker(row)) {
    try {
      const value = tableRowBinding(row)
      return state.registeredRole !== value.prepared_role ? rowChanged : state.tracked ? rowSourceOnly : value.source_sha256 !== state.sha256 ? rowChanged : undefined
    } catch (error) { return rowFailure(error) }
  }
  if (row.target_json == null) return protectedWord ? "Prepare this older proposal against the current Word revision." : state.tracked ? trackedChanges : undefined
  let value:ReturnType<typeof binding>
  try {value = binding(row)} catch {return "This proposal has an invalid Word revision binding. Prepare a fresh change list."}
  // No bound proposal applies to this revision, and a fresh change list cannot either.
  if (state.tracked) return trackedChanges
  if (value.version === 1) {
    if (value.kind === "legacy_clause" && state.invalidClauseIds?.includes(row.id)) return "This proposal has an invalid Word revision binding. Prepare a fresh change list."
    if (value.kind === "legacy_clause" && (value.prepared_role !== state.registeredRole || value.prepared_path !== state.documentPath)) return "The document's registered role changed. Prepare a fresh change list."
    return value.source_sha256 === state.sha256 ? undefined : "This proposal targets an earlier Word revision. Prepare a fresh change list."
  }
  if (!state.targets) return unreadableRevision
  // Never rebased: the pointed paragraph must still read exactly as prepared.
  const target = boundParagraph(value,state.targets)
  if (!target) return paragraphChanged
  // The touch rule (othersPending): a later pending change by anyone but our
  // own side (every name doc.haus signs with, and the proposal's author) holds
  // the paragraph, one that leaves its text as it was (a formatting change) included.
  const others = target.pending.length ? othersPending(target.pending,ours([row.author])) : []
  return others.length ? heldParagraph(row.find_text,others) : undefined
}
async function checkedTableRow(original: Uint8Array, rows: RedlineRow[], state?: RedlineState) {
  assertTableRowSingleton(rows)
  const row = rows[0]!, value = tableRowBinding(row)
  if (state?.registeredRole !== value.prepared_role) throw new RedlineAcceptConflictError(rowChanged)
  if (await tableRowSourceHasRevisions(original)) throw new RedlineAcceptConflictError(rowSourceOnly)
  if (createHash("sha256").update(original).digest("hex") !== value.source_sha256) throw new RedlineAcceptConflictError(rowChanged)
  return row
}
async function checkedInsertion(original:Uint8Array,rows:RedlineRow[],state?:RedlineState) {
  assertExclusiveRedlineSingleton(rows)
  const row=rows[0]!,value=paragraphInsertionBinding(row)
  if(state?.registeredRole!==value.prepared_role)throw new RedlineAcceptConflictError(insertionChanged)
  if(await tableRowSourceHasRevisions(original))throw new RedlineAcceptConflictError(insertionSourceOnly)
  if(createHash("sha256").update(original).digest("hex")!==value.source_sha256)throw new RedlineAcceptConflictError(insertionChanged)
  return row
}
function insertionFailure(error:unknown) {
  return error instanceof ParagraphInsertionProposalError||error instanceof RedlineAcceptConflictError||error instanceof RedlineProposalConflictError?error.message:insertionUnsupported
}
async function checkedSpacing(original: Uint8Array, rows: RedlineRow[], state?: RedlineState) {
  assertExclusiveRedlineSingleton(rows)
  const row = rows[0]!, value = paragraphSpacingBinding(row)
  if (state?.registeredRole !== value.prepared_role) throw new RedlineAcceptConflictError(spacingChanged)
  if (await tableRowSourceHasRevisions(original)) throw new RedlineAcceptConflictError(spacingSourceOnly)
  if (createHash("sha256").update(original).digest("hex") !== value.source_sha256) throw new RedlineAcceptConflictError(spacingChanged)
  return row
}
function spacingFailure(error: unknown) {
  if (error instanceof ParagraphSpacingProposalError || error instanceof RedlineAcceptConflictError || error instanceof RedlineProposalConflictError) return error.message
  return spacingUnsupported
}
function rowFailure(error: unknown) {
  if (error instanceof TableRowProposalError || error instanceof RedlineAcceptConflictError) return error.message
  return "This table row change cannot be safely prepared. Reject it and prepare a fresh change list, or make the change directly in Word."
}
/** Accept all's refusal for its stale proposals, in list order: the first by
 * its quote and reason, and how many others are stale too. */
export function staleBatchReason(stale:{row:RedlineRow;reason:string}[]) {
  const {row,reason} = stale[0]!
  // The file's own reason holds every proposal alike.
  if (reason === trackedChanges || reason === unreadableRevision) return reason
  const change = `The change to ${quoted(row.find_text)}`
  // A held paragraph's reason already names its change.
  const named = reason.startsWith(change) ? reason : `${change} can no longer be accepted. ${reason}`
  const others = stale.length - 1
  return others ? `${named} ${others} other proposal${others === 1 ? "" : "s"} can no longer be accepted either; the change list says why.` : named
}
function heldParagraph(text:string,authors:string[]) {
  const names = [...new Set(authors.map((author) => normalizeUntrustedText(author).text.trim()).filter(Boolean))]
  const many = authors.length > 1
  return `The change to ${quoted(text)} can no longer be accepted: its paragraph now has ${many ? "pending changes" : "a pending change"}${names.length ? ` by ${listNames(names)}` : ""}. Reject this proposal, decide ${many ? "those changes" : "that change"} under Tracked changes, then request edits again.`
}
// One all-or-nothing write. A version 1 row is located by its stored handle
// through whole-revision localization of the version 1 rows. A version 2 row is
// located directly: its bound paragraph, unchanged, in these bytes and the
// stored occurrence of its quote inside that paragraph, with no whole-document
// candidate search. Overlapping spans in one paragraph are refused. `options`
// are the reader's DocxRevisionOptions; with a revisionAuthor the edits are
// written as tracked changes.
function requireCompleteClause(packet: Awaited<ReturnType<typeof createDocxLocalization>>, row: RedlineRow, targetHandle: string) {
  const matches = packet.candidates.filter(candidate => candidate.amendment_id === String(row.id) && candidate.target_handle === targetHandle)
  if (matches.length !== 1 || matches[0]!.paragraph.text !== row.find_text || matches[0]!.context_before || matches[0]!.context_after)
    throw new Error("The proposal's Word revision binding is invalid.")
}
async function bakeBound(original:Uint8Array,rows:RedlineRow[],state?:RedlineState,options:DocxRevisionOptions = {}) {
  const current = createHash("sha256").update(original).digest("hex")
  const bindings = rows.map(binding)
  for (const value of bindings) if (value.version === 1 && value.kind === "legacy_clause" && (value.prepared_role !== state?.registeredRole || value.prepared_path !== state?.documentPath))
    throw new Error("The document's registered role changed. Prepare a fresh change list.")
  if (bindings.some((value) => value.version === 1 && value.source_sha256 !== current)) throw new Error("This proposal targets an earlier Word revision. Prepare changes against the current document.")
  const targets = !bindings.some((value) => value.version === 2) ? []
    : state?.sha256 === current && state.targets?.length ? state.targets : await paragraphTargets(original,options)
  // Use stored row identity here: separate proposal batches can reuse model-local IDs.
  const amendment = (row:RedlineRow) => ({id:String(row.id),before_quote:row.find_text,replacement:row.new_text})
  const revisionBound = rows.flatMap((row,index) => {
    const value = bindings[index]!
    return value.version === 1 ? [{row,target_handle:value.target_handle}] : []
  })
  const localized = revisionBound.length ? {
    packet:await createDocxLocalization(original,revisionBound.map(({row}) => amendment(row)),{},options),
    selections:revisionBound.map(({row,target_handle}) => ({amendment_id:String(row.id),target_handle})),
  } : undefined
  if (localized) for (const { row, target_handle } of revisionBound) {
    const value = binding(row)
    if (value.version !== 1 || value.kind !== "legacy_clause") continue
    requireCompleteClause(localized.packet, row, target_handle)
  }
  const edits = rows.flatMap((row,index) => {
    const value = bindings[index]!
    if (value.version === 1) return []
    const paragraph = boundParagraph(value,targets)
    if (!paragraph) throw new Error(paragraphChanged)
    return [{...amendment(row),paragraph_handle:paragraph.paragraph_handle,occurrence:value.occurrence}]
  })
  const result = await applyDocxParagraphEdits(original,edits,localized,options)
  // Each row's applied span and paragraph, in row order, for comment mapping.
  return { ...result, bindings:rows.map((row) => result.report.amendments.find((applied) => applied.amendment_id === String(row.id))!) }
}
