import { docxodus } from "./docxodus"
import { compareDocxWithPointers, paragraphTargets, resolveParagraphTarget, withStableRevisionDates, type ParagraphTarget } from "../../../dochaus/lib/drafting-amendment"
import type { RedlineRow } from "./db"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { createDocxLocalization, applyDocxParagraphEdits, DocxTrackedRevisionsError, openRevision, pendingRevisionAnchors, pendingRevisionRefusal, type DocxRevisionOptions } from "../../../dochaus/lib/docx-amendments"
import { createHash } from "node:crypto"
import { draftingAttribution } from "../../../dochaus/lib/docx-attribution"
import { normalizeUntrustedText } from "../../../dochaus/lib/untrusted-text"
import { listNames } from "./document-versions"
import { exportDocxComments, preserveDocxComments, readDocxCommentParagraphPositions, readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"
import { assertLayeredRoundTrip, assertRedlineRoundTrip } from "../../../dochaus/lib/redline-round-trip"
import { discloseSourceRevisions } from "../../../dochaus/lib/docx-revision-views"

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
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

// A replay session's paragraphs with a pending revision, read once: its
// untracked edits add none. Ingest's own install of the same Docxodus package
// declares the session type again; the reader only calls the session's methods.
const pendingRevisions = new WeakMap<Session, Map<string, string[]>>()
const pendingIn = (session: Session) => pendingRevisions.get(session)
  ?? pendingRevisions.set(session, pendingRevisionAnchors(session as unknown as Parameters<typeof pendingRevisionAnchors>[0])).get(session)!

// Reproduce one proposal's edit on an open session. Returns an error message
// (rather than throwing) when the anchor text can no longer be found — e.g. an
// earlier proposal in the same replay already rewrote the passage this one
// anchors to. Callers decide whether that is fatal: the redlined view skips the
// row and renders the rest; accept must fail loudly rather than bake a partial set.
// It is refused in a paragraph with a pending revision, doc.haus's own included
// (the touch rule): untracked, doc.haus's words would become that revision's
// author's, and a whole-clause rewrite would rewrite that revision.
function applyProposal(session: Session, row: RedlineRow): { ok: true } | { ok: false; error: string } {
  if (row.scope === "clause") {
    const target = session.findByText(row.find_text, { ignoreWhitespace: true })
    if (!target) return { ok: false, error: `Clause no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` }
    const refused = pendingRevisionRefusal(pendingIn(session), [target.id], row.find_text, "applied")
    if (refused) return { ok: false, error: refused }
    const result = session.replaceText(target.id, row.new_text)
    if (!result.success) return { ok: false, error: `Redline #${row.id} failed: ${result.error?.message ?? "unknown error"}` }
    return { ok: true }
  }

  const targets = session.findAllByText(row.find_text)
  if (!targets.length) return { ok: false, error: `Text no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` }
  const refused = pendingRevisionRefusal(pendingIn(session), targets.map((t) => t.id), row.find_text, "applied")
  if (refused) return { ok: false, error: refused }
  const results = targets.flatMap((t) => session.replaceTextRange(t.id, row.find_text, row.new_text))
  const failed = results.find((r) => !r.success)
  if (failed) return { ok: false, error: `Redline #${row.id} failed: ${failed.error?.message ?? "unknown error"}` }
  return { ok: true }
}

// The redlined view the viewer renders: clean document compared against the same
// document with every pending proposal applied, so Docxodus emits native w:ins/w:del
// the browser paints green/red. With no pending rows the compare is a no-op and the
// clean document round-trips unchanged.
//
// Older proposals without review notes retain the tolerant preview replay.
// Revision-bound proposals and annotated proposals fail closed: exporting only
// part of their tracked changes or silently losing a stored note is misleading.
// `state` is the request's redlineState for these bytes, when already read.
// `options` is how the document is read. In layer mode (an uploaded source)
// with pending revisions, revision-bound proposals are written into it instead
// (layeredExport).
export async function buildRedlined(original: Uint8Array, rows: RedlineRow[], state?: RedlineState, options: DocxRevisionOptions = {}): Promise<Uint8Array> {
  if (!rows.length) return original
  const annotated = rows.some(row => row.review_comment.trim())
  // Paragraph comments cannot safely follow a structural replay by ordinal.
  if (annotated && rows.some(row => /[\r\n\t]/.test(row.find_text + row.new_text)))
    throw new RedlinePreviewConflictError("Word review comments require paragraph-preserving proposals. Prepare a fresh change list without structural replacements.")
  if (rows.some((row) => row.target_json != null)) {
    if (rows.some((row) => row.target_json == null))
      throw new RedlinePreviewConflictError("Older and revision-bound proposals cannot share one redline export. Resolve the older proposals or prepare a fresh change list.")
    try {
      if (layeredRedline(rows, options) && (await (await docxodus()).getRevisions(original)).length)
        return await layeredExport(original, rows, state, options)
      const result = await bakeBound(original, rows, state)
      const paragraphs = result.revision.paragraphs.filter(row => row.scope === "body")
      const source = readDocxCommentParagraphs(original)
      if (annotated && JSON.stringify(paragraphs.map(row => row.text)) !== JSON.stringify(source.map(row => row.paragraph_text)))
        throw new Error("unsupported_comment_paragraph_mapping")
      const targets = result.bindings.map(binding => paragraphs.findIndex(paragraph => paragraph.paragraph_handle === binding.paragraph_handle))
      return commentRedline(original, result.bytes, rows, targets)
    } catch (error) {
      if (error instanceof RedlinePreviewConflictError) throw error
      throw new RedlinePreviewConflictError("A proposal no longer resolves against this Word revision. Prepare a fresh change list before exporting its tracked changes and review comments.")
    }
  }
  const dx = await docxodus()
  const session = dx.openDocxSession(original, {})
  try {
    const targets: number[] = []
    const applied = rows.filter((row) => {
      const before = row.review_comment.trim() ? readDocxCommentParagraphs(session.save()) : []
      if (row.review_comment.trim() && before.filter(paragraph => paragraph.paragraph_text.includes(row.find_text)).length !== 1)
        throw new RedlinePreviewConflictError(`Review comment for proposal #${row.id} has a missing or ambiguous paragraph target. Prepare a fresh change list before exporting.`)
      const result = applyProposal(session, row)
      if (!result.ok) {
        if (row.review_comment.trim()) throw new RedlinePreviewConflictError(`Proposal #${row.id} cannot be replayed with its review comment. Prepare a fresh change list before exporting.`)
        console.warn("[redline] skipping in redlined view", diagnosticErrorMetadata(result.error))
        return false
      }
      const after = row.review_comment.trim() ? readDocxCommentParagraphs(session.save()) : []
      const changed = before.filter((paragraph, index) => paragraph.paragraph_text !== after[index]?.paragraph_text)
      if (row.review_comment.trim() && (before.length !== after.length || changed.length !== 1))
        throw new RedlinePreviewConflictError(`Review comment for proposal #${row.id} requires one exact changed paragraph. Prepare a fresh change list before exporting.`)
      targets.push(changed[0]?.paragraph_index ?? -1)
      return true
    })
    if (!applied.length) return original
    return commentRedline(original, session.save(), applied, targets)
  } finally { session.close() }
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

// The other side's markup with the proposals written into it in place, as
// native tracked changes under the proposals' author, and their review
// comments. Their revisions and comments are kept as they are; comparison
// would accept them. Track Changes is left on, as a returned redline travels.
async function layeredExport(original: Uint8Array, rows: RedlineRow[], state: RedlineState | undefined, options: DocxRevisionOptions) {
  const author = previewAuthor(rows)
  const baked = await bakeBound(original, rows, state, { ...options, ownAuthor: author })
  const tracked = await bakeBound(original, rows, state, { ...options, ownAuthor: author, revisionAuthor: author })
  let output = tracked.bytes
  if (rows.some(row => row.review_comment.trim())) {
    // An edit keeps every w:p in place, so the edited paragraph's position in
    // the original finds its accepted row, which a deleted mark may join to another.
    const positions = await bodyParagraphPositions(original, options)
    const accepted = readDocxCommentParagraphPositions(output)
    output = withReviewComments(output, rows, tracked.bindings.map(binding => {
      const at = positions.get(binding.paragraph_handle)
      return at === undefined ? undefined : accepted.find(row => row.xml_paragraph_indexes.includes(at))
    }))
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

function previewAuthor(rows: RedlineRow[]) {
  const fallback = draftingAttribution().author
  const authors = [...new Set(rows.map(row => row.author || fallback))]
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
// pending and takes the edits beside them as native tracked changes under the
// proposals' author, so the accepted file is itself the redline that goes back;
// the round trip proves rejecting only those gives the document as it was, and
// accepting them the same edits untracked.
export async function bake(original: Uint8Array, rows: RedlineRow[], state?: RedlineState, options: DocxRevisionOptions = {}): Promise<Uint8Array> {
  const bound = rows.some((row) => row.target_json != null)
  if (bound && rows.some((row) => row.target_json == null)) throw new Error("Revision-bound and older proposals require separate review. Prepare a fresh change list.")
  const write = async (author?: string, revisionAuthor?: string) => bound
    ? (await bakeBound(original,rows,state,{...options,...author ? {ownAuthor:author} : {},...revisionAuthor ? {revisionAuthor} : {}})).bytes
    : bakeOlder(original,rows,revisionAuthor)
  if (options.pendingRevisions !== "layer" || !(await (await docxodus()).getRevisions(original)).length) return write()
  const author = previewAuthor(rows)
  const [plain,tracked] = [await write(author),await write(author,author)]
  await assertLayeredRoundTrip(tracked,original,plain)
  return tracked
}

// Older proposals, replayed by their text; with a revisionAuthor, as native
// tracked changes under that name.
async function bakeOlder(original: Uint8Array, rows: RedlineRow[], revisionAuthor?: string) {
  const dx = await docxodus()
  const session = dx.openDocxSession(original, revisionAuthor ? { trackedChanges: "render_inline", revisionAuthor } : {})
  try {
    for (const row of rows) {
      const result = applyProposal(session, row)
      if (!result.ok) throw new Error(result.error)
    }
    return session.save()
  } finally {
    session.close()
  }
}

// Version 1 binds a proposal to the whole Word revision it was prepared against.
// Version 2 binds it to one body paragraph's exact text and the selected
// occurrence of its quote inside that paragraph; base_sha256 is diagnostic only.
// The paragraph is the one its doc.haus pointer names (pid) or, on an uploaded
// source where it has none, the one Word's own unique w14:paraId names (para_id).
type RevisionBinding = { version:1; source_sha256:string; amendment_id:string; target_handle:string }
type ParagraphBinding = { version:2; amendment_id:string; paragraph_sha256:string; occurrence:number; base_sha256:string } & ({ pid:string } | { para_id:string })
export function binding(row:RedlineRow):RevisionBinding|ParagraphBinding {
  if (!row.target_json) throw new Error("The proposal has no exact Word revision binding.")
  const parsed:unknown = JSON.parse(row.target_json)
  const value = parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string,unknown> : {}
  const hex = (text:unknown) => typeof text === "string" && /^[a-f0-9]{64}$/.test(text)
  const named = (text:unknown) => typeof text === "string" && text.length > 0
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
export type RedlineState = { sha256:string; targets:ParagraphTarget[]|null; tracked?:true }
// Read once per request from the bytes being reviewed. Paragraph targets are
// read only when a paragraph-bound proposal needs them, or to explain why a
// whole-revision proposal is stale; an unreadable revision leaves those
// proposals stale rather than failing the whole change list.
export async function redlineState(bytes:Uint8Array,rows:RedlineRow[],options:DocxRevisionOptions = {}):Promise<RedlineState> {
  const sha256 = createHash("sha256").update(bytes).digest("hex")
  const bindings = rows.flatMap((row) => {try {return [binding(row)]} catch {return []}})
  if (!bindings.some((value) => value.version === 2 || value.source_sha256 !== sha256)) return {sha256,targets:[]}
  try {return {sha256,targets:await paragraphTargets(bytes,options)}}
  catch (error) {
    if (error instanceof DocxTrackedRevisionsError) return {sha256,targets:null,tracked:true}
    console.warn("[redline] paragraph targets unavailable",diagnosticErrorMetadata(error))
    return {sha256,targets:null}
  }
}
// A target's para_id is null unless no other body paragraph repeats it.
function boundParagraph(value:ParagraphBinding,targets:ParagraphTarget[]) {
  const target = "pid" in value ? resolveParagraphTarget(targets,value.pid) : targets.find((row) => row.para_id === value.para_id)
  return target?.sha256 === value.paragraph_sha256 ? target : undefined
}
export function redlineStaleReason(row:RedlineRow,state:RedlineState,protectedWord:boolean) {
  if (row.target_json == null) return protectedWord ? "Prepare this older proposal against the current Word revision." : undefined
  let value:ReturnType<typeof binding>
  try {value = binding(row)} catch {return "This proposal has an invalid Word revision binding. Prepare a fresh change list."}
  // No bound proposal applies to this revision, and a fresh change list cannot either.
  if (state.tracked) return trackedChanges
  if (value.version === 1) return value.source_sha256 === state.sha256 ? undefined : "This proposal targets an earlier Word revision. Prepare a fresh change list."
  if (!state.targets) return "This Word revision cannot be read paragraph by paragraph. Prepare a fresh change list."
  // Never rebased: the pointed paragraph must still read exactly as prepared.
  const target = boundParagraph(value,state.targets)
  if (!target) return paragraphChanged
  // The touch rule (othersPending): a later pending change by anyone but the
  // proposal's author holds the paragraph, one that leaves its text as it was
  // (a formatting change) included.
  const others = target.pending.filter((author) => author !== (row.author || draftingAttribution().author))
  return others.length ? heldParagraph(row.find_text,others) : undefined
}
function heldParagraph(text:string,authors:string[]) {
  const names = [...new Set(authors.map((author) => normalizeUntrustedText(author).text.trim()).filter(Boolean))]
  const many = authors.length > 1
  return `The change to ${JSON.stringify(text)} can no longer be accepted: its paragraph now has ${many ? "pending changes" : "a pending change"}${names.length ? ` by ${listNames(names)}` : ""}. Reject this proposal, decide ${many ? "those changes" : "that change"} under Tracked changes, then request edits again.`
}
// One all-or-nothing write. A version 1 row is located by its stored handle
// through whole-revision localization of the version 1 rows. A version 2 row is
// located directly: its bound paragraph, unchanged, in these bytes and the
// stored occurrence of its quote inside that paragraph, with no whole-document
// candidate search. Overlapping spans in one paragraph are refused. `options`
// are the reader's DocxRevisionOptions; with a revisionAuthor the edits are
// written as tracked changes.
async function bakeBound(original:Uint8Array,rows:RedlineRow[],state?:RedlineState,options:DocxRevisionOptions = {}) {
  const current = createHash("sha256").update(original).digest("hex")
  const bindings = rows.map(binding)
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
