import { docxodus } from "./docxodus"
import { compareDocxWithPointers, paragraphTargets, resolveParagraphTarget, type ParagraphTarget } from "../../../dochaus/lib/drafting-amendment"
import type { RedlineRow } from "./db"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { createDocxLocalization, applyDocxParagraphEdits } from "../../../dochaus/lib/docx-amendments"
import { createHash } from "node:crypto"
import { draftingAttribution } from "../../../dochaus/lib/docx-attribution"
import { exportDocxComments, preserveDocxComments, readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"

// Replay and bake pending redline proposals against a clean .docx.
//
// The canonical document on disk is always the accepted ("clean") state. A
// proposal records how to reproduce one edit — a surgical find/replace ('phrase')
// or a whole-paragraph rewrite ('clause') — located the same way the dochaus tools
// located it originally. We never store offsets: text is re-resolved against the
// live document every time, so edits stay valid as the document changes underneath.

type Session = ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
export class RedlinePreviewConflictError extends Error {}

// Reproduce one proposal's edit on an open session. Returns an error message
// (rather than throwing) when the anchor text can no longer be found — e.g. an
// earlier proposal in the same replay already rewrote the passage this one
// anchors to. Callers decide whether that is fatal: the redlined view skips the
// row and renders the rest; accept must fail loudly rather than bake a partial set.
function applyProposal(session: Session, row: RedlineRow): { ok: true } | { ok: false; error: string } {
  if (row.scope === "clause") {
    const target = session.findByText(row.find_text, { ignoreWhitespace: true })
    if (!target) return { ok: false, error: `Clause no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` }
    const result = session.replaceText(target.id, row.new_text)
    if (!result.success) return { ok: false, error: `Redline #${row.id} failed: ${result.error?.message ?? "unknown error"}` }
    return { ok: true }
  }

  const targets = session.findAllByText(row.find_text)
  if (!targets.length) return { ok: false, error: `Text no longer found for redline #${row.id}: ${JSON.stringify(row.find_text)}` }
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
export async function buildRedlined(original: Uint8Array, rows: RedlineRow[], state?: RedlineState): Promise<Uint8Array> {
  if (!rows.length) return original
  const annotated = rows.some(row => row.review_comment.trim())
  // Paragraph comments cannot safely follow a structural replay by ordinal.
  if (annotated && rows.some(row => /[\r\n\t]/.test(row.find_text + row.new_text)))
    throw new RedlinePreviewConflictError("Word review comments require paragraph-preserving proposals. Prepare a fresh change list without structural replacements.")
  if (rows.some((row) => row.target_json != null)) {
    if (rows.some((row) => row.target_json == null))
      throw new RedlinePreviewConflictError("Older and revision-bound proposals cannot share one redline export. Resolve the older proposals or prepare a fresh change list.")
    try {
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
  const compared = await compareDocxWithPointers(original, modified, { authorName: previewAuthor(rows) })
  const redline = (() => {
    try { return preserveDocxComments(original, compared) }
    catch { throw new RedlinePreviewConflictError("An existing Word comment could not be preserved at its original unchanged anchor or with its original links and formatting dependencies. Review comments on edited ranges, complex fields, objects, threaded comments and attachments separately before exporting.") }
  })()
  if (!rows.some(row => row.review_comment.trim())) return redline
  const paragraphs = readDocxCommentParagraphs(modified)
  const exported = readDocxCommentParagraphs(redline)
  if (JSON.stringify(paragraphs.map(row => row.paragraph_text)) !== JSON.stringify(exported.map(row => row.paragraph_text)))
    throw new RedlinePreviewConflictError("The Word comparison changed paragraph boundaries; review comments cannot be attached safely. Prepare a fresh change list.")
  const comments = rows.flatMap((row, index) => {
    if (!row.review_comment.trim()) return []
    const paragraph = paragraphs[targets[index]]
    if (!paragraph || !row.author.trim())
      throw new RedlinePreviewConflictError(`Review comment for proposal #${row.id} has no supported body paragraph or stored author. Prepare a fresh change list before exporting.`)
    return [{ id: String(row.id), text: row.review_comment, author: row.author,
      anchor: { paragraph_index: paragraph.paragraph_index, paragraph_text: paragraph.paragraph_text } }]
  })
  try {
    const output = exportDocxComments(redline, comments)
    if (output.unresolved.length)
      throw new RedlinePreviewConflictError(`Word review comments could not be attached: ${output.unresolved.map(row => `proposal #${row.id}: ${row.reason}`).join(" ")}`)
    return output.bytes
  } catch (error) {
    if (error instanceof RedlinePreviewConflictError) throw error
    throw new RedlinePreviewConflictError("The stored Word review comments could not be exported safely. Check their text, authors and document anchors before exporting again.")
  }
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

// Bake the given proposals into the clean document, returning new canonical bytes
// with the edits applied and no tracked changes — the new accepted state. Accept
// fails loudly: a proposal that no longer resolves throws rather than silently
// baking a partial set into the canonical document. `state` is the request's
// redlineState for these bytes, when already read for the stale check.
export async function bake(original: Uint8Array, rows: RedlineRow[], state?: RedlineState): Promise<Uint8Array> {
  if (rows.some((row) => row.target_json != null)) {
    if (rows.some((row) => row.target_json == null)) throw new Error("Revision-bound and older proposals require separate review. Prepare a fresh change list.")
    return (await bakeBound(original,rows,state)).bytes
  }
  const dx = await docxodus()
  const session = dx.openDocxSession(original, {})
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
// Version 2 binds it to one pointed body paragraph's exact text and the selected
// occurrence of its quote inside that paragraph; base_sha256 is diagnostic only.
type RevisionBinding = { version:1; source_sha256:string; amendment_id:string; target_handle:string }
type ParagraphBinding = { version:2; amendment_id:string; pid:string; paragraph_sha256:string; occurrence:number; base_sha256:string }
export function binding(row:RedlineRow):RevisionBinding|ParagraphBinding {
  if (!row.target_json) throw new Error("The proposal has no exact Word revision binding.")
  const parsed:unknown = JSON.parse(row.target_json)
  const value = parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string,unknown> : {}
  const hex = (text:unknown) => typeof text === "string" && /^[a-f0-9]{64}$/.test(text)
  const named = (text:unknown) => typeof text === "string" && text.length > 0
  if (value.version === 1 && hex(value.source_sha256) && named(value.target_handle) && named(value.amendment_id))
    return value as RevisionBinding
  if (value.version === 2 && named(value.amendment_id) && named(value.pid) && hex(value.paragraph_sha256) && hex(value.base_sha256) &&
    typeof value.occurrence === "number" && Number.isSafeInteger(value.occurrence) && value.occurrence >= 0)
    return value as ParagraphBinding
  throw new Error("The proposal's Word revision binding is invalid.")
}

const paragraphChanged = "The paragraph this change targets has changed since it was prepared. Prepare a fresh change list."
export type RedlineState = { sha256:string; targets:ParagraphTarget[]|null }
// Read once per request from the bytes being reviewed. Paragraph targets are
// read only when a paragraph-bound proposal needs them; an unreadable revision
// leaves those proposals stale rather than failing the whole change list.
export async function redlineState(bytes:Uint8Array,rows:RedlineRow[]):Promise<RedlineState> {
  const sha256 = createHash("sha256").update(bytes).digest("hex")
  if (!rows.some((row) => {try {return binding(row).version === 2} catch {return false}})) return {sha256,targets:[]}
  try {return {sha256,targets:await paragraphTargets(bytes)}}
  catch (error) {
    console.warn("[redline] paragraph targets unavailable",diagnosticErrorMetadata(error))
    return {sha256,targets:null}
  }
}
function boundParagraph(value:ParagraphBinding,targets:ParagraphTarget[]) {
  const target = resolveParagraphTarget(targets,value.pid)
  return target?.sha256 === value.paragraph_sha256 ? target : undefined
}
export function redlineStaleReason(row:RedlineRow,state:RedlineState,protectedWord:boolean) {
  if (row.target_json == null) return protectedWord ? "Prepare this older proposal against the current Word revision." : undefined
  let value:ReturnType<typeof binding>
  try {value = binding(row)} catch {return "This proposal has an invalid Word revision binding. Prepare a fresh change list."}
  if (value.version === 1) return value.source_sha256 === state.sha256 ? undefined : "This proposal targets an earlier Word revision. Prepare a fresh change list."
  if (!state.targets) return "This Word revision cannot be read paragraph by paragraph. Prepare a fresh change list."
  // Never rebased: the pointed paragraph must still read exactly as prepared.
  return boundParagraph(value,state.targets) ? undefined : paragraphChanged
}
// One all-or-nothing write. A version 1 row is located by its stored handle
// through whole-revision localization of the version 1 rows. A version 2 row is
// located directly: its pointer's unchanged paragraph in these bytes and the
// stored occurrence of its quote inside that paragraph, with no whole-document
// candidate search. Overlapping spans in one paragraph are refused.
async function bakeBound(original:Uint8Array,rows:RedlineRow[],state?:RedlineState) {
  const current = createHash("sha256").update(original).digest("hex")
  const bindings = rows.map(binding)
  if (bindings.some((value) => value.version === 1 && value.source_sha256 !== current)) throw new Error("This proposal targets an earlier Word revision. Prepare changes against the current document.")
  const targets = !bindings.some((value) => value.version === 2) ? []
    : state?.sha256 === current && state.targets?.length ? state.targets : await paragraphTargets(original)
  // Use stored row identity here: separate proposal batches can reuse model-local IDs.
  const amendment = (row:RedlineRow) => ({id:String(row.id),before_quote:row.find_text,replacement:row.new_text})
  const revisionBound = rows.flatMap((row,index) => {
    const value = bindings[index]!
    return value.version === 1 ? [{row,target_handle:value.target_handle}] : []
  })
  const localized = revisionBound.length ? {
    packet:await createDocxLocalization(original,revisionBound.map(({row}) => amendment(row))),
    selections:revisionBound.map(({row,target_handle}) => ({amendment_id:String(row.id),target_handle})),
  } : undefined
  const edits = rows.flatMap((row,index) => {
    const value = bindings[index]!
    if (value.version === 1) return []
    const paragraph = boundParagraph(value,targets)
    if (!paragraph) throw new Error(paragraphChanged)
    return [{...amendment(row),paragraph_handle:paragraph.paragraph_handle,occurrence:value.occurrence}]
  })
  const result = await applyDocxParagraphEdits(original,edits,localized)
  // Each row's applied span and paragraph, in row order, for comment mapping.
  return { ...result, bindings:rows.map((row) => result.report.amendments.find((applied) => applied.amendment_id === String(row.id))!) }
}
