import { docxodus } from "./docxodus"
import { compareDocxWithPointers } from "../../../dochaus/lib/drafting-amendment"
import type { RedlineRow } from "./db"
import { diagnosticErrorMetadata } from "../../../dochaus/lib/diagnostic-metadata"
import { createDocxLocalization, applyDocxLocalization } from "../../../dochaus/lib/docx-amendments"
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
export async function buildRedlined(original: Uint8Array, rows: RedlineRow[]): Promise<Uint8Array> {
  if (!rows.length) return original
  const annotated = rows.some(row => row.review_comment.trim())
  // Paragraph comments cannot safely follow a structural replay by ordinal.
  if (annotated && rows.some(row => /[\r\n\t]/.test(row.find_text + row.new_text)))
    throw new RedlinePreviewConflictError("Word review comments require paragraph-preserving proposals. Prepare a fresh change list without structural replacements.")
  if (rows.some((row) => row.target_json != null)) {
    if (rows.some((row) => row.target_json == null))
      throw new RedlinePreviewConflictError("Older and revision-bound proposals cannot share one redline export. Resolve the older proposals or prepare a fresh change list.")
    try {
      const result = await bakeBound(original, rows)
      const paragraphs = result.packet.revision.paragraphs.filter(row => row.scope === "body")
      const source = readDocxCommentParagraphs(original)
      if (annotated && JSON.stringify(paragraphs.map(row => row.text)) !== JSON.stringify(source.map(row => row.paragraph_text)))
        throw new Error("unsupported_comment_paragraph_mapping")
      const targets = rows.map((row, index) => {
        const candidate = result.packet.candidates.find(candidate => candidate.amendment_id === String(row.id)
          && candidate.target_handle === result.bindings[index].target_handle)!
        return paragraphs.findIndex(paragraph => paragraph.paragraph_handle === candidate.paragraph.paragraph_handle)
      })
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
// baking a partial set into the canonical document.
export async function bake(original: Uint8Array, rows: RedlineRow[]): Promise<Uint8Array> {
  if (rows.some((row) => row.target_json != null)) {
    if (rows.some((row) => row.target_json == null)) throw new Error("Revision-bound and older proposals require separate review. Prepare a fresh change list.")
    return (await bakeBound(original,rows)).bytes
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

export function binding(row:RedlineRow) {
  if (!row.target_json) throw new Error("The proposal has no exact Word revision binding.")
  const value:unknown = JSON.parse(row.target_json)
  if (!value || typeof value !== "object" || !("version" in value) || value.version !== 1 ||
    !("source_sha256" in value) || typeof value.source_sha256 !== "string" || !/^[a-f0-9]{64}$/.test(value.source_sha256) ||
    !("target_handle" in value) || typeof value.target_handle !== "string" || !value.target_handle ||
    !("amendment_id" in value) || typeof value.amendment_id !== "string" || !value.amendment_id)
    throw new Error("The proposal's Word revision binding is invalid.")
  return value
}
export function redlineStaleReason(row:RedlineRow,currentHash:string,protectedWord:boolean) {
  if (row.target_json == null) return protectedWord ? "Prepare this older proposal against the current Word revision." : undefined
  try {
    if (binding(row).source_sha256 !== currentHash) return "This proposal targets an earlier Word revision. Prepare a fresh change list."
  } catch { return "This proposal has an invalid Word revision binding. Prepare a fresh change list." }
}
async function bakeBound(original:Uint8Array,rows:RedlineRow[]) {
  const current = createHash("sha256").update(original).digest("hex")
  const bindings = rows.map(binding)
  if (bindings.some((value) => value.source_sha256 !== current)) throw new Error("This proposal targets an earlier Word revision. Prepare changes against the current document.")
  // Use stored row identity here: separate proposal batches can reuse model-local IDs.
  const packet = await createDocxLocalization(original,rows.map((row) => ({id:String(row.id),before_quote:row.find_text,replacement:row.new_text})))
  const result = await applyDocxLocalization(original,packet,rows.map((row,index) => ({amendment_id:String(row.id),target_handle:bindings[index].target_handle})))
  return { ...result, packet, bindings }
}
