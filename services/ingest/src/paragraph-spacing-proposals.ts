import { createHash } from "node:crypto"
import { readFileSync } from "node:fs"
import { isDeepStrictEqual } from "node:util"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { inspectExactParagraphSpacing, type ExactSpacingRequest } from "../../../dochaus/lib/docx-exact-spacing"
import { authorExactParagraphSpacingReview } from "../../../dochaus/lib/docx-exact-spacing-review"
import { validateDocxRowReviewReason } from "../../../dochaus/lib/docx-table-row-review"
import { frozenParagraphSpacingView, paragraphSpacingProposalSchema, type ParagraphSpacingProposalBinding, type ParagraphSpacingProposalView } from "../../../dochaus/lib/paragraph-spacing-proposal"
import { assertRedlineProposalCompatibility } from "../../../dochaus/lib/table-row-proposal"
import { getDocumentRole, listPendingRedlines, openDb, setPreparedChanges, type RedlineRow } from "./db"
import { resolveMatterDocumentTarget } from "./document-identity"
import { registeredDocumentName } from "./registered-document"
import { tableRowSourceHasRevisions } from "./table-row-proposals"

const hash = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
export class ParagraphSpacingProposalError extends Error {}
export const spacingSourceOnly = "Paragraph spacing changes currently require an uploaded Word source without tracked changes."
export const spacingChanged = "This spacing change targets an earlier Word file or document role. Reject it and prepare a fresh change list."
export const spacingUnsupported = "This paragraph's spacing cannot be safely prepared as a tracked proposal. Make this change directly in Word."
export const spacingInvalid = "This spacing change has an invalid stored operation. Reject it and prepare a fresh change list."
export const spacingSnapshotChanged = "The paragraph, its context or its direct spacing does not match this proposal. Reject it and prepare a fresh change list."
export function paragraphSpacingBinding(row: RedlineRow): ParagraphSpacingProposalBinding {
  try {
    const value = paragraphSpacingProposalSchema.parse(JSON.parse(row.target_json ?? "null"))
    if (row.scope !== "paragraph_spacing" || row.formatting_json != null || row.anchor_id != null || !row.author.trim() || row.author !== row.author.trim() || row.author.length > 100 || !row.author.isWellFormed() || /[\u0000-\u001f\u007f]/.test(row.author) ||
      row.find_text !== value.paragraph_text || row.old_text !== value.paragraph_text || row.new_text !== value.paragraph_text) throw new Error()
    validateDocxRowReviewReason(row.review_comment)
    return value
  } catch { throw new ParagraphSpacingProposalError(spacingInvalid) }
}
/** Used by recording, fresh display, preview and acceptance. Neither a matching
 * handle alone nor a normalized old value can replace the frozen source data. */
export async function checkedParagraphSpacingOperation(bytes: Uint8Array, row: RedlineRow): Promise<{ request: ExactSpacingRequest; view: ParagraphSpacingProposalView & { table_context: ExactSpacingRequest["tableContext"] } }> {
  const value = paragraphSpacingBinding(row)
  if (hash(bytes) !== value.source_sha256) throw new ParagraphSpacingProposalError(spacingChanged)
  let packet: Awaited<ReturnType<typeof inspectExactParagraphSpacing>>
  try { packet = await inspectExactParagraphSpacing(bytes, value.paragraph_text) }
  catch { throw new ParagraphSpacingProposalError(spacingUnsupported) }
  const matches = packet.matches.filter(match => match.paragraphHandle === value.paragraph_handle)
  if (packet.sourceSha256 !== value.source_sha256 || matches.length !== 1) throw new ParagraphSpacingProposalError(spacingSnapshotChanged)
  const target = matches[0]!
  if (!target.eligibility.eligible) throw new ParagraphSpacingProposalError(spacingUnsupported)
  if (target.text !== value.paragraph_text || !isDeepStrictEqual(target.tableContext, value.table_context) || !isDeepStrictEqual(target.directAfter, value.before_direct_after)) throw new ParagraphSpacingProposalError(spacingSnapshotChanged)
  const request: ExactSpacingRequest = { sourceSha256: value.source_sha256, paragraphHandle: value.paragraph_handle, paragraphText: value.paragraph_text,
    tableContext: target.tableContext, spacingAfterPoints: Number(value.after_direct_after_twips) / 20, author: row.author }
  return { request, view: { ...frozenParagraphSpacingView(value), table_context: target.tableContext } }
}
export async function paragraphSpacingView(bytes: Uint8Array, row: RedlineRow) { return (await checkedParagraphSpacingOperation(bytes, row)).view }

/** Owns the document lock. Model work has already frozen the whole source and
 * exact selected paragraph; this call never writes canonical document bytes. */
export async function recordParagraphSpacingProposal(directory: string, name: string, input: ParagraphSpacingProposalBinding,
  review: { author: string; reason: string; openItems: string[]; preparedPath: string }) {
  if (registeredDocumentName(name) !== name || !name.toLowerCase().endsWith(".docx")) throw new ParagraphSpacingProposalError(spacingSourceOnly)
  let value: ParagraphSpacingProposalBinding
  try { value = paragraphSpacingProposalSchema.parse(input) } catch { throw new ParagraphSpacingProposalError(spacingInvalid) }
  const author = review.author.trim()
  return withArtifactMutationLocksAsync(directory, [name], async () => {
    const target = resolveMatterDocumentTarget(directory, name)
    if (target.docPath !== review.preparedPath) throw new ParagraphSpacingProposalError(spacingChanged)
    const bytes = new Uint8Array(readFileSync(target.docPath)), db = openDb(directory)
    try {
      if (getDocumentRole(db, target.docPath) !== "source" || await tableRowSourceHasRevisions(bytes)) throw new ParagraphSpacingProposalError(spacingSourceOnly)
      if (hash(bytes) !== value.source_sha256) throw new ParagraphSpacingProposalError(spacingChanged)
      assertRedlineProposalCompatibility(listPendingRedlines(db, target.docPath), [{ scope: "paragraph_spacing" }])
      const targetJson = JSON.stringify(value)
      const row: RedlineRow = { id: 0, doc_path: target.docPath, doc_name: target.name, scope: "paragraph_spacing", find_text: value.paragraph_text,
        old_text: value.paragraph_text, new_text: value.paragraph_text, author, review_comment: review.reason, anchor_id: null, formatting_json: null,
        target_json: targetJson, created_at: Date.now(), status: "pending" }
      try {
        const operation = await checkedParagraphSpacingOperation(bytes, row)
        await authorExactParagraphSpacingReview(bytes, operation.request, row.review_comment)
      } catch (error) {
        if (error instanceof ParagraphSpacingProposalError) throw error
        throw new ParagraphSpacingProposalError(spacingUnsupported)
      }
      return db.transaction(() => {
        // No await can separate the final identity/source/role check from the
        // compatibility check, INSERT and prepared-metadata publication.
        if (resolveMatterDocumentTarget(directory, name).docPath !== review.preparedPath || getDocumentRole(db, target.docPath) !== value.prepared_role || hash(readFileSync(target.docPath)) !== value.source_sha256) throw new ParagraphSpacingProposalError(spacingChanged)
        assertRedlineProposalCompatibility(listPendingRedlines(db, target.docPath), [{ scope: "paragraph_spacing" }])
        const id = Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,review_comment,anchor_id,target_json,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)", [
          row.doc_path, row.doc_name, row.scope, row.find_text, row.old_text, row.new_text, row.author, row.review_comment, null, targetJson, row.created_at,
        ]).lastInsertRowid)
        setPreparedChanges(db, target.docPath, { redline_ids: [id], open_items: review.openItems })
        return { id }
      }).immediate()
    } finally { db.close() }
  })
}
