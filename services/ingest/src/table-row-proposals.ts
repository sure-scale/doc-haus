import { createHash } from "node:crypto"
import { readFileSync } from "node:fs"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { docxTableRowInsertionTargets, type DocxTableRowInsertion } from "../../../dochaus/lib/docx-table-row-insertion"
import { authorDocxTableRowReview } from "../../../dochaus/lib/docx-table-row-review"
import { docxTableRowDeletionTargets, type DocxTableRowDeletion } from "../../../dochaus/lib/docx-table-row-deletion"
import { authorDocxTableRowDeletionReview } from "../../../dochaus/lib/docx-table-row-deletion-review"
import { assertRedlineProposalCompatibility, tableRowProposalMarker, tableRowProposalSchema, type TableRowProposalBinding, type TableRowInsertionProposalBinding, type TableRowInsertionProposalView, type TableRowDeletionProposalView } from "../../../dochaus/lib/table-row-proposal"
import { getDocumentRole, listPendingRedlines, openDb, setPreparedChanges, type RedlineRow } from "./db"
import { resolveMatterDocumentTarget } from "./document-identity"
import { registeredDocumentName } from "./registered-document"
import { docxodus } from "./docxodus"

const hash = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
export class TableRowProposalError extends Error {}
export const rowSourceOnly = "Table row changes currently require an uploaded Word source without tracked changes."
export const rowChanged = "This row change targets an earlier Word file or document role. Reject it and prepare a fresh change list."
export const rowMixed = "A table row change must be reviewed on its own. Reject the mixed change list and prepare separate requests."
export const rowUnsupported = "This row insertion cannot be safely prepared for this table. Use a simple unmerged table, or make the structural change directly in Word."
export const rowDeletionUnsupported = "This row cannot be deleted safely as a tracked proposal. Make this change directly in Word."
export const rowSnapshotChanged = "The complete row contents do not match this deletion proposal. Reject it and prepare a fresh change list."
export async function tableRowSourceHasRevisions(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes,{persistAnchorIds:false})
  try { return session.listRevisions().length > 0 } finally { session.close() }
}

export function tableRowBinding(row: RedlineRow): TableRowProposalBinding {
  try {
    if (row.scope !== "table_row" || row.formatting_json != null || row.anchor_id != null || !row.author.trim() || row.author !== row.author.trim() || !row.review_comment.trim()) throw new Error()
    return tableRowProposalSchema.parse(JSON.parse(row.target_json ?? "null"))
  } catch { throw new TableRowProposalError("This row change has an invalid stored operation. Reject it and prepare a fresh change list.") }
}
export function assertTableRowSingleton(rows: readonly RedlineRow[]) {
  if (rows.some(tableRowProposalMarker) && rows.length !== 1) throw new TableRowProposalError(rowMixed)
}
export function tableRowRequest(binding: TableRowInsertionProposalBinding, author: string): DocxTableRowInsertion {
  return { source_sha256: binding.source_sha256, target_handle: binding.target_handle, position: binding.position, cells: [...binding.cells], author }
}
type CheckedTableRowOperation =
  | {kind:"table_row_insert";request:DocxTableRowInsertion;view:TableRowInsertionProposalView & {table:number;row:number;reference_cells:string[]}}
  | {kind:"table_row_delete";request:DocxTableRowDeletion;view:TableRowDeletionProposalView & {table:number;row:number}}
/** Every recorder/view/writer uses this check. A deletion's frozen complete
 * cells must still equal the exact eligible source row, not only its handle. */
export async function checkedTableRowOperation(bytes: Uint8Array, row: RedlineRow): Promise<CheckedTableRowOperation> {
  const binding = tableRowBinding(row)
  if (hash(bytes) !== binding.source_sha256) throw new TableRowProposalError(rowChanged)
  if (binding.kind === "table_row_delete") {
    let target: Awaited<ReturnType<typeof docxTableRowDeletionTargets>>["targets"][number] | undefined
    try { target = (await docxTableRowDeletionTargets(bytes)).targets.find(target => target.target_handle === binding.target_handle) }
    catch { throw new TableRowProposalError(rowDeletionUnsupported) }
    if (!target) throw new TableRowProposalError(rowDeletionUnsupported)
    if (JSON.stringify(target.cells) !== JSON.stringify(binding.cells)) throw new TableRowProposalError(rowSnapshotChanged)
    return {kind:binding.kind,request:{source_sha256:binding.source_sha256,target_handle:binding.target_handle,author:row.author},view:{kind:binding.kind,cells:[...binding.cells],table:target.table,row:target.row}}
  }
  const target = (await docxTableRowInsertionTargets(bytes)).targets.find(target => target.target_handle === binding.target_handle)
  if (!target || target.columns !== binding.cells.length) throw new TableRowProposalError(rowUnsupported)
  return {kind:binding.kind,request:tableRowRequest(binding,row.author),view:{ kind: binding.kind, position: binding.position, table: target.table, row: target.row, reference_cells: target.cells, cells: [...binding.cells] }}
}
export async function tableRowView(bytes: Uint8Array, row: RedlineRow) { return (await checkedTableRowOperation(bytes,row)).view }

/** Model work ends before this call. It records only the frozen, verified row;
 * no canonical bytes are written. The shared transaction gate also covers
 * proposals admitted by a tool after its permission prompt resumes. */
export async function recordTableRowProposal(directory: string, name: string, value: TableRowProposalBinding,
  review: { author: string; reason: string; openItems: string[]; preparedPath: string }) {
  if (registeredDocumentName(name) !== name || !name.toLowerCase().endsWith(".docx")) throw new TableRowProposalError(rowSourceOnly)
  const binding = tableRowProposalSchema.parse(value), author = review.author.trim()
  return withArtifactMutationLocksAsync(directory, [name], async () => {
    const target = resolveMatterDocumentTarget(directory, name)
    if (target.docPath !== review.preparedPath) throw new TableRowProposalError(rowChanged)
    const bytes = new Uint8Array(readFileSync(target.docPath))
    const db = openDb(directory)
    try {
      if (getDocumentRole(db, target.docPath) !== "source" || await tableRowSourceHasRevisions(bytes)) throw new TableRowProposalError(rowSourceOnly)
      if (hash(bytes) !== binding.source_sha256) throw new TableRowProposalError(rowChanged)
      assertRedlineProposalCompatibility(listPendingRedlines(db, target.docPath), [{ scope: "table_row" }])
      let view: Awaited<ReturnType<typeof tableRowView>>
      try {
        const row = { scope: "table_row", author, review_comment: review.reason, anchor_id: null, target_json: JSON.stringify(binding) } as RedlineRow
        const operation = await checkedTableRowOperation(bytes, row)
        view = operation.view
        if (operation.kind === "table_row_delete") await authorDocxTableRowDeletionReview(bytes, operation.request, review.reason)
        else await authorDocxTableRowReview(bytes, operation.request, review.reason)
      } catch (error) {
        if (error instanceof TableRowProposalError) throw error
        throw new TableRowProposalError(binding.kind === "table_row_delete" ? rowDeletionUnsupported : rowUnsupported)
      }
      // No await separates this last source/role check from the INSERT.
      return db.transaction(() => {
        if (resolveMatterDocumentTarget(directory,name).docPath !== review.preparedPath || getDocumentRole(db, target.docPath) !== binding.prepared_role || hash(readFileSync(target.docPath)) !== binding.source_sha256) throw new TableRowProposalError(rowChanged)
        assertRedlineProposalCompatibility(listPendingRedlines(db, target.docPath), [{ scope: "table_row" }])
        const id = Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,review_comment,anchor_id,target_json,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)", [
          target.docPath, target.name, "table_row", `Table ${view.table}, row ${view.row}`, (view.kind === "table_row_insert" ? view.reference_cells : view.cells).join(" | "), binding.kind === "table_row_insert" ? binding.cells.join(" | ") : "", author, review.reason, null, JSON.stringify(binding), Date.now(),
        ]).lastInsertRowid)
        setPreparedChanges(db, target.docPath, { redline_ids: [id], open_items: review.openItems })
        return { id }
      }).immediate()
    } finally { db.close() }
  })
}
