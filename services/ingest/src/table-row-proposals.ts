import { createHash } from "node:crypto"
import { readFileSync } from "node:fs"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { docxTableRowInsertionTargets, type DocxTableRowInsertion } from "../../../dochaus/lib/docx-table-row-insertion"
import { authorDocxTableRowReview } from "../../../dochaus/lib/docx-table-row-review"
import { assertRedlineProposalCompatibility, tableRowProposalMarker, tableRowProposalSchema, type TableRowProposalBinding, type TableRowProposalView } from "../../../dochaus/lib/table-row-proposal"
import { getDocumentRole, listPendingRedlines, openDb, setPreparedChanges, type RedlineRow } from "./db"
import { resolveMatterDocumentTarget } from "./document-identity"
import { registeredDocumentName } from "./registered-document"
import { docxodus } from "./docxodus"

const hash = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
export class TableRowProposalError extends Error {}
export const rowSourceOnly = "Row insertions currently require an uploaded Word source without tracked changes."
export const rowChanged = "This row insertion targets an earlier Word file or document role. Reject it and prepare a fresh change list."
export const rowMixed = "A row insertion must be reviewed on its own. Reject the mixed change list and prepare separate requests."
export const rowUnsupported = "This row insertion cannot be safely prepared for this table. Use a simple unmerged table, or make the structural change directly in Word."
export async function tableRowSourceHasRevisions(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes,{persistAnchorIds:false})
  try { return session.listRevisions().length > 0 } finally { session.close() }
}

export function tableRowBinding(row: RedlineRow): TableRowProposalBinding {
  try {
    if (row.scope !== "table_row" || row.formatting_json != null || row.anchor_id != null || !row.author.trim() || row.author !== row.author.trim() || !row.review_comment.trim()) throw new Error()
    return tableRowProposalSchema.parse(JSON.parse(row.target_json ?? "null"))
  } catch { throw new TableRowProposalError("This row insertion has an invalid stored operation. Reject it and prepare a fresh change list.") }
}
export function assertTableRowSingleton(rows: readonly RedlineRow[]) {
  if (rows.some(tableRowProposalMarker) && rows.length !== 1) throw new TableRowProposalError(rowMixed)
}
export function tableRowRequest(binding: TableRowProposalBinding, author: string): DocxTableRowInsertion {
  return { source_sha256: binding.source_sha256, target_handle: binding.target_handle, position: binding.position, cells: [...binding.cells], author }
}
export async function tableRowView(bytes: Uint8Array, row: RedlineRow): Promise<TableRowProposalView & { table: number; row: number; reference_cells: string[] }> {
  const binding = tableRowBinding(row)
  if (hash(bytes) !== binding.source_sha256) throw new TableRowProposalError(rowChanged)
  const target = (await docxTableRowInsertionTargets(bytes)).targets.find(target => target.target_handle === binding.target_handle)
  if (!target || target.columns !== binding.cells.length) throw new TableRowProposalError(rowUnsupported)
  return { kind: "table_row_insert", position: binding.position, table: target.table, row: target.row, reference_cells: target.cells, cells: [...binding.cells] }
}

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
        view = await tableRowView(bytes, row)
        await authorDocxTableRowReview(bytes, tableRowRequest(binding, author), review.reason)
      } catch (error) {
        if (error instanceof TableRowProposalError) throw error
        throw new TableRowProposalError(rowUnsupported)
      }
      // No await separates this last source/role check from the INSERT.
      return db.transaction(() => {
        if (resolveMatterDocumentTarget(directory,name).docPath !== review.preparedPath || getDocumentRole(db, target.docPath) !== binding.prepared_role || hash(readFileSync(target.docPath)) !== binding.source_sha256) throw new TableRowProposalError(rowChanged)
        assertRedlineProposalCompatibility(listPendingRedlines(db, target.docPath), [{ scope: "table_row" }])
        const id = Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,review_comment,anchor_id,target_json,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)", [
          target.docPath, target.name, "table_row", `Table ${view.table}, row ${view.row}`, view.reference_cells.join(" | "), binding.cells.join(" | "), author, review.reason, null, JSON.stringify(binding), Date.now(),
        ]).lastInsertRowid)
        setPreparedChanges(db, target.docPath, { redline_ids: [id], open_items: review.openItems })
        return { id }
      }).immediate()
    } finally { db.close() }
  })
}
