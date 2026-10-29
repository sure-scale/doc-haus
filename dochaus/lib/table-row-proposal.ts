import { z } from "zod"
import { paragraphSpacingExclusiveReason, paragraphSpacingProposalMarker } from "./paragraph-spacing-proposal"
import { paragraphInsertionExclusiveReason, paragraphInsertionProposalMarker } from "./paragraph-insertion-proposal"

// A structural proposal binds the complete uploaded source. Display strings or
// coordinates never locate the write; only the verified source/target pair does.
export const tableRowInsertionProposalSchema = z.object({
  version: z.literal(3), kind: z.literal("table_row_insert"), prepared_role: z.literal("source"),
  source_sha256: z.string().regex(/^[a-f0-9]{64}$/), target_handle: z.string().min(1),
  position: z.enum(["before", "after"]), cells: z.array(z.string().max(4000)).min(1).max(32),
}).strict().refine(value => value.cells.some(cell => cell.length) && value.cells.join("").length <= 24000 &&
  value.cells.every(cell => cell === cell.trim() && cell.isWellFormed() && !/[\u0000-\u001f\u007f]/.test(cell)), "Unsupported literal cell content")
// Deleted cells are original-source evidence. Preserve their exact spaces;
// never apply insertion's trim rule or turn them into replacement content.
export const tableRowDeletionCellsSchema = z.array(z.string().max(4000)).min(1).max(32).refine(cells => cells.some(cell => cell.length) && cells.join("").length <= 24000 &&
  cells.every(cell => cell.isWellFormed() && !/[\u0000-\u001f\u007f]/.test(cell)), "Unsupported original cell content")
export const tableRowDeletionProposalSchema = z.object({
  version: z.literal(3), kind: z.literal("table_row_delete"), prepared_role: z.literal("source"),
  source_sha256: z.string().regex(/^[a-f0-9]{64}$/), target_handle: z.string().min(1), cells: tableRowDeletionCellsSchema,
}).strict()
export const tableRowProposalSchema = z.union([tableRowInsertionProposalSchema, tableRowDeletionProposalSchema])
export type TableRowInsertionProposalBinding = z.infer<typeof tableRowInsertionProposalSchema>
export type TableRowDeletionProposalBinding = z.infer<typeof tableRowDeletionProposalSchema>
export type TableRowProposalBinding = z.infer<typeof tableRowProposalSchema>
export type TableRowInsertionProposalView = { kind: "table_row_insert"; position: "before" | "after"; cells: string[] } &
  ({ table: number; row: number; reference_cells: string[] } | { table?: never; row?: never; reference_cells?: never })
export type TableRowDeletionProposalView = { kind: "table_row_delete"; cells: string[] } &
  ({ table: number; row: number } | { table?: never; row?: never })
export type TableRowProposalView = TableRowInsertionProposalView | TableRowDeletionProposalView
export function frozenTableRowView(binding: TableRowProposalBinding): TableRowProposalView {
  return binding.kind === "table_row_insert" ? {kind:binding.kind,position:binding.position,cells:[...binding.cells]} : {kind:binding.kind,cells:[...binding.cells]}
}

export function tableRowProposalMarker(row: { scope: string; target_json?: string | null }) {
  if (row.scope === "table_row") return true
  try {
    const value: unknown = JSON.parse(row.target_json ?? "null")
    return !!value && typeof value === "object" && !Array.isArray(value) &&
      ((value as { version?: unknown }).version === 3 || ["table_row_insert","table_row_delete"].includes(String((value as { kind?: unknown }).kind)))
  } catch { return false }
}

export function legacyClauseProposalMarker(row: { scope: string; target_json?: string | null }) {
  try { return JSON.parse(row.target_json ?? "null")?.kind === "legacy_clause" } catch { return false }
}
const legacyClauseExclusiveReason = "Review or reject the pending changes before preparing this clause change. This change must be reviewed on its own."
export const tableRowExclusiveReason = "Review or reject the pending changes before preparing a table row change. A table row change must be reviewed on its own."
export class RedlineProposalConflictError extends Error {}
export function exclusiveRedlineProposalMarker(row: { scope: string; target_json?: string | null }) {
  return tableRowProposalMarker(row) || paragraphSpacingProposalMarker(row) || paragraphInsertionProposalMarker(row)
}
export function assertExclusiveRedlineSingleton(rows: readonly { scope: string; target_json?: string | null }[]) {
  if (rows.some(legacyClauseProposalMarker) && rows.length !== 1) throw new RedlineProposalConflictError(legacyClauseExclusiveReason)
  if (rows.some(exclusiveRedlineProposalMarker) && rows.length !== 1)
    throw new RedlineProposalConflictError(rows.some(paragraphInsertionProposalMarker) ? paragraphInsertionExclusiveReason : rows.some(paragraphSpacingProposalMarker) ? paragraphSpacingExclusiveReason : tableRowExclusiveReason)
}
/** Call inside each recorder's immediate transaction, before supersession. */
export function assertRedlineProposalCompatibility(
  pending: readonly { scope: string; target_json?: string | null }[], incoming: readonly { scope: string; target_json?: string | null }[],
) {
  if ((pending.some(legacyClauseProposalMarker) && incoming.length) || (incoming.some(legacyClauseProposalMarker) && (pending.length || incoming.length !== 1)))
    throw new RedlineProposalConflictError(legacyClauseExclusiveReason)
  if ((pending.some(exclusiveRedlineProposalMarker) && incoming.length) ||
      (incoming.some(exclusiveRedlineProposalMarker) && (pending.length || incoming.length !== 1)))
    throw new RedlineProposalConflictError([...pending, ...incoming].some(paragraphInsertionProposalMarker) ? paragraphInsertionExclusiveReason : [...pending, ...incoming].some(paragraphSpacingProposalMarker) ? paragraphSpacingExclusiveReason : tableRowExclusiveReason)
}
