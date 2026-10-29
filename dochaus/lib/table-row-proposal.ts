import { z } from "zod"

// A structural proposal binds the complete uploaded source. Display strings or
// coordinates never locate the write; only the verified source/target pair does.
export const tableRowProposalSchema = z.object({
  version: z.literal(3), kind: z.literal("table_row_insert"), prepared_role: z.literal("source"),
  source_sha256: z.string().regex(/^[a-f0-9]{64}$/), target_handle: z.string().min(1),
  position: z.enum(["before", "after"]), cells: z.array(z.string().max(4000)).min(1).max(32),
}).strict().refine(value => value.cells.some(cell => cell.length) && value.cells.join("").length <= 24000 &&
  value.cells.every(cell => cell === cell.trim() && cell.isWellFormed() && !/[\u0000-\u001f\u007f]/.test(cell)), "Unsupported literal cell content")
export type TableRowProposalBinding = z.infer<typeof tableRowProposalSchema>
export type TableRowProposalView = { kind: "table_row_insert"; position: "before" | "after"; cells: string[] } &
  ({ table: number; row: number; reference_cells: string[] } | { table?: never; row?: never; reference_cells?: never })

export function tableRowProposalMarker(row: { scope: string; target_json?: string | null }) {
  if (row.scope === "table_row") return true
  try {
    const value: unknown = JSON.parse(row.target_json ?? "null")
    return !!value && typeof value === "object" && !Array.isArray(value) &&
      ((value as { version?: unknown }).version === 3 || (value as { kind?: unknown }).kind === "table_row_insert")
  } catch { return false }
}

export const tableRowExclusiveReason = "Review or reject the pending changes before preparing a row insertion. A row insertion must be reviewed on its own."
export class RedlineProposalConflictError extends Error {}
/** Call inside each recorder's immediate transaction, before supersession. */
export function assertRedlineProposalCompatibility(
  pending: readonly { scope: string; target_json?: string | null }[], incoming: readonly { scope: string; target_json?: string | null }[],
) {
  if ((pending.some(tableRowProposalMarker) && incoming.length) ||
      (incoming.some(tableRowProposalMarker) && (pending.length || incoming.length !== 1))) throw new RedlineProposalConflictError(tableRowExclusiveReason)
}
