import { z } from "zod"

const literal = z.string().refine(text => text.isWellFormed() && !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(text), "Unsupported literal text")
const verticalMerge = z.enum(["none", "restart", "continue"])
const tableContext = z.object({
  table: z.number().int().positive(), row: z.number().int().positive().nullable(), column: z.number().int().positive().nullable(),
  row_context: z.array(literal.max(65536)).max(64),
  row_span: z.number().int().nonnegative().optional(), column_span: z.number().int().positive().optional(), vertical_merge: verticalMerge.optional(),
  row_cells: z.array(z.object({ column: z.number().int().positive(), column_span: z.number().int().positive(), row_span: z.number().int().nonnegative(), vertical_merge: verticalMerge, text: literal.max(65536) }).strict()).max(64).optional(),
  unavailable_reason: z.enum(["pending_table_structure", "unsupported_table_geometry"]).optional(),
}).strict().refine(value => JSON.stringify(value).length <= 65536, "Table context is too large")
export const paragraphSpacingBeforeSchema = z.union([
  z.object({ kind: z.literal("absent") }).strict(),
  z.object({ kind: z.literal("twips"), raw: z.string().regex(/^\d{1,32}$/), normalized: z.string().regex(/^(0|[1-9]\d{0,31})$/) }).strict()
    .refine(value => /^\d{1,32}$/.test(value.raw) && BigInt(value.raw).toString() === value.normalized, "Direct spacing values disagree"),
])
export const paragraphSpacingAfterSchema = z.string().regex(/^(0|[1-9]\d{0,2})$/).refine(value => Number(value) <= 960 && String(Number(value) / 20 * 20) === value, "Unsupported spacing after")
export const paragraphSpacingProposalSchema = z.object({
  version: z.literal(4), kind: z.literal("paragraph_spacing_after"), prepared_role: z.literal("source"),
  source_sha256: z.string().regex(/^[a-f0-9]{64}$/), paragraph_handle: z.string().min(1).max(200),
  paragraph_text: literal.min(1).max(8192), table_context: tableContext.nullable(),
  before_direct_after: paragraphSpacingBeforeSchema, after_direct_after_twips: paragraphSpacingAfterSchema,
}).strict().refine(value => value.before_direct_after.kind === "absent" || value.before_direct_after.normalized !== value.after_direct_after_twips, "The direct spacing is already set to that value")
export type ParagraphSpacingProposalBinding = z.infer<typeof paragraphSpacingProposalSchema>
export type ParagraphSpacingBefore = z.infer<typeof paragraphSpacingBeforeSchema>
export type ParagraphSpacingTableContext = z.infer<typeof tableContext>
export type ParagraphSpacingProposalView = {
  kind: "paragraph_spacing_after"; paragraph_text: string; before_direct_after: ParagraphSpacingBefore; after_direct_after_twips: string;
  /** Frozen context remains readable when the location is stale. Physical cell order. */
  row_context: string[] | null;
} & ({ table_context: ParagraphSpacingTableContext | null } | { table_context?: never })
export function frozenParagraphSpacingView(value: ParagraphSpacingProposalBinding): ParagraphSpacingProposalView {
  return { kind: value.kind, paragraph_text: value.paragraph_text, before_direct_after: { ...value.before_direct_after }, after_direct_after_twips: value.after_direct_after_twips, row_context: value.table_context ? [...value.table_context.row_context] : null }
}
export function paragraphSpacingProposalMarker(row: { scope: string; target_json?: string | null }) {
  if (row.scope === "paragraph_spacing") return true
  try {
    const value: unknown = JSON.parse(row.target_json ?? "null")
    return !!value && typeof value === "object" && !Array.isArray(value) &&
      ((value as { version?: unknown }).version === 4 || (value as { kind?: unknown }).kind === "paragraph_spacing_after")
  } catch { return false }
}
export const paragraphSpacingExclusiveReason = "A paragraph spacing change must be reviewed on its own. Review or reject the pending changes before preparing another change."
