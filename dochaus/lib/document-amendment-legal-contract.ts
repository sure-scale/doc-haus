import { z } from "zod"

export const DOCUMENT_AMENDMENT_LEGAL_OUTPUT = "document-amendment-legal-v1"
const boundedText = z.string().max(8192)
const exactText = boundedText.min(1)
const reason = z.string().min(1).max(4000)
const cells = z.array(z.string().max(4000)).min(1).max(32)
const amendment = z.object({
  before_quote: exactText, replacement: z.string().max(16384),
  target_context: z.string().min(1), reason: z.string(),
}).strict()
const paragraphInsertion = z.object({
  reference_text: exactText, position: z.enum(["before", "after"]),
  text: z.string().min(1).max(4000), target_context: exactText, reason,
}).strict()
const rowInsertion = z.object({
  reference_cells: cells, position: z.enum(["before", "after"]), cells, target_context: exactText, reason,
}).strict()
const rowDeletion = z.object({ reference_cells: cells, target_context: exactText, reason }).strict()
const formatting = z.union([
  z.object({
    kind: z.literal("paragraph_spacing_after"), paragraph_text: exactText,
    spacing_after_points: z.number().min(0).max(48), target_context: exactText, reason,
  }).strict(),
  z.object({ kind: z.literal("unsupported_formatting"), request: exactText, reason }).strict(),
])

// A model transport contract, shared with legal admission. It supplies required
// fields and bounded operation shapes; exact source, literal-cell and operation
// eligibility checks still belong to the document-amendment writer.
export const documentAmendmentLegalSchema = z.object({
  amendments: z.array(amendment).max(24),
  open_items: z.array(boundedText).max(24),
  paragraph_insertions: z.array(paragraphInsertion).max(24).optional(),
  table_rows: z.array(rowInsertion).max(24).optional(),
  table_row_deletions: z.array(rowDeletion).max(24).optional(),
  paragraph_formats: z.array(formatting).max(24).optional(),
}).strict()
export const documentAmendmentLegalJsonSchema = z.toJSONSchema(documentAmendmentLegalSchema)
export const decodeDocumentAmendmentLegal = (value: unknown) => documentAmendmentLegalSchema.parse(value)
