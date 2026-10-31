import { z } from "zod"

export const paragraphInsertionProposalSchema=z.object({
  version:z.literal(5),kind:z.literal("paragraph_insert"),prepared_role:z.literal("source"),source_sha256:z.string().regex(/^[a-f0-9]{64}$/),
  source_document_id:z.number().int().positive(),target_handle:z.string().min(1),reference_text:z.string().min(1).max(8192),position:z.enum(["before","after"]),text:z.string().min(1).max(4000),
  location:z.object({headings:z.array(z.string()),previous_paragraph:z.string().nullable(),next_paragraph:z.string().nullable(),body_paragraph_number:z.number().int().positive()}).strict(),
}).strict().refine(value=>value.text===value.text.trim()&&value.text.isWellFormed()&&!/[\u0000-\u001f\u007f]/.test(value.text),"Unsupported literal paragraph")
export type ParagraphInsertionProposalBinding=z.infer<typeof paragraphInsertionProposalSchema>
export type ParagraphInsertionProposalView={kind:"paragraph_insert";reference_text:string;position:"before"|"after";text:string;location:ParagraphInsertionProposalBinding["location"]}
export function frozenParagraphInsertionView(value:ParagraphInsertionProposalBinding):ParagraphInsertionProposalView {
  return {kind:value.kind,reference_text:value.reference_text,position:value.position,text:value.text,location:value.location}
}
export function paragraphInsertionProposalMarker(row:{scope:string;target_json?:string|null}) {
  if(row.scope==="paragraph_insert")return true
  try {const value:unknown=JSON.parse(row.target_json??"null");return !!value&&typeof value==="object"&&!Array.isArray(value)&&
    ((value as {version?:unknown}).version===5||(value as {kind?:unknown}).kind==="paragraph_insert" )}catch{return false}
}
export const paragraphInsertionExclusiveReason="Review or reject pending changes before preparing a paragraph insertion. This change must be reviewed on its own."
