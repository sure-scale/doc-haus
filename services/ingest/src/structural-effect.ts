import { authorParagraphInsertion, paragraphInsertionTargets, type DocxParagraphInsertion } from "../../../dochaus/lib/docx-paragraph-insertion"
import { authorDocxTableRowInsertion, docxTableRowInsertionTargets, verifyDocxTableRowInsertion, type DocxTableRowInsertion } from "../../../dochaus/lib/docx-table-row-insertion"
import { authorDocxTableRowDeletion, docxTableRowDeletionTargets, verifyDocxTableRowDeletion, type DocxTableRowDeletion } from "../../../dochaus/lib/docx-table-row-deletion"
import { authorExactParagraphSpacing, inspectExactParagraphSpacing, type ExactSpacingRequest } from "../../../dochaus/lib/docx-exact-spacing"

type Operation = {kind:"paragraph_insert";request:DocxParagraphInsertion} | {kind:"row_insert";request:DocxTableRowInsertion} | {kind:"row_delete";request:DocxTableRowDeletion} | {kind:"spacing_after";request:ExactSpacingRequest}

// Disposable authoring proves both selective decisions before returning; the
// returned operation is a detached copy of the exact request actually proved.
export async function structuralEffect(bytes:Uint8Array,input:Operation) {
  const source=new Uint8Array(bytes), operation=structuredClone(input)
  if(operation.kind === "paragraph_insert") {
    const packet=await paragraphInsertionTargets(source), target=packet.targets.find(row=>row.target_handle===operation.request.target_handle)
    if(!target || packet.source_sha256!==operation.request.source_sha256) throw new Error("structural_effect_target_changed")
    const authored=await authorParagraphInsertion(source,operation.request)
    return {operation,effect:{kind:operation.kind,location:target,position:operation.request.position,before:null,after:authored.paragraph.text},tracked:authored.bytes}
  }
  if(operation.kind === "row_insert") {
    const packet=await docxTableRowInsertionTargets(source),target=packet.targets.find(row=>row.target_handle===operation.request.target_handle)
    if(!target || packet.source_sha256!==operation.request.source_sha256) throw new Error("structural_effect_target_changed")
    const authored=await authorDocxTableRowInsertion(source,operation.request),proof=await verifyDocxTableRowInsertion(source,authored.bytes,operation.request)
    return {operation,effect:{kind:operation.kind,location:target,position:operation.request.position,before:null,after:proof.paragraphs.map(row=>row.text)},tracked:authored.bytes}
  }
  if(operation.kind === "row_delete") {
    const packet=await docxTableRowDeletionTargets(source),target=packet.targets.find(row=>row.target_handle===operation.request.target_handle)
    if(!target || packet.source_sha256!==operation.request.source_sha256) throw new Error("structural_effect_target_changed")
    const authored=await authorDocxTableRowDeletion(source,operation.request),proof=await verifyDocxTableRowDeletion(source,authored.bytes,operation.request)
    return {operation,effect:{kind:operation.kind,location:target,before:proof.paragraphs.map(row=>row.text),after:null},tracked:authored.bytes}
  }
  const packet=await inspectExactParagraphSpacing(source,operation.request.paragraphText),target=packet.matches.find(row=>row.paragraphHandle===operation.request.paragraphHandle)
  if(!target || packet.sourceSha256!==operation.request.sourceSha256 || JSON.stringify(target.tableContext)!==JSON.stringify(operation.request.tableContext)) throw new Error("structural_effect_target_changed")
  const authored=await authorExactParagraphSpacing(source,operation.request)
  return {operation,effect:{kind:operation.kind,location:target,before:authored.beforeDirect,after:authored.afterDirect,unit:"twips" as const},tracked:authored.trackedBytes}
}
