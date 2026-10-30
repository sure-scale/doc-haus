import { createHash } from "node:crypto"
import { lstatSync, readFileSync, realpathSync } from "node:fs"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { paragraphInsertionTargets, type DocxParagraphInsertion } from "../../../dochaus/lib/docx-paragraph-insertion"
import { authorParagraphInsertionReview } from "../../../dochaus/lib/docx-paragraph-insertion-review"
import { validateDocxRowReviewReason } from "../../../dochaus/lib/docx-table-row-review"
import { paragraphInsertionProposalSchema, frozenParagraphInsertionView, type ParagraphInsertionProposalBinding } from "../../../dochaus/lib/paragraph-insertion-proposal"
import { assertRedlineProposalCompatibility } from "../../../dochaus/lib/table-row-proposal"
import { getRedline, listPendingRedlines, openDb, setPreparedChanges, type RedlineRow } from "./db"
import { resolveMatterDocumentTarget } from "./document-identity"
import { registeredDocumentName } from "./registered-document"
import { tableRowSourceHasRevisions } from "./table-row-proposals"
const hash=(bytes:Uint8Array)=>createHash("sha256").update(bytes).digest("hex")
export class ParagraphInsertionProposalError extends Error {}
export const insertionSourceOnly="Paragraph insertion currently requires an uploaded Word source without tracked changes."
export const insertionChanged="This paragraph insertion targets an earlier Word file or document role. Reject it and prepare a fresh change list."
export const insertionUnsupported="This paragraph cannot be safely inserted here as a tracked proposal. Make the structural change in Word."
export const insertionInvalid="This paragraph insertion has an invalid stored operation. Reject it and prepare a fresh change list."
export function sourceDocumentId(db:ReturnType<typeof openDb>,file:string):number|null {
  const row=db.query("SELECT id, role FROM documents WHERE doc_path = ?").get(file) as {id:number;role:string}|null
  return row?.role==="source"?row.id:null
}
// Called within the writer's registry transaction after staged bytes are installed.
// Do not hash the staged file here: the pre-write check covers the original bytes.
export function checkParagraphInsertionAuthority(db:ReturnType<typeof openDb>,rows:readonly RedlineRow[],file:string) {
  for(const row of rows.filter(r=>r.scope==="paragraph_insert")) {
    const value=paragraphInsertionBinding(row),current=getRedline(db,row.id)
    const pending=listPendingRedlines(db,file)
    try {
      if(row.doc_path!==file||sourceDocumentId(db,file)!==value.source_document_id||lstatSync(file).isSymbolicLink()||realpathSync.native(file)!==file||
        pending.length!==1||pending[0]?.id!==row.id||!current||current.status!=="pending"||JSON.stringify(current)!==JSON.stringify(row)) throw new Error()
    }catch{throw new ParagraphInsertionProposalError(insertionChanged)}
  }
}
export function paragraphInsertionBinding(row:RedlineRow):ParagraphInsertionProposalBinding {
  try {
    const value=paragraphInsertionProposalSchema.parse(JSON.parse(row.target_json??"null"))
    if(row.scope!=="paragraph_insert"||row.formatting_json!=null||row.anchor_id!=null||!row.author.trim()||row.author!==row.author.trim()||row.author.length>100||!row.author.isWellFormed()||/[\u0000-\u001f\u007f]/.test(row.author)||
      row.find_text!==value.reference_text||row.old_text!==""||row.new_text!==value.text)throw new Error()
    validateDocxRowReviewReason(row.review_comment);return value
  }catch{throw new ParagraphInsertionProposalError(insertionInvalid)}
}
export async function checkedParagraphInsertionOperation(bytes:Uint8Array,row:RedlineRow) {
  const value=paragraphInsertionBinding(row)
  if(hash(bytes)!==value.source_sha256)throw new ParagraphInsertionProposalError(insertionChanged)
  let packet:Awaited<ReturnType<typeof paragraphInsertionTargets>>
  try{packet=await paragraphInsertionTargets(bytes)}catch{throw new ParagraphInsertionProposalError(insertionUnsupported)}
  const targets=packet.targets.filter(target=>target.target_handle===value.target_handle)
  if(packet.source_sha256!==value.source_sha256||targets.length!==1||targets[0]!.text!==value.reference_text||
    JSON.stringify({headings:targets[0]!.headings,previous_paragraph:targets[0]!.previous_paragraph,next_paragraph:targets[0]!.next_paragraph,body_paragraph_number:targets[0]!.body_paragraph_number})!==JSON.stringify(value.location))throw new ParagraphInsertionProposalError(insertionChanged)
  if(!targets[0]!.supported||!targets[0]![value.position])throw new ParagraphInsertionProposalError(insertionUnsupported)
  const request:DocxParagraphInsertion={source_sha256:value.source_sha256,target_handle:value.target_handle,position:value.position,text:value.text,author:row.author}
  return {request,view:frozenParagraphInsertionView(value)}
}
export async function recordParagraphInsertionProposal(directory:string,name:string,input:ParagraphInsertionProposalBinding,
  review:{author:string;reason:string;openItems:string[];preparedPath:string}) {
  if(registeredDocumentName(name)!==name||!name.toLowerCase().endsWith(".docx"))throw new ParagraphInsertionProposalError(insertionSourceOnly)
  let value:ParagraphInsertionProposalBinding
  try{value=paragraphInsertionProposalSchema.parse(input)}catch{throw new ParagraphInsertionProposalError(insertionInvalid)}
  const author=review.author.trim()
  return withArtifactMutationLocksAsync(directory,[name],async()=>{
    const target=resolveMatterDocumentTarget(directory,name)
    if(target.docPath!==review.preparedPath)throw new ParagraphInsertionProposalError(insertionChanged)
    const bytes=new Uint8Array(readFileSync(target.docPath)),db=openDb(directory)
    try{
      if(sourceDocumentId(db,target.docPath)!==value.source_document_id||await tableRowSourceHasRevisions(bytes))throw new ParagraphInsertionProposalError(insertionSourceOnly)
      if(hash(bytes)!==value.source_sha256)throw new ParagraphInsertionProposalError(insertionChanged)
      assertRedlineProposalCompatibility(listPendingRedlines(db,target.docPath),[{scope:"paragraph_insert"}])
      const targetJson=JSON.stringify(value),row:RedlineRow={id:0,doc_path:target.docPath,doc_name:target.name,scope:"paragraph_insert",find_text:value.reference_text,old_text:"",new_text:value.text,
        author,review_comment:review.reason,anchor_id:null,formatting_json:null,target_json:targetJson,created_at:Date.now(),status:"pending"}
      try{const operation=await checkedParagraphInsertionOperation(bytes,row);await authorParagraphInsertionReview(bytes,operation.request,row.review_comment)}
      catch(error){if(error instanceof ParagraphInsertionProposalError)throw error;throw new ParagraphInsertionProposalError(insertionUnsupported)}
      return db.transaction(()=>{
        if(resolveMatterDocumentTarget(directory,name).docPath!==review.preparedPath||sourceDocumentId(db,target.docPath)!==value.source_document_id||hash(readFileSync(target.docPath))!==value.source_sha256)throw new ParagraphInsertionProposalError(insertionChanged)
        assertRedlineProposalCompatibility(listPendingRedlines(db,target.docPath),[{scope:"paragraph_insert"}])
        const id=Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,review_comment,anchor_id,target_json,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
          [row.doc_path,row.doc_name,row.scope,row.find_text,row.old_text,row.new_text,row.author,row.review_comment,null,targetJson,row.created_at]).lastInsertRowid)
        setPreparedChanges(db,target.docPath,{redline_ids:[id],open_items:review.openItems});return {id}
      }).immediate()
    }finally{db.close()}
  })
}
