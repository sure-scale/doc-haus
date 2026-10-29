import { unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxCommentParagraphPositions, readDocxComments } from "./docx-comments"
import { assertExactSpacingPackage } from "./docx-exact-spacing"
import { assertDocxRowReviewPackage, peelAcceptedParagraphReviewComment, peelRemovedRowReviewComment, validateDocxRowReviewReason } from "./docx-table-row-review"
import { authorParagraphInsertion, resolveParagraphInsertion, verifyParagraphInsertion, normalizeParagraphInsertionDecisionSettings, type DocxParagraphInsertion } from "./docx-paragraph-insertion"

// Offline group review; not a Word or application approval action.
const BODY="word/document.xml", RELS="word/_rels/document.xml.rels", TYPES="[Content_Types].xml"
type Session=ReturnType<Awaited<ReturnType<typeof docxodus>>["openDocxSession"]>
type Revision=ReturnType<Session["listRevisions"]>[number]
function require(value:unknown,reason:string):asserts value {if(!value)throw new Error(`docx_paragraph_insertion_review:${reason}`)}
function copy(value:DocxParagraphInsertion) {
  require(typeof value.author==="string"&&value.author.trim()&&value.author.length<=100&&value.author.isWellFormed()&&!/[\u0000-\u001f\u007f]/.test(value.author),"invalid_author")
  return {...value,author:value.author.trim()}
}
function identity(row:Revision) {const {anchorId:_anchor,affectedAnchors:_affected,...rest}=row;return JSON.stringify(rest)}
function native(row:Revision) {const {id:_id,anchorId:_anchor,affectedAnchors:_affected,...rest}=row;return JSON.stringify(rest)}
async function registry(bytes:Uint8Array) {
  assertExactSpacingPackage(bytes);const session=(await docxodus()).openDocxSession(bytes,{persistAnchorIds:false})
  try{return session.listRevisions()}finally{session.close()}
}
const numeric=(id:string|null)=>{require(id&&/^[+-]?\d+$/.test(id),"invalid_comment_id");return BigInt(id).toString()}
function anchor(core:Uint8Array,paragraph:{xml_paragraph_index:number;text:string}) {
  const positions=readDocxCommentParagraphPositions(core).filter(p=>p.xml_paragraph_indexes.length===1&&p.xml_paragraph_indexes[0]===paragraph.xml_paragraph_index&&p.paragraph_text===paragraph.text)
  require(positions.length===1,"inserted_paragraph_unavailable")
  return {paragraph_index:positions[0]!.paragraph_index,paragraph_text:paragraph.text}
}
async function inspect(source:Uint8Array,bytes:Uint8Array,input:DocxParagraphInsertion,reason:string) {
  assertExactSpacingPackage(source);assertExactSpacingPackage(bytes)
  const old=new Set(readDocxComments(source).map(comment=>numeric(comment.comment_id)))
  const added=readDocxComments(bytes).filter(comment=>!old.has(numeric(comment.comment_id)))
  require(added.length===1&&added[0]!.markers.length===3,"owned_comment_unavailable")
  const index=added[0]!.markers[0]!.xml_paragraph_index
  require(index!==null&&Number.isSafeInteger(index)&&index>=0,"inserted_paragraph_unavailable")
  const peeled=peelAcceptedParagraphReviewComment(source,bytes,input.author,reason,{kind:"accepted",xml_paragraph_index:index,text:input.text})
  assertExactSpacingPackage(peeled.core)
  const verified=await verifyParagraphInsertion(source,peeled.core,input)
  require(verified.paragraph.xml_paragraph_index===index,"reason_not_on_inserted_paragraph")
  const [core,final]=await Promise.all([registry(peeled.core),registry(bytes)])
  require(JSON.stringify(core.map(identity).sort())===JSON.stringify(final.map(identity).sort()),"comment_changed_revision_membership")
  const expected=verified.revisions[0]!, owned=final.filter(row=>row.partUri===expected.partUri&&JSON.stringify(row.constituentKeys.toSorted())===JSON.stringify(expected.constituentKeys.toSorted()))
  require(owned.length===1&&native(owned[0]!)===native(expected),"native_group_unavailable")
  return {...peeled,verified,owned:owned[0]!,anchor:anchor(peeled.core,verified.paragraph)}
}
function member(session:Session,wanted:Revision) {
  const rows=session.listRevisions().filter(row=>row.partUri===wanted.partUri&&JSON.stringify(row.constituentKeys.toSorted())===JSON.stringify(wanted.constituentKeys.toSorted()))
  require(rows.length===1&&native(rows[0]!)===native(wanted),"native_group_changed")
  return rows[0]!
}
function digest(session:Session) {
  const d=session.getPackageManifest().rawPackageBytesDigest
  require(d?.algorithm==="SHA-256"&&typeof d.value==="string"&&/^[a-f0-9]{64}$/.test(d.value),"package_digest_unavailable")
  return d.value
}
function restore(source:Uint8Array,bytes:Uint8Array,commentPart:string) {
  assertExactSpacingPackage(source);assertExactSpacingPackage(bytes)
  const original=unzipSync(source), output=unzipSync(bytes)
  for(const [name,part] of Object.entries(original))if(name!==BODY&&name!==commentPart&&(![RELS,TYPES].includes(name)||original[commentPart]))output[name]=part
  const result=zipSync(output,{mtime:new Date(1980,0,1)});assertExactSpacingPackage(result);assertDocxRowReviewPackage(result,bytes)
  return result
}
export async function verifyParagraphInsertionReview(sourceBytes:Uint8Array,artifact:Uint8Array,value:DocxParagraphInsertion,storedReason:string) {
  const source=new Uint8Array(sourceBytes),bytes=new Uint8Array(artifact),input=copy(value),reason=validateDocxRowReviewReason(storedReason)
  const checked=await inspect(source,bytes,input,reason)
  return {commentId:checked.commentId,paragraph:checked.verified.paragraph,revisions:checked.verified.revisions}
}
export async function resolveParagraphInsertionReview(sourceBytes:Uint8Array,artifact:Uint8Array,value:DocxParagraphInsertion,storedReason:string,decision:"accept"|"reject") {
  require(decision==="accept"||decision==="reject","invalid_decision")
  const source=new Uint8Array(sourceBytes),bytes=new Uint8Array(artifact),input=copy(value),reason=validateDocxRowReviewReason(storedReason)
  const checked=await inspect(source,bytes,input,reason),expectedCore=await resolveParagraphInsertion(source,checked.core,input,decision)
  let expected=source
  if(decision==="accept") {
    const output=exportDocxComments(expectedCore,[{id:"paragraph-review",text:reason,author:input.author,anchor:checked.anchor}])
    require(output.unresolved.length===0&&output.attached.length===1&&output.attached[0]!.comment_id===checked.commentId,"accepted_reason_unresolved")
    expected=output.bytes
  }
  assertExactSpacingPackage(expected)
  const session=(await docxodus()).openDocxSession(bytes,{persistAnchorIds:false})
  try {
    member(session,checked.owned);session.listComments();session.listAnnotations();session.save()
    const before=digest(session),version=session.getVersion()
    const resolve={tool:"paragraph",action:decision,mutation:()=>{const found=member(session,checked.owned);return decision==="accept"?session.acceptRevision(found.id):session.rejectRevision(found.id)}}
    const steps=decision==="accept"?[resolve]:[{tool:"comment",action:"remove_reason",mutation:()=>{const rows=session.listComments().filter(row=>String(row.id)===checked.commentId);require(rows.length===1,"owned_comment_unavailable");return session.removeComment(rows[0]!.anchorId)}},resolve]
    const result=session.executeBatch(steps,"atomic")
    if(!result.success){require(session.getVersion()===version&&digest(session)===before,"failed_review_not_rolled_back");throw new Error("docx_paragraph_insertion_review:review_failed_rolled_back")}
    const saved=session.save();assertExactSpacingPackage(saved)
    const peeled=decision==="accept"?peelAcceptedParagraphReviewComment(source,saved,input.author,reason,{kind:"accepted",...checked.verified.paragraph}):peelRemovedRowReviewComment(source,saved,input.author,reason,checked.commentId)
    assertExactSpacingPackage(peeled.core);require(peeled.commentId===checked.commentId,"owned_comment_changed")
    // Peeling inspected the actual annotated package first. Validate the only
    // possible settings transition before restoring the accepted baseline.
    const normalized=await normalizeParagraphInsertionDecisionSettings(source,checked.core,saved,input)
    const normalizedCore=await normalizeParagraphInsertionDecisionSettings(source,checked.core,peeled.core,input)
    assertDocxRowReviewPackage(normalizedCore,expectedCore)
    if(decision==="accept")assertDocxRowReviewPackage(normalized,expected)
    const [actual,intended]=await Promise.all([registry(normalizedCore),registry(expectedCore)])
    require(JSON.stringify(actual.map(identity).sort())===JSON.stringify(intended.map(identity).sort()),"final_revision_membership_changed")
    return decision==="reject"?source:restore(expectedCore,normalized,checked.commentPart)
  }finally{session.close()}
}
export async function authorParagraphInsertionReview(sourceBytes:Uint8Array,value:DocxParagraphInsertion,storedReason:string) {
  const source=new Uint8Array(sourceBytes),input=copy(value),reason=validateDocxRowReviewReason(storedReason),core=await authorParagraphInsertion(source,input)
  const output=exportDocxComments(core.bytes,[{id:"paragraph-review",text:reason,author:input.author,anchor:anchor(core.bytes,core.paragraph)}])
  require(output.unresolved.length===0&&output.attached.length===1,"reason_unresolved")
  const checked=await inspect(source,output.bytes,input,reason),bytes=restore(core.bytes,output.bytes,checked.commentPart)
  const verified=await verifyParagraphInsertionReview(source,bytes,input,reason)
  await resolveParagraphInsertionReview(source,bytes,input,reason,"accept");await resolveParagraphInsertionReview(source,bytes,input,reason,"reject")
  return {bytes,...verified}
}
