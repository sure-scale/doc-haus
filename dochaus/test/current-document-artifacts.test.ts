import { expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { ensureParagraphPointers } from "../lib/drafting-amendment"
import { currentDocumentArtifacts } from "../lib/current-document-artifacts"
import { currentDocumentView } from "../lib/current-document-retrieval"
import { exportDocxComments, readDocxCommentParagraphs } from "../lib/docx-comments"
import { renderMarkdownToDocx } from "../lib/render"
import { buildChecklistDocumentContext } from "../workflows/lean/checklist-document-context"
import { writeLeanPlan } from "../workflows/lean/plan"
import { LEAN_PROMPT_VERSION, stageInputs } from "../workflows/lean/stages"
import { hashStageInputs, leanFile, LEAN_STATE_SCHEMA, snapshotStage, writeLeanBytes, writeLeanText } from "../workflows/lean/state"
test("new draft invalidates old applied body and annotations in context and targeted retrieval",async()=>{
 const directory=mkdtempSync(path.join(tmpdir(),"artifact-epoch-"))
 try{
  writeLeanPlan(directory,[{title:"Record",purpose:"Requested record",file:"record.docx",parts:[{heading:"Record",scope:"Record"}]}])
  let original=(await renderMarkdownToDocx("Initial violet receipt.")).bytes,applied=(await renderMarkdownToDocx("Applied violet receipt.")).bytes
  writeLeanBytes(directory,"d1/body.docx",original);await ensureParagraphPointers(leanFile(directory,"d1/body.docx"));original=readFileSync(leanFile(directory,"d1/body.docx"));writeLeanBytes(directory,"d1/body-applied.docx",applied)
  const anchor=readDocxCommentParagraphs(applied)[0]!,annotated=exportDocxComments(applied,[{id:"history",author:"Reviewer",text:"Old applied explanation.",anchor:{paragraph_text:anchor.paragraph_text,paragraph_index:anchor.paragraph_index}}]).bytes
  writeLeanBytes(directory,"d1/redline.docx",annotated)
  snapshotStage(directory,1,"d1.draft",["d1/body.docx"])
  snapshotStage(directory,2,"d1.apply",["d1/body-applied.docx","d1/redline.docx"])
  const key=hashStageInputs("d1.apply",LEAN_PROMPT_VERSION,stageInputs(directory,"d1.apply"))
  writeLeanText(directory,"state.json",JSON.stringify({schema:LEAN_STATE_SCHEMA,profile:"lean",stages:{"d1.apply":{status:"done",input_sha:key}}}))
  expect(currentDocumentArtifacts(directory,"d1")).toMatchObject({body:"d1/body-applied.docx",redline:"d1/redline.docx",applied:true})
  expect(currentDocumentView(directory,"d1","comments")).toMatchObject({status:"available",entries:[expect.objectContaining({text:"Old applied explanation."})]})
  // A pending cumulative repair changes apply inputs, not the accepted body.
  writeLeanText(directory,"d1/changes.json",JSON.stringify({amendments:[],questions:[],pending:"new correction"}))
  expect(currentDocumentArtifacts(directory,"d1")).toMatchObject({body:"d1/body-applied.docx",redline:"d1/redline.docx",applied:true})
  expect(currentDocumentView(directory,"d1","current_body")).toMatchObject({status:"available",entries:[expect.objectContaining({text:"Applied violet receipt."})]})
  writeLeanBytes(directory,"d1/body.docx",(await renderMarkdownToDocx("Initial violet receipt.")).bytes);await ensureParagraphPointers(leanFile(directory,"d1/body.docx"));const next=readFileSync(leanFile(directory,"d1/body.docx"))
  expect(readDocxCommentParagraphs(next).map(row=>row.pids)).not.toEqual(readDocxCommentParagraphs(original).map(row=>row.pids))
  const selected=currentDocumentArtifacts(directory,"d1");expect(selected.body).toBe("d1/body.docx");expect(selected.applied).toBe(false);expect(selected.redline).toBeUndefined()
  expect(currentDocumentView(directory,"d1","current_body")).toMatchObject({status:"available",entries:[expect.objectContaining({text:"Initial violet receipt.",artifact:".dochaus/lean/d1/body.docx"})]})
  expect(currentDocumentView(directory,"d1","comments")).toMatchObject({status:"available",entries:[]})
  const context=buildChecklistDocumentContext(directory,{item_id:"one",item:{check:"Verify violet receipt.",basis:"directive",source_refs:[{document:"request",locator:"request",quote:"Verify violet receipt."}],qualifications:[],target:{document:"d1",role:"Record",treatment:"operative"}}})
  expect(context.artifacts[0]!.value.paragraphs).toEqual(readDocxCommentParagraphs(next));expect(context.artifacts[0]!.value.annotations!.some(row=>row.artifact.endsWith("redline.docx"))).toBe(false)
  expect(context.warnings.join(" ")).toContain("not a current accepted apply result")
  expect(readFileSync(leanFile(directory,"d1/body-applied.docx"))).toEqual(Buffer.from(applied));expect(readFileSync(leanFile(directory,"d1/redline.docx"))).toEqual(Buffer.from(annotated))
 }finally{rmSync(directory,{recursive:true,force:true})}
})
