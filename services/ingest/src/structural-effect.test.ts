import {expect,test} from "bun:test"
import {Document,Packer,Paragraph,Table,TableCell,TableRow} from "docx"
import {strFromU8,strToU8,unzipSync,zipSync} from "fflate"
import {createDocxRevision} from "../../../dochaus/lib/docx-amendments"
import {renderMarkdownToDocx} from "../../../dochaus/lib/render"
import {paragraphInsertionTargets,resolveParagraphInsertion} from "../../../dochaus/lib/docx-paragraph-insertion"
import {docxTableRowInsertionTargets,resolveDocxTableRowInsertion} from "../../../dochaus/lib/docx-table-row-insertion"
import {docxTableRowDeletionTargets,resolveDocxTableRowDeletion} from "../../../dochaus/lib/docx-table-row-deletion"
import {inspectExactParagraphSpacing,resolveExactParagraphSpacing} from "../../../dochaus/lib/docx-exact-spacing"
import {structuralEffect} from "./structural-effect"

test("structural effects describe accepted paragraph and direct spacing, with rejection restoring source",async()=>{
 const source=(await renderMarkdownToDocx("# Terms\n\nReference paragraph.\n\nOther obligation.")).bytes,original=new Uint8Array(source)
 const packet=await paragraphInsertionTargets(source),target=packet.targets.find(row=>row.text==="Reference paragraph.")!
 const request={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"after" as const,text:"A literal *new* obligation.",author:"Review Counsel"}
 const projected=await structuralEffect(source,{kind:"paragraph_insert",request})
 expect(projected.effect.after).toBe(request.text)
 expect(projected.effect).toMatchObject({position:"after",location:{target_handle:target.target_handle,body_paragraph_number:target.body_paragraph_number}})
 await expect(structuralEffect(source,{kind:"paragraph_insert",request:{...request,target_handle:"unknown"}})).rejects.toThrow("structural_effect_target_changed")
 await expect(structuralEffect(source,{kind:"paragraph_insert",request:{...request,source_sha256:"0".repeat(64)}})).rejects.toThrow("structural_effect_target_changed")
 const accepted=await resolveParagraphInsertion(source,projected.tracked,request,"accept")
 const paragraphs=(await createDocxRevision(accepted)).paragraphs.map(row=>row.text)
 expect(paragraphs.indexOf(request.text)).toBe(paragraphs.indexOf(target.text)+1)
 expect(await resolveParagraphInsertion(source,projected.tracked,request,"reject")).toEqual(source)
 const spacingSource=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes,spacingBefore=new Uint8Array(spacingSource)
 const spacing=await inspectExactParagraphSpacing(spacingSource,"Same"),match=spacing.matches[0]!
 const format={sourceSha256:spacing.sourceSha256,paragraphHandle:match.paragraphHandle,paragraphText:match.text,tableContext:match.tableContext,spacingAfterPoints:6,author:"Review Counsel"}
 const formatEffect=await structuralEffect(spacingSource,{kind:"spacing_after",request:format})
 expect(formatEffect.effect).toMatchObject({before:match.directAfter.kind==="twips"?match.directAfter.raw:null,after:"120",unit:"twips"})
 const formatted=await resolveExactParagraphSpacing(spacingSource,formatEffect.tracked,format,"accept")
 expect((await inspectExactParagraphSpacing(formatted,"Same")).matches[0]!.directAfter).toMatchObject({kind:"twips",normalized:"120"})
 expect(await resolveExactParagraphSpacing(spacingSource,formatEffect.tracked,format,"reject")).toEqual(spacingSource)
 expect(spacingSource).toEqual(spacingBefore)
 expect(source).toEqual(original)
})

test("structural effects describe accepted row insertion and deletion",async()=>{
 const cells=(values:string[])=>new TableRow({children:values.map(text=>new TableCell({children:[new Paragraph(text)]}))})
 const source=new Uint8Array(await Packer.toBuffer(new Document({sections:[{children:[new Paragraph("Schedule"),new Table({rows:[cells(["Item","Fee"]),cells(["Service","Ten pounds"]),cells(["Support","Five pounds"]) ]})]}]}))),original=new Uint8Array(source)
 const targets=await docxTableRowInsertionTargets(source),target=targets.targets.find(row=>row.cells[0]==="Service")!
 expect(target).toBeTruthy()
 await expect(structuralEffect(source,{kind:"row_insert",request:{source_sha256:targets.source_sha256,target_handle:"unknown",position:"after",cells:["A","B"],author:"Review Counsel"}})).rejects.toThrow("structural_effect_target_changed")
 const insert={source_sha256:targets.source_sha256,target_handle:target.target_handle,position:"after" as const,cells:["Archive","Two pounds"],author:"Review Counsel"}
 const projected=await structuralEffect(source,{kind:"row_insert",request:insert})
 expect(projected.effect.after).toEqual(insert.cells)
 const accepted=await resolveDocxTableRowInsertion(source,projected.tracked,insert,"accept")
 expect((await docxTableRowInsertionTargets(accepted)).targets.map(row=>row.cells)).toContainEqual(insert.cells)
 expect(await resolveDocxTableRowInsertion(source,projected.tracked,insert,"reject")).toEqual(source)
 const deletions=await docxTableRowDeletionTargets(source),deleted=deletions.targets.find(row=>row.cells[0]==="Service")!
 const remove={source_sha256:deletions.source_sha256,target_handle:deleted.target_handle,author:"Review Counsel"}
 await expect(structuralEffect(source,{kind:"row_delete",request:{...remove,source_sha256:"0".repeat(64)}})).rejects.toThrow("structural_effect_target_changed")
 const deletion=await structuralEffect(source,{kind:"row_delete",request:remove})
 expect(deletion.effect.before).toEqual(deleted.cells)
 const reduced=await resolveDocxTableRowDeletion(source,deletion.tracked,remove,"accept")
 expect((await docxTableRowInsertionTargets(reduced)).targets.map(row=>row.cells)).not.toContainEqual(deleted.cells)
 expect(await resolveDocxTableRowDeletion(source,deletion.tracked,remove,"reject")).toEqual(source)
 expect(source).toEqual(original)
})


test("duplicate source locations remain distinct and direct spacing distinguishes absent from zero",async()=>{
 const source=(await renderMarkdownToDocx("Before.\n\nSame.\n\nAfter.\n\nBefore.\n\nSame.\n\nAfter.")).bytes
 const packet=await paragraphInsertionTargets(source),matches=packet.targets.filter(row=>row.text==="Same.")
 expect(matches).toHaveLength(2)
 for(const target of matches){
  const request={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"before" as const,text:"New term.",author:"Review Counsel"}
  const projected=await structuralEffect(source,{kind:"paragraph_insert",request})
  expect(projected.effect).toMatchObject({position:"before",location:{target_handle:target.target_handle,body_paragraph_number:target.body_paragraph_number}})
 }
 const raw=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes
 for(const [attribute,expected] of [["",null],[" w:after=\"0\"","0"]] as const){
  const parts=unzipSync(raw),doc=strFromU8(parts["word/document.xml"]!)
  parts["word/document.xml"]=strToU8(doc.replaceAll(' w:after="80"',attribute))
  const bytes=zipSync(parts),packet=await inspectExactParagraphSpacing(bytes,"Same"),target=packet.matches[0]!
  const request={sourceSha256:packet.sourceSha256,paragraphHandle:target.paragraphHandle,paragraphText:target.text,tableContext:target.tableContext,spacingAfterPoints:6,author:"Review Counsel"}
  const projection=await structuralEffect(bytes,{kind:"spacing_after",request})
  expect(projection.effect.kind).toBe("spacing_after")
  if (!("directAfter" in projection.effect.location)) throw new Error("Expected a direct-spacing target")
  expect(projection.effect.before).toBe(expected)
  expect(projection.effect.location.directAfter.kind).toBe(expected===null?"absent":"twips")
  await expect(structuralEffect(bytes,{kind:"spacing_after",request:{...request,paragraphHandle:"unknown"}})).rejects.toThrow("structural_effect_target_changed")
 }
})
