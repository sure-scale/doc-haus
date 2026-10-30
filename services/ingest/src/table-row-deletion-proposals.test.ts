import { afterEach, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { tmpdir } from "node:os"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "../../../dochaus/lib/docxodus"
import { exportDocxComments, readDocxComments } from "../../../dochaus/lib/docx-comments"
import { docxTableRowDeletionTargets } from "../../../dochaus/lib/docx-table-row-deletion"
import { tableRowDeletionProposalSchema, tableRowInsertionProposalSchema, type TableRowDeletionProposalBinding } from "../../../dochaus/lib/table-row-proposal"
import { getRedline, listPendingRedlines, openDb, upsertDocument, type RedlineRow } from "./db"
import { resolveMatterDocumentTarget } from "./document-identity"
import { bake, buildRedlined, RedlineAcceptConflictError, RedlinePreviewConflictError, redlineState } from "./redline"
import { checkedTableRowOperation, recordTableRowProposal, tableRowView } from "./table-row-proposals"

const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main", O="http://schemas.openxmlformats.org/officeDocument/2006/relationships", R="http://schemas.openxmlformats.org/package/2006/relationships", CT="http://schemas.openxmlformats.org/package/2006/content-types"
const BODY="word/document.xml", FILE="working.docx", AUTHOR="Synthetic Reviewer", REASON="Deletes the expressly identified obsolete row."
const CELLS=["  Keep original spaces  ","Literal **words** | <value>","😀"]
const escape=(value:string)=>value.replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
const p=(value:string)=>`<w:p><w:r><w:t xml:space="preserve">${escape(value)}</w:t></w:r></w:p>`
const cell=(value:string)=>`<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(value)}</w:tc>`
const row=(values:string[],header=false)=>`<w:tr>${header?"<w:trPr><w:tblHeader/></w:trPr>":""}${values.map(cell).join("")}</w:tr>`
const parse=(bytes:Uint8Array)=>new DOMParser().parseFromString(strFromU8(bytes),"application/xml")
const hash=(bytes:Uint8Array)=>createHash("sha256").update(bytes).digest("hex")
function source(){
  const bytes=zipSync({
    "[Content_Types].xml":strToU8(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${BODY}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`),
    "_rels/.rels":strToU8(`<Relationships xmlns="${R}"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels":strToU8(`<Relationships xmlns="${R}"><Relationship Id="styles" Type="${O}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${O}/settings" Target="settings.xml"/></Relationships>`),
    "word/styles.xml":strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/settings.xml":strToU8(`<w:settings xmlns:w="${W}"/>`),
    [BODY]:strToU8(`<w:document xmlns:w="${W}" xmlns:r="${O}"><w:body>${p("Before")}<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(3)}</w:tblGrid>${row(["Heading A","Heading B","Heading C"],true)}${row(CELLS)}${row(["Last A","Last B","Last C"])}</w:tbl>${p("After")}<w:sectPr/></w:body></w:document>`),
  },{mtime:new Date(1980,0,1)})
  return exportDocxComments(bytes,[{id:"original",text:"Preserve the existing note.",author:"Prior Reviewer",anchor:{paragraph_text:"After"}}]).bytes
}
function changed(bytes:Uint8Array,part:string,change:(doc:Document)=>void){const parts=unzipSync(bytes),doc=parse(parts[part]!);change(doc);parts[part]=strToU8(new XMLSerializer().serializeToString(doc));return zipSync(parts)}
const directories:string[]=[]
afterEach(()=>directories.splice(0).forEach(dir=>rmSync(dir,{recursive:true,force:true})))
function matter(bytes=source()){
  const directory=realpathSync.native(mkdtempSync(path.join(tmpdir(),"delete-row-proposal-")));directories.push(directory)
  const file=path.join(directory,FILE);writeFileSync(file,bytes);const db=openDb(directory)
  try{upsertDocument(db,file,FILE,Date.now(),null)}finally{db.close()}
  return{directory,file,bytes}
}
async function binding(bytes:Uint8Array):Promise<TableRowDeletionProposalBinding>{const packet=await docxTableRowDeletionTargets(bytes),target=packet.targets.find(row=>row.row===2)!;return{version:3,kind:"table_row_delete",prepared_role:"source",source_sha256:packet.source_sha256,target_handle:target.target_handle,cells:[...target.cells]}}
const review=(file:string)=>({author:AUTHOR,reason:REASON,openItems:[],preparedPath:file})
function pending(directory:string,file:string){const db=openDb(directory);try{return listPendingRedlines(db,file)}finally{db.close()}}
function count(directory:string){const db=openDb(directory);try{return(db.query("SELECT COUNT(*) AS n FROM document_versions").get() as{n:number}).n}finally{db.close()}}

test("strict delete bindings preserve full source spaces and cannot carry insertion/replacement fields",async()=>{
  const op=await binding(source())
  expect(tableRowDeletionProposalSchema.parse(op).cells).toEqual(CELLS)
  for(const extra of [{position:"after"},{reference_cells:CELLS},{replacement:""},{new_cells:[]}])expect(tableRowDeletionProposalSchema.safeParse({...op,...extra}).success).toBe(false)
  expect(tableRowDeletionProposalSchema.safeParse({...op,cells:["  ","",""]}).success).toBe(true)
  expect(tableRowInsertionProposalSchema.safeParse({...op,kind:"table_row_insert",position:"after"}).success).toBe(false)
})

test("record, fresh view, preview and bake share exact frozen source-cell equality before native deletion",async()=>{
  const m=matter(),op=await binding(m.bytes),{id}=await recordTableRowProposal(m.directory,FILE,op,review(m.file)),row=pending(m.directory,m.file)[0]!
  expect(row.id).toBe(id);expect(row.new_text).toBe("")
  expect(await tableRowView(m.bytes,row)).toEqual({kind:"table_row_delete",table:1,row:2,cells:CELLS})
  const state=await redlineState(m.bytes,[row],{},"source"),preview=await buildRedlined(m.bytes,[row],state)
  expect(readDocxComments(preview).map(c=>c.text)).toEqual(["Preserve the existing note.",REASON])
  expect(readDocxComments(preview).find(c=>c.text===REASON)!.markers.every(m=>m.deleted_paragraph_text===CELLS[0])).toBe(true)
  const accepted=await bake(m.bytes,[row],state)
  expect(parse(unzipSync(accepted)[BODY]!).getElementsByTagNameNS(W,"tr").length).toBe(2)
  expect(readDocxComments(accepted).map(c=>c.text)).toEqual(["Preserve the existing note."])
  expect(readFileSync(m.file)).toEqual(Buffer.from(m.bytes));expect(row.status).toBe("pending")
  const n=matter(m.bytes),sdk=await docxodus(),probe=sdk.openDocxSession(m.bytes,{}),prototype=Object.getPrototypeOf(probe) as typeof probe;probe.close()
  const deletion=spyOn(prototype,"deleteTableRow")
  try{
    for(const cells of [["Changed",...CELLS.slice(1)],[CELLS[0]!.trim(),...CELLS.slice(1)]]){
      const bad={...op,cells},forged={...row,target_json:JSON.stringify(bad)}
      await expect(recordTableRowProposal(n.directory,FILE,bad,review(n.file))).rejects.toThrow("complete row contents")
      await expect(tableRowView(m.bytes,forged)).rejects.toThrow("complete row contents")
      await expect(buildRedlined(m.bytes,[forged],state)).rejects.toThrow("complete row contents")
      await expect(bake(m.bytes,[forged],state)).rejects.toThrow("complete row contents")
    }
    expect(deletion).not.toHaveBeenCalled();expect(pending(n.directory,n.file)).toEqual([])
  }finally{deletion.mockRestore()}
})

test("selected fields, hyperlinks, controls, malformed properties and original empty row properties stay ineligible",async()=>{
  const original=source(),sdk=await docxodus(),probe=sdk.openDocxSession(original,{}),prototype=Object.getPrototypeOf(probe) as typeof probe;probe.close()
  const deletion=spyOn(prototype,"deleteTableRow")
  try{
    for(const kind of ["field","hyperlink","control","property","empty_row_properties"]){
      const bytes=changed(original,BODY,doc=>{
        const row=doc.getElementsByTagNameNS(W,"tr")[1]!,paragraph=row.getElementsByTagNameNS(W,"p")[0]!
        if(kind==="empty_row_properties"){row.insertBefore(doc.createElementNS(W,"w:trPr"),row.firstChild);return}
        if(kind==="property"){const props=doc.createElementNS(W,"w:pPr"),spacing=doc.createElementNS(W,"w:spacing");spacing.appendChild(doc.createElementNS(W,"w:t"));props.appendChild(spacing);paragraph.insertBefore(props,paragraph.firstChild);return}
        const template=kind==="field"?`<w:fldSimple xmlns:w="${W}" w:instr=" DATE "><w:r><w:t>Cached value</w:t></w:r></w:fldSimple>`:kind==="hyperlink"?`<w:hyperlink xmlns:w="${W}" w:anchor="MissingTarget"><w:r><w:t>Linked words</w:t></w:r></w:hyperlink>`:`<w:sdt xmlns:w="${W}"><w:sdtPr><w:tag w:val="owned"/></w:sdtPr><w:sdtContent><w:r><w:t>Controlled words</w:t></w:r></w:sdtContent></w:sdt>`
        paragraph.appendChild(doc.importNode(parse(strToU8(template)).documentElement!,true))
      })
      const m=matter(bytes),s=sdk.openDocxSession(bytes,{persistAnchorIds:false})
      let op:TableRowDeletionProposalBinding
      try{const id=Object.entries(s.project().anchorIndex).find(([,a])=>a.kind==="tbl"&&a.partUri==="/word/document.xml")![0],meta=s.getTableMetadata(id).metadata!.rows[1]!,target=parse(unzipSync(bytes)[BODY]!).getElementsByTagNameNS(W,"tr")[1]!,cells=Array.from(target.getElementsByTagNameNS(W,"tc")).map(cell=>Array.from(cell.getElementsByTagNameNS(W,"t")).map(text=>text.textContent??"").join(""));op={version:3,kind:"table_row_delete",prepared_role:"source",source_sha256:hash(bytes),target_handle:createHash("sha256").update(JSON.stringify([hash(bytes),id,meta.anchor.id,meta.cells[0]!.anchor.id])).digest("hex"),cells}}finally{s.close()}
      await expect(recordTableRowProposal(m.directory,FILE,op,review(m.file))).rejects.toThrow("cannot be deleted safely")
      expect(pending(m.directory,m.file)).toEqual([]);expect(readFileSync(m.file)).toEqual(Buffer.from(bytes))
    }
    expect(deletion).not.toHaveBeenCalled()
  }finally{deletion.mockRestore()}
})

test("unexpected deletion proof failures use safe row-change guidance for preview and Accept",async()=>{
  const m=matter(),op=await binding(m.bytes)
  await recordTableRowProposal(m.directory,FILE,op,review(m.file))
  const rows=pending(m.directory,m.file),state=await redlineState(m.bytes,rows,{},"source")
  expect(readDocxComments(await buildRedlined(m.bytes,rows,state)).some(c=>c.text===REASON)).toBe(true)
  expect(parse(unzipSync(await bake(m.bytes,rows,state))[BODY]!).getElementsByTagNameNS(W,"tr").length).toBe(2)
  const sdk=await docxodus(),probe=sdk.openDocxSession(m.bytes,{}),prototype=Object.getPrototypeOf(probe) as typeof probe;probe.close()
  const failure="Private SDK diagnostic: deletion proof failed",deletion=spyOn(prototype,"deleteTableRow").mockImplementation(()=>{throw new Error(failure)})
  try{
    const guidance="This table row change cannot be safely prepared. Reject it and prepare a fresh change list, or make the change directly in Word."
    for(const [run,kind] of [[()=>buildRedlined(m.bytes,rows,state),RedlinePreviewConflictError],[()=>bake(m.bytes,rows,state),RedlineAcceptConflictError]] as const){
      const error=await run().then(()=>undefined,error=>error)
      expect(error).toBeInstanceOf(kind);expect(error.message).toBe(guidance)
      expect(error.message).not.toContain("insertion");expect(error.message).not.toContain(failure)
    }
    expect(deletion).toHaveBeenCalledTimes(2)
    expect(readFileSync(m.file)).toEqual(Buffer.from(m.bytes));expect(pending(m.directory,m.file)[0]!.status).toBe("pending")
  }finally{deletion.mockRestore()}
})

test("real delete list/preview/Accept retain full cells, exact clean bytes and one version; tampered/stale/mixed/role cases write nothing",async()=>{
  const temporary=realpathSync.native(mkdtempSync(path.join(tmpdir(),"delete-row-http-")));directories.push(temporary)
  const workspace=path.join(temporary,"workspace"),config=path.join(temporary,"config");mkdirSync(workspace);mkdirSync(config);writeFileSync(path.join(config,"opencode.json"),"{}")
  const probe=Bun.serve({port:0,fetch:()=>new Response()});const port=probe.port!;probe.stop(true)
  const service=Bun.spawn([process.execPath,"run","src/server.ts"],{cwd:import.meta.dir.replace(/\/src$/,""),env:{...process.env,WORKSPACE_ROOT:workspace,DOCHAUS_DIR:config,INGEST_PORT:String(port),INGEST_HOST:"127.0.0.1",OPENCODE_URL:"http://127.0.0.1:1",OPENCODE_SERVER_PASSWORD:""},stdout:Bun.file(path.join(temporary,"server.log")),stderr:Bun.file(path.join(temporary,"server.log"))})
  const request=(route:string,options?:RequestInit)=>fetch(`http://127.0.0.1:${port}${route}`,{...options,signal:AbortSignal.timeout(30000)})
  try{
    const deadline=Date.now()+20000;while(!await request("/matters").then(r=>r.ok).catch(()=>false)){if(service.exitCode!==null||Date.now()>deadline)throw new Error(await Bun.file(path.join(temporary,"server.log")).text());await Bun.sleep(50)}
    for(const mode of ["single","all","snapshot","stale","mixed","role"] as const){
      const created=await(await request("/matters",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({title:`Synthetic deletion ${mode}`})})).json() as{id:string;dir:string},route=`/matters/${created.id}`,bytes=source(),form=new FormData();form.append("file",new File([bytes],FILE))
      expect((await request(`${route}/documents`,{method:"POST",body:form})).status).toBe(200)
      const file=resolveMatterDocumentTarget(created.dir,FILE).docPath,op=await binding(bytes),{id}=await recordTableRowProposal(created.dir,FILE,op,review(file)),row=pending(created.dir,file)[0]!,expected=await bake(bytes,[row],await redlineState(bytes,[row],{},"source"))
      const preview=await request(`${route}/documents/redlined?name=${FILE}`);expect(preview.status).toBe(200)
      expect(readDocxComments(new Uint8Array(await preview.arrayBuffer())).find(c=>c.text===REASON)!.markers.every(m=>m.deleted_paragraph_text===CELLS[0])).toBe(true)
      expect(readFileSync(file)).toEqual(Buffer.from(bytes));expect(count(created.dir)).toBe(1)
      let textId:number|undefined
      if(mode==="snapshot"){const db=openDb(created.dir);try{db.run("UPDATE redlines SET target_json=? WHERE id=?",[JSON.stringify({...op,cells:["Tampered",...CELLS.slice(1)]}),id])}finally{db.close()}}
      if(mode==="stale")writeFileSync(file,changed(bytes,"word/styles.xml",doc=>doc.documentElement!.setAttribute("changed","yes")))
      if(mode==="role"){const db=openDb(created.dir);try{db.run("UPDATE documents SET role='deliverable'")}finally{db.close()}}
      if(mode==="mixed"){const db=openDb(created.dir);try{textId=Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,created_at) VALUES (?,?,?,?,?,?,?,?)",[file,FILE,"phrase","Before","Before","Earlier",AUTHOR,Date.now()]).lastInsertRowid)}finally{db.close()}}
      const before=readFileSync(file),listed=await(await request(`${route}/redlines?name=${FILE}`)).json() as Array<RedlineRow&{table_row:unknown;stale?:boolean}>
      expect(listed[0]!.table_row).toEqual(mode==="single"||mode==="all"?{kind:"table_row_delete",cells:CELLS,table:1,row:2}:{kind:"table_row_delete",cells:mode==="snapshot"?["Tampered",...CELLS.slice(1)]:CELLS})
      if(mode!=="single"&&mode!=="all")expect(listed[0]!.stale).toBe(true)
      const accepted=await request(mode==="all"?`${route}/redlines/accept-all?name=${FILE}`:`${route}/redlines/${id}/accept`,{method:"POST"})
      expect(accepted.status).toBe(mode==="single"||mode==="all"?200:409)
      expect(readFileSync(file)).toEqual(mode==="single"||mode==="all"?Buffer.from(expected):before);expect(count(created.dir)).toBe(mode==="single"||mode==="all"?2:1)
      const db=openDb(created.dir);try{expect(getRedline(db,id)!.status).toBe(mode==="single"||mode==="all"?"accepted":"pending")}finally{db.close()}
      if(mode==="mixed"){expect((await request(`${route}/redlines/${textId}/accept`,{method:"POST"})).status).toBe(409);expect((await request(`${route}/redlines/accept-all?name=${FILE}`,{method:"POST"})).status).toBe(409);expect((await request(`${route}/documents/redlined?name=${FILE}`)).status).toBe(409)}
      if(mode==="snapshot"){expect((await request(`${route}/documents/redlined?name=${FILE}`)).status).toBe(409);expect((await request(`${route}/redlines/accept-all?name=${FILE}`,{method:"POST"})).status).toBe(409)}
    }
  }finally{service.kill();await service.exited}
},120000)
