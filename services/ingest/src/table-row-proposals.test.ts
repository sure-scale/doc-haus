import { afterEach, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, renameSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxTableRowInsertionTargets } from "../../../dochaus/lib/docx-table-row-insertion"
import { docxodus as sharedDocxodus } from "../../../dochaus/lib/docxodus"
import { readDocxComments } from "../../../dochaus/lib/docx-comments"
import { recordRedline, recordRedlineProposal } from "../../../dochaus/lib/redlines"
import { tableRowProposalSchema, type TableRowInsertionProposalBinding } from "../../../dochaus/lib/table-row-proposal"
import { getPreparedChanges, getRedline, listPendingRedlines, openDb, upsertDocument, type RedlineRow } from "./db"
import { bake, buildRedlined, redlineStaleReason, redlineState } from "./redline"
import { docxodus } from "./docxodus"
import { recordTableRowProposal, tableRowView } from "./table-row-proposals"
import { resolveMatterDocumentTarget } from "./document-identity"

const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main", R="http://schemas.openxmlformats.org/package/2006/relationships", O="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const BODY="word/document.xml", TYPES="[Content_Types].xml", FILE="working.docx", AUTHOR="Synthetic Reviewer", REASON="Records the expressly requested row."
const hash=(bytes:Uint8Array)=>createHash("sha256").update(bytes).digest("hex")
const text=(value:string)=>value.replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
const p=(value:string)=>`<w:p><w:r><w:t>${text(value)}</w:t></w:r></w:p>`
const parse=(bytes:Uint8Array)=>new DOMParser().parseFromString(strFromU8(bytes),"application/xml")
function source() {
  const table=`<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(3)}</w:tblGrid>${[1,2,3].map(row=>`<w:tr><w:trPr/>${[1,2,3].map(col=>`<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(`Row ${row} cell ${col}`)}</w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`
  return zipSync({
    [TYPES]:strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${BODY}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`),
    "_rels/.rels":strToU8(`<Relationships xmlns="${R}"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels":strToU8(`<Relationships xmlns="${R}"><Relationship Id="styles" Type="${O}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${O}/settings" Target="settings.xml"/></Relationships>`),
    "word/styles.xml":strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/settings.xml":strToU8(`<w:settings xmlns:w="${W}"/>`),
    [BODY]:strToU8(`<w:document xmlns:w="${W}"><w:body>${p("Before")}${table}${p("After")}<w:sectPr/></w:body></w:document>`),
  },{mtime:new Date(1980,0,1)})
}
function modified(bytes:Uint8Array,part:string,change:(doc:Document)=>void) {
  const parts=unzipSync(bytes), doc=parse(parts[part]!);change(doc)
  parts[part]=strToU8(new XMLSerializer().serializeToString(doc));return zipSync(parts)
}
const rowTexts=(bytes:Uint8Array)=>Array.from(parse(unzipSync(bytes)[BODY]!).getElementsByTagNameNS(W,"tr")).map(row=>Array.from(row.getElementsByTagNameNS(W,"tc")).map(cell=>Array.from(cell.getElementsByTagNameNS(W,"t")).map(t=>t.textContent).join("")))
const directories:string[]=[]
afterEach(()=>directories.splice(0).forEach(dir=>rmSync(dir,{recursive:true,force:true})))
function matter(bytes=source()) {
  const directory=realpathSync.native(mkdtempSync(path.join(tmpdir(),"row-proposal-")));directories.push(directory)
  const file=path.join(directory,FILE);writeFileSync(file,bytes)
  const db=openDb(directory);try {upsertDocument(db,file,FILE,Date.now(),null)} finally {db.close()}
  return {directory,file,bytes}
}
async function operation(bytes:Uint8Array,cells=["Literal **bold** | <value>","","Emoji 😀"]):Promise<TableRowInsertionProposalBinding> {
  const offered=await docxTableRowInsertionTargets(bytes)
  return {version:3,kind:"table_row_insert",prepared_role:"source",source_sha256:offered.source_sha256,target_handle:offered.targets[1]!.target_handle,position:"after",cells}
}
const review={author:AUTHOR,reason:REASON,openItems:["Confirm the stated values."]}
function pending(directory:string,file:string) {const db=openDb(directory);try{return listPendingRedlines(db,file)}finally{db.close()}}
const legacy=(file:string)=>({docPath:file,docName:FILE,scope:"phrase" as const,findText:"Before",oldText:"Before",newText:"Earlier",author:AUTHOR,reviewComment:"A text change.",anchorId:"same"})

test("recorded row uses structured cells and native-comment preview; canonical Accept result remains clean and exact",async()=>{
  const m=matter(), binding=await operation(m.bytes), {id}=await recordTableRowProposal(m.directory,FILE,binding,{...review,preparedPath:m.file}), rows=pending(m.directory,m.file)
  expect(rows).toHaveLength(1);expect(rows[0]!.id).toBe(id)
  expect(JSON.parse(rows[0]!.target_json!)).toEqual(binding)
  expect(await tableRowView(m.bytes,rows[0]!)).toEqual({kind:"table_row_insert",position:"after",table:1,row:2,reference_cells:["Row 2 cell 1","Row 2 cell 2","Row 2 cell 3"],cells:binding.cells})
  const state=await redlineState(m.bytes,rows,{},"source"), preview=await buildRedlined(m.bytes,rows,state), baked=await bake(m.bytes,rows,state)
  expect(readDocxComments(preview).map(row=>row.text)).toEqual([REASON])
  expect(rowTexts(preview)).toEqual([...rowTexts(m.bytes).slice(0,2),binding.cells,...rowTexts(m.bytes).slice(2)])
  expect(rowTexts(baked)).toEqual(rowTexts(preview));expect(readDocxComments(baked)).toEqual([])
  expect((await (await docxodus()).getRevisions(baked))).toEqual([])
  expect(readFileSync(m.file)).toEqual(Buffer.from(m.bytes));expect(pending(m.directory,m.file)[0]!.status).toBe("pending")
})

test("source-only and whole-file binding refuse formatting-only staleness, changed roles, and native property revisions",async()=>{
  const m=matter(), binding=await operation(m.bytes), {id}=await recordTableRowProposal(m.directory,FILE,binding,{...review,preparedPath:m.file}), row=pending(m.directory,m.file)[0]!
  const changed=modified(m.bytes,"word/styles.xml",doc=>doc.getElementsByTagNameNS(W,"name")[0]!.setAttributeNS(W,"w:val","Different"))
  const stale=await redlineState(changed,[row],{},"source")
  expect(redlineStaleReason(row,stale,false)).toContain("earlier Word file")
  await expect(bake(changed,[row],stale)).rejects.toThrow("earlier Word file")
  await expect(buildRedlined(m.bytes,[row],await redlineState(m.bytes,[row],{},"deliverable"))).rejects.toThrow("document role")
  expect(pending(m.directory,m.file)[0]!.id).toBe(id)
  const tracked=modified(source(),BODY,doc=>{
    const first=doc.getElementsByTagNameNS(W,"p")[0]!, props=doc.createElementNS(W,"w:pPr"), change=doc.createElementNS(W,"w:pPrChange")
    change.setAttributeNS(W,"w:id","19");change.setAttributeNS(W,"w:author","Earlier Reviewer");change.setAttributeNS(W,"w:date","2026-01-01T00:00:00Z");change.appendChild(doc.createElementNS(W,"w:pPr"));props.appendChild(change);first.insertBefore(props,first.firstChild)
  })
  const native=matter(tracked)
  await expect(recordTableRowProposal(native.directory,FILE,await operation(tracked),{...review,preparedPath:native.file})).rejects.toThrow("without tracked changes")
  expect(pending(native.directory,native.file)).toEqual([])
})

test("shared recorders refuse row coexistence before superseding; malformed kind/scope cannot become a text proposal",async()=>{
  const m=matter(), binding=await operation(m.bytes)
  await recordTableRowProposal(m.directory,FILE,binding,{...review,preparedPath:m.file})
  for(const record of [recordRedline,recordRedlineProposal]) expect(()=>record(m.directory,legacy(m.file))).toThrow("reviewed on its own")
  expect(pending(m.directory,m.file).map(row=>row.status)).toEqual(["pending"])
  const row=pending(m.directory,m.file)[0]!, disguised={...row,scope:"phrase" as const}
  await expect(buildRedlined(m.bytes,[disguised],await redlineState(m.bytes,[disguised],{},"source"))).rejects.toThrow("invalid stored operation")
  await expect(bake(m.bytes,[row,{...disguised,id:row.id+1}],await redlineState(m.bytes,[row],{},"source"))).rejects.toThrow("on its own")
  const n=matter();recordRedline(n.directory,legacy(n.file))
  await expect(recordTableRowProposal(n.directory,FILE,await operation(n.bytes),{...review,preparedPath:n.file})).rejects.toThrow("reviewed on its own")
  expect(pending(n.directory,n.file)).toHaveLength(1)
  expect(tableRowProposalSchema.safeParse({...binding,cells:[" unsafe ","","ok"]}).success).toBe(false)
})

test("final persistence check refuses a source/path/role or pending-proposal race after verified authoring",async()=>{
  const sdk=await sharedDocxodus(), probe=sdk.openDocxSession(source(),{}), prototype=Object.getPrototypeOf(probe) as typeof probe;probe.close()
  const insert=prototype.insertTableRow
  for(const race of ["source","path","role","pending"] as const){
    const m=matter(), binding=await operation(m.bytes);let once=false
    const spy=spyOn(prototype,"insertTableRow").mockImplementation(function(this:typeof probe,...args:Parameters<typeof insert>){
      const result=insert.apply(this,args)
      if(!once){once=true;if(race==="source")writeFileSync(m.file,modified(m.bytes,"word/styles.xml",doc=>doc.documentElement!.setAttribute("raced","yes")))
        else if(race==="path")renameSync(m.file,path.join(m.directory,"Working.docx"))
        else if(race==="role"){const db=openDb(m.directory);try{db.run("UPDATE documents SET role='deliverable'")}finally{db.close()}}
        else recordRedline(m.directory,legacy(m.file))}
      return result
    })
    try{await expect(recordTableRowProposal(m.directory,FILE,binding,{...review,preparedPath:m.file})).rejects.toThrow(race==="pending"?"reviewed on its own":"earlier Word file")}
    finally{spy.mockRestore()}
    expect(once).toBe(true)
    expect(pending(m.directory,m.file).filter(row=>row.scope==="table_row")).toEqual([])
    const db=openDb(m.directory);try{expect(getPreparedChanges(db,m.file)).toBeNull()}finally{db.close()}
  }
})

test("real list, preview, singleton Accept/all and stale/mixed refusal preserve versions and exact source bytes",async()=>{
  const temporary=mkdtempSync(path.join(tmpdir(),"row-http-"));directories.push(temporary)
  const workspace=path.join(temporary,"workspace"), config=path.join(temporary,"config");mkdirSync(workspace);mkdirSync(config);writeFileSync(path.join(config,"opencode.json"),"{}")
  const probe=Bun.serve({port:0,fetch:()=>new Response()});const port=probe.port!;probe.stop(true)
  const service=Bun.spawn([process.execPath,"run","src/server.ts"],{cwd:import.meta.dir.replace(/\/src$/,""),env:{...process.env,WORKSPACE_ROOT:workspace,DOCHAUS_DIR:config,INGEST_PORT:String(port),INGEST_HOST:"127.0.0.1",OPENCODE_URL:"http://127.0.0.1:1",OPENCODE_SERVER_PASSWORD:""},stdout:Bun.file(path.join(temporary,"server.log")),stderr:Bun.file(path.join(temporary,"server.log"))})
  const request=(route:string,options?:RequestInit)=>fetch(`http://127.0.0.1:${port}${route}`,{...options,signal:AbortSignal.timeout(30000)})
  try{
    const deadline=Date.now()+20000
    while(!await request("/matters").then(r=>r.ok).catch(()=>false)){if(service.exitCode!==null||Date.now()>deadline)throw new Error(await Bun.file(path.join(temporary,"server.log")).text());await Bun.sleep(50)}
    for(const mode of ["single","all","mixed","stale","role"] as const){
      const created=await(await request("/matters",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({title:`Synthetic row ${mode}`})})).json() as {id:string;dir:string}
      const route=`/matters/${created.id}`, bytes=source(), form=new FormData();form.append("file",new File([bytes],FILE))
      const upload=await request(`${route}/documents`,{method:"POST",body:form});expect(upload.status).toBe(200)
      const file=resolveMatterDocumentTarget(created.dir,FILE).docPath, binding=await operation(bytes), {id}=await recordTableRowProposal(created.dir,FILE,binding,{...review,preparedPath:file})
      const rows=pending(created.dir,file), expected=await bake(bytes,rows,await redlineState(bytes,rows,{},"source"))
      const versionCount=()=>{const db=openDb(created.dir);try{return (db.query("SELECT COUNT(*) AS n FROM document_versions").get() as {n:number}).n}finally{db.close()}}
      expect(versionCount()).toBe(1)
      const preview=await request(`${route}/documents/redlined?name=${FILE}`);expect(preview.status).toBe(200);expect(readDocxComments(new Uint8Array(await preview.arrayBuffer())).map(c=>c.text)).toEqual([REASON]);expect(versionCount()).toBe(1)
      if(mode==="mixed"){const db=openDb(created.dir);try{db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,created_at) VALUES (?,?,?,?,?,?,?,?)",[file,FILE,"phrase","Before","Before","Earlier",AUTHOR,Date.now()])}finally{db.close()}}
      if(mode==="stale")writeFileSync(file,modified(bytes,"word/styles.xml",doc=>doc.documentElement!.setAttribute("raced","true")))
      if(mode==="role"){const db=openDb(created.dir);try{db.run("UPDATE documents SET role='deliverable'")}finally{db.close()}}
      const before=readFileSync(file), listed=await(await request(`${route}/redlines?name=${FILE}`)).json() as Array<RedlineRow & {table_row?:unknown;stale?:boolean}>
      if(mode==="single"||mode==="all")expect(listed[0]!.table_row).toEqual({kind:"table_row_insert",position:"after",table:1,row:2,reference_cells:["Row 2 cell 1","Row 2 cell 2","Row 2 cell 3"],cells:binding.cells})
      else{expect(listed[0]!.stale).toBe(true);expect(listed[0]!.table_row).toEqual({kind:"table_row_insert",position:"after",cells:binding.cells})}
      const accept=await request(mode==="all"?`${route}/redlines/accept-all?name=${FILE}`:`${route}/redlines/${id}/accept`,{method:"POST"})
      expect(accept.status).toBe(mode==="single"||mode==="all"?200:409)
      expect(readFileSync(file)).toEqual(mode==="single"||mode==="all"?Buffer.from(expected):before)
      expect(versionCount()).toBe(mode==="single"||mode==="all"?2:1)
      const db=openDb(created.dir);try{expect(getRedline(db,id)!.status).toBe(mode==="single"||mode==="all"?"accepted":"pending")}finally{db.close()}
      if(mode==="mixed"){expect((await request(`${route}/documents/redlined?name=${FILE}`)).status).toBe(409);expect((await request(`${route}/redlines/accept-all?name=${FILE}`,{method:"POST"})).status).toBe(409)}
    }
  }finally{service.kill();await service.exited}
},120000)
