import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtemp, mkdir, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { openDb, getRedline, listSourceDocuments } from "./db"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"
import { createDocxRevision } from "../../../dochaus/lib/docx-amendments"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import wordIntegration from "../../../dochaus/tool/word-integration"
import { acceptedWorkingDocxUpdate } from "./document-amendments"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { buildRedlined, bake } from "./redline"
import { docxodus } from "./docxodus"
import { readDocxComments } from "../../../dochaus/lib/docx-comments"

test("two-pass HTTP proposals require approval, preserve atomic revision metadata and refuse stale or unbound writes", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(),"dochaus-word-edit-"))
  const workspace = path.join(temporary,"workspace")
  const config = path.join(temporary,"config")
  await mkdir(workspace);await mkdir(config)
  await mkdir(path.join(workspace,".preferences"))
  await Bun.write(path.join(workspace,".preferences/preferences.json"),JSON.stringify({attorney:"Saved Reviewing Lawyer",firm:"Saved Reviewing Firm"}))
  await Bun.write(path.join(config,"opencode.json"),"{}")
  const generationRequests:Array<{path:string;directory:string|null;prompt:string}> = []
  // Protocol fixture only. The live provider proof is a separate product run.
  const engine = Bun.serve({port:0,fetch:async (request) => {
    const body = await request.json() as {prompt:string}
    generationRequests.push({path:new URL(request.url).pathname,directory:request.headers.get("x-opencode-directory"),prompt:body.prompt})
    const text = body.prompt.includes("<draft_view") ? JSON.stringify({amendments:[
      {before_quote:"five days",replacement:"seven days",target_context:"current operative notice",reason:"Requested notice update."},
      {before_quote:"three months",replacement:"six months",target_context:"retention period",reason:"Requested retention update."},
    ],open_items:[]}) : (() => {
      const candidates = JSON.parse(body.prompt.match(/<candidates>([\s\S]*)<\/candidates>/)![1]) as Array<{amendment_id:string;target_handle:string}>
      return JSON.stringify({selections:["amendment-1","amendment-2"].map((id) => ({amendment_id:id,target_handle:candidates.find((value) => value.amendment_id === id)!.target_handle}))})
    })()
    return Response.json({data:{text,model:{providerID:"fixture",id:"isolated"},elapsed_ms:1,usage:null}})
  }})
  const probe = Bun.serve({port:0,fetch:() => new Response()});const port=probe.port;probe.stop(true)
  const process = Bun.spawn([globalThis.process.execPath,"run","src/server.ts"],{cwd:path.resolve(import.meta.dir,".."),
    env:{...globalThis.process.env,WORKSPACE_ROOT:workspace,DOCHAUS_DIR:config,INGEST_PORT:String(port),INGEST_HOST:"127.0.0.1",OPENCODE_URL:`http://127.0.0.1:${engine.port}`,OPENCODE_SERVER_PASSWORD:""},
    stdout:Bun.file(path.join(temporary,"ingest.log")),stderr:Bun.file(path.join(temporary,"ingest.log"))})
  const request=(route:string,options?:RequestInit) => fetch(`http://127.0.0.1:${port}${route}`,{...options,signal:AbortSignal.timeout(30000)})
  const hash=(bytes:Uint8Array) => createHash("sha256").update(bytes).digest("hex")
  const form=(bytes:Uint8Array) => {const body=new FormData();body.append("file",new File([new Uint8Array(bytes)],"working.docx"));return body}
  try {
    const deadline=Date.now()+20000
    while (!await request("/matters").then((response) => response.ok).catch(() => false)) {
      if (Date.now()>deadline) throw new Error("Isolated ingest did not start")
      await Bun.sleep(50)
    }
    const matter = await (await request("/matters",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({title:"Independent Word review"})})).json() as {id:string;dir:string}
    const route=`/matters/${matter.id}`
    const original = Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({sections:[{children:[
      new Paragraph("Current notice must arrive within five days."),new Paragraph("Retain records for three months."),new Paragraph("Historical note: notice once arrived within five days."),
    ]}]}))))
    expect((await request(`${route}/deliverables`,{method:"POST",body:form(original)})).status).toBe(200)
    const prepare=await request(`${route}/documents/amendments`,{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({name:"working.docx",instructions:"Update current notice and retention; leave the historical note unchanged."})})
    expect(prepare.status).toBe(200)
    const proposed=await prepare.json() as {proposed:number;redline_ids:number[]}
    expect(proposed.proposed).toBe(2)
    expect(generationRequests.map((row) => row.path)).toEqual(["/api/dochaus/generate","/api/dochaus/generate"])
    expect(generationRequests.every((row) => row.directory === matter.dir)).toBe(true)
    expect(generationRequests[1].prompt).not.toContain("<draft_view")
    const file=path.join(matter.dir,"working.docx")
    expect(hash(await Bun.file(file).bytes())).toBe(hash(original))
    expect(readWorkingDocx(matter.dir,file)).toBeUndefined()
    const db=openDb(matter.dir)
    expect(proposed.redline_ids.map((id) => getRedline(db,id)?.author)).toEqual(["Saved Reviewing Lawyer","Saved Reviewing Lawyer"])
    db.run("CREATE TRIGGER refuse_projection BEFORE INSERT ON meta WHEN NEW.key LIKE 'working-docx:%' BEGIN SELECT RAISE(ABORT,'isolated registry failure'); END")
    expect((await request(`${route}/redlines/${proposed.redline_ids[0]}/accept`,{method:"POST"})).status).toBe(409)
    expect(hash(await Bun.file(file).bytes())).toBe(hash(original))
    expect(getRedline(db,proposed.redline_ids[0])?.status).toBe("pending")
    expect(readWorkingDocx(matter.dir,file)).toBeUndefined()
    db.run("DROP TRIGGER refuse_projection")
    expect((await request(`${route}/redlines/${proposed.redline_ids[0]}/accept`,{method:"POST"})).status).toBe(200)
    const accepted=await Bun.file(file).bytes()
    const revision=await createDocxRevision(accepted)
    expect(revision.markdown).toContain("Current notice must arrive within seven days.")
    expect(revision.markdown).toContain("Historical note: notice once arrived within five days.")
    expect(readWorkingDocx(matter.dir,file)?.source_sha256).toBe(hash(accepted))
    expect(readWorkingDocx(matter.dir,file)?.projection_sha256).toBe(revision.projection_sha256)
    expect(getRedline(db,proposed.redline_ids[0])?.status).toBe("accepted")
    expect(listSourceDocuments(db)).toEqual([])
    const pending=await (await request(`${route}/redlines?name=working.docx`)).json() as Array<{stale:boolean}>
    expect(pending).toHaveLength(1);expect(pending[0].stale).toBe(true)
    expect((await request(`${route}/redlines/${proposed.redline_ids[1]}/accept`,{method:"POST"})).status).toBe(409)
    expect(hash(await Bun.file(file).bytes())).toBe(hash(accepted))
    expect((await request(`${route}/deliverables`,{method:"POST",body:form(original)})).status).toBe(409)
    expect((await request(`${route}/documents/revisions`,{method:"POST",body:form(original)})).status).toBe(409)
    expect((await request(`${route}/documents/revisions`,{method:"POST",headers:{"x-document-prior-sha256":hash(original)},body:form(original)})).status).toBe(409)
    expect(hash(await Bun.file(file).bytes())).toBe(hash(accepted))
    expect((await request(`${route}/documents/revisions`,{method:"POST",headers:{"x-document-prior-sha256":hash(accepted)},body:form(original)})).status).toBe(200)
    expect(readWorkingDocx(matter.dir,file)?.source_sha256).toBe(hash(original))
    // Malformed new targeting cannot activate legacy text matching.
    db.run("UPDATE redlines SET target_json = ? WHERE id = ?",["{invalid",proposed.redline_ids[1]])
    expect((await request(`${route}/redlines/${proposed.redline_ids[1]}/accept`,{method:"POST"})).status).toBe(409)
    expect(hash(await Bun.file(file).bytes())).toBe(hash(original))
    db.close()
  } finally {
    process.kill();await process.exited;engine.stop(true);await rm(temporary,{recursive:true,force:true})
  }
},120000)

test("queued legacy writers cannot overwrite a newly accepted working Word revision",async () => {
  const directory = await mkdtemp(path.join(tmpdir(),"dochaus-word-adoption-"))
  const bytes=Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({sections:[{children:[new Paragraph("Original wording remains.")]}]}))))
  const file=path.join(directory,"working.docx")
  try {
    await writeDeliverable(directory,"working.docx",bytes)
    const update=await acceptedWorkingDocxUpdate(directory,file,bytes)
    let acquired!:() => void;let release!:() => void;let asked!:() => void
    const ready=new Promise<void>((resolve) => {acquired=resolve})
    const barrier=new Promise<void>((resolve) => {release=resolve})
    const permission=new Promise<void>((resolve) => {asked=resolve})
    const owner=withArtifactMutationLocksAsync(directory,[file],async () => {
      acquired();await barrier
      const db=openDb(directory);try {db.transaction(() => update(db)).immediate()} finally {db.close()}
    })
    await ready
    const revision=writeDeliverableRevision(directory,"working.docx",Buffer.from(bytes))
    const legacy=wordIntegration.execute({document:"working.docx",action:"replace",find:"Original",replace:"Unexpected"},
      {directory,worktree:directory,agent:"test",sessionID:"test",messageID:"test",callID:"test",abort:new AbortController().signal,metadata() {},async ask() {asked()}})
    // Bun's promise matcher may synchronously drain a pending promise. Consume
    // rejections without invoking matchers while the owner awaits this barrier.
    const revisionRefusal=revision.then(() => undefined,(error:unknown) => error)
    const legacyRefusal=legacy.then(() => undefined,(error:unknown) => error)
    await permission;release();await owner
    const revisionError=await revisionRefusal;const legacyError=await legacyRefusal
    expect(revisionError instanceof Error ? revisionError.message : "").toContain("exact prior hash")
    expect(legacyError instanceof Error ? legacyError.message : "").toContain("accepted Word revision")
    expect(await Bun.file(file).bytes()).toEqual(new Uint8Array(bytes))
    expect(readWorkingDocx(directory,file)?.source_sha256).toBe(createHash("sha256").update(bytes).digest("hex"))
  } finally {await rm(directory,{recursive:true,force:true})}
})

test("mixed and malformed targeting never silently export or apply legacy proposals",async () => {
  const original=await Packer.toBuffer(new Document({sections:[{children:[new Paragraph("Original wording remains.")]}]}))
  const legacy={id:1,doc_path:"working.docx",doc_name:"working.docx",scope:"phrase" as const,find_text:"Original",old_text:"Original",new_text:"Changed",author:"Review",review_comment:"",anchor_id:null,created_at:1,status:"pending" as const}
  const bound={...legacy,id:2,target_json:"{invalid"}
  await expect(buildRedlined(original,[legacy,bound])).rejects.toThrow("cannot share one redline export")
  await expect(bake(original,[legacy,bound])).rejects.toThrow("separate review")
  await expect(bake(original,[bound])).rejects.toThrow()
})

test("stale proposals stay pending and visible while plain download and fresh tracked changes remain available", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(),"dochaus-word-stale-"))
  const workspace = path.join(temporary,"workspace")
  const config = path.join(temporary,"config")
  await mkdir(workspace);await mkdir(config)
  await mkdir(path.join(workspace,".preferences"))
  await Bun.write(path.join(workspace,".preferences/preferences.json"),JSON.stringify({attorney:"Saved Reviewing Lawyer",firm:"Saved Reviewing Firm"}))
  await Bun.write(path.join(config,"opencode.json"),"{}")
  // Scripted local protocol fixture: one batch of amendments per Prepare request.
  const batches = [[
    {before_quote:"five days",replacement:"seven days",target_context:"notice period",reason:"Requested notice update."},
    {before_quote:"three months",replacement:"six months",target_context:"retention period",reason:"Requested retention update."},
  ],[
    {before_quote:"thirty days",replacement:"forty-five days",target_context:"payment period",reason:"Requested payment update."},
  ]]
  const engine = Bun.serve({port:0,fetch:async (request) => {
    const body = await request.json() as {prompt:string}
    const text = body.prompt.includes("<draft_view") ? JSON.stringify({amendments:batches.shift(),open_items:[]}) : (() => {
      const candidates = JSON.parse(body.prompt.match(/<candidates>([\s\S]*)<\/candidates>/)![1]) as Array<{amendment_id:string;target_handle:string}>
      return JSON.stringify({selections:[...new Set(candidates.map((row) => row.amendment_id))].map((id) => ({amendment_id:id,target_handle:candidates.find((row) => row.amendment_id === id)!.target_handle}))})
    })()
    return Response.json({data:{text,model:{providerID:"fixture",id:"isolated"},elapsed_ms:1,usage:null}})
  }})
  const probe = Bun.serve({port:0,fetch:() => new Response()});const port=probe.port;probe.stop(true)
  const process = Bun.spawn([globalThis.process.execPath,"run","src/server.ts"],{cwd:path.resolve(import.meta.dir,".."),
    env:{...globalThis.process.env,WORKSPACE_ROOT:workspace,DOCHAUS_DIR:config,INGEST_PORT:String(port),INGEST_HOST:"127.0.0.1",OPENCODE_URL:`http://127.0.0.1:${engine.port}`,OPENCODE_SERVER_PASSWORD:""},
    stdout:Bun.file(path.join(temporary,"ingest.log")),stderr:Bun.file(path.join(temporary,"ingest.log"))})
  const request=(route:string,options?:RequestInit) => fetch(`http://127.0.0.1:${port}${route}`,{...options,signal:AbortSignal.timeout(30000)})
  const hash=(bytes:Uint8Array) => createHash("sha256").update(bytes).digest("hex")
  const form=(bytes:Uint8Array) => {const body=new FormData();body.append("file",new File([new Uint8Array(bytes)],"working.docx"));return body}
  try {
    const deadline=Date.now()+20000
    while (!await request("/matters").then((response) => response.ok).catch(() => false)) {
      if (Date.now()>deadline) throw new Error("Isolated ingest did not start")
      await Bun.sleep(50)
    }
    const matter = await (await request("/matters",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({title:"Stale proposal review"})})).json() as {id:string;dir:string}
    const route=`/matters/${matter.id}`
    const file=path.join(matter.dir,"working.docx")
    const original = Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({sections:[{children:[
      new Paragraph("Notice must arrive within five days."),new Paragraph("Retain records for three months."),new Paragraph("Pay each invoice within thirty days."),
    ]}]}))))
    expect((await request(`${route}/deliverables`,{method:"POST",body:form(original)})).status).toBe(200)
    const prepare=async () => {
      const response=await request(`${route}/documents/amendments`,{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({name:"working.docx",instructions:"Update the stated periods."})})
      expect(response.status).toBe(200)
      return (await response.json() as {redline_ids:number[]}).redline_ids
    }
    const [acceptedID,staleID]=await prepare()
    expect((await request(`${route}/redlines/${acceptedID}/accept`,{method:"POST"})).status).toBe(200)
    const accepted=await Bun.file(file).bytes()
    expect(hash(accepted)).not.toBe(hash(original))
    const staleReason="This proposal targets an earlier Word revision. Prepare a fresh change list."
    const pending=async () => await (await request(`${route}/redlines?name=working.docx`)).json() as Array<{id:number;status:string;stale?:boolean;stale_reason?:string}>
    const exported=async () => {
      const response=await request(`${route}/documents/redlined?name=working.docx`)
      expect(response.status).toBe(200)
      return new Uint8Array(await response.arrayBuffer())
    }
    const download=async () => {
      const response=await request(`${route}/documents/content?name=working.docx`)
      expect(response.status).toBe(200)
      return new Uint8Array(await response.arrayBuffer())
    }
    // Only a stale proposal is pending: the export is the accepted file, unchanged.
    expect(hash(await exported())).toBe(hash(accepted))
    expect(hash(await download())).toBe(hash(accepted))
    const [freshID]=await prepare()
    expect(await pending()).toMatchObject([
      {id:staleID,status:"pending",stale:true,stale_reason:staleReason},
      {id:freshID,status:"pending"},
    ])
    expect((await pending())[1].stale).toBeUndefined()
    expect(hash(await download())).toBe(hash(accepted))
    // The tracked changes and review notes carry the fresh proposal only.
    const redline=await exported()
    const dx=await docxodus()
    const revised=(await dx.getRevisions(redline)).map((row) => row.text).join(" ")
    expect(revised).toContain("forty-five")
    expect(revised).not.toContain("six")
    expect(revised).not.toContain("three")
    expect(readDocxComments(redline).map((row) => row.text)).toEqual(["Requested payment update."])
    // Exporting changes nothing: no proposal is accepted, rejected or dropped.
    expect(await pending()).toMatchObject([{id:staleID,status:"pending",stale:true},{id:freshID,status:"pending"}])
    expect(hash(await Bun.file(file).bytes())).toBe(hash(accepted))
    expect(readWorkingDocx(matter.dir,file)?.source_sha256).toBe(hash(accepted))
    const db=openDb(matter.dir)
    try {
      expect([acceptedID,staleID,freshID].map((id) => getRedline(db,id)?.status)).toEqual(["accepted","pending","pending"])
    } finally {db.close()}
    expect((await request(`${route}/redlines/${staleID}/accept`,{method:"POST"})).status).toBe(409)
    expect(hash(await Bun.file(file).bytes())).toBe(hash(accepted))
  } finally {
    process.kill();await process.exited;engine.stop(true);await rm(temporary,{recursive:true,force:true})
  }
},120000)
