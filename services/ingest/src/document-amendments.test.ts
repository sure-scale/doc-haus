import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtemp, mkdir, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { AlignmentType, Document, LevelFormat, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { openDb, getRedline, listSourceDocuments } from "./db"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"
import { createDocxRevision, openRevision, verifyPreservation } from "../../../dochaus/lib/docx-amendments"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import wordIntegration from "../../../dochaus/tool/word-integration"
import { acceptedWorkingDocxUpdate, documentAmendmentRoutes } from "./document-amendments"
import { writeDeliverable, writeDeliverableRevision } from "./deliverable"
import { buildRedlined, bake } from "./redline"

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

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
type Candidate = {amendment_id:string;target_handle:string}
type LegalPass = {amendments:Array<{before_quote:string;replacement:string;target_context:string;reason:string}>;open_items:string[]}

// Protocol fixture only: a local stand-in for the engine's generate route. The
// first pass returns frozen legal proposals; the second returns locator rows.
function syntheticEngine(legal:() => LegalPass,locate:(candidates:Candidate[]) => unknown[]) {
  const prompts:string[] = []
  const server = Bun.serve({port:0,fetch:async (request) => {
    const {prompt} = await request.json() as {prompt:string}
    prompts.push(prompt)
    const text = prompt.includes("<draft_view") ? JSON.stringify(legal())
      : JSON.stringify({selections:locate(JSON.parse(prompt.match(/<candidates>([\s\S]*)<\/candidates>/)![1]) as Candidate[])})
    return Response.json({data:{text,model:{providerID:"fixture",id:"isolated"},elapsed_ms:1,usage:null}})
  }})
  return {server,prompts}
}
const firstCandidate = (candidates:Candidate[]) => [...new Set(candidates.map((row) => row.amendment_id))]
  .map((id) => ({amendment_id:id,target_handle:candidates.find((row) => row.amendment_id === id)!.target_handle}))
const sha256 = (bytes:Uint8Array) => createHash("sha256").update(bytes).digest("hex")

// In-process preparation route bound to a synthetic engine; restores the env.
async function withPreparation<T>(directory:string,engine:ReturnType<typeof syntheticEngine>,run:(prepare:(instructions:string) => Promise<Response>) => Promise<T>) {
  const prior = {url:process.env.OPENCODE_URL,password:process.env.OPENCODE_SERVER_PASSWORD}
  process.env.OPENCODE_URL = `http://127.0.0.1:${engine.server.port}`
  process.env.OPENCODE_SERVER_PASSWORD = ""
  const app = documentAmendmentRoutes({directory:() => directory})
  try {
    return await run(async (instructions) => app.request("http://local/matter/documents/amendments",{method:"POST",headers:{"content-type":"application/json"},
      body:JSON.stringify({name:"working.docx",instructions})}))
  } finally {
    if (prior.url === undefined) delete process.env.OPENCODE_URL; else process.env.OPENCODE_URL = prior.url
    if (prior.password === undefined) delete process.env.OPENCODE_SERVER_PASSWORD; else process.env.OPENCODE_SERVER_PASSWORD = prior.password
    engine.server.stop(true)
  }
}

test("an accepted revision, its stale sibling proposal and fresh revision-bound preparation survive an ingest restart", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(),"dochaus-word-restart-"))
  const workspace = path.join(temporary,"workspace")
  const config = path.join(temporary,"config")
  await mkdir(workspace);await mkdir(config)
  await Bun.write(path.join(config,"opencode.json"),"{}")
  let fresh = false
  const engine = syntheticEngine(() => fresh
    ? {amendments:[{before_quote:"three months",replacement:"six months",target_context:"retention period",reason:"Requested retention update."}],open_items:[]}
    : {amendments:[
      {before_quote:"five days",replacement:"seven days",target_context:"current notice period",reason:"Requested notice update."},
      {before_quote:"three months",replacement:"six months",target_context:"retention period",reason:"Requested retention update."},
    ],open_items:[]},firstCandidate)
  const services:Array<ReturnType<typeof Bun.spawn>> = []
  let port = 0
  // Each instance gets a fresh port and log; only the workspace root is shared.
  const start = async (label:string) => {
    const probe = Bun.serve({port:0,fetch:() => new Response()});port=probe.port!;probe.stop(true)
    const log = path.join(temporary,`ingest-${label}.log`)
    const service = Bun.spawn([process.execPath,"run","src/server.ts"],{cwd:path.resolve(import.meta.dir,".."),
      env:{...process.env,WORKSPACE_ROOT:workspace,DOCHAUS_DIR:config,INGEST_PORT:String(port),INGEST_HOST:"127.0.0.1",OPENCODE_URL:`http://127.0.0.1:${engine.server.port}`,OPENCODE_SERVER_PASSWORD:""},
      stdout:Bun.file(log),stderr:Bun.file(log)})
    services.push(service)
    const deadline = Date.now()+20000
    while (!await request("/matters").then((response) => response.ok).catch(() => false)) {
      if (service.exitCode !== null || Date.now()>deadline) throw new Error(`Isolated ingest ${label} did not start:\n${await Bun.file(log).text()}`)
      await Bun.sleep(50)
    }
    return service
  }
  const request = (route:string,options?:RequestInit) => fetch(`http://127.0.0.1:${port}${route}`,{...options,signal:AbortSignal.timeout(30000)})
  const form = (bytes:Uint8Array) => {const body=new FormData();body.append("file",new File([new Uint8Array(bytes)],"working.docx"));return body}
  const prepare = () => request(`${route}/documents/amendments`,{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({name:"working.docx",instructions:"Update the notice period and the retention period."})})
  const status = (id:number) => {const db=openDb(matter.dir);try {return getRedline(db,id)?.status} finally {db.close()}}
  let matter!:{id:string;dir:string};let route=""
  try {
    const first = await start("first")
    matter = await (await request("/matters",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({title:"Restart durability review"})})).json() as {id:string;dir:string}
    route = `/matters/${matter.id}`
    const original = Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({sections:[{children:[
      new Paragraph("Notice must arrive within five days."),new Paragraph("Retain records for three months."),new Paragraph("Unrelated closing wording stays."),
    ]}]}))))
    expect((await request(`${route}/deliverables`,{method:"POST",body:form(original)})).status).toBe(200)
    const file = path.join(matter.dir,"working.docx")
    const proposed = await (await prepare()).json() as {proposed:number;redline_ids:number[];base_sha256:string}
    expect(proposed.proposed).toBe(2)
    const [acceptedId,staleId] = proposed.redline_ids
    expect((await request(`${route}/redlines/${acceptedId}/accept`,{method:"POST"})).status).toBe(200)
    const accepted = await Bun.file(file).bytes()
    expect(sha256(accepted)).not.toBe(proposed.base_sha256)
    const record = readWorkingDocx(matter.dir,file)!
    expect(record.source_sha256).toBe(sha256(accepted))
    expect(record.projection_sha256).toBe((await createDocxRevision(accepted)).projection_sha256)

    first.kill();await first.exited
    await start("second")
    // Durable state only: the respawned service holds no in-memory carry-over.
    expect(sha256(await Bun.file(file).bytes())).toBe(sha256(accepted))
    expect(readWorkingDocx(matter.dir,file)).toEqual(record)
    const pending = await (await request(`${route}/redlines?name=working.docx`)).json() as Array<{id:number;stale?:boolean;stale_reason?:string}>
    expect(pending.map((row) => [row.id,row.stale])).toEqual([[staleId,true]])
    expect(pending[0].stale_reason).toContain("earlier Word revision")
    expect((await request(`${route}/redlines/${staleId}/accept`,{method:"POST"})).status).toBe(409)
    expect(sha256(await Bun.file(file).bytes())).toBe(sha256(accepted))
    expect(readWorkingDocx(matter.dir,file)).toEqual(record)
    expect(status(staleId)).toBe("pending")

    fresh = true
    const rebound = await (await prepare()).json() as {proposed:number;redline_ids:number[];base_sha256:string}
    expect(rebound.proposed).toBe(1)
    expect(rebound.base_sha256).toBe(sha256(accepted))
    const db = openDb(matter.dir)
    try {
      expect(JSON.parse(getRedline(db,rebound.redline_ids[0])!.target_json!).source_sha256).toBe(sha256(accepted))
    } finally {db.close()}
    expect((await request(`${route}/redlines/${rebound.redline_ids[0]}/accept`,{method:"POST"})).status).toBe(200)
    const final = await Bun.file(file).bytes()
    const revision = await createDocxRevision(final)
    expect(revision.markdown).toContain("Notice must arrive within seven days.")
    expect(revision.markdown).toContain("Retain records for six months.")
    expect(revision.markdown).toContain("Unrelated closing wording stays.")
    expect(readWorkingDocx(matter.dir,file)?.source_sha256).toBe(sha256(final))
    expect(readWorkingDocx(matter.dir,file)?.projection_sha256).toBe(revision.projection_sha256)
    expect(status(rebound.redline_ids[0])).toBe("accepted")
    // The older stale proposal never becomes acceptable against a later revision.
    expect(status(staleId)).toBe("pending")
    expect((await request(`${route}/redlines/${staleId}/accept`,{method:"POST"})).status).toBe(409)
    expect(sha256(await Bun.file(file).bytes())).toBe(sha256(final))
  } finally {
    for (const service of services) {service.kill();await service.exited}
    engine.server.stop(true);await rm(temporary,{recursive:true,force:true})
  }
},120000)

test("a locator unresolved_reason becomes a visible open item without a pending proposal or byte change", async () => {
  const directory = await mkdtemp(path.join(tmpdir(),"dochaus-word-unresolved-"))
  const file = path.join(directory,"working.docx")
  const ambiguous = "Two notice clauses could match; the instruction does not say which one governs."
  let mixed = false
  const engine = syntheticEngine(() => ({amendments:[
    {before_quote:"five days",replacement:"seven days",target_context:"governing notice clause",reason:"Requested notice update."},
    {before_quote:"three months",replacement:"six months",target_context:"retention period",reason:"Requested retention update."},
  ],open_items:["Confirm the client's retention policy."]}),(candidates) => firstCandidate(candidates).map((row) =>
    row.amendment_id === "amendment-1" || !mixed ? {amendment_id:row.amendment_id,unresolved_reason:ambiguous} : row))
  try {
    await writeDeliverable(directory,"working.docx",Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({sections:[{children:[
      new Paragraph("Notice must arrive within five days."),new Paragraph("Retain records for three months."),
    ]}]})))))
    const original = await Bun.file(file).bytes()
    const rows = () => {const db=openDb(directory);try {return db.query("SELECT * FROM redlines").all() as Array<{find_text:string;status:string;target_json:string}>} finally {db.close()}}
    await withPreparation(directory,engine,async (prepare) => {
      const refused = await prepare("Update the governing notice period and the retention period.")
      expect(refused.status).toBe(200)
      expect(await refused.json()).toMatchObject({proposed:0,open_items:["Confirm the client's retention policy.",ambiguous,ambiguous]})
      expect(rows()).toEqual([])
      expect(await Bun.file(file).bytes()).toEqual(original)
      expect(readWorkingDocx(directory,file)).toBeUndefined()

      mixed = true
      const partial = await prepare("Update the governing notice period and the retention period.")
      expect(partial.status).toBe(200)
      const body = await partial.json() as {proposed:number;redline_ids:number[];open_items:string[]}
      expect(body.proposed).toBe(1)
      expect(body.open_items).toEqual(["Confirm the client's retention policy.",ambiguous])
      // Only the resolved amendment is recorded; the unresolved quote has no row.
      expect(rows().map((row) => [row.find_text,row.status,JSON.parse(row.target_json).amendment_id])).toEqual([["three months","pending","amendment-2"]])
      expect(await Bun.file(file).bytes()).toEqual(original)
      expect(readWorkingDocx(directory,file)).toBeUndefined()
    })
    expect(engine.prompts).toHaveLength(4)
  } finally {engine.server.stop(true);await rm(directory,{recursive:true,force:true})}
})

test("accepting an amendment to a numbered list paragraph preserves numbering, tables, sibling paragraphs and untouched runs", async () => {
  const directory = await mkdtemp(path.join(tmpdir(),"dochaus-word-numbered-"))
  const file = path.join(directory,"working.docx")
  const engine = syntheticEngine(() => ({amendments:[
    {before_quote:"thirty days",replacement:"ten days",target_context:"first numbered delivery clause",reason:"Requested delivery update."},
  ],open_items:[]}),firstCandidate)
  try {
    await writeDeliverable(directory,"working.docx",await Packer.toBuffer(new Document({
      numbering:{config:[{reference:"clauses",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.START}]}]},
      sections:[{children:[
        new Paragraph("Supply terms"),
        new Paragraph({numbering:{reference:"clauses",level:0},children:[new TextRun("The Supplier shall deliver "),new TextRun({text:"all Goods",bold:true}),new TextRun({text:" within thirty days of the order date.",italics:true})]}),
        new Paragraph({numbering:{reference:"clauses",level:0},children:[new TextRun("Payment is due within "),new TextRun({text:"sixty days",bold:true}),new TextRun(".")]}),
        new Table({rows:[new TableRow({children:[new TableCell({children:[new Paragraph("Notice")]}),new TableCell({children:[new Paragraph({children:[new TextRun({text:"In writing",underline:{}})]})]})]})]}),
        new Paragraph("Closing paragraph."),
      ]}],
    })))
    const original = await Bun.file(file).bytes()
    const ids = await withPreparation(directory,engine,async (prepare) => {
      const response = await prepare("Shorten the delivery period in the first numbered clause to ten days.")
      expect(response.status).toBe(200)
      return (await response.json() as {redline_ids:number[]}).redline_ids
    })
    expect(ids).toHaveLength(1)
    const db = openDb(directory)
    const row = (() => {try {return getRedline(db,ids[0])!} finally {db.close()}})()
    // The accept route's baking step for a revision-bound proposal.
    const baked = await bake(original,[row])
    const before = await openRevision(original)
    const after = await openRevision(baked)
    try {
      const listOrdinal = before.paragraphs.findIndex((paragraph) => paragraph.match.text.includes("thirty days"))
      expect(before.paragraphs.flatMap((paragraph,index) => paragraph.match.text === after.paragraphs[index].match.text ? [] : [index])).toEqual([listOrdinal])
      expect(after.paragraphs[listOrdinal].match.text).toBe("The Supplier shall deliver all Goods within ten days of the order date.")
      expect(() => verifyPreservation(before,after,new Set([listOrdinal]))).not.toThrow()
      // Element markup is compared after one shared XML serializer, so the
      // engine's self-closing tag spelling is not mistaken for a change.
      const markup = (opened:typeof before,name:string) => Array.from(opened.documents["word/document.xml"].getElementsByTagNameNS(W,name)).map(String)
      expect(markup(before,"numPr")).toHaveLength(2)
      expect(markup(before,"numPr")[0]).toContain("numId")
      expect(markup(before,"tc")).toHaveLength(2)
      for (const name of ["numPr","tbl","tc","sectPr"]) expect(markup(after,name)).toEqual(markup(before,name))
      expect(after.paragraphs.map((paragraph,index) => index === listOrdinal ? "" : String(paragraph.node)))
        .toEqual(before.paragraphs.map((paragraph,index) => index === listOrdinal ? "" : String(paragraph.node)))
      const pPr = (node:Element) => Array.from(node.childNodes).filter((child) => (child as Element).localName === "pPr").map(String)
      expect(pPr(after.paragraphs[listOrdinal].node)).toEqual(pPr(before.paragraphs[listOrdinal].node))
      const runs = (node:Element) => Array.from(node.getElementsByTagNameNS(W,"r")).map((run) => ({
        text:Array.from(run.getElementsByTagNameNS(W,"t")).map((text) => text.textContent).join(""),
        rPr:Array.from(run.getElementsByTagNameNS(W,"rPr")).map(String).join(""),markup:String(run)}))
      const beforeRuns = runs(before.paragraphs[listOrdinal].node)
      const afterRuns = runs(after.paragraphs[listOrdinal].node)
      expect(afterRuns.filter((run) => !run.text.includes("ten days")).map((run) => run.markup))
        .toEqual(beforeRuns.filter((run) => !run.text.includes("thirty days")).map((run) => run.markup))
      expect(afterRuns.find((run) => run.text.includes("ten days"))!.rPr).toBe(beforeRuns.find((run) => run.text.includes("thirty days"))!.rPr)
      for (const part of ["word/numbering.xml","word/styles.xml"]) expect(after.parts[part]).toEqual(before.parts[part])
    } finally {before.session.close();after.session.close()}
  } finally {engine.server.stop(true);await rm(directory,{recursive:true,force:true})}
})
