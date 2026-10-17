import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtemp, mkdir, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { pathToFileURL } from "node:url"
import { paragraphTargets } from "../../../dochaus/lib/drafting-amendment"
import { createDocxRevision } from "../../../dochaus/lib/docx-amendments"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"
import { getRedline, openDb } from "./db"
import { docxodus } from "./docxodus"

// An accepted doc.haus working file saved by LibreOffice keeps every paragraph
// pointer and text, so proposals bound to unchanged paragraphs stay valid after
// the saved file is imported. Fixtures and their provenance: fixtures/README.md.

type Candidate = {amendment_id:string;target_handle:string}
type Amendment = {before_quote:string;replacement:string;target_context:string;reason:string}

const pendingChanges:Amendment[] = [
  {before_quote:"sixty days",replacement:"forty days",target_context:"numbered payment clause",reason:"Requested payment update."},
  {before_quote:"in writing",replacement:"in writing or by email",target_context:"notice table detail",reason:"Requested notice method."},
  {before_quote:"three months",replacement:"six months",target_context:"retention period",reason:"Requested retention update."},
  {before_quote:"REF-ABC-1",replacement:"REF-ABC-2",target_context:"header reference",reason:"Requested header update."},
]
const closingChange:Amendment[] = [
  {before_quote:"Closing wording stays.",replacement:"Closing wording remains.",target_context:"closing paragraph",reason:"Requested closing update."},
]
const paragraphChanged = "The paragraph this change targets has changed since it was prepared. Prepare a fresh change list."
const earlierRevision = "This proposal targets an earlier Word revision. Prepare a fresh change list."
const sha256 = (bytes:Uint8Array) => createHash("sha256").update(bytes).digest("hex")
const readBytes = async (file:string) => new Uint8Array(await Bun.file(file).arrayBuffer())
const fixture = (name:string) => readBytes(path.join(import.meta.dir,"fixtures",name))
const wordForm = (bytes:Uint8Array) => {const body=new FormData();body.append("file",new File([new Uint8Array(bytes)],"working.docx"));return body}

// Body paragraphs with their pointer, exact-text hash and text.
async function bodyParagraphs(bytes:Uint8Array) {
  const revision = await createDocxRevision(bytes)
  return (await paragraphTargets(bytes)).map((row) => ({pid:row.pid,sha256:row.sha256,text:revision.paragraphs[row.ordinal]!.text}))
}

// Protocol fixture only: a local stand-in for the engine's generate route. The
// first pass returns the next frozen batch; the second picks each first candidate.
function syntheticEngine(batches:Amendment[][]) {
  return Bun.serve({port:0,fetch:async (request) => {
    const {prompt} = await request.json() as {prompt:string}
    const text = prompt.includes("<draft_view") ? JSON.stringify({amendments:batches.shift() ?? [],open_items:[]}) : (() => {
      const candidates = JSON.parse(prompt.match(/<candidates>([\s\S]*)<\/candidates>/)![1]) as Candidate[]
      return JSON.stringify({selections:[...new Set(candidates.map((row) => row.amendment_id))]
        .map((id) => ({amendment_id:id,target_handle:candidates.find((row) => row.amendment_id === id)!.target_handle}))})
    })()
    return Response.json({data:{text,model:{providerID:"fixture",id:"isolated"},elapsed_ms:1,usage:null}})
  }})
}

// One isolated ingest child process and one matter, driven through the real routes.
async function isolatedReview(temporary:string,batches:Amendment[][]) {
  const engine = syntheticEngine(batches)
  const workspace = path.join(temporary,"workspace")
  const config = path.join(temporary,"config")
  const log = path.join(temporary,"ingest.log")
  await mkdir(workspace,{recursive:true});await mkdir(config,{recursive:true})
  await Bun.write(path.join(config,"opencode.json"),"{}")
  const probe = Bun.serve({port:0,fetch:() => new Response()});const port=probe.port!;probe.stop(true)
  const service = Bun.spawn([process.execPath,"run","src/server.ts"],{cwd:path.resolve(import.meta.dir,".."),
    env:{...process.env,WORKSPACE_ROOT:workspace,DOCHAUS_DIR:config,INGEST_PORT:String(port),INGEST_HOST:"127.0.0.1",OPENCODE_URL:`http://127.0.0.1:${engine.port}`,OPENCODE_SERVER_PASSWORD:""},
    stdout:Bun.file(log),stderr:Bun.file(log)})
  const stop = async () => {service.kill();await service.exited;engine.stop(true)}
  const request = (route:string,options?:RequestInit) => fetch(`http://127.0.0.1:${port}${route}`,{...options,signal:AbortSignal.timeout(30000)})
  try {
    const deadline = Date.now()+20000
    while (!await request("/matters").then((response) => response.ok).catch(() => false)) {
      if (service.exitCode !== null || Date.now()>deadline) throw new Error(`Isolated ingest did not start:\n${await Bun.file(log).text()}`)
      await Bun.sleep(50)
    }
    const created = await (await request("/matters",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({title:"LibreOffice round trip"})})).json() as {id:string;dir:string}
    const route = `/matters/${created.id}`
    const file = path.join(created.dir,"working.docx")
    return {stop,file,
      current:async () => new Uint8Array(await Bun.file(file).arrayBuffer()),
      deliver:async (bytes:Uint8Array) => (await request(`${route}/deliverables`,{method:"POST",body:wordForm(bytes)})).status,
      prepare:async (instructions:string) => {
        const response = await request(`${route}/documents/amendments`,{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({name:"working.docx",instructions})})
        expect(response.status).toBe(200)
        return (await response.json() as {redline_ids:number[]}).redline_ids
      },
      revise:async (bytes:Uint8Array,prior:string) => (await request(`${route}/documents/revisions`,{method:"POST",headers:{"x-document-prior-sha256":prior},body:wordForm(bytes)})).status,
      accept:async (id:number) => (await request(`${route}/redlines/${id}/accept`,{method:"POST"})).status,
      post:(suffix:string,init?:RequestInit) => request(`${route}${suffix}`,{...init,method:"POST"}),
      redlined:() => request(`${route}/documents/redlined?name=working.docx`),
      pending:async () => (await (await request(`${route}/redlines?name=working.docx`)).json() as Array<{id:number;stale?:boolean;stale_reason?:string}>)
        .map((row) => [row.id,row.stale,row.stale_reason]),
      bindings:(ids:number[]) => {const db=openDb(created.dir);try {return ids.map((id) => JSON.parse(getRedline(db,id)!.target_json!) as Record<string,unknown>)} finally {db.close()}},
      rows:(ids:number[]) => {const db=openDb(created.dir);try {return ids.map((id) => getRedline(db,id))} finally {db.close()}},
      working:() => readWorkingDocx(created.dir,file),
    }
  } catch (error) {
    await stop();throw error
  }
}
type Review = Awaited<ReturnType<typeof isolatedReview>>

// Delivers the accepted file, prepares proposals on it, then imports the saved
// file through the revisions route with the delivered file's hash. Asserts that
// pointers and per-paragraph hashes survive and that only proposals on an
// edited paragraph, plus the whole-revision header proposal, go stale.
async function importSaved(review:Review,accepted:Uint8Array,saved:Uint8Array,edited:string|null) {
  expect(await review.deliver(accepted)).toBe(200)
  expect(sha256(await review.current())).toBe(sha256(accepted))
  const ids = await review.prepare("Update the payment period, the notice method, the retention period and the header reference.")
  expect(ids).toHaveLength(4)
  const before = await bodyParagraphs(accepted)
  const pid = (quote:string) => before.find((row) => row.text.includes(quote))!.pid
  // Body proposals bind their paragraph; the header proposal binds the whole revision.
  expect(review.bindings(ids).map((row) => [row.version,row.pid ?? null])).toEqual([[2,pid("sixty days")],[2,pid("in writing")],[2,pid("three months")],[1,null]])
  expect(await review.pending()).toEqual(ids.map((id) => [id,undefined,undefined]))

  const after = await bodyParagraphs(saved)
  // Every paragraph keeps its own pointer.
  expect(before.every((row) => row.pid)).toBe(true)
  expect(after.map((row) => row.pid)).toEqual(before.map((row) => row.pid))
  // Every paragraph keeps its exact text hash, except one edited in LibreOffice.
  expect(after.flatMap((row,index) => row.sha256 === before[index]!.sha256 ? [] : [row.text])).toEqual(edited ? [edited] : [])

  expect(await review.revise(saved,sha256(accepted))).toBe(200)
  expect(sha256(await review.current())).toBe(sha256(saved))
  const [payment,notice,retention,header] = ids as [number,number,number,number]
  expect(await review.pending()).toEqual([[payment,undefined,undefined],[notice,undefined,undefined],
    edited ? [retention,true,paragraphChanged] : [retention,undefined,undefined],[header,true,earlierRevision]])
  return {ids:{payment,notice,retention,header},before,after}
}

for (const {saved,edited} of [
  {saved:"synthetic-word-libreoffice-saved.docx",edited:null},
  {saved:"synthetic-word-libreoffice-edited.docx",edited:"Records must be retained securely for three months."},
]) test(`a LibreOffice save ${edited ? "with one edited paragraph " : ""}keeps paragraph-bound proposals valid through import, sibling accept, new Prepare and export`, async () => {
  const temporary = await mkdtemp(path.join(tmpdir(),"dochaus-word-libreoffice-"))
  try {
    const review = await isolatedReview(temporary,[pendingChanges,closingChange])
    try {
      const accepted = await fixture("synthetic-word-accepted.docx")
      const {ids,before,after} = await importSaved(review,accepted,await fixture(saved),edited)

      // The imported LibreOffice file exports the fresh proposals only.
      const exported = await review.redlined()
      expect(exported.status).toBe(200)
      const revisions = (await (await docxodus()).getRevisions(new Uint8Array(await exported.arrayBuffer()))).map((row) => row.text)
      for (const text of ["sixty","forty"," or by email"]) expect(revisions).toContain(text)
      expect(revisions.includes("six")).toBe(!edited)
      expect(revisions.join(" ")).not.toContain("REF-ABC")

      // A sibling accept applies on the LibreOffice file and keeps every pointer.
      expect(await review.accept(ids.payment)).toBe(200)
      const accepting = await bodyParagraphs(await review.current())
      expect(accepting.map((row) => row.pid)).toEqual(before.map((row) => row.pid))
      expect(accepting.map((row) => row.text)).toEqual(after.map((row) => row.text.replace("within sixty days","within forty days")))
      expect(await review.pending()).toEqual([[ids.notice,undefined,undefined],
        edited ? [ids.retention,true,paragraphChanged] : [ids.retention,undefined,undefined],[ids.header,true,earlierRevision]])

      // A new Prepare binds the LibreOffice file's pointer and accepts.
      const base = await review.current()
      const closing = accepting.find((row) => row.text === "Closing wording stays.")!
      const [fresh] = await review.prepare("Change the closing wording.") as [number]
      expect(review.bindings([fresh])).toEqual([{version:2,amendment_id:"amendment-1",pid:closing.pid,paragraph_sha256:closing.sha256,occurrence:0,base_sha256:sha256(base)}])
      expect(await review.accept(fresh)).toBe(200)
      expect((await bodyParagraphs(await review.current())).at(-1)).toEqual({...closing,text:"Closing wording remains.",
        sha256:sha256(new TextEncoder().encode("Closing wording remains."))})
      expect((await review.redlined()).status).toBe(200)
    } finally {await review.stop()}
  } finally {await rm(temporary,{recursive:true,force:true})}
},120000)

// One edit to the retention paragraph, as Word saves it with Track Changes on
// (a w:ins revision) or off (a plain run).
async function retentionEdit(bytes:Uint8Array,tracked:boolean) {
  const run = `<w:r><w:t xml:space="preserve"> Securely.</w:t></w:r>`
  const session = (await docxodus()).openDocxSession(bytes,{})
  try {
    const block = session.findByText("Records must be retained for three months.")!
    const result = session.raw.replaceXml(block.id,session.raw.getXml(block.id).replace(/<\/w:p>\s*$/,
      `${tracked ? `<w:ins w:id="901" w:author="Reviewer" w:date="2026-10-17T12:00:00Z">${run}</w:ins>` : run}</w:p>`))
    if (!result.success) throw new Error(result.error?.message)
    return session.save()
  } finally {session.close()}
}

test("a Word file with tracked changes is refused plainly at import and at Accept, changing nothing, while the same edit untracked imports", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(),"dochaus-word-tracked-"))
  try {
    const review = await isolatedReview(temporary,[pendingChanges])
    try {
      expect(await review.deliver(await fixture("synthetic-word-accepted.docx"))).toBe(200)
      const [payment,notice,retention,header] = await review.prepare("Update the payment period, the notice method, the retention period and the header reference.") as [number,number,number,number]
      // The first accept records the working Word revision that later imports must name.
      expect(await review.accept(payment)).toBe(200)
      const accepted = await review.current()
      const record = review.working()
      expect(record?.source_sha256).toBe(sha256(accepted))
      const ids = [notice,retention,header]
      const rows = review.rows(ids)
      const valid = [[notice,undefined,undefined],[retention,undefined,undefined],[header,true,earlierRevision]]
      expect(await review.pending()).toEqual(valid)
      const refused = async (response:Response,error:string) => {
        expect(response.status).toBe(409)
        expect(await response.json()).toEqual({error})
        expect(review.working()).toEqual(record)
        expect(review.rows(ids)).toEqual(rows)
      }
      const tracked = await retentionEdit(accepted,true)
      await expect(createDocxRevision(tracked)).rejects.toThrow("unsupported_existing_tracked_revisions")

      // Import through the revisions route: refused before anything is written.
      await refused(await review.post("/documents/revisions",{headers:{"x-document-prior-sha256":sha256(accepted)},body:wordForm(tracked)}),
        "This Word file contains tracked changes. Accept or reject them in Word, then upload it again.")
      expect(sha256(await review.current())).toBe(sha256(accepted))
      expect(await review.pending()).toEqual(valid)

      // A direct Word save left tracked changes in the stored file: every
      // revision-bound proposal says so, and Accept and Accept all refuse.
      const trackedAccept = "This Word file contains tracked changes. Accept or reject them in Word before changes can be applied."
      await Bun.write(review.file,tracked)
      expect(await review.pending()).toEqual(ids.map((id) => [id,true,trackedAccept]))
      for (const id of [notice,header]) await refused(await review.post(`/redlines/${id}/accept`),trackedAccept)
      await refused(await review.post("/redlines/accept-all?name=working.docx"),trackedAccept)
      expect(sha256(await review.current())).toBe(sha256(tracked))

      // Restored, the same edit without tracked changes imports and stales only its paragraph.
      await Bun.write(review.file,accepted)
      expect(await review.pending()).toEqual(valid)
      const untracked = await retentionEdit(accepted,false)
      expect(await review.revise(untracked,sha256(accepted))).toBe(200)
      expect(sha256(await review.current())).toBe(sha256(untracked))
      expect(review.working()?.source_sha256).toBe(sha256(untracked))
      expect(await review.pending()).toEqual([[notice,undefined,undefined],[retention,true,paragraphChanged],[header,true,earlierRevision]])
    } finally {await review.stop()}
  } finally {await rm(temporary,{recursive:true,force:true})}
},120000)

test("Prepare on a Word file with tracked changes and no working revision says how to resolve them", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(),"dochaus-word-tracked-prepare-"))
  try {
    const review = await isolatedReview(temporary,[pendingChanges])
    try {
      expect(await review.deliver(await retentionEdit(await fixture("synthetic-word-accepted.docx"),true))).toBe(200)
      expect(review.working()).toBeUndefined()
      const response = await review.post("/documents/amendments",{headers:{"content-type":"application/json"},body:JSON.stringify({name:"working.docx",instructions:"Update the retention period."})})
      expect(response.status).toBe(409)
      expect((await response.json() as {error:string}).error).toBe("This Word file contains tracked changes. Accept or reject them in Word before preparing changes.")
      expect(await review.pending()).toEqual([])
    } finally {await review.stop()}
  } finally {await rm(temporary,{recursive:true,force:true})}
},120000)

// Live check against the installed LibreOffice; skipped where it is absent.
test.skipIf(!Bun.which("soffice"))("a live LibreOffice save of the accepted file keeps every pointer and hash, and its proposals import fresh", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(),"dochaus-word-soffice-"))
  try {
    const accepted = await fixture("synthetic-word-accepted.docx")
    const input = path.join(temporary,"input")
    const output = path.join(temporary,"output")
    await mkdir(input);await mkdir(output)
    await Bun.write(path.join(input,"working.docx"),accepted)
    const soffice = Bun.spawn(["soffice","--headless","--norestore",`-env:UserInstallation=${pathToFileURL(path.join(temporary,"lo-profile")).href}`,
      "--convert-to","docx:MS Word 2007 XML","--outdir",output,path.join(input,"working.docx")],{stdout:"pipe",stderr:"pipe"})
    const timer = setTimeout(() => soffice.kill(),150000)
    const [exit,stdout,stderr] = await Promise.all([soffice.exited,new Response(soffice.stdout).text(),new Response(soffice.stderr).text()])
    clearTimeout(timer)
    if (exit !== 0 || !await Bun.file(path.join(output,"working.docx")).exists()) throw new Error(`soffice exited ${exit}:\n${stdout}\n${stderr}`)
    const saved = await readBytes(path.join(output,"working.docx"))
    // LibreOffice rewrote the package rather than copying it.
    expect(sha256(saved)).not.toBe(sha256(accepted))
    const review = await isolatedReview(path.join(temporary,"ingest"),[pendingChanges])
    try {await importSaved(review,accepted,saved,null)} finally {await review.stop()}
  } finally {await rm(temporary,{recursive:true,force:true})}
},180000)
