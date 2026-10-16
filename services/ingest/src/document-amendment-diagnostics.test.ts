import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtemp, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { documentAmendmentRoutes, readDocumentAmendmentDiagnostic } from "./document-amendments"
import { getDocumentRole, listSourceDocuments, openDb } from "./db"
import { writeDeliverable } from "./deliverable"
import { readWorkingDocx } from "../../../dochaus/lib/working-docx"

test("failed preparation retains bounded private phase evidence without publishing or changing pending approvals", async () => {
  const directory = await mkdtemp(path.join(tmpdir(),"dochaus-private-preparation-"))
  const bytes = Buffer.from(withParagraphPointers(await Packer.toBuffer(new Document({sections:[{children:[new Paragraph("The current notice period is five days.")]}]}))))
  const file = path.join(directory,"working.docx")
  const instructions = "Update the current notice period to seven days."
  const requests:Array<{prompt:string;directory:string|null;authorization:string|null}> = []
  const outputs:string[] = []
  let mode:"locator"|"quote"|"oversize" = "locator"
  const engine = Bun.serve({port:0,fetch:async (request) => {
    const body = await request.json() as {prompt:string}
    requests.push({prompt:body.prompt,directory:request.headers.get("x-opencode-directory"),authorization:request.headers.get("authorization")})
    const text = mode === "oversize" ? " ".repeat(2*1024*1024)+"{" : body.prompt.includes("<draft_view") ? JSON.stringify({amendments:[{
      before_quote:mode === "quote" ? "words absent from this document" : "five days",replacement:"seven days",target_context:"current notice period",reason:"Requested update.",
    }],open_items:[]}) : JSON.stringify({selections:[{amendment_id:"amendment-1",target_handle:"unoffered-handle",replacement:"unauthorized extra field"}]})
    outputs.push(text)
    return Response.json({data:{text,model:{providerID:"synthetic",id:"local"},elapsed_ms:2,usage:null}})
  }})
  const priorUrl = process.env.OPENCODE_URL
  const priorPassword = process.env.OPENCODE_SERVER_PASSWORD
  process.env.OPENCODE_URL = `http://127.0.0.1:${engine.port}`
  process.env.OPENCODE_SERVER_PASSWORD = "private-synthetic-secret"
  const app = documentAmendmentRoutes({directory:() => directory})
  const prepare = () => app.request("http://local/matter/documents/amendments",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({name:"working.docx",instructions})})
  try {
    await writeDeliverable(directory,"working.docx",bytes)
    const db = openDb(directory)
    db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,created_at) VALUES (?,?,?,?,?,?,?,?)",[file,"working.docx","phrase","current","current","operative","Reviewer",Date.now()])
    const pendingBefore = db.query("SELECT * FROM redlines").all()
    const response = await prepare()
    const publicFailure = await response.json() as {error:string;diagnostic_ref:string;diagnostic_hint:string}
    const receipt = readDocumentAmendmentDiagnostic(directory,"working.docx")!
    expect(response.status).toBe(409)
    expect(receipt.status).toBe("failed")
    expect(receipt.phase).toBe("locator_validation")
    expect(publicFailure.diagnostic_ref).toBe(receipt.attempt_ref)
    expect(publicFailure.diagnostic_hint).toContain("private")
    expect(JSON.stringify(publicFailure)).not.toContain("five days")
    expect(receipt.base_sha256).toBe(createHash("sha256").update(bytes).digest("hex"))
    expect(receipt.captures.first_prompt.value).toBe(requests[0].prompt)
    expect(receipt.captures.second_prompt.value).toBe(requests[1].prompt)
    expect(receipt.captures.first_text.value).toBe(outputs[0])
    expect(receipt.captures.second_text.value).toBe(outputs[1])
    expect(receipt.captures.frozen_legal_proposals.value).toEqual(JSON.parse(outputs[0]))
    expect((receipt.captures.candidates.value as unknown[]).length).toBe(1)
    expect(receipt.captures.selections.value).toEqual(JSON.parse(outputs[1]).selections)
    expect(receipt.timings.some((row) => row.phase === "legal_generation")).toBe(true)
    expect(receipt.timings.some((row) => row.phase === "locator_generation")).toBe(true)
    expect(receipt.timings.reduce((sum,row) => sum+row.elapsed_ms,0)).toBe(receipt.elapsed_ms)
    expect(Object.keys(receipt.error!).sort()).toEqual(["error_kind","error_ref"])
    expect(requests).toHaveLength(2)
    expect(requests.every((row) => row.directory === directory && Boolean(row.authorization))).toBe(true)
    expect(JSON.stringify(receipt)).not.toContain("private-synthetic-secret")
    expect(readDocumentAmendmentDiagnostic(directory,"working.docx")).toEqual(receipt)
    expect(await Bun.file(file).bytes()).toEqual(new Uint8Array(bytes))
    expect(db.query("SELECT * FROM redlines").all()).toEqual(pendingBefore)
    expect(listSourceDocuments(db)).toEqual([])
    expect(getDocumentRole(db,file)).toBe("deliverable")
    expect(readWorkingDocx(directory,file)).toBeUndefined()
    mode = "quote"
    const unmatched = await prepare()
    const unmatchedBody = await unmatched.json() as {diagnostic_ref:string}
    const latest = readDocumentAmendmentDiagnostic(directory,"working.docx")!
    expect(unmatched.status).toBe(409)
    expect(latest.attempt_ref).not.toBe(receipt.attempt_ref)
    expect(unmatchedBody.diagnostic_ref).toBe(latest.attempt_ref)
    expect(latest.phase).toBe("candidate_extraction")
    expect(latest.captures.first_text.value).toBe(outputs[2])
    expect(latest.captures.second_text).toBeUndefined()
    expect(requests).toHaveLength(3)
    expect(db.query("SELECT key FROM meta WHERE key LIKE 'document-amendment-diagnostic:%'").all()).toHaveLength(1)
    mode = "oversize"
    expect((await prepare()).status).toBe(409)
    const bounded = readDocumentAmendmentDiagnostic(directory,"working.docx")!
    expect(bounded.phase).toBe("legal_parse")
    expect(bounded.captures.first_text.value).toBeUndefined()
    expect(bounded.captures.first_text.omitted_reason).toContain("size limit")
    expect(bounded.captures.first_text.sha256).toBe(createHash("sha256").update(JSON.stringify(outputs[3])).digest("hex"))
    expect(Buffer.byteLength(JSON.stringify(bounded))).toBeLessThan(8*1024*1024+65536)
    expect(requests).toHaveLength(4)
    expect(await Bun.file(file).bytes()).toEqual(new Uint8Array(bytes))
    expect(db.query("SELECT * FROM redlines").all()).toEqual(pendingBefore)
    db.close()
  } finally {
    engine.stop(true)
    if (priorUrl === undefined) delete process.env.OPENCODE_URL
    else process.env.OPENCODE_URL = priorUrl
    if (priorPassword === undefined) delete process.env.OPENCODE_SERVER_PASSWORD
    else process.env.OPENCODE_SERVER_PASSWORD = priorPassword
    await rm(directory,{recursive:true,force:true})
  }
},30000)
