// Only the actual ReviewGrid extract/recheck producer -> session visibility contract.
// Run from apps/web: bun test/review-grid-ownership.browser.ts
import assert from "node:assert/strict"
import { mkdtemp, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { fileURLToPath } from "node:url"
import { isUserConversation } from "../../../dochaus/lib/session-visibility"
import { isConversation } from "../src/components/dashboard/data"
import type { Grid } from "../src/api/ingest/grid"

const {chromium,expect}=(await import(process.env.PLAYWRIGHT_MODULE??"@playwright/test")) as typeof import("@playwright/test")
const directory=await mkdtemp(path.join(tmpdir(),"grid-ownership-")), entry=path.join(directory,"harness.tsx")
await Bun.write(entry,`
import {createElement} from ${JSON.stringify(fileURLToPath(import.meta.resolve("react")))}
import {createRoot} from ${JSON.stringify(fileURLToPath(import.meta.resolve("react-dom/client")))}
import ReviewGrid from ${JSON.stringify(path.resolve(import.meta.dir,"../src/components/ReviewGrid.tsx"))}
createRoot(document.getElementById("root")).render(createElement(ReviewGrid,{matterId:"one",title:"Fixture",directory:"/synthetic/matter",documents:[{id:1,name:"Draft.docx",doc_path:"Draft.docx",created_at:0}],documentsRevision:0,onClose:()=>{},onAsk:()=>{}}))
`)
const build=await Bun.build({entrypoints:[entry],target:"browser",define:{"import.meta.env":JSON.stringify({VITE_INGEST_URL:"http://grid.invalid",VITE_OPENCODE_URL:"http://engine.invalid"})}})
assert.equal(build.success,true,build.logs.join("\n"))
const script=await build.outputs.find(output=>output.kind==="entry-point")!.text()
const server=Bun.serve({port:0,hostname:"127.0.0.1",fetch:request=>new URL(request.url).pathname==="/app.js"
  ?new Response(script,{headers:{"content-type":"text/javascript"}})
  :new Response('<div id="root"></div><script type="module" src="/app.js"></script>',{headers:{"content-type":"text/html"}})})
const browser=await chromium.launch({headless:true,...(process.env.PLAYWRIGHT_CHROMIUM_PATH?{executablePath:process.env.PLAYWRIGHT_CHROMIUM_PATH}:{})})
try {
  let grid:Grid={version:0,columns:[{id:"question",question:"When is payment due?",compare:false}],cells:{}}
  const created:{id:string;title?:string;metadata?:unknown;permissions?:unknown}[]=[], unexpected:string[]=[]
  const page=await browser.newPage();page.setDefaultTimeout(8000)
  await page.route("http://grid.invalid/**",async route=>{
    const request=route.request(),pathname=new URL(request.url()).pathname
    const json=(body:unknown)=>route.fulfill({status:200,contentType:"application/json",body:JSON.stringify(body)})
    if(pathname==="/matters/one/grid"){
      if(request.method()==="PUT")grid={...request.postDataJSON() as Grid,version:(grid.version??0)+1}
      return json(grid)
    }
    if(pathname==="/matters/one")return json({documents:[{id:1,name:"Draft.docx",doc_path:"Draft.docx",created_at:0}]})
    if(pathname==="/matters/one/grid/sources")return json({})
    if(pathname==="/matters/one/grid/recheck")return json({citations:[{documentName:"Draft.docx",docPath:"Draft.docx",section:"1",excerpt:"Payment is due tomorrow.",charStart:0,charEnd:24}],passages:"[1] Payment is due tomorrow.",scope:{passages:1,read:0,candidates:1,reread:0,terms:["payment"],channels:["exact"]}})
    if(pathname==="/preferences")return json({attorney:"Fixture"})
    if(pathname==="/engine/config")return json({model:"fixture/selected",model_concurrency:1})
    unexpected.push(`${request.method()} ${pathname}`);return route.fulfill({status:404})
  })
  await page.route("http://engine.invalid/**",route=>{
    const request=route.request(),pathname=new URL(request.url()).pathname,key=`${request.method()} ${pathname}`
    const json=(body:unknown)=>route.fulfill({status:200,contentType:"application/json",body:JSON.stringify(body)})
    if(key==="POST /api/session"){
      const body=request.postDataJSON() as (typeof created)[number], session={...body,id:`grid-${created.length}`};created.push(session)
      return json({data:{...session,location:{directory:"/synthetic/matter"}}})
    }
    const detail=created.find(item=>pathname===`/api/session/${item.id}`)
    if(request.method()==="GET"&&detail)return json({data:{...detail,location:{directory:"/synthetic/matter"}}})
    const session=created.find(item=>pathname.includes(`/session/${item.id}/`))
    if(session){
      if(/\/(agent|model|wait)$/.test(pathname))return route.fulfill({status:204})
      if(pathname.endsWith("/prompt"))return json({data:{id:"user",sessionID:session.id}})
      if(pathname.endsWith("/message")){
        const answer={type:"assistant",id:"answer",agent:"extract",time:{created:1,completed:2},content:[{type:"text",text:session.id==="grid-0"?"Not addressed":"Due tomorrow.\nSource: [1]"}]}
        const idle={type:"idle",id:"idle",outcome:"succeeded",time:{created:3}}
        return json({data:new URL(request.url()).searchParams.get("order")==="desc"?[idle,answer]:[answer,idle],cursor:{}})
      }
    }
    unexpected.push(key);return route.fulfill({status:404})
  })
  await page.goto(server.url.href)
  await page.locator(".review-cell-fill").click()
  await expect.poll(()=>created.length).toBe(2)
  await expect.poll(()=>grid.cells["Draft.docx::question"]?.recheck?.outcome).toBe("found")
  assert.deepEqual(created.map(session=>session.metadata),[{dochaus_role:"review-grid"},{dochaus_role:"review-grid"}])
  assert.deepEqual(created[1]!.permissions,[{action:"*",resource:"*",effect:"deny"}])
  assert.equal(created.every(session=>!isUserConversation(session)&&!isConversation(session)),true)
  assert.deepEqual(unexpected,[])
  console.log("PASS actual ReviewGrid extract + no-tools recheck producers emit ownership and stay outside conversation/sidebar/dashboard counts")
}finally{await browser.close();server.stop(true);await rm(directory,{recursive:true,force:true})}
