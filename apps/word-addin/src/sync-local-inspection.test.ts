import {expect,test} from "bun:test"
import {WordSync,type Transport,type Acknowledgement} from "./sync"
import type {WordConnectionState} from "../../../dochaus/lib/word-addin-contract"

function fixture(options:{refuseSettle?:boolean;owner?:"available"|"owned"|"other"}={}){
 const events:string[]=[]
 let owner=options.owner??"owned"
 let operation:{sequence:number;operationId:string|null;state:"idle"|"pending"|"confirmed"|"not-dispatched"}={sequence:0,operationId:null,state:"idle"}
 const state=():WordConnectionState=>({version:{base:{id:"base",number:1},current:{id:"base",number:1},state:"current"},nativeOperation:{sequence:operation.sequence,state:operation.state},lease:owner==="available"?{state:"available",expiresAt:null}:{state:owner,generation:"generation",expiresAt:999999}})
 const checkpoint:Acknowledgement={base:"base",fingerprint:"1",number:1,at:1,overlap:false}
 const transport:Transport={
  async lease(action){events.push(action);if(action==="claim")owner="owned";return state()},
  async nativeOperation(_base,_generation,input){events.push(input.action==="settle"?input.outcome:input.action);if(input.action==="begin")operation={sequence:input.sequence+1,operationId:input.operationId,state:"pending"};if(input.action==="settle"&&!options.refuseSettle)operation={sequence:input.sequence,operationId:input.operationId,state:input.outcome};return{operation,connection:state()}},
  async save(){events.push("upload");throw Error("Read-only check must not upload")},async heartbeat(){return state()},async done(){return state()},
 }
 const host={async saved(){return true},async save(){events.push("native-save")},async export(){events.push("export");return new Uint8Array([1])}}
 const sync=new WordSync("base",host,transport,{read:()=>checkpoint,write(){events.push("checkpoint")}},()=>{},async()=>"1",async()=>{})
 return{sync,events}
}

test("read-only inspection authorizes preparation but always settles not-dispatched without dirtying the document",async()=>{
 for(const owner of ["owned","available"] as const){
  const f=fixture({owner});await f.sync.start();f.events.length=0
  const before=f.sync.snapshot()
  expect(await f.sync.localInspection({conversationId:"conversation",turnId:"turn"},async identity=>{f.events.push("read-only");expect(identity).toEqual({base:"base",generation:"generation"});return{status:"ready"}})).toEqual({status:"ready"})
  expect(f.events).toEqual(["inspect",...(owner==="available"?["claim"]:[]),"begin","read-only","not-dispatched"])
  const after=f.sync.snapshot()
  expect(after.content).toBe(before.content);expect(after.status).toBe(before.status);expect(after.busy).toBe(false)
  expect(f.events).not.toContain("confirmed");expect(f.events).not.toContain("native-save");expect(f.events).not.toContain("upload");expect(f.events).not.toContain("checkpoint")
  f.sync.dispose()
 }
})

test("read-only inspection refusals settle truthfully and retain document state",async()=>{
 const f=fixture();await f.sync.start();f.events.length=0;const before=f.sync.snapshot()
 const cause=new Error("Safe preimage refusal")
 await expect(f.sync.localInspection({conversationId:"conversation",turnId:"turn"},async()=>{throw cause})).rejects.toBe(cause)
 expect(f.events).toEqual(["inspect","begin","not-dispatched"])
 expect(f.sync.snapshot().content).toBe(before.content);expect(f.sync.snapshot().status).toBe(before.status)
 f.sync.dispose()
 const blocked=fixture({refuseSettle:true});await blocked.sync.start();blocked.events.length=0
 await expect(blocked.sync.localInspection({conversationId:"conversation",turnId:"turn"},async()=>({status:"ready"}))).rejects.toMatchObject({mutationState:"uncertain"})
 expect(blocked.events).not.toContain("confirmed");expect(blocked.events).toContain("not-dispatched")
 blocked.sync.dispose()
})
