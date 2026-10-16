import { afterEach, expect, test } from "bun:test"
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { checklistRefinementSourceBasis, readEffectiveChecklistItems, reconcileChecklistParentCoverage, runChecklistRefinementFlow, type ChecklistCheckResult } from "./checklist-flow"
import { buildChecklistCheckerPrompt, type ChecklistItem } from "./checklist-contract"
import { writeLeanPlan } from "./plan"
import { readLeanJson, writeLeanText } from "./state"
import type { StageSessionPort } from "./stages"
const dirs:string[]=[]
afterEach(()=>dirs.splice(0).forEach(dir=>rmSync(dir,{recursive:true,force:true})))
const ref={document:"request",locator:"request",quote:"Propose a revision subject to approval."}
const item:ChecklistItem={check:"Address the source concerns.",basis:"proposal",source_refs:[ref],qualifications:["Subject to approval."],target:{document:"d1",role:"Agreement",treatment:"operative"}}
function setup(){const dir=mkdtempSync(path.join(tmpdir(),"refinement-"));dirs.push(dir);writeLeanPlan(dir,[{title:"Agreement",purpose:"Record instructions",file:"agreement.docx",parts:[{heading:"Agreement",scope:"Record instructions"}]}]);writeLeanText(dir,"request.md",ref.quote);writeLeanText(dir,"checklist/extraction.json",JSON.stringify({expected_sources:["request"],sources:[{key:"request",input_sha:"old",status:"succeeded",items:[{id:"parent",item}],issues:[]}]}));return dir}
function store(dir:string,units:unknown[],expected=["a"]){writeLeanText(dir,"checklist/refinements.json",JSON.stringify({parents:[{parent_id:"parent",parent:item,scope:"cited_sources",source_basis:checklistRefinementSourceBasis(dir),expected_units:expected,units}]}))}
const child={id:"child",item:{...item,required_action:"adopt" as const}}
function unit(){return {unit_id:"a",input_sha:"a",status:"succeeded",result:{parent_id:"parent",unit_id:"a",disposition:"refined",reason:"Source-directed check.",items:[child.item],source_refs:[ref],limitations:[]},children:[child],issues:[]}}
function finding(kind:"body"|"comment"="body"):ChecklistCheckResult{return {item_id:"child",input_sha:"x",status:"succeeded",issues:[],finding:{item_id:"child",status:"satisfied",action_status:"fulfilled",open_decisions:[],reason:"Present.",evidence:[{document:"d1",locator:"p1",text:"A revision.",kind}],source_refs:[ref],inspection_scope:"One paragraph.",limitation:"None."}}}
test("empty or incomplete assigned source scope blocks and preserves originals",()=>{const dir=setup();const original=readLeanJson(dir,"checklist/extraction.json");store(dir,[]);expect(readEffectiveChecklistItems(dir).complete).toBe(false);expect(reconcileChecklistParentCoverage(dir,[])[0].status).toBe("indeterminate");expect(readLeanJson(dir,"checklist/extraction.json")).toEqual(original);store(dir,[unit()],["a","b"]);expect(reconcileChecklistParentCoverage(dir,[finding()])[0].status).toBe("indeterminate")})
test("complete units plus native findings provide scoped coverage; comments do not establish adoption",()=>{const dir=setup();store(dir,[unit()]);expect(readEffectiveChecklistItems(dir).items).toEqual([child]);expect(reconcileChecklistParentCoverage(dir,[finding()])[0]).toMatchObject({status:"satisfied",scope:"cited_sources"});expect(reconcileChecklistParentCoverage(dir,[finding("comment")])[0].status).toBe("indeterminate");const open=finding();open.finding={...open.finding!,status:"open",open_decisions:["Client approval"]};expect(reconcileChecklistParentCoverage(dir,[open])[0].status).toBe("open");writeLeanText(dir,"request.md","Changed instruction.");expect(readEffectiveChecklistItems(dir).complete).toBe(false)})
test("source-only jobs cache completed units and source changes invalidate without rewriting parent",async()=>{const dir=setup();let calls=0;const sessions:StageSessionPort={async createSession(){return `s${calls}`},async prompt(input){calls++;const packet=JSON.parse(input.text.split("REFINEMENT_INPUT:\n\n")[1]);return JSON.stringify({parent_id:packet.parent_id,unit_id:packet.source_unit.id,disposition:"refined",reason:"Proposal only.",items:[{...item,required_action:"propose"}],source_refs:[ref],limitations:[]})}};const selection={parent_id:"parent",scope:"cited_sources" as const,units:[{id:"a",role:"Original instruction",passages:[ref]}],instructions:[ref],targets:[{document:"d1",role:"Agreement"}],retrieval:"Exact source pointer."};const deps={sessions,model:null};await runChecklistRefinementFlow(dir,deps,selection);await runChecklistRefinementFlow(dir,deps,selection);expect(calls).toBe(1);expect(readEffectiveChecklistItems(dir).items[0].item.required_action).toBe("propose");expect((readLeanJson(dir,"checklist/extraction.json") as any).sources[0].items[0].item).toEqual(item);writeLeanText(dir,"request.md",ref.quote+" Additional context.");await runChecklistRefinementFlow(dir,deps,selection);expect(calls).toBe(2)})
test("new action instructions are conditional and approval stays distinct",()=>{const packet={item_id:"parent",item,current_evidence:[],settled_instructions:[],retrieval:"Pointer"};expect(buildChecklistCheckerPrompt(packet)).not.toContain("Required action is distinct");const prompt=buildChecklistCheckerPrompt({...packet,item:{...item,required_action:"propose"}});expect(prompt).toContain("fulfilled while client approval remains open");expect(prompt).toContain('"required_action"')})
test("explicit no-applicable source disposition is counted without silently dropping unresolved units",()=>{const dir=setup();const none={...unit(),unit_id:"b",children:[],result:{...unit().result,unit_id:"b",disposition:"no_applicable_check",items:[],reason:"No additional applicable obligation in this unit."}};store(dir,[unit(),none],["a","b"]);expect(readEffectiveChecklistItems(dir).complete).toBe(true);expect(reconcileChecklistParentCoverage(dir,[finding()])[0].status).toBe("satisfied");store(dir,[unit(),{...none,result:{...none.result,disposition:"unresolved"}}],["a","b"]);expect(reconcileChecklistParentCoverage(dir,[finding()])[0].status).toBe("indeterminate")})


test("open approval cannot hide an unfulfilled or uninspected required action", () => {
  const dir = setup(); store(dir, [unit()])
  for (const action_status of ["not_fulfilled", "indeterminate"] as const) {
    const row = finding()
    row.finding = { ...row.finding!, status: "open", action_status, open_decisions: ["Approval"] }
    expect(reconcileChecklistParentCoverage(dir, [row])[0]!.status).toBe(action_status === "not_fulfilled" ? "unsatisfied" : "indeterminate")
    row.finding.status = "not_applicable"
    expect(reconcileChecklistParentCoverage(dir, [row])[0]!.status).toBe("indeterminate")
  }
  const fulfilled = finding()
  fulfilled.finding = { ...fulfilled.finding!, status: "open", open_decisions: ["Approval"], action_status: "fulfilled" }
  expect(reconcileChecklistParentCoverage(dir, [fulfilled])[0]!.status).toBe("open")
})

test("retargeted children never certify the original parent even when their own findings pass", () => {
  const dir = setup()
  for (const target of [{ ...item.target, document: "d2" }, { ...item.target, treatment: "advice" as const }]) {
    const changed = { ...child, item: { ...child.item, target } }
    store(dir, [{ ...unit(), children: [changed], result: { ...unit().result, items: [changed.item] } }])
    expect(readEffectiveChecklistItems(dir).complete).toBe(false)
    expect(reconcileChecklistParentCoverage(dir, [finding()])[0]!.status).toBe("indeterminate")
  }
})

test("refined checker final output example declares exactly the required action fields", () => {
  const packet = { item_id: "child", item, current_evidence: [], settled_instructions: [], retrieval: "Pointer" }
  const output = (prompt: string) => JSON.parse(prompt.split("Return only JSON ")[1]!.split(". Choose one enum value.")[0]!)
  expect(output(buildChecklistCheckerPrompt(packet)).action_status).toBeUndefined()
  const shape = output(buildChecklistCheckerPrompt({ ...packet, item: { ...item, required_action: "propose" } }))
  expect(shape.action_status).toBe("fulfilled|not_fulfilled|indeterminate|not_applicable")
  expect(shape.open_decisions).toEqual([])
})


test("open approval does not let comments prove fulfilled operative adoption", () => {
  const dir = setup(); store(dir, [unit()])
  const row = finding("comment")
  row.finding = { ...row.finding!, status: "open", action_status: "fulfilled", open_decisions: ["Approval"] }
  expect(reconcileChecklistParentCoverage(dir, [row])[0]!.status).toBe("indeterminate")
  row.finding.evidence[0]!.kind = "body"
  expect(reconcileChecklistParentCoverage(dir, [row])[0]!.status).toBe("open")
  row.finding.action_status = "not_applicable"
  row.finding.evidence[0]!.kind = "comment"
  expect(reconcileChecklistParentCoverage(dir, [row])[0]!.status).toBe("open")
})
