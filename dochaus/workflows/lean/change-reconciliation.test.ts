import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { reconciliationWork, reconciliationPrompt, runChangeReconciliation } from "./change-reconciliation"
import { type ChangePlan, type ChangePreparation } from "./change-pipeline"
import { type ChangeList, readChangeList } from "./checks"
import { checkChanges } from "./change-checks"
import { writeLeanPlan } from "./plan"
import { packageCompositionBlock } from "./prompts"
import { readLeanState, writeLeanText } from "./state"
import { cleanupMatters, matter, requestMemo, scriptedModel, SOURCE, SUPPORT } from "./test-support"
const p1 = "p111111111111", p2 = "p222222222222", p3 = "p333333333333"
const support = [SUPPORT[0]!]
afterEach(cleanupMatters)
function setup() {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Record", purpose: "Record supplied terms", file: "record.docx", parts: [{ heading: "Terms", scope: "Terms" }] }])
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([p1, p2, p3].map((pid, index) => ({ pid, text: `Initial provision ${index + 1}.`, sha256: paragraphSha256(`Initial provision ${index + 1}.`) }))))
  return directory
}
function amendment(pid: string, updated: string) { return { id: "a1", pid, existing_sha256: paragraphSha256(`Initial provision ${[p1,p2,p3].indexOf(pid)+1}.`), updated, reason: "Carry supplied instruction.", sources: support } }
function changes(amendments: ReturnType<typeof amendment>[] = []): ChangeList { return { amendments, questions: [], shared_paragraphs: [] } }
function inputs(groups: { id: string; pids: string[] }[], ops: ChangeList[]) {
  const plan: ChangePlan = { shared_decisions: [], dismissed: [], findings: [], questions: [], groups: groups.map(row => ({ ...row, goal: "Preserve the supplied instruction and its conditions.", sources: support, limitations: [] })) }
  const preparation: ChangePreparation = { groups: groups.map((row, index) => ({ group: row.id, changes: ops[index]! })) }
  return { plan, preparation }
}
function save(directory: string, input: ReturnType<typeof inputs>) {
  writeLeanText(directory, "d1/change-plan.json", JSON.stringify(input.plan));writeLeanText(directory, "d1/change-preparation.json", JSON.stringify(input.preparation))
}
test("disjoint prepared operations assemble without native calls, retain current untouched history, and exclude old-draft operations", async () => {
  const dir = setup(), input = inputs([{id:"one",pids:[p1]},{id:"two",pids:[p2]}], [changes([amendment(p1,"First supplied term.")]),changes([amendment(p2,"Second supplied term.")])])
  const untouched=amendment(p3,"Unrelated accepted correction.")
  writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes([untouched,{...amendment("p999999999999","Old epoch."),existing_sha256:"old"}]),explanations:[{id:"i1",pid:p3,reason:"Initial revision explanation.",sources:support}]}))
  save(dir,input)
  const model=scriptedModel({})
  const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
  expect(result.status).toBe("done");expect(model.prompts).toEqual([])
  const accepted=readChangeList(dir,"d1")!
  expect(accepted.amendments.map(row=>[row.pid,row.updated])).toEqual([[p1,"First supplied term."],[p2,"Second supplied term."],[p3,untouched.updated]])
  expect(accepted.explanations?.[0]?.reason).toBe("Initial revision explanation.")
  expect(accepted.amendments.some(row=>row.pid==="p999999999999")).toBe(false)
})
test("overlapping groups union transitively; incompatible same-PID candidates go to one focused job",()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1,p2]},{id:"two",pids:[p2,p3]},{id:"three",pids:[p3]}],[changes([amendment(p2,"Candidate A.")]),changes([amendment(p2,"Candidate B.")]),changes([amendment(p3,"Dependent term.")])])
 const work=reconciliationWork(dir,"d1",input.plan,input.preparation)
 expect(work.conflicts).toHaveLength(1);expect(new Set(work.conflicts[0]!.pids)).toEqual(new Set([p1,p2,p3]));expect(work.direct.amendments).toEqual([])
})
test("identical operations coalesce reasons and source evidence without silent priority",()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]}],[changes([amendment(p1,"Same supplied term.")])])
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([{...amendment(p1,"Same supplied term."),reason:"Earlier supported rationale."}])))
 const work=reconciliationWork(dir,"d1",input.plan,input.preparation)
 expect(work.conflicts).toEqual([]);expect(work.direct.amendments).toHaveLength(1);expect(work.direct.amendments[0]!.reason).toContain("Earlier supported rationale.");expect(work.direct.amendments[0]!.reason).toContain("Carry supplied instruction.")
})
test("current changed explanations and withheld candidates remain explicit conflict work",()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]}],[changes([amendment(p1,"New effect.")])])
 writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes(),explanations:[{id:"i1",pid:p1,reason:"Earlier effect.",sources:support}],withheld:[{amendment:amendment(p1,"Unverified intent."),issues:["Source support unresolved."]}]}))
 const work=reconciliationWork(dir,"d1",input.plan,input.preparation)
 expect(work.conflicts).toHaveLength(1);expect(work.conflicts[0]!.changes.explanations).toHaveLength(1);expect(work.conflicts[0]!.changes.withheld).toHaveLength(1)
})
test("conflict resolution preserves source questions and stops on newly discovered dependency instead of publishing omissions",async()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]}],[changes([amendment(p1,"New proposal.")])]);writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([amendment(p1,"Earlier proposal.")])));save(dir,input)
 const previous=JSON.stringify(readChangeList(dir,"d1"))
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[],questions:[],dependencies:["The proposed condition depends on another group's provision."]})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("failed");expect(JSON.stringify(readChangeList(dir,"d1"))).toBe(previous);expect(model.prompts).toHaveLength(1)
})
test("conflict prompts omit unrelated matter history and retain actual source/instruction boundaries",()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]}],[changes([amendment(p1,"New proposal.")])]);writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([amendment(p1,"Earlier proposal.")])));save(dir,input)
 for(const file of ["spine.md","d1/review.md","d1/review-independent.md"])writeLeanText(dir,file,"UNRELATED_HISTORY_SENTINEL ".repeat(3000))
 const work=reconciliationWork(dir,"d1",input.plan,input.preparation),prompt=reconciliationPrompt(dir,"d1",work.conflicts[0]!)
 expect(prompt.length).toBeLessThan(12000);expect(prompt).not.toContain("UNRELATED_HISTORY_SENTINEL");expect(prompt).not.toContain("Initial provision 3.");expect(prompt).toContain(SOURCE);expect(prompt).toContain("proposal is not approval");expect(prompt).toContain("view=source")
})

test("indispensable source question survives focused reconciliation and blocks before apply",async()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]}],[changes([amendment(p1,"New proposal.")])]);writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([amendment(p1,"Earlier proposal.")])));save(dir,input)
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[{...amendment(p1,"Supported term. [open item: confirm timing choice]"),sources:support}],questions:[{pid:p1,question:"Which timing choice should be used?",context:"The supplied sources do not settle this essential choice."}],dependencies:[]})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result).toMatchObject({status:"blocked"});expect(readChangeList(dir,"d1")!.questions).toHaveLength(1);expect(model.prompts).toHaveLength(1)
})
test("a changed explanation cannot silently disappear from a conflicting component",async()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]}],[changes([amendment(p1,"New proposal.")])]);writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes(),explanations:[{id:"i1",pid:p1,reason:"Original substantive revision.",sources:support}]}));save(dir,input)
 let calls=0
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[amendment(p1,"Combined correction.")],questions:[],...(calls++?{explanations:[{pid:p1,reason:"Original revision and its corrected resulting effect.",sources:support}]}:{})})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("done");expect(model.prompts).toHaveLength(2);expect(model.prompts[1]!.text).toContain("explanation was omitted");expect(readChangeList(dir,"d1")!.explanations).toHaveLength(1)
})

test("distinct indispensable questions from independent accepted preparations both survive final validation",async()=>{
 const dir=setup(), groups=[{id:"one",pids:[p1]},{id:"two",pids:[p2]}]
 const lists=groups.map((group,index)=>{
  const checked=checkChanges(dir,JSON.stringify({amendments:[amendment(group.pids[0]!,`Supported term. [open item: confirm choice ${index+1}]`)],questions:[{pid:group.pids[0],question:`Which essential choice ${index+1} should be used?`,context:"Source does not supply this choice."}]}),[p1,p2,p3].map((pid,index)=>({pid,text:`Initial provision ${index+1}.`,sha256:paragraphSha256(`Initial provision ${index+1}.`)})),{document:"d1"})
  expect(checked.issues).toEqual([]);return checked.changes!
 })
 writeLeanText(dir,"questions.json",JSON.stringify({questions:lists.flatMap(row=>row.questions)}))
 const input=inputs(groups,lists);save(dir,input);const model=scriptedModel({})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("blocked");expect(readChangeList(dir,"d1")!.questions).toHaveLength(2);expect(readChangeList(dir,"d1")!.amendments).toHaveLength(2);expect(model.prompts).toEqual([])
})
test("a retained global open question is not silently discarded when it lacks a paragraph carrier",async()=>{
 const dir=setup(),input=inputs([],[]),question={id:"q-existing",stage:"changes",question:"Which essential choice applies?",context:"Unresolved supplied choice.",status:"open"}
 writeLeanText(dir,"questions.json",JSON.stringify({questions:[question]}));writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes(),questions:[question]}));save(dir,input)
 const work=reconciliationWork(dir,"d1",input.plan,input.preparation);expect(work.direct.questions).toHaveLength(1)
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:scriptedModel({}).port},"d1",()=>{})
 expect(result.status).toBe("failed");expect("reason" in result&&result.reason).toContain("tied to no paragraph");expect(readChangeList(dir,"d1")!.questions).toHaveLength(1)
})
test("one conflict result plus unaffected prepared operations retain complete operator coverage",async()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]},{id:"two",pids:[p2]}],[changes([amendment(p1,"New correction.")]),changes([amendment(p2,"Independent correction.")])])
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([amendment(p1,"Previous correction."),amendment(p3,"Unrelated retained correction.")])));save(dir,input)
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[amendment(p1,"Previous and new correction combined.")],questions:[],dependencies:[]})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("done");expect(readChangeList(dir,"d1")!.amendments.map(row=>row.pid)).toEqual([p1,p2,p3]);expect(model.prompts).toHaveLength(1);expect(readChangeList(dir,"d1")!.amendments[1]!.updated).toBe("Independent correction.")
})

test("restoring initial text requires an explicit source-grounded rejection instead of resurrecting or silently dropping an old operation",async()=>{
 const dir=setup(),input=inputs([{id:"restore",pids:[p1]}],[changes()]);writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([amendment(p1,"Earlier unsupported removal.")])));save(dir,input)
 expect(reconciliationWork(dir,"d1",input.plan,input.preparation).conflicts).toHaveLength(1)
 let call=0
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[],questions:[],...(call++?{retained_initial:[{pid:p1,reason:"The source supports the initial provision, not the earlier removal.",sources:support}]}:{})})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("done");expect(model.prompts).toHaveLength(2);expect(model.prompts[1]!.text).toContain("candidate was omitted");expect(readChangeList(dir,"d1")!.amendments).toEqual([]);expect(readChangeList(dir,"d1")!.retained_initial?.[0]?.reason).toContain("earlier removal")
})

test("withdrawing a prior operation requires a nonempty verified canonical source reference",async()=>{
 const dir=setup(),input=inputs([{id:"restore",pids:[p1]}],[changes()]);writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([amendment(p1,"Earlier supported operation.")])));save(dir,input)
 let call=0
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[],questions:[],retained_initial:[{pid:p1,reason:"Source-grounded restoration of initial wording.",sources:call++===0?[]:call===2?[{source_id:SOURCE,quote:"Invented words absent from source."}]:support}]})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("done");expect(model.prompts).toHaveLength(3);expect(model.prompts[1]!.text).toContain("source-grounded reason");expect(model.prompts[2]!.text).toContain("supporting quote is not verified");expect(readChangeList(dir,"d1")!.retained_initial?.[0]?.sources).toEqual(support)
})
test("a requested memo reaches the conflict prompt; without one it is unchanged",()=>{
 const dir=setup(),input=inputs([{id:"one",pids:[p1]}],[changes([amendment(p1,"New proposal.")])]);writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([amendment(p1,"Earlier proposal.")])));save(dir,input)
 const prompt=()=>reconciliationPrompt(dir,"d1",reconciliationWork(dir,"d1",input.plan,input.preparation).conflicts[0]!)
 const without=prompt()
 requestMemo(dir)
 const block=packageCompositionBlock(dir,"d1"),memo=prompt()
 expect(block).toStartWith("<package_composition>\n");expect(memo.split(block).length-1).toBe(1);expect(memo.replace(`\n\n${block}`,"")).toBe(without);expect(without).not.toContain("package_composition")
})
