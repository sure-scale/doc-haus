import { afterEach, expect, test } from "bun:test"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { reconciliationWork, reconciliationPrompt, runChangeReconciliation } from "./change-reconciliation"
import { type ChangePlan, type ChangePreparation } from "./change-pipeline"
import { type ChangeList, readChangeList } from "./checks"
import { checkChanges } from "./change-checks"
import { writeLeanPlan } from "./plan"
import { packageCompositionBlock } from "./prompts"
import { LEAN_QUESTIONS_SCHEMA, readLeanState, writeLeanText } from "./state"
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
function changes(amendments: ReturnType<typeof amendment>[] = []): ChangeList { return { amendments, questions: [] } }
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
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[{...amendment(p1,"Supported term. [open item: confirm timing choice]"),sources:support}],questions:[{pid:p1,question:"Which timing choice should be used?",kind:"judgment",context:"The supplied sources do not settle this essential choice."}],dependencies:[]})})
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
  const checked=checkChanges(dir,JSON.stringify({amendments:[amendment(group.pids[0]!,`Supported term. [open item: confirm choice ${index+1}]`)],questions:[{pid:group.pids[0],question:`Which essential choice ${index+1} should be used?`,kind:"judgment",context:"Source does not supply this choice."}]}),[p1,p2,p3].map((pid,index)=>({pid,text:`Initial provision ${index+1}.`,sha256:paragraphSha256(`Initial provision ${index+1}.`)})),{document:"d1"})
  expect(checked.issues).toEqual([]);return checked.changes!
 })
 writeLeanText(dir,"questions.json",JSON.stringify({schema:LEAN_QUESTIONS_SCHEMA,questions:lists.flatMap(row=>row.questions)}))
 const input=inputs(groups,lists);save(dir,input);const model=scriptedModel({})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("blocked");expect(readChangeList(dir,"d1")!.questions).toHaveLength(2);expect(readChangeList(dir,"d1")!.amendments).toHaveLength(2);expect(model.prompts).toEqual([])
})
test("a retained global open question is not silently discarded when it lacks a paragraph carrier",async()=>{
 const dir=setup(),input=inputs([],[]),question={id:"q-existing",stage:"changes",question:"Which essential choice applies?",kind:"judgment",context:"Unresolved supplied choice.",status:"open"}
 writeLeanText(dir,"questions.json",JSON.stringify({schema:LEAN_QUESTIONS_SCHEMA,questions:[question]}));writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes(),questions:[question]}));save(dir,input)
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

function targets(prompt: string): string[] {
 return (JSON.parse(/<initial_targets>\n(.*)\n<\/initial_targets>/s.exec(prompt)![1]!) as { pid: string }[]).map(row => row.pid)
}
function provisions(directory: string, texts: string[]) {
 const rows = texts.map((text, index) => ({ pid: `p${(index + 1).toString(16).padStart(12, "0")}`, text, sha256: paragraphSha256(text) }))
 writeLeanPlan(directory, [{ title: "Record", purpose: "Record supplied terms", file: "record.docx", parts: [{ heading: "Terms", scope: "Terms" }] }])
 writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(rows))
 return rows
}
function op(row: { pid: string; sha256: string }, updated: string) { return { id: "a1", pid: row.pid, existing_sha256: row.sha256, updated, reason: "Carry supplied instruction.", sources: support } }
test("an oversized conflict component is reconciled in bounded slices, never refused",async()=>{
 const dir=matter(),filler="The supplier keeps the delivery schedule under review. ".repeat(7)
 const rows=provisions(dir,Array.from({length:200},(_,index)=>`Initial provision ${index+1}. ${filler}`))
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes(rows.map((row,index)=>op(row,`Earlier correction ${index+1}. ${filler}`)))))
 save(dir,inputs([{id:"whole",pids:rows.map(row=>row.pid)}],[changes(rows.map((row,index)=>op(row,`Prepared correction ${index+1}. ${filler}`)))]))
 const model=scriptedModel({"lean-changes":prompt=>JSON.stringify({amendments:targets(prompt).map(pid=>({pid,updated:`Reconciled correction for ${pid}.`,reason:"Combined supported effect.",sources:support})),questions:[],dependencies:[]})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("done");expect(model.prompts.length).toBeGreaterThan(1)
 for(const prompt of model.prompts)expect(prompt.text.length).toBeLessThanOrEqual(60_000)
 const targeted=model.prompts.flatMap(prompt=>targets(prompt.text))
 expect(targeted.toSorted()).toEqual(rows.map(row=>row.pid).toSorted())
 expect(readChangeList(dir,"d1")!.amendments.map(row=>[row.pid,row.updated])).toEqual(rows.map(row=>[row.pid,`Reconciled correction for ${row.pid}.`]))
})
test("a paragraph too large for one bounded step keeps its earlier change and withholds the new one as an open item",async()=>{
 const dir=matter(),rows=provisions(dir,["Initial provision 1.",`Long provision. ${"The supplier keeps the delivery schedule under review. ".repeat(1200)}`])
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([op(rows[0]!,"Earlier first correction."),op(rows[1]!,"Earlier shortened provision.")])))
 save(dir,inputs([{id:"one",pids:rows.map(row=>row.pid)}],[changes([op(rows[0]!,"New first correction."),op(rows[1]!,"New shortened provision.")])]))
 const model=scriptedModel({"lean-changes":prompt=>JSON.stringify({amendments:targets(prompt).map(pid=>({pid,updated:"Combined first correction.",reason:"Combined supported effect.",sources:support})),questions:[],dependencies:[]})})
 const result=await runChangeReconciliation(dir,readLeanState(dir)!,{sessions:model.port},"d1",()=>{})
 expect(result.status).toBe("done");expect(model.prompts).toHaveLength(1);expect(targets(model.prompts[0]!.text)).toEqual([rows[0]!.pid])
 const accepted=readChangeList(dir,"d1")!
 expect(accepted.amendments.map(row=>[row.pid,row.updated])).toEqual([[rows[0]!.pid,"Combined first correction."],[rows[1]!.pid,"Earlier shortened provision."]])
 expect(accepted.withheld?.map(row=>[row.amendment.pid,row.amendment.updated])).toEqual([[rows[1]!.pid,"New shortened provision."]])
 expect("openItems" in result&&result.openItems.filter(item=>item.kind==="unapplied_amendment").map(item=>[item.ref.pid,item.detail])).toEqual([[rows[1]!.pid,expect.stringContaining("too large to reconcile in one bounded step")]])
})
const tooLong=`Long provision. ${"The supplier keeps the delivery schedule under review. ".repeat(1200)}`
function tagged(prompt: string, tag: string) { return JSON.parse(new RegExp(`<${tag}>\\n(.*)\\n</${tag}>`, "s").exec(prompt)![1]!) }
function echo(edits: (pid: string) => boolean = () => true, explained: string[] = []) {
 return scriptedModel({"lean-changes":prompt=>{const listed=targets(prompt);return JSON.stringify({amendments:listed.filter(edits).map(pid=>({pid,updated:`Reconciled correction for ${pid}.`,reason:"Combined supported effect.",sources:support})),explanations:listed.filter(pid=>explained.includes(pid)).map(pid=>({pid,reason:"Initial revision explanation, corrected.",sources:support})),questions:[],dependencies:[]})}})
}
async function reconcile(dir: string, port: ReturnType<typeof scriptedModel>["port"]) { return runChangeReconciliation(dir,readLeanState(dir)!,{sessions:port},"d1",()=>{}) }
test("explanations, withheld rows and earlier-only paragraphs stay with their own slice across slice boundaries",async()=>{
 const dir=matter(),filler="The supplier keeps the delivery schedule under review. ".repeat(7)
 const rows=provisions(dir,Array.from({length:200},(_,index)=>`Initial provision ${index+1}. ${filler}`))
 const explained=rows.filter((_,index)=>index%9===0).map(row=>row.pid),pending=rows.filter((_,index)=>index%13===5)
 writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes(rows.map((row,index)=>op(row,`Earlier correction ${index+1}. ${filler}`))),explanations:explained.map((pid,index)=>({id:`i${index+1}`,pid,reason:"Initial revision explanation.",sources:support})),withheld:pending.map(row=>({amendment:op(row,"Unverified intent."),issues:["Source support unresolved."]}))}))
 save(dir,inputs([{id:"whole",pids:rows.map(row=>row.pid)}],[changes(rows.filter((_,index)=>index%7).map((row,index)=>op(row,`Prepared correction ${index+1}. ${filler}`)))]))
 const model=echo(undefined,explained),result=await reconcile(dir,model.port)
 expect(result.status).toBe("done");expect(model.prompts.length).toBeGreaterThan(1)
 for(const prompt of model.prompts){
  const listed=targets(prompt.text),own=tagged(prompt.text,"competing_operations_and_explanations") as ChangeList
  expect([...own.amendments.map(row=>row.pid),...own.explanations!.map(row=>row.pid),...own.withheld!.map(row=>row.amendment.pid)].filter(pid=>!listed.includes(pid))).toEqual([])
 }
 expect(model.prompts.flatMap(prompt=>targets(prompt.text)).toSorted()).toEqual(rows.map(row=>row.pid).toSorted())
 const accepted=readChangeList(dir,"d1")!
 expect(accepted.amendments).toHaveLength(200);expect(accepted.explanations!.map(row=>row.pid).toSorted()).toEqual(explained.toSorted());expect(accepted.withheld).toEqual([])
})
test("a paragraph too large for one bounded step keeps its earlier change, every explanation and its earlier withheld rows",async()=>{
 const dir=matter(),rows=provisions(dir,["Initial provision 1.",tooLong]),long=rows[1]!
 writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes([op(rows[0]!,"Earlier first correction."),op(long,"Earlier shortened provision.")]),explanations:[{id:"i1",pid:long.pid,reason:"Earlier initial revision explanation.",sources:support}],withheld:[{amendment:op(long,"Unverified intent."),issues:["Source support unresolved."]}]}))
 save(dir,inputs([{id:"one",pids:rows.map(row=>row.pid)}],[{...changes([op(rows[0]!,"New first correction."),op(long,"New shortened provision.")]),explanations:[{id:"i1",pid:long.pid,reason:"Prepared initial revision explanation.",sources:support}]}]))
 const model=echo(),result=await reconcile(dir,model.port)
 expect(result.status).toBe("done");expect(model.prompts.map(prompt=>targets(prompt.text))).toEqual([[rows[0]!.pid]])
 const accepted=readChangeList(dir,"d1")!
 expect(accepted.amendments.map(row=>[row.pid,row.updated])).toEqual([[rows[0]!.pid,`Reconciled correction for ${rows[0]!.pid}.`],[long.pid,"Earlier shortened provision."]])
 expect(accepted.explanations!.map(row=>[row.pid,row.reason])).toEqual([[long.pid,"Earlier initial revision explanation."],[long.pid,"Prepared initial revision explanation."]])
 expect(accepted.withheld!.map(row=>[row.amendment.pid,row.amendment.updated,row.issues])).toEqual([[long.pid,"Unverified intent.",["Source support unresolved."]],[long.pid,"New shortened provision.",[expect.stringContaining("too large to reconcile in one bounded step")]]])
 expect("openItems" in result&&result.openItems.filter(item=>item.kind==="unapplied_amendment").map(item=>item.ref.pid)).toEqual([long.pid,long.pid])
})
test("a paragraph too large for one bounded step keeps its own question and still blocks for the answer",async()=>{
 const dir=matter(),rows=provisions(dir,["Initial provision 1.",tooLong]),long=rows[1]!
 const question={id:"q-length",stage:"change-prepare" as const,question:"Which length applies?",kind:"judgment" as const,context:"The sources do not settle it.",pid:long.pid,document:"d1" as const,status:"open" as const}
 writeLeanText(dir,"questions.json",JSON.stringify({schema:LEAN_QUESTIONS_SCHEMA,questions:[question]}))
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([op(rows[0]!,"Earlier first correction."),op(long,"Earlier shortened provision. [open item: confirm length]")])))
 save(dir,inputs([{id:"one",pids:rows.map(row=>row.pid)}],[{...changes([op(rows[0]!,"New first correction."),op(long,"New shortened provision. [open item: confirm length]")]),questions:[{...question,stage:"changes"}]}]))
 const result=await reconcile(dir,echo().port)
 expect(result.status).toBe("blocked");expect(readChangeList(dir,"d1")!.questions.map(row=>[row.id,row.pid])).toEqual([["q-length",long.pid]])
})
test("a paragraph too large for one bounded step with one prepared change and nothing earlier is applied as prepared",async()=>{
 const dir=matter(),rows=provisions(dir,["Initial provision 1.",tooLong]),long=rows[1]!
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([op(rows[0]!,"Earlier first correction.")])))
 save(dir,inputs([{id:"one",pids:rows.map(row=>row.pid)}],[changes([op(rows[0]!,"New first correction."),op(long,"New shortened provision.")])]))
 const result=await reconcile(dir,echo().port),accepted=readChangeList(dir,"d1")!
 expect(result.status).toBe("done");expect(accepted.withheld).toEqual([])
 expect(accepted.amendments.map(row=>[row.pid,row.updated])).toEqual([[rows[0]!.pid,`Reconciled correction for ${rows[0]!.pid}.`],[long.pid,"New shortened provision."]])
})
test("a paragraph too large for one bounded step with several earlier changes withholds them all as open items instead of failing",async()=>{
 const dir=matter(),rows=provisions(dir,["Initial provision 1.",tooLong]),long=rows[1]!
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes([op(rows[0]!,"Earlier first correction."),op(long,"Earlier shortened provision."),{...op(long,"Earlier other provision."),id:"a3"}])))
 save(dir,inputs([{id:"one",pids:rows.map(row=>row.pid)}],[changes([op(rows[0]!,"New first correction."),op(long,"New shortened provision.")])]))
 const result=await reconcile(dir,echo().port),accepted=readChangeList(dir,"d1")!
 expect(result.status).toBe("done");expect(accepted.amendments.map(row=>row.pid)).toEqual([rows[0]!.pid])
 expect(accepted.withheld!.map(row=>row.amendment.updated)).toEqual(["Earlier shortened provision.","Earlier other provision.","New shortened provision."])
 expect("openItems" in result&&result.openItems.filter(item=>item.kind==="unapplied_amendment").map(item=>item.ref.pid)).toEqual([long.pid,long.pid,long.pid])
})
test("a split goal's slices point to the goal's other prepared changes instead of calling them independent",async()=>{
 const dir=matter(),filler="The supplier keeps the delivery schedule under review. ".repeat(7)
 const rows=provisions(dir,Array.from({length:40},(_,index)=>`Initial provision ${index+1}. ${filler}`))
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes(rows.map((row,index)=>op(row,`Earlier correction ${index+1}. ${filler}`)))))
 save(dir,inputs([{id:"whole",pids:rows.map(row=>row.pid)}],[changes(rows.map((row,index)=>op(row,`Prepared correction ${index+1}. ${filler}`)))]))
 const model=echo(),result=await reconcile(dir,model.port)
 expect(result.status).toBe("done");expect(model.prompts.length).toBeGreaterThan(1)
 for(const prompt of model.prompts){expect(prompt.text.length).toBeLessThanOrEqual(60_000);expect(prompt.text).toContain("read-record .dochaus/lean/d1/change-preparation.json");expect(prompt.text).not.toContain("that is not a dependency")}
})
test("when the goal and instructions alone exceed the bound, the withheld changes say so rather than blaming the paragraph",async()=>{
 const dir=matter(),rows=provisions(dir,Array.from({length:5},(_,index)=>`Initial provision ${index+1}.`))
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes(rows.map((row,index)=>op(row,`Earlier correction ${index+1}.`)))))
 const input=inputs([{id:"whole",pids:rows.map(row=>row.pid)}],[changes(rows.map((row,index)=>op(row,`Prepared correction ${index+1}.`)))])
 input.plan.groups[0]!.goal=`Shorten every provision. ${"Keep every supplied condition. ".repeat(2000)}`;save(dir,input)
 const model=echo(),result=await reconcile(dir,model.port),accepted=readChangeList(dir,"d1")!
 expect(result.status).toBe("done");expect(model.prompts).toEqual([])
 expect(accepted.amendments.map(row=>row.updated)).toEqual(rows.map((_,index)=>`Earlier correction ${index+1}.`))
 expect(accepted.withheld!.map(row=>[row.amendment.updated,row.issues])).toEqual(rows.map((_,index)=>[`Prepared correction ${index+1}.`,[expect.stringContaining("instructions, goals and shared decisions")]]))
 expect(JSON.stringify(accepted.withheld)).not.toContain("paragraph is too large")
})
test("slices with nothing to reconcile make no model call, and results keep job order whichever slice finishes first",async()=>{
 const dir=matter(),filler="The supplier keeps the delivery schedule under review. ".repeat(20)
 const rows=provisions(dir,Array.from({length:200},(_,index)=>`Initial provision ${index+1}. ${filler}`))
 const edited=[...rows.slice(0,5),...rows.slice(195)],first=edited[0]!.pid,last=edited.at(-1)!.pid
 writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes(edited.map((row,index)=>op(row,`Earlier correction ${index+1}.`))),explanations:[first,last].map((pid,index)=>({id:`i${index+1}`,pid,reason:`Initial revision explanation ${index+1}.`,sources:support}))}))
 save(dir,inputs([{id:"whole",pids:rows.map(row=>row.pid)}],[changes(edited.map((row,index)=>op(row,`Prepared correction ${index+1}.`)))]))
 const model=echo(pid=>edited.some(row=>row.pid===pid),[first,last])
 const port={...model.port,prompt:async(input:Parameters<typeof model.port.prompt>[0])=>{if(input.text.includes(`"pid":"${first}"`))await Bun.sleep(25);return model.port.prompt(input)}}
 const result=await reconcile(dir,port)
 expect(result.status).toBe("done");expect(model.prompts).toHaveLength(2)
 expect(readChangeList(dir,"d1")!.explanations!.map(row=>[row.id,row.pid])).toEqual([["i1",first],["i2",last]])
})
const splitFiller="The supplier keeps the delivery schedule under review. ".repeat(7).trim()
function splitGoal(dir: string, extra = 0, adjust: (earlier: ChangeList, prepared: ChangeList, first: { pid: string; sha256: string }) => void = () => {}) {
 const filler=splitFiller
 const rows=provisions(dir,Array.from({length:40+extra},(_,index)=>`Initial provision ${index+1}. ${filler}`)),goal=rows.slice(0,40)
 const earlier=changes(goal.map((row,index)=>op(row,`Earlier correction ${index+1}. ${filler}`))),prepared=changes(goal.map((row,index)=>op(row,`Prepared correction ${index+1}. ${filler}`)))
 adjust(earlier,prepared,goal[0]!)
 writeLeanText(dir,"d1/changes.json",JSON.stringify(earlier))
 save(dir,inputs([{id:"whole",pids:goal.map(row=>row.pid)}],[prepared]))
 return rows
}
function dependent(pid: string, on: string, reporter = pid) {
 return scriptedModel({"lean-changes":prompt=>{const listed=targets(prompt),reports=listed.includes(reporter);return JSON.stringify({amendments:listed.filter(row=>!reports||row!==pid).map(row=>({pid:row,updated:`Reconciled correction for ${row}.`,reason:"Combined supported effect.",sources:support})),questions:[],dependencies:reports?[{pid,depends_on:[on],issue:"The renamed term must match the definition reconciled elsewhere."}]:[]})}})
}
test("a sliced job's dependency on another slice of the same goal withholds that change and keeps its earlier one instead of failing",async()=>{
 const dir=matter(),rows=splitGoal(dir),model=dependent(rows[0]!.pid,rows[39]!.pid),result=await reconcile(dir,model.port)
 expect(result.status).toBe("done");expect(model.prompts.length).toBeGreaterThan(1)
 expect(model.prompts.every(prompt=>prompt.text.includes('"depends_on"'))).toBe(true)
 const accepted=readChangeList(dir,"d1")!
 expect(accepted.amendments.map(row=>[row.pid,row.updated])).toEqual(rows.map((row,index)=>[row.pid,index?`Reconciled correction for ${row.pid}.`:`Earlier correction 1. ${splitFiller}`]))
 expect(accepted.withheld!.map(row=>[row.amendment.pid,row.amendment.updated,row.issues])).toEqual([[rows[0]!.pid,`Prepared correction 1. ${splitFiller}`,[expect.stringContaining("depends on another paragraph of the same edit")]]])
 expect(accepted.withheld![0]!.issues[0]).toContain('It depends on the paragraph beginning "Initial provision 40. The supplier keeps the delivery…". Reason given: The renamed term must match the definition reconciled elsewhere.')
 expect(accepted.withheld![0]!.issues[0]).not.toContain(rows[39]!.pid)
 for(const prompt of model.prompts)expect(prompt.text).not.toContain("dependencies holds a newly discovered cross-component dependency")
 expect("openItems" in result&&result.openItems.filter(item=>item.kind==="unapplied_amendment").map(item=>item.ref.pid)).toEqual([rows[0]!.pid])
})
test("a sliced job's dependency outside its own goal still stops the run",async()=>{
 const dir=matter(),rows=splitGoal(dir,1),before=JSON.stringify(readChangeList(dir,"d1"))
 const result=await reconcile(dir,dependent(rows[0]!.pid,rows[40]!.pid).port)
 expect(result.status).toBe("failed");expect(JSON.stringify(readChangeList(dir,"d1"))).toBe(before)
})
test("a sliced job cannot defer a paragraph that has no new change to withhold; the run stops instead of losing the report",async()=>{
 const shapes:[string,Parameters<typeof splitGoal>[2],boolean][]=[
  ["no candidate",(earlier,prepared)=>{earlier.amendments.shift();prepared.amendments.shift()},false],
  ["earlier change only",(_,prepared)=>{prepared.amendments.shift()},false],
  ["prepared change equal to the earlier one",(earlier,prepared)=>{prepared.amendments[0]={...earlier.amendments[0]!}},false],
  ["earlier withheld row only",(earlier,prepared,first)=>{earlier.amendments.shift();prepared.amendments.shift();earlier.withheld=[{amendment:op(first,"Unverified intent."),issues:["Source support unresolved."]}]},false],
  ["paragraph listed in another job",()=>{},true],
 ]
 for(const [name,adjust,elsewhere] of shapes){
  const dir=matter(),rows=splitGoal(dir,0,adjust),before=JSON.stringify(readChangeList(dir,"d1"))
  const model=elsewhere?dependent(rows[39]!.pid,rows[38]!.pid,rows[0]!.pid):dependent(rows[0]!.pid,rows[39]!.pid)
  const result=await reconcile(dir,model.port)
  expect([name,result.status]).toEqual([name,"failed"]);expect(JSON.stringify(readChangeList(dir,"d1"))).toBe(before)
 }
})
test("a sliced job's deferred paragraph with only a prepared change has it withheld, not applied",async()=>{
 const dir=matter(),rows=splitGoal(dir,0,earlier=>{earlier.amendments.shift()}),result=await reconcile(dir,dependent(rows[0]!.pid,rows[39]!.pid).port)
 expect(result.status).toBe("done")
 const accepted=readChangeList(dir,"d1")!
 expect(accepted.amendments.some(row=>row.pid===rows[0]!.pid)).toBe(false)
 expect(accepted.withheld!.map(row=>[row.amendment.pid,row.amendment.updated])).toEqual([[rows[0]!.pid,`Prepared correction 1. ${splitFiller}`]])
})
test("a sliced job's deferred paragraph keeps its earlier withheld row once beside the new one",async()=>{
 const dir=matter(),rows=splitGoal(dir,0,(earlier,_,first)=>{earlier.withheld=[{amendment:op(first,"Unverified intent."),issues:["Source support unresolved."]}]})
 const result=await reconcile(dir,dependent(rows[0]!.pid,rows[39]!.pid).port)
 expect(result.status).toBe("done")
 expect(readChangeList(dir,"d1")!.withheld!.map(row=>[row.amendment.updated,row.issues.length])).toEqual([["Unverified intent.",1],[`Prepared correction 1. ${splitFiller}`,1]])
})
test("a paragraph too large for one bounded step keeps its earlier retained-initial disposition",async()=>{
 const dir=matter(),rows=provisions(dir,["Initial provision 1.",tooLong]),long=rows[1]!
 const retained={pid:long.pid,reason:"The source supports the initial provision.",sources:support}
 writeLeanText(dir,"d1/changes.json",JSON.stringify({...changes([op(rows[0]!,"Earlier first correction.")]),retained_initial:[retained]}))
 save(dir,inputs([{id:"one",pids:rows.map(row=>row.pid)}],[changes([op(rows[0]!,"New first correction."),op(long,"New shortened provision.")])]))
 const result=await reconcile(dir,echo().port),accepted=readChangeList(dir,"d1")!
 expect(result.status).toBe("done");expect(accepted.retained_initial).toEqual([retained])
 expect(accepted.withheld!.map(row=>[row.amendment.pid,row.amendment.updated])).toEqual([[long.pid,"New shortened provision."]])
})
test("failure reasons list jobs in document order whichever job fails first",async()=>{
 const dir=matter(),filler="The supplier keeps the delivery schedule under review. ".repeat(7)
 const rows=provisions(dir,Array.from({length:200},(_,index)=>`Initial provision ${index+1}. ${filler}`)),first=rows[0]!.pid
 writeLeanText(dir,"d1/changes.json",JSON.stringify(changes(rows.map((row,index)=>op(row,`Earlier correction ${index+1}. ${filler}`)))))
 save(dir,inputs([{id:"whole",pids:rows.map(row=>row.pid)}],[changes(rows.map((row,index)=>op(row,`Prepared correction ${index+1}. ${filler}`)))]))
 const model=scriptedModel({"lean-changes":()=>JSON.stringify({amendments:[],questions:[],dependencies:["The change depends on another document."]})})
 const port={...model.port,prompt:async(input:Parameters<typeof model.port.prompt>[0])=>{if(input.text.includes(`"pid":"${first}"`))await Bun.sleep(25);return model.port.prompt(input)}}
 const result=await reconcile(dir,port)
 expect(result.status).toBe("failed")
 const parts=[...((result as {reason:string}).reason.matchAll(/whole part (\d+) of/g))].map(match=>Number(match[1]))
 expect(parts.length).toBeGreaterThan(2);expect(parts).toEqual(parts.toSorted((a,b)=>a-b))
})
test("a deferred change's reported reasons name other paragraphs by their opening words, never by id",async()=>{
 const dir=matter(),rows=splitGoal(dir),[first,before,last]=[rows[0]!.pid,rows[38]!.pid,rows[39]!.pid]
 const model=scriptedModel({"lean-changes":prompt=>{const listed=targets(prompt),reports=listed.includes(first);return JSON.stringify({amendments:listed.filter(row=>!reports||row!==first).map(row=>({pid:row,updated:`Reconciled correction for ${row}.`,reason:"Combined supported effect.",sources:support})),questions:[],dependencies:reports?[{pid:first,depends_on:[last],issue:`Term must match ${last}.`},{pid:first,depends_on:[before],issue:`Cross-reference to ${before} moves.`}]:[]})}})
 const result=await reconcile(dir,model.port),reason=readChangeList(dir,"d1")!.withheld![0]!.issues[0]!
 expect(result.status).toBe("done");expect(reason).not.toMatch(/p[0-9a-f]{12}/)
 expect(reason).toContain('Reason given: Term must match the paragraph beginning "Initial provision 40. The supplier keeps the delivery…"; Cross-reference to the paragraph beginning "Initial provision 39. The supplier keeps the delivery…" moves.')
})
