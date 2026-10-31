import { expect, test } from "bun:test"
import { applyDocxLocalization, createDocxLocalization } from "../../../dochaus/lib/docx-amendments"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import { wordingEffects, requalifyWordingSurvivors } from "./wording-effect"
import { counterpartyMarkupDocx } from "../../../dochaus/test/support/counterparty-markup-fixture"

test("verified wording effects group a paragraph and requalify independent survivors",async()=>{
  const source=(await renderMarkdownToDocx("# Terms\n\nDelivery is thirty days and payment is monthly.\n\nNotices are written.\n\n| Item | Fee |\n| --- | --- |\n| Service | Ten pounds | ")).bytes
  const amendments=[{id:"delivery",before_quote:"thirty days",replacement:"ten days"},{id:"payment",before_quote:"monthly",replacement:"weekly"},{id:"notice",before_quote:"written",replacement:"by email"},{id:"fee",before_quote:"Ten pounds",replacement:"Twenty pounds"}]
  const packet=await createDocxLocalization(source,amendments)
  const selections=amendments.map(row=>({amendment_id:row.id,target_handle:packet.candidates.find(candidate=>candidate.amendment_id===row.id)!.target_handle}))
  const effects=await wordingEffects(source,packet,selections)
  expect(effects.groups.map(group=>group.ids)).toEqual([["delivery","payment"],["notice"],["fee"]])
  expect(effects.groups[0]).toMatchObject({before:"Delivery is thirty days and payment is monthly.",after:"Delivery is ten days and payment is weekly."})
  expect(effects.groups[2]!.context.table).toBeTruthy()
  const retained=await requalifyWordingSurvivors(source,packet,selections,effects,["notice","fee"])
  expect(retained.effects.groups).toEqual(effects.groups.slice(1))
  expect(retained.packet!.amendments.map(row=>row.id)).toEqual(["notice","fee"])
  expect(retained.selections).toEqual(selections.slice(2))
  expect((await applyDocxLocalization(source,retained.packet!,retained.selections)).report.amendments.map(row=>row.amendment_id)).toEqual(["notice","fee"])
  await expect(requalifyWordingSurvivors(source,packet,selections,effects,["delivery","notice"])).rejects.toThrow("wording_survivor_partial_group")
  await expect(requalifyWordingSurvivors(source,packet,selections,effects,["unknown"])).rejects.toThrow("wording_survivor_invalid")
  const reordered=await requalifyWordingSurvivors(source,packet,[...selections].reverse(),effects,["fee","notice"])
  expect(reordered.effects.groups).toEqual(effects.groups.slice(1))
  await expect(requalifyWordingSurvivors(source,packet,selections,effects,["notice","notice"])).rejects.toThrow("wording_survivor_invalid")
  const duplicatePacket=structuredClone(packet);duplicatePacket.amendments.push(duplicatePacket.amendments[0]!)
  await expect(wordingEffects(source,duplicatePacket,selections)).rejects.toThrow()
  const duplicateSelection=[...selections];duplicateSelection[3]=duplicateSelection[2]!
  await expect(wordingEffects(source,packet,duplicateSelection)).rejects.toThrow()
  const forgedSource=structuredClone(effects);forgedSource.source_sha256="forged"
  await expect(requalifyWordingSurvivors(source,packet,selections,forgedSource,["notice"])).rejects.toThrow("wording_survivor_invalid")
  const forgedContext=structuredClone(effects);forgedContext.groups[1]!.context.text="forged"
  await expect(requalifyWordingSurvivors(source,packet,selections,forgedContext,["notice"])).rejects.toThrow("wording_effect_mapping_invalid")
  const forged=structuredClone(effects);forged.groups[1]!.after="Misleading result."
  await expect(requalifyWordingSurvivors(source,packet,selections,forged,["notice"])).rejects.toThrow("wording_effect_mapping_invalid")
  const other=(await renderMarkdownToDocx("Changed source.")).bytes
  await expect(wordingEffects(other,packet,selections)).rejects.toThrow("stale_docx_revision")
  await expect(wordingEffects(source,packet,[...selections.slice(0,3),{...selections[3]!,target_handle:selections[2]!.target_handle}])).rejects.toThrow()
})


test("layered source with prior comments and revisions stays immutable",async()=>{
  const source=counterpartyMarkupDocx(), before=new Uint8Array(source)
  const options={pendingRevisions:"layer" as const,revisionAuthor:"Review Counsel"}
  const amendments=[{id:"law",before_quote:"Province",replacement:"Region"}]
  const packet=await createDocxLocalization(source,amendments,{},options)
  const selections=[{amendment_id:"law",target_handle:packet.candidates.find(row=>row.amendment_id==="law" && row.paragraph.scope!=="cmt")!.target_handle}]
  const effects=await wordingEffects(source,packet,selections,options)
  expect(effects.groups[0]!.before).toContain("Province")
  expect(effects.groups[0]!.after).toContain("Region")
  const result=await requalifyWordingSurvivors(source,packet,selections,effects,["law"],options)
  expect(result.effects).toEqual(effects)
  expect(source).toEqual(before)
})
