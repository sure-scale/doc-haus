import { applyDocxLocalization, createDocxLocalization, type DocxLocalizationPacket, type DocxRevisionOptions } from "../../../dochaus/lib/docx-amendments"

type Selection = {amendment_id:string;target_handle:string}
export type WordingEffect = {paragraph_handle:string;ids:string[];before:string;after:string;context:DocxLocalizationPacket["candidates"][number]["paragraph"]}
export type WordingEffects = {source_sha256:string;groups:WordingEffect[]}

// Ephemeral verified effects: the writer checks source, selections, saved text
// and preservation. No model-authored description becomes effect authority.
export async function wordingEffects(bytes:Uint8Array,packet:DocxLocalizationPacket,selections:Selection[],options:DocxRevisionOptions = {}):Promise<WordingEffects> {
  const applied = await applyDocxLocalization(bytes,packet,selections,options)
  const reports = applied.report.amendments
  if (reports.length !== selections.length || new Set(reports.map(row => row.amendment_id)).size !== selections.length ||
      applied.report.source_sha256 !== packet.revision.source_sha256 || !applied.report.package_structure_verified) throw new Error("wording_effect_mapping_invalid")
  const groups = new Map<string,WordingEffect>()
  for (const row of reports) {
    const selection = selections.find(value => value.amendment_id === row.amendment_id && value.target_handle === row.target_handle)
    const amendment = packet.amendments.find(value => value.id === row.amendment_id)
    const candidates = packet.candidates.filter(value => value.amendment_id === row.amendment_id && value.target_handle === row.target_handle)
    if (!selection || !amendment || candidates.length !== 1 || candidates[0]!.paragraph.paragraph_handle !== row.paragraph_handle ||
      candidates[0]!.paragraph.text !== row.paragraph_before || amendment.before_quote !== row.before_quote || amendment.replacement !== row.replacement)
      throw new Error("wording_effect_mapping_invalid")
    const group = groups.get(row.paragraph_handle)
    if (group) {
      if (group.before !== row.paragraph_before || group.after !== row.paragraph_after) throw new Error("wording_effect_mapping_invalid")
      group.ids.push(row.amendment_id)
    } else groups.set(row.paragraph_handle,{paragraph_handle:row.paragraph_handle,ids:[row.amendment_id],before:row.paragraph_before,after:row.paragraph_after,context:candidates[0]!.paragraph})
  }
  // Writer report order follows selection order; normalize to frozen packet order.
  const index = new Map(packet.amendments.map((row,at) => [row.id,at]))
  const ordered = [...groups.values()]
  for (const group of ordered) group.ids.sort((a,b) => index.get(a)! - index.get(b)!)
  ordered.sort((a,b) => index.get(a.ids[0]!)! - index.get(b.ids[0]!)!)
  return {source_sha256:applied.report.source_sha256,groups:ordered}
}

// A reviewer must decide whole paragraph groups. Rebuild the selected subset
// from the same source and compare its verified effects before persistence.
export async function requalifyWordingSurvivors(bytes:Uint8Array,packet:DocxLocalizationPacket,selections:Selection[],original:WordingEffects,ids:string[],options:DocxRevisionOptions = {}) {
  if (new Set(ids).size !== ids.length || ids.some(id => !packet.amendments.some(row => row.id === id)) || original.source_sha256 !== packet.revision.source_sha256)
    throw new Error("wording_survivor_invalid")
  const verified = await wordingEffects(bytes,packet,selections,options)
  if (JSON.stringify(verified) !== JSON.stringify(original)) throw new Error("wording_effect_mapping_invalid")
  const selected = new Set(ids)
  if (original.groups.some(group => group.ids.some(id => selected.has(id)) && !group.ids.every(id => selected.has(id)))) throw new Error("wording_survivor_partial_group")
  const expected = original.groups.filter(group => group.ids.every(id => selected.has(id)))
  if (!ids.length) return {effects:{source_sha256:original.source_sha256,groups:[]},packet:null,selections:[] as Selection[]}
  const amendments = packet.amendments.filter(row => selected.has(row.id))
  const subset = await createDocxLocalization(bytes,amendments,packet.limits,options)
  const survivingSelections = selections.filter(row => selected.has(row.amendment_id))
  const effects = await wordingEffects(bytes,subset,survivingSelections,options)
  if (JSON.stringify(effects.groups) !== JSON.stringify(expected)) throw new Error("wording_survivor_effect_changed")
  return {effects,packet:subset,selections:survivingSelections}
}
