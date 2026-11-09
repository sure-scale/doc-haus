import {expect,test} from "bun:test"
import {createHash} from "node:crypto"
import {strToU8,zipSync} from "fflate"
import {WordChatHost,type WordChatWord, type WordChatHostError} from "./word-chat-host"

const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main",P="http://schemas.microsoft.com/office/2006/xmlPackage",R="http://schemas.openxmlformats.org/package/2006/relationships",O="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const MAIN="word/document.xml", SETTINGS="word/settings.xml"
const para=(text:string)=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const cell=(content:string)=>`<w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr>${content}</w:tc>`
const row=(a:string,b:string)=>`<w:tr>${cell(para(a))}${cell(para(b))}</w:tr>`
const retained=row("Label","Value"),removed=row("Remove","Entry"),section='<w:sectPr><w:pgSz w:w="12000" w:h="16000"/></w:sectPr>'
const mark=(id:number)=>`w:id="${id}" w:author="doc.haus" w:date="2026-03-01T00:00:00Z"`
const deletedCell=(text:string,id:number)=>cell(`<w:p><w:pPr><w:rPr><w:del ${mark(id)}/></w:rPr></w:pPr><w:del ${mark(id+1)}><w:r><w:delText>${text}</w:delText></w:r></w:del></w:p>`)
const deleted=`<w:tr><w:trPr><w:del ${mark(1)}/></w:trPr>${deletedCell("Remove",2)}${deletedCell("Entry",4)}</w:tr>`
const table=(tracked=false)=>`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>${retained}${tracked?deleted:removed}</w:tbl>`
const main=(content:string)=>`<w:document xmlns:w="${W}"><w:body>${content}</w:body></w:document>`
const fullMain=(tracked=false)=>main(para("Before")+table(tracked)+para("After")+section)
const relation=(id:string,kind:string,target:string)=>`<Relationship Id="${id}" Type="${O}/${kind}" Target="${target}"/>`
const rels=(content:string)=>`<Relationships xmlns="${R}">${content}</Relationships>`
const rootrels=rels(relation("main","officeDocument",MAIN))
const styles=`<w:styles xmlns:w="${W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:profile="urn:font-profile:before" mc:Ignorable="profile"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`
const types:Record<string,string>={[MAIN]:"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",[SETTINGS]:"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml","word/styles.xml":"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"}
const ctype=(name:string)=>name.endsWith(".rels")?"application/vnd.openxmlformats-package.relationships+xml":types[name]??"application/xml"
const base:Record<string,string>={"_rels/.rels":rootrels,"word/_rels/document.xml.rels":rels(relation("styles","styles","styles.xml")+relation("settings","settings","settings.xml")),"word/styles.xml":styles,[SETTINGS]:`<w:settings xmlns:w="${W}"/>`,"[Content_Types].xml":`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${Object.entries(types).map(([name,type])=>`<Override PartName="/${name}" ContentType="${type}"/>`).join("")}</Types>`}
const flat=(parts:Record<string,string>)=>`<pkg:package xmlns:pkg="${P}">${Object.entries(parts).map(([name,value])=>`<pkg:part pkg:name="/${name}" pkg:contentType="${ctype(name)}"><pkg:xmlData>${value}</pkg:xmlData></pkg:part>`).join("")}</pkg:package>`
const scoped=(tracked=false)=>flat({[MAIN]:main(table(tracked)),"_rels/.rels":rootrels,"word/_rels/document.xml.rels":rels(relation("styles","styles","styles.xml")),"word/styles.xml":styles})
const pack=(tracked=false)=>zipSync(Object.fromEntries(Object.entries({...base,[MAIN]:fullMain(tracked),[SETTINGS]:`<w:settings xmlns:w="${W}">${tracked?"<w:trackRevisions/>":""}</w:settings>`}).map(([name,value])=>[name,strToU8(value)])))
const source=pack(),authored=pack(true),hash=createHash("sha256").update(source).digest("hex")
const texts=["Before","Label","Value","Remove","Entry","After"],ids=texts.map((_,i)=>`original-${i}`)
const proposal={kind:"table_row_delete" as const,id:"review-parent",documentId:"document",explanation:"Delete row",openItems:[],createdAt:"now",table:{tableIndex:0,rows:[texts.slice(1,3),texts.slice(3,5)].map((row,ri)=>row.map((text,ci)=>({id:ids[1+ri*2+ci]!,text})))},rowIndex:1,cells:texts.slice(3,5)}
const oracle=()=>({kind:"table_row_delete" as const,proposalId:proposal.id,sourceSha256:hash,table:proposal.table,rowIndex:1,cells:proposal.cells,author:"doc.haus" as const,writerRequest:{source_sha256:hash,target_handle:"source-bound",author:"doc.haus"},authoredBase64:Buffer.from(authored).toString("base64"),revisions:[{id:"row-delete",family:"row_delete"}],tableImport:{beforeOoxml:scoped(),ooxml:scoped(true)}})

/** An Office batch model: replacement runs on sync; Table proxies are invalidated
 * by replacement. These deterministic mocks do not establish native acceptance. */
function office(options:{failure?:string;tracking?:string;formatting?:boolean}={}) {
 const state={applied:false,queued:false,failed:false,writes:[] as Array<{xml:string;tracking:string;location:string}>,modes:[] as string[],formats:[] as boolean[],authorCalls:0,runs:0,scopedReads:0,rangeLocations:[] as string[]}
 let tracking=options.tracking??"TrackAll",formatting=options.formatting??true
 const failure=options.failure??"none",collection=<T>(items:()=>T[])=>({get items(){return items()},load(_props:string){}})
 const current=(i:number)=>state.applied&&(i===3||i===4)?"":texts[i]!
 const localId=(i:number)=>state.applied&&((i>0&&i<5)||(failure==="outside-id"&&i===0))?`renewed-${i}`:ids[i]!
 const tableXml=()=>{
  let value=table(state.applied)
  if(failure==="stale-table"&&!state.applied||failure==="post-geometry"&&state.applied)value=value.replace('w:w="1800"','w:w="2100"')
  if(failure==="late-table"&&state.scopedReads>=3&&!state.applied)value=value.replace('w:w="1800"','w:w="2100"')
  return value
 }
 const full=()=>flat({...base,[MAIN]:main(para("Before")+tableXml()+para("After")+section),"word/styles.xml":failure==="stale-style"?styles.replace('w:val="Normal"','w:val="Changed"'):failure==="post-style-binding"&&state.applied?styles.replace("urn:font-profile:before","urn:font-profile:changed"):styles})
 const scopedXml=()=>{
  state.scopedReads++
  const extra=failure==="leak-paragraph"?para("Before"):failure==="leak-table"?table():failure==="leak-section"?section.replace('w:w="12000"','w:w="9999"'):section
  return flat({...base,[MAIN]:main(tableXml()+extra)})
 }
 const paragraphs=texts.map((text,i)=>({get uniqueLocalId(){return localId(i)},get text(){return current(i)},tableNestingLevel:i>0&&i<5?1:0,styleBuiltIn:"Normal",spaceBefore:0,spaceAfter:0,lineSpacing:12,load(_properties:string){},getRange(_location?:string){return paragraphRange(i,text)}}))
 const paragraphRange=(i:number,text:string)=>({kind:"paragraph",index:i,getReviewedText(view:string){return {value:view==="Original"?text:current(i)}},getOoxml(){return {value:flat({...base,[MAIN]:main(state.applied&&(i===3||i===4)?deletedCell(text,i===3?2:4).replace(/^<w:tc><w:tcPr>.*?<\/w:tcPr>/,"").replace(/<\/w:tc>$/,""):para(text))})}}})
 const range={paragraphs:collection(()=>failure==="partial-ids"?paragraphs.slice(1,4):failure==="leak-ids"?paragraphs.slice(0,5):paragraphs.slice(1,5)),getOoxml(){return {value:scopedXml()}},compareLocationWith(other:{kind?:string;index?:number}){return {value:failure==="overlap"&&other.kind==="paragraph"?"Overlap":other.kind==="paragraph"?(other.index===0?"AdjacentAfter":"AdjacentBefore"):"Inside"}},insertOoxml(xml:string,location:string){state.writes.push({xml,location,tracking});if(failure==="insert-throw")throw Error("PRIVATE native insert");state.queued=true;return range}}
 const nativeTable=(createdAfter:boolean)=>({rowCount:2,isUniform:true,get values(){return [[current(1),current(2)],[current(3),current(4)]]},rows:collection(()=>[0,1].map(ri=>({rowIndex:ri,cellCount:2,isHeader:false,values:[[current(1+ri*2),current(2+ri*2)]],delete(){throw Error("Legacy deletion must not dispatch")}}))),load(_properties:string){},getRange(location:"Whole"){state.rangeLocations.push(location);return range},getCell(ri:number,ci:number){if(state.applied&&!createdAfter)throw Error("Stale table proxy");const i=1+ri*2+ci;return {body:{paragraphs:collection(()=>[paragraphs[i]!]),getRange(){return paragraphRange(i,texts[i]!)}}}}})
 const bodyRange={kind:"body",getOoxml(){return {value:full()}},getReviewedText(view:string){return {value:(view==="Original"?texts:texts.map((_,i)=>current(i))).join("\r")}},getTrackedChanges(){return collection(()=>[])},getComments(){return collection(()=>state.applied&&failure==="post-comments"?[{author:"Reviewer",content:"Changed",resolved:false}]:[])}}
 const document={body:{paragraphs:collection(()=>paragraphs),tables:collection(()=>[nativeTable(state.applied)]),getRange(){return bodyRange}},get changeTrackingMode(){return tracking},set changeTrackingMode(value:string){state.modes.push(value);tracking=value},get trackFormatting(){return formatting},set trackFormatting(value:boolean){state.formats.push(value);formatting=value},load(_props:string){},saved:false}
 const word={async run(fn:(context:unknown)=>Promise<unknown>){state.runs++;return fn({document,async sync(){if(state.queued){state.queued=false;state.applied=true;if(failure==="restore-ignored")formatting=!formatting;if(failure==="lost-sync"&&!state.failed){state.failed=true;throw Error("PRIVATE lost sync")}}}})}} as WordChatWord
 const host=new WordChatHost(word,()=>true)
 const apply=(payload=oracle(),mayDispatch?:()=>boolean)=>host.applyRowDeletion(proposal,async()=>source,async()=>{state.authorCalls++;return payload},mayDispatch)
 return {state,document,apply}
}

test("verified parent import replaces one Whole table with tracking Off, restores settings and permits only inside ID renewal",async()=>{
 for(const tracking of ["Off","TrackAll","TrackMineOnly"])for(const formatting of [false,true]){
  const fixture=office({tracking,formatting})
  expect((await fixture.apply()).status).toBe("applied")
  expect((await fixture.apply()).status).toBe("already-applied")
  expect(fixture.state.writes).toEqual([{xml:scoped(true),location:"Replace",tracking:"Off"}])
  expect(fixture.state.rangeLocations).toEqual(["Whole","Whole"])
  expect(fixture.document.changeTrackingMode).toBe(tracking);expect(fixture.document.trackFormatting).toBe(formatting)
  expect(fixture.state.modes).toEqual(["Off",tracking]);expect(fixture.state.authorCalls).toBe(1)
 }
})

test("forged, oversized or mismatched parent payloads refuse before opening Word",async()=>{
 for(const mutate of [(v:ReturnType<typeof oracle>)=>{v.tableImport.ooxml=v.tableImport.ooxml.replace("Remove","Forged")},(v:ReturnType<typeof oracle>)=>{v.tableImport.ooxml=v.tableImport.ooxml.replace('w:val="Normal"','w:val="Changed"')},(v:ReturnType<typeof oracle>)=>{v.tableImport.ooxml="x".repeat(1024*1024+1)},(v:ReturnType<typeof oracle>)=>{v.proposalId="different"}]){
  const fixture=office(),payload=oracle();mutate(payload)
  await expect(fixture.apply(payload)).rejects.toMatchObject({mutationState:"unchanged"})
  expect(fixture.state.runs).toBe(0);expect(fixture.state.writes).toHaveLength(0)
 }
})

test("only the pair strings verified before Office reads can reach dispatch",async()=>{
 const fixture=office(),payload=oracle()
 expect((await fixture.apply(payload,()=>{payload.tableImport.ooxml="unverified replacement";return true})).status).toBe("applied")
 expect(fixture.state.writes[0]!.xml).toBe(scoped(true))
})

test("stale geometry/defaults, partial or leaking ranges and overlapping boundaries refuse without mutation",async()=>{
 for(const failure of ["stale-table","stale-style","late-table","partial-ids","leak-ids","leak-paragraph","leak-table","leak-section","overlap"]){
  const fixture=office({failure})
  await expect(fixture.apply()).rejects.toMatchObject({mutationState:"unchanged"})
  expect(fixture.state.writes).toHaveLength(0);expect(fixture.state.modes).toEqual([])
 }
 const lease=office()
 await expect(lease.apply(oracle(),()=>false)).rejects.toMatchObject({mutationState:"unchanged"})
 expect(lease.state.writes).toHaveLength(0);expect(lease.state.modes).toEqual([])
})

test("dispatch or post-proof failure restores both settings and fences an uncertain import against replay",async()=>{
 for(const failure of ["insert-throw","lost-sync","outside-id","post-geometry","post-comments","post-style-binding","restore-ignored"]){
  const fixture=office({failure,tracking:"TrackMineOnly",formatting:false})
  let error:WordChatHostError|undefined
  try{await fixture.apply()}catch(value){error=value as WordChatHostError}
  expect(error?.mutationState).toBe("uncertain")
  expect(error?.message).not.toContain("PRIVATE")
  expect(fixture.document.changeTrackingMode).toBe("TrackMineOnly");expect(fixture.document.trackFormatting).toBe(false)
  await expect(fixture.apply()).rejects.toMatchObject({mutationState:"uncertain"})
  expect(fixture.state.authorCalls).toBe(1);expect(fixture.state.writes).toHaveLength(1)
 }
})
