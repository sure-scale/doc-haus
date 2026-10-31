import { expect, spyOn, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxComments } from "./docx-comments"
import { authorDocxTableRowDeletion, docxTableRowDeletionTargets, resolveDocxTableRowDeletion, type DocxTableRowDeletion } from "./docx-table-row-deletion"
import { authorDocxTableRowDeletionReview, resolveDocxTableRowDeletionReview } from "./docx-table-row-deletion-review"

const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main", O="http://schemas.openxmlformats.org/officeDocument/2006/relationships", R="http://schemas.openxmlformats.org/package/2006/relationships", CT="http://schemas.openxmlformats.org/package/2006/content-types"
const BODY="word/document.xml", COMMENTS="word/comments.xml", RELS="word/_rels/document.xml.rels", TYPES="[Content_Types].xml", AUTHOR="Review Editor"
const REASON="Explains the requested row deletion.\nLiteral **note** & <review> 😀."
const escape=(value:string)=>value.replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
const p=(value:string)=>`<w:p><w:r><w:t xml:space="preserve">${escape(value)}</w:t></w:r></w:p>`
const cell=(value:string)=>`<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p(value)}</w:tc>`
const row=(a:string,b:string)=>`<w:tr>${cell(a)}${cell(b)}</w:tr>`
const parse=(value:Uint8Array)=>new DOMParser().parseFromString(strFromU8(value),"application/xml")
const xml=(node:Node)=>new XMLSerializer().serializeToString(node)
function fixture(prior=false,deleted="Same") {
  const stamp=`w:author="${AUTHOR}" w:date="2026-01-01T00:00:00Z"`
  const before=prior?`<w:p><w:del w:id="70" ${stamp}><w:r><w:delText xml:space="preserve">${escape(deleted)}</w:delText></w:r></w:del></w:p>`:p("Before")
  const after=prior?`<w:p><w:ins w:id="71" ${stamp}><w:r><w:t>After</w:t></w:r></w:ins></w:p>`:p("After")
  let bytes=zipSync({
    [TYPES]:strToU8(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${BODY}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`),
    "_rels/.rels":strToU8(`<Relationships xmlns="${R}"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    [RELS]:strToU8(`<Relationships xmlns="${R}"><Relationship Id="styles" Type="${O}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${O}/settings" Target="settings.xml"/></Relationships>`),
    "word/styles.xml":strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/settings.xml":strToU8(`<w:settings xmlns:w="${W}"/>`),
    [BODY]:strToU8(`<w:document xmlns:w="${W}" xmlns:r="${O}"><w:body>${before}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${row("Reference","Current")}${row("",deleted)}${row("Last","Stable")}</w:tbl>${after}<w:sectPr/></w:body></w:document>`),
  },{mtime:new Date(1980,0,1)})
  if(!prior)return bytes
  bytes=exportDocxComments(bytes,[{id:"old",text:"Keep original note.",author:"Prior Reviewer",anchor:{paragraph_text:"After"}}]).bytes
  const parts=unzipSync(bytes),comments=parse(parts[COMMENTS]!),body=parse(parts[BODY]!),rels=parse(parts[RELS]!),types=parse(parts[TYPES]!)
  const w14="http://schemas.microsoft.com/office/word/2010/wordml",w15="http://schemas.microsoft.com/office/word/2012/wordml"
  comments.documentElement!.setAttribute("xmlns:w14",w14);comments.documentElement!.setAttribute("xmlns:r",O)
  const first=comments.getElementsByTagNameNS(W,"p")[0]!;first.setAttributeNS(w14,"w14:paraId","ABCDEF01")
  const link=comments.createElementNS(W,"w:hyperlink"),run=comments.createElementNS(W,"w:r"),text=comments.createElementNS(W,"w:t")
  link.setAttributeNS(O,"r:id","cLink");text.textContent=" Link";run.appendChild(text);link.appendChild(run);first.appendChild(link)
  comments.documentElement!.appendChild(comments.importNode(parse(strToU8(`<w:comment xmlns:w="${W}" xmlns:w14="${w14}" w:id="1" w:author="Prior Reviewer"><w:p w14:paraId="ABCDEF02"><w:r><w:t>Keep original reply.</w:t></w:r></w:p></w:comment>`)).documentElement!,true))
  const host=Array.from(body.getElementsByTagNameNS(W,"p")).find(p=>p.textContent==="After")!,refRun=body.createElementNS(W,"w:r"),ref=body.createElementNS(W,"w:commentReference")
  ref.setAttributeNS(W,"w:id","1");refRun.appendChild(ref);host.appendChild(refRun)
  const rel=rels.createElementNS(R,"Relationship");rel.setAttribute("Id","thread");rel.setAttribute("Type","http://schemas.microsoft.com/office/2011/relationships/commentsExtended");rel.setAttribute("Target","commentsExtended.xml");rels.documentElement!.appendChild(rel)
  const override=types.createElementNS(CT,"Override");override.setAttribute("PartName","/word/commentsExtended.xml");override.setAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml");types.documentElement!.appendChild(override)
  parts["word/commentsExtended.xml"]=strToU8(`<w15:commentsEx xmlns:w15="${w15}"><w15:commentEx w15:paraId="ABCDEF01" w15:done="0"/><w15:commentEx w15:paraId="ABCDEF02" w15:paraIdParent="ABCDEF01" w15:done="0"/></w15:commentsEx>`)
  parts["word/_rels/comments.xml.rels"]=strToU8(`<Relationships xmlns="${R}"><Relationship Id="cLink" Type="${O}/hyperlink" Target="https://example.com/review" TargetMode="External"/></Relationships>`)
  for(const [name,doc] of [[COMMENTS,comments],[BODY,body],[RELS,rels],[TYPES,types]] as const)parts[name]=strToU8(xml(doc))
  return zipSync(parts)
}
async function input(bytes:Uint8Array):Promise<DocxTableRowDeletion>{const offered=await docxTableRowDeletionTargets(bytes);return{source_sha256:offered.source_sha256,target_handle:offered.targets.find(row=>row.row===2)!.target_handle,author:AUTHOR}}
function changed(bytes:Uint8Array,part:string,change:(doc:Document)=>void){const parts=unzipSync(bytes),doc=parse(parts[part]!);change(doc);parts[part]=strToU8(xml(doc));return zipSync(parts)}
async function records(bytes:Uint8Array){const s=(await docxodus()).openDocxSession(bytes,{persistAnchorIds:false});try{return s.listRevisions().map(({anchorId:_anchor,affectedAnchors:_affected,...r})=>r)}finally{s.close()}}
function physicalComments(bytes:Uint8Array,ids:Set<string>){
  const parts=unzipSync(bytes),doc=parse(parts[BODY]!),definitions=parse(parts[COMMENTS]!)
  const markers=["commentRangeStart","commentRangeEnd","commentReference"].flatMap(kind=>Array.from(doc.getElementsByTagNameNS(W,kind)).filter(marker=>ids.has(marker.getAttributeNS(W,"id")!)).map(marker=>{
    const path:number[]=[];let node:Node=marker,paragraph:Element|undefined
    while(node.parentNode&&node.parentNode.nodeType!==9){if(node.nodeType===1&&(node as Element).namespaceURI===W&&(node as Element).localName==="p")paragraph=node as Element;path.unshift(Array.from(node.parentNode.childNodes).indexOf(node as ChildNode));node=node.parentNode}
    return{kind,id:marker.getAttributeNS(W,"id"),path,subtree:xml(marker),host_text:paragraph?.textContent,host_xml:paragraph?xml(paragraph):undefined}
  }))
  return{definitions:Array.from(definitions.getElementsByTagNameNS(W,"comment")).filter(comment=>ids.has(comment.getAttributeNS(W,"id")!)).map(xml),markers}
}

test("final deleted-row comment uses exact physical paragraph, clean core acceptance and byte-exact rejection",async()=>{
  const source=fixture(),original=new Uint8Array(source),request=await input(source),result=await authorDocxTableRowDeletionReview(source,request,REASON)
  const comments=readDocxComments(result.bytes);expect(comments).toHaveLength(1);expect(comments[0]!.text).toBe(REASON)
  expect(comments[0]!.markers.map(m=>[m.kind,m.xml_paragraph_index,m.accepted_paragraph_text,m.deleted_paragraph_text])).toEqual(["commentRangeStart","commentRangeEnd","commentReference"].map(kind=>[kind,4,"","Same"]))
  const core=await authorDocxTableRowDeletion(source,request),expected=await resolveDocxTableRowDeletion(source,core.bytes,request,"accept")
  expect(await resolveDocxTableRowDeletionReview(source,result.bytes,request,REASON,"accept")).toEqual(expected)
  expect(readDocxComments(expected)).toEqual([]);expect(await records(expected)).toEqual([])
  expect(await resolveDocxTableRowDeletionReview(source,result.bytes,request,REASON,"reject")).toEqual(source)
  expect(source).toEqual(original)
})

test("duplicate deleted words and exact whitespace select the reviewed cell and preserve prior threads and native revisions",async()=>{
  for(const deleted of ["Same","  "]){
    const source=fixture(true,deleted),request={...await input(source),author:` ${AUTHOR} `},result=await authorDocxTableRowDeletionReview(source,request,REASON)
    const old=readDocxComments(source),comments=readDocxComments(result.bytes),own=comments.find(c=>c.comment_id===result.commentId)!
    // Pending deletion changes accepted-view ordinals. Its physical paragraph,
    // marker locations, payloads and all other reader evidence remain exact.
    const physicalReader=(values:ReturnType<typeof readDocxComments>)=>values.map(comment=>({...comment,markers:comment.markers.map(({paragraph_index:_acceptedOrdinal,...marker})=>marker)}))
    expect(physicalReader(comments.slice(0,2))).toEqual(physicalReader(old))
    const oldIds=new Set(old.map(comment=>comment.comment_id))
    expect(physicalComments(result.bytes,oldIds)).toEqual(physicalComments(source,oldIds))
    expect(own.markers.every(m=>m.xml_paragraph_index===4&&m.deleted_paragraph_text===deleted&&m.accepted_paragraph_text==="")).toBe(true)
    const raw=parse(unzipSync(result.bytes)[COMMENTS]!);expect(Array.from(raw.getElementsByTagNameNS(W,"comment")).find(c=>c.getAttributeNS(W,"id")===result.commentId)!.getAttributeNS(W,"author")).toBe(AUTHOR)
    const accepted=await resolveDocxTableRowDeletionReview(source,result.bytes,request,REASON,"accept")
    expect((await records(accepted))).toEqual(await records(source));expect(readDocxComments(accepted).map(c=>c.text)).toEqual(old.map(c=>c.text))
    for(const name of ["word/commentsExtended.xml","word/_rels/comments.xml.rels",RELS,TYPES]){expect(unzipSync(result.bytes)[name]).toEqual(unzipSync(source)[name]);expect(unzipSync(accepted)[name]).toEqual(unzipSync(source)[name])}
    expect(await resolveDocxTableRowDeletionReview(source,result.bytes,request,REASON,"reject")).toEqual(source)
  }
})

test("foreign comments, IDs, moved/duplicated markers and changed dependencies refuse the final reviewed artifact",async()=>{
  const source=fixture(true),request=await input(source),result=await authorDocxTableRowDeletionReview(source,request,REASON)
  const comment=changed(result.bytes,COMMENTS,doc=>Array.from(doc.getElementsByTagNameNS(W,"comment")).find(c=>c.getAttributeNS(W,"id")===result.commentId)!.setAttributeNS(W,"w:author","Other"))
  const duplicate=changed(result.bytes,BODY,doc=>{const ref=Array.from(doc.getElementsByTagNameNS(W,"commentReference")).find(c=>c.getAttributeNS(W,"id")===result.commentId)!;ref.parentNode!.appendChild(ref.cloneNode(true))})
  const moved=changed(result.bytes,BODY,doc=>{for(const name of ["commentRangeStart","commentRangeEnd","commentReference"]){const marker=Array.from(doc.getElementsByTagNameNS(W,name)).find(c=>c.getAttributeNS(W,"id")===result.commentId)!;doc.getElementsByTagNameNS(W,"p")[0]!.appendChild(name==="commentReference"?marker.parentNode!:marker)}})
  const id=changed(result.bytes,COMMENTS,doc=>Array.from(doc.getElementsByTagNameNS(W,"comment")).find(c=>c.getAttributeNS(W,"id")===result.commentId)!.setAttributeNS(W,"w:id","0"))
  const rel=changed(result.bytes,"word/_rels/comments.xml.rels",doc=>doc.getElementsByTagNameNS(R,"Relationship")[0]!.setAttribute("Target","https://example.com/changed"))
  const type=changed(result.bytes,TYPES,doc=>Array.from(doc.getElementsByTagNameNS(CT,"Override")).find(c=>c.getAttribute("PartName")==="/word/comments.xml")!.setAttribute("ContentType","invalid"))
  const movedOld=changed(result.bytes,BODY,doc=>{const marker=Array.from(doc.getElementsByTagNameNS(W,"commentRangeStart")).find(c=>c.getAttributeNS(W,"id")==="0")!,end=Array.from(doc.getElementsByTagNameNS(W,"commentRangeEnd")).find(c=>c.getAttributeNS(W,"id")==="0")!;marker.parentNode!.insertBefore(marker,end)})
  const oldAuthor=changed(result.bytes,COMMENTS,doc=>Array.from(doc.getElementsByTagNameNS(W,"comment")).find(c=>c.getAttributeNS(W,"id")==="0")!.setAttributeNS(W,"w:author","Changed Prior Reviewer"))
  expect(readDocxComments(movedOld).slice(0,2)).toEqual(readDocxComments(result.bytes).slice(0,2))
  expect(physicalComments(movedOld,new Set(["0","1"]))).not.toEqual(physicalComments(result.bytes,new Set(["0","1"])))
  for(const bytes of [comment,duplicate,moved,id,rel,type,movedOld,oldAuthor])await expect(resolveDocxTableRowDeletionReview(source,bytes,request,REASON,"accept")).rejects.toThrow()
  await expect(resolveDocxTableRowDeletionReview(source,result.bytes,request,"A different reason.","accept")).rejects.toThrow("comment_content_or_prior_comments_changed")
  expect(await resolveDocxTableRowDeletionReview(source,result.bytes,request,REASON,"reject")).toEqual(source)
})

test("invalid explanation refuses before row mutation",async()=>{
  const source=fixture(),request=await input(source),s=(await docxodus()).openDocxSession(source,{}),prototype=Object.getPrototypeOf(s) as typeof s;s.close();const deletion=spyOn(prototype,"deleteTableRow")
  try{for(const reason of [""," ","x".repeat(4001),"Bad\u0000reason"])await expect(authorDocxTableRowDeletionReview(source,request,reason)).rejects.toThrow("invalid_reason");expect(deletion).not.toHaveBeenCalled()}finally{deletion.mockRestore()}
})

test("both decisions atomically restore the owned reason and row when native decision fails after mutation",async()=>{
  const source=fixture(true),request=await input(source),result=await authorDocxTableRowDeletionReview(source,request,REASON),s=(await docxodus()).openDocxSession(source,{}),prototype=Object.getPrototypeOf(s) as typeof s;s.close()
  for(const decision of ["accept","reject"] as const){
    const method=decision==="accept"?"acceptRevision":"rejectRevision",original=prototype[method],remove=prototype.removeComment,batch=prototype.executeBatch,active=new WeakSet<object>();let removed=0,mutated=false,restored=false
    const removal=spyOn(prototype,"removeComment").mockImplementation(function(this:typeof s,...args:Parameters<typeof remove>){const result=remove.apply(this,args);if(result.success){removed++;active.add(this)}return result})
    const failure=spyOn(prototype,method).mockImplementation(function(this:typeof s,id:string){const result=original.call(this,id);if(!active.has(this))return result;mutated=result.success;return{...result,success:false,error:{code:"internal_error",message:"Synthetic post-mutation failure"}}})
    const observation=spyOn(prototype,"executeBatch").mockImplementation(function(this:typeof s,...args:Parameters<typeof batch>){const before=this.getPackageManifest().rawPackageBytesDigest,version=this.getVersion(),result=batch.apply(this,args),after=this.getPackageManifest().rawPackageBytesDigest;if(active.has(this)&&!result.success)restored=before.algorithm===after.algorithm&&before.value===after.value&&version===this.getVersion();return result})
    try{await expect(resolveDocxTableRowDeletionReview(source,result.bytes,request,REASON,decision)).rejects.toThrow("review_failed_rolled_back");expect(removed).toBe(1);expect(mutated).toBe(true);expect(restored).toBe(true)}finally{removal.mockRestore();failure.mockRestore();observation.mockRestore()}
  }
})
