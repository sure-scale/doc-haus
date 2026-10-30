import { expect, spyOn, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync, Zip, ZipPassThrough } from "fflate"
import { renderMarkdownToDocx } from "./render"
import { exportDocxComments, readDocxComments } from "./docx-comments"
import { docxodus } from "./docxodus"
import { paragraphInsertionTargets, authorParagraphInsertion, resolveParagraphInsertion, verifyParagraphInsertion, normalizeParagraphInsertionDecisionSettings } from "./docx-paragraph-insertion"
import { authorParagraphInsertionReview, resolveParagraphInsertionReview, verifyParagraphInsertionReview } from "./docx-paragraph-insertion-review"
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
test("identical target and neighbor text still have distinct source paragraph positions",async()=>{
  const source=(await renderMarkdownToDocx("Before.\n\nSame.\n\nAfter.\n\nBefore.\n\nSame.\n\nAfter.")).bytes
  const targets=(await paragraphInsertionTargets(source)).targets.filter(target=>target.text==="Same.")
  expect(targets.map(target=>[target.previous_paragraph,target.next_paragraph,target.body_paragraph_number])).toEqual([["Before.","After.",2],["Before.","After.",5]])
  expect(targets[0]!.target_handle).not.toBe(targets[1]!.target_handle)
})
async function fixture() {
  const base=(await renderMarkdownToDocx("Reference text.\n\nReference text.\n\nEarlier clause.")).bytes
  const withComment=exportDocxComments(base,[{id:"prior",text:"Review the earlier clause.",author:"Reviewer",anchor:{paragraph_text:"Earlier clause."}}])
  expect(withComment.unresolved).toHaveLength(0)
  const parts=unzipSync(withComment.bytes),doc=new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!),"application/xml")
  const text=Array.from(doc.getElementsByTagNameNS(W,"t")).find(node=>node.textContent==="Earlier clause.")!,run=text.parentNode!,parent=run.parentNode!
  const revision=doc.createElementNS(W,"w:ins");revision.setAttributeNS(W,"w:id","990");revision.setAttributeNS(W,"w:author","Reviewer");revision.setAttributeNS(W,"w:date","2026-01-01T00:00:00Z")
  parent.replaceChild(revision,run);revision.appendChild(run)
  const body=doc.getElementsByTagNameNS(W,"body")[0]!,numbered=new DOMParser().parseFromString(`<w:p xmlns:w="${W}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Numbered sibling.</w:t></w:r></w:p>`,"application/xml")
  body.insertBefore(doc.importNode(numbered.documentElement,true),body.getElementsByTagNameNS(W,"sectPr")[0] ?? null)
  parts["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(doc))
  parts["word/numbering.xml"]=strToU8(`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`)
  const rels=new DOMParser().parseFromString(strFromU8(parts["word/_rels/document.xml.rels"]!),"application/xml"),rel=rels.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships","Relationship")
  rel.setAttribute("Id","numbering-test");rel.setAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");rel.setAttribute("Target","numbering.xml");rels.documentElement!.appendChild(rel);parts["word/_rels/document.xml.rels"]=strToU8(new XMLSerializer().serializeToString(rels))
  const types=new DOMParser().parseFromString(strFromU8(parts["[Content_Types].xml"]!),"application/xml"),override=types.createElementNS("http://schemas.openxmlformats.org/package/2006/content-types","Override")
  override.setAttribute("PartName","/word/numbering.xml");override.setAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");types.documentElement!.appendChild(override);parts["[Content_Types].xml"]=strToU8(new XMLSerializer().serializeToString(types))
  return zipSync(parts)
}
test("one source-bound plain paragraph preserves other comments and revisions through both decisions",async()=>{
  const source=await fixture(),original=new Uint8Array(source),packet=await paragraphInsertionTargets(source), duplicates=packet.targets.filter(t=>t.text==="Reference text.")
  expect(duplicates).toHaveLength(2);expect(duplicates[0]!.target_handle).not.toBe(duplicates[1]!.target_handle)
  const dx=await docxodus(),priorSession=dx.openDocxSession(source,{persistAnchorIds:false}),prior=priorSession.listRevisions();priorSession.close()
  for(const position of ["before","after"] as const){
    const request={source_sha256:packet.source_sha256,target_handle:duplicates[1]!.target_handle,position,text:"Additional term: A * B [C].",author:"Reviewer"}
    const authored=await authorParagraphInsertion(source,request)
    expect(authored.revisions).toHaveLength(1);expect(authored.revisions[0]!.constituentKeys).toHaveLength(2)
    expect((await verifyParagraphInsertion(source,authored.bytes,request)).paragraph.text).toBe(request.text)
    await expect(verifyParagraphInsertion(source,authored.bytes,{...request,target_handle:duplicates[0]!.target_handle})).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,authored.bytes,{...request,position:position==="before"?"after":"before"})).rejects.toThrow()
    const mutate = (change:(doc:Document)=>void) => {
      const parts=unzipSync(authored.bytes),doc=new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!),"application/xml")
      change(doc);parts["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(doc));return zipSync(parts)
    }
    const inserted = (doc:Document) => Array.from(doc.getElementsByTagNameNS(W,"p")).find(p=>Array.from(p.getElementsByTagNameNS(W,"t")).some(t=>t.textContent===request.text))!
    await expect(verifyParagraphInsertion(source,mutate(doc=>{
      const p=inserted(doc),pp=p.getElementsByTagNameNS(W,"pPr")[0]!,style=doc.createElementNS(W,"w:pStyle");style.setAttributeNS(W,"w:val","Heading1");pp.appendChild(style)
    }),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,mutate(doc=>{
      const mark=inserted(doc).getElementsByTagNameNS(W,"ins")[0]!;mark.getAttributeNodeNS(W,"id")!.value="990"
    }),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,mutate(doc=>{
      const p=inserted(doc),t=p.getElementsByTagNameNS(W,"t")[0]!;t.textContent="Unapproved text."
    }),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,mutate(doc=>{const mark=inserted(doc).getElementsByTagNameNS(W,"ins")[0]!;mark.getAttributeNodeNS(W,"author")!.value="Other"}),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,mutate(doc=>{const p=inserted(doc),mark=p.getElementsByTagNameNS(W,"ins")[0]!;p.appendChild(mark)}),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,mutate(doc=>{const p=inserted(doc),extra=doc.createElementNS(W,"w:ins");p.appendChild(extra)}),request)).rejects.toThrow()
    const changedPart=(name:string,change:(doc:Document)=>void)=>{const p=unzipSync(authored.bytes),doc=new DOMParser().parseFromString(strFromU8(p[name]!),"application/xml");change(doc);p[name]=strToU8(new XMLSerializer().serializeToString(doc));return zipSync(p)}
    await expect(verifyParagraphInsertion(source,changedPart("word/comments.xml",doc=>{doc.getElementsByTagNameNS(W,"t")[0]!.textContent="Changed comment."}),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,changedPart("word/settings.xml",doc=>{doc.documentElement!.appendChild(doc.createElementNS(W,"w:evenAndOddHeaders"))}),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,changedPart("word/_rels/document.xml.rels",doc=>{const rel=Array.from(doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships","Relationship")).find(r=>r.getAttribute("Id")==="numbering-test")!;rel.setAttribute("Target","changed-numbering.xml")}),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,mutate(doc=>{const prior=Array.from(doc.getElementsByTagNameNS(W,"ins")).find(n=>n.getAttributeNS(W,"id")==="990")!;prior.getAttributeNodeNS(W,"author")!.value="Other"}),request)).rejects.toThrow()
    await expect(verifyParagraphInsertion(source,mutate(doc=>{const first=doc.getElementsByTagNameNS(W,"p")[0]!;first.setAttributeNS("http://www.w3.org/2000/xmlns/","xmlns:x","urn:changed")}),request)).rejects.toThrow()
    const accepted=await resolveParagraphInsertion(source,authored.bytes,request,"accept"),rejected=await resolveParagraphInsertion(source,authored.bytes,request,"reject")
    expect(rejected).toEqual(source);expect(readDocxComments(accepted).map(c=>[c.text,c.author])).toEqual(readDocxComments(source).map(c=>[c.text,c.author]))
    expect(unzipSync(accepted)["word/numbering.xml"]).toEqual(unzipSync(source)["word/numbering.xml"])
    const reopened=dx.openDocxSession(accepted,{persistAnchorIds:false});try{expect(reopened.listRevisions().map(r=>[r.id,r.author,r.text])).toEqual(prior.map(r=>[r.id,r.author,r.text]))}finally{reopened.close()}
    expect(source).toEqual(original)
  }
  await expect(authorParagraphInsertion(source,{source_sha256:"0".repeat(64),target_handle:duplicates[0]!.target_handle,position:"after",text:"New.",author:"Reviewer"})).rejects.toThrow("stale_source")
})

test("default style ambiguity and document protection refuse before native insertion",async()=>{
  const base=await fixture(), dx=await docxodus()
  for(const mode of ["unused","default_number","default_outline","default_pstyle","ancestor_number","ancestor_missing","docdefaults_outline","missing","duplicate","cycle","protected","writeprotected"] as const){
    const parts=unzipSync(base),styles=new DOMParser().parseFromString(strFromU8(parts["word/styles.xml"]!),"application/xml"),body=new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!),"application/xml")
    const make=(name:string)=>styles.createElementNS(W,`w:${name}`)
    const set=(node:Element,name:string,value:string)=>node.setAttributeNS(W,`w:${name}`,value)
    const normal=Array.from(styles.getElementsByTagNameNS(W,"style")).find(n=>n.getAttributeNS(W,"styleId")==="Normal")!
    const safe=make("style");set(safe,"type","paragraph");set(safe,"styleId","SafeReference");styles.documentElement!.appendChild(safe)
    for(const t of Array.from(body.getElementsByTagNameNS(W,"t")).filter(n=>n.textContent==="Reference text.")){
      const p=t.parentNode!.parentNode as Element,properties=body.createElementNS(W,"w:pPr"),style=body.createElementNS(W,"w:pStyle")
      style.setAttributeNS(W,"w:val","SafeReference");properties.appendChild(style);p.insertBefore(properties,p.firstChild)
    }
    const num=(owner:Element)=>{const pp=make("pPr"),n=make("numPr"),id=make("numId");set(id,"val","1");n.appendChild(id);pp.appendChild(n);owner.appendChild(pp)}
    const outline=(owner:Element)=>{const pp=make("pPr"),n=make("outlineLvl");set(n,"val","0");pp.appendChild(n);owner.appendChild(pp)}
    if(mode==="unused"){const unused=make("style");set(unused,"type","paragraph");set(unused,"styleId","Unused");num(unused);styles.documentElement!.appendChild(unused)}
    if(mode==="default_number")num(normal)
    if(mode==="default_outline")outline(normal)
    if(mode==="default_pstyle"){const pp=make("pPr"),style=make("pStyle");set(style,"val","Heading1");pp.appendChild(style);normal.appendChild(pp)}
    if(mode==="ancestor_missing"){const ref=make("basedOn");set(ref,"val","MissingStyle");normal.appendChild(ref)}
    if(mode==="ancestor_number"){const parent=make("style"),ref=make("basedOn");set(parent,"type","paragraph");set(parent,"styleId","Inherited");num(parent);styles.documentElement!.appendChild(parent);set(ref,"val","Inherited");normal.appendChild(ref)}
    if(mode==="docdefaults_outline"){const def=styles.getElementsByTagNameNS(W,"pPrDefault")[0]!,pp=def.getElementsByTagNameNS(W,"pPr")[0]!,n=make("outlineLvl");set(n,"val","0");pp.appendChild(n)}
    if(mode==="missing")normal.removeAttributeNS(W,"default")
    if(mode==="duplicate")styles.documentElement!.appendChild(normal.cloneNode(true))
    if(mode==="cycle"){const ref=make("basedOn");set(ref,"val","Normal");normal.appendChild(ref)}
    if(mode==="protected"||mode==="writeprotected"){
      const settings=new DOMParser().parseFromString(strFromU8(parts["word/settings.xml"]!),"application/xml")
      settings.documentElement!.appendChild(settings.createElementNS(W,mode==="protected"?"w:documentProtection":"w:writeProtection"))
      parts["word/settings.xml"]=strToU8(new XMLSerializer().serializeToString(settings))
    }
    parts["word/styles.xml"]=strToU8(new XMLSerializer().serializeToString(styles));parts["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(body))
    const source=zipSync(parts),snapshot=new Uint8Array(source),packet=await paragraphInsertionTargets(source),target=packet.targets.find(t=>t.text==="Reference text.")!
    expect(target).toBeTruthy();expect(target.supported).toBe(mode==="unused")
    const request={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"before" as const,text:"New clause.",author:"Reviewer"}
    if(mode==="unused"){expect((await authorParagraphInsertion(source,request)).bytes.length).toBeGreaterThan(0)}
    else {
      const probe=dx.openDocxSession(source,{persistAnchorIds:false}),proto=Object.getPrototypeOf(probe),original=proto.insertParagraph;probe.close();let calls=0
      proto.insertParagraph=function(...args:unknown[]){calls++;return original.apply(this,args)}
      try {await expect(authorParagraphInsertion(source,request)).rejects.toThrow(mode.includes("protected")?"protected_document":"unresolved_default_style");expect(calls).toBe(0)}
      finally{proto.insertParagraph=original}
    }
    expect(source).toEqual(snapshot)
  }
})

test("crossing ranges and nonliteral targets are not exposed as safe insertion boundaries",async()=>{
  const source=await fixture(),parts=unzipSync(source),doc=new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!),"application/xml")
  const body=doc.getElementsByTagNameNS(W,"body")[0]!,paragraphs=Array.from(body.childNodes).filter((n):n is Element=>n.nodeType===1&&(n as Element).namespaceURI===W&&(n as Element).localName==="p")
  const start=doc.createElementNS(W,"w:bookmarkStart"),end=doc.createElementNS(W,"w:bookmarkEnd")
  start.setAttributeNS(W,"w:id","9000");start.setAttributeNS(W,"w:name","Range");end.setAttributeNS(W,"w:id","9000")
  body.insertBefore(start,paragraphs[0]!);body.insertBefore(end,paragraphs[1]!.nextSibling)
  parts["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(doc));const crossed=zipSync(parts),packet=await paragraphInsertionTargets(crossed)
  expect(packet.targets.filter(t=>t.text==="Reference text.")).toHaveLength(0)
  body.removeChild(end);parts["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(doc))
  await expect(paragraphInsertionTargets(zipSync(parts))).rejects.toThrow("unpaired_range")
  const original=unzipSync(source),normal=new DOMParser().parseFromString(strFromU8(original["word/document.xml"]!),"application/xml")
  const t=Array.from(normal.getElementsByTagNameNS(W,"t")).find(node=>node.textContent==="Reference text.")!
  t.parentNode!.appendChild(normal.createElementNS(W,"w:tab"));original["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(normal))
  const unsupported=await paragraphInsertionTargets(zipSync(original))
  expect(unsupported.targets.filter(row=>row.text==="Reference text.")).toHaveLength(1)
  const clean=await paragraphInsertionTargets(source),target=clean.targets.find(row=>row.text==="Reference text.")!
  for(const text of ["One\nTwo","One\rTwo","One\tTwo"])
    await expect(authorParagraphInsertion(source,{source_sha256:clean.source_sha256,target_handle:target.target_handle,position:"before",text,author:"Reviewer"})).rejects.toThrow("unsupported_literal")
})

test("malformed source ranges and direct paragraph text refuse before insertion",async()=>{
  const source=await fixture(),baseline=await paragraphInsertionTargets(source)
  expect(baseline.targets.filter(t=>t.text==="Reference text.")).toHaveLength(2)
  const dx=await docxodus(),session=dx.openDocxSession(source,{persistAnchorIds:false}),proto=Object.getPrototypeOf(session),original=proto.insertParagraph;session.close();let calls=0
  proto.insertParagraph=function(...args:unknown[]){calls++;return original.apply(this,args)}
  try {
    for(const mode of ["nonempty","column","missing_name","duplicate_alias","reversed","malformed_id"] as const){
      const parts=unzipSync(source),doc=new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!),"application/xml")
      const starts=Array.from(doc.getElementsByTagNameNS(W,"bookmarkStart")),ends=Array.from(doc.getElementsByTagNameNS(W,"bookmarkEnd"))
      expect(starts.length).toBeGreaterThanOrEqual(2);const first=starts[0]!,end=ends[0]!
      if(mode==="nonempty")first.appendChild(doc.createElementNS(W,"w:r"))
      if(mode==="column")first.setAttributeNS(W,"w:colFirst","0")
      if(mode==="missing_name")first.removeAttributeNS(W,"name")
      if(mode==="duplicate_alias"){const alias=`00${first.getAttributeNS(W,"id")}`;starts[1]!.getAttributeNodeNS(W,"id")!.value=alias;ends[1]!.getAttributeNodeNS(W,"id")!.value=alias}
      if(mode==="reversed")first.parentNode!.insertBefore(end,first)
      if(mode==="malformed_id")first.getAttributeNodeNS(W,"id")!.value="not-a-number"
      parts["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(doc));const malformed=zipSync(parts),snapshot=new Uint8Array(malformed)
      await expect(paragraphInsertionTargets(malformed)).rejects.toThrow()
      expect(malformed).toEqual(snapshot)
    }
    for(const mode of ["paragraph","run"] as const){
      const parts=unzipSync(source),doc=new DOMParser().parseFromString(strFromU8(parts["word/document.xml"]!),"application/xml")
      const text=Array.from(doc.getElementsByTagNameNS(W,"t")).find(t=>t.textContent==="Reference text.")!,run=text.parentNode!,parent=mode==="run"?run:run.parentNode!
      parent.appendChild(doc.createTextNode("unexpected"));parts["word/document.xml"]=strToU8(new XMLSerializer().serializeToString(doc))
      const malformed=zipSync(parts),packet=await paragraphInsertionTargets(malformed)
      expect(packet.targets.filter(t=>t.text==="Reference text.")).toHaveLength(1)
    }
    expect(calls).toBe(0)
  }finally{proto.insertParagraph=original}
})


test("paragraph insertion reason belongs to the new paragraph through both review decisions",async()=>{
  const source=await fixture(),packet=await paragraphInsertionTargets(source),target=packet.targets.find(t=>t.text==="Reference text.")!
  const request={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"after" as const,text:"A new literal obligation.",author:"Reviewer"},reason="Explain the inserted obligation."
  const artifact=await authorParagraphInsertionReview(source,request,reason)
  expect((await verifyParagraphInsertionReview(source,artifact.bytes,request,reason)).commentId).toBe(artifact.commentId)
  const comments=readDocxComments(artifact.bytes);expect(comments.map(c=>c.text)).toContain(reason)
  const review=comments.find(c=>c.text===reason)!;expect(review.markers).toHaveLength(3)
  expect(review.markers.every(m=>m.accepted_paragraph_text===request.text)).toBe(true)
  expect(await resolveParagraphInsertionReview(source,artifact.bytes,request,reason,"reject")).toEqual(source)
  const accepted=await resolveParagraphInsertionReview(source,artifact.bytes,request,reason,"accept")
  expect(readDocxComments(accepted).map(c=>c.text)).toContain(reason)
  await expect(verifyParagraphInsertionReview(source,artifact.bytes,request,"A different reason.")).rejects.toThrow()
  const parts=unzipSync(artifact.bytes),doc=new DOMParser().parseFromString(strFromU8(parts["word/comments.xml"]!),"application/xml")
  const old=Array.from(doc.getElementsByTagNameNS(W,"t")).find(t=>t.textContent==="Review the earlier clause.")!;old.textContent="Changed old comment."
  parts["word/comments.xml"]=strToU8(new XMLSerializer().serializeToString(doc))
  await expect(verifyParagraphInsertionReview(source,zipSync(parts),request,reason)).rejects.toThrow()
  const relParts=unzipSync(artifact.bytes),rels=new DOMParser().parseFromString(strFromU8(relParts["word/_rels/document.xml.rels"]!),"application/xml")
  const relationship=Array.from(rels.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships","Relationship")).find(row=>row.getAttribute("Id")==="numbering-test")!
  relationship.setAttribute("Target","other-numbering.xml");relParts["word/_rels/document.xml.rels"]=strToU8(new XMLSerializer().serializeToString(rels))
  await expect(verifyParagraphInsertionReview(source,zipSync(relParts),request,reason)).rejects.toThrow()
})


test("reason decisions prove tracking settings before restoration",async()=>{
  for(const enabled of [false,true]) {
    const base=await fixture(),parts=unzipSync(base)
    if(enabled){const doc=new DOMParser().parseFromString(strFromU8(parts["word/settings.xml"]!),"application/xml");doc.documentElement!.appendChild(doc.createElementNS(W,"w:trackRevisions"));parts["word/settings.xml"]=strToU8(new XMLSerializer().serializeToString(doc))}
    const source=zipSync(parts),snapshot=new Uint8Array(source),packet=await paragraphInsertionTargets(source),target=packet.targets.find(t=>t.text==="Reference text.")!
    const request={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"before" as const,text:"Additional duty.",author:"Reviewer"},reason="Explain the duty."
    const core=await authorParagraphInsertion(source,request),review=await authorParagraphInsertionReview(source,request,reason)
    for(const decision of ["accept","reject"] as const)expect((await resolveParagraphInsertionReview(source,review.bytes,request,reason,decision)).length).toBeGreaterThan(0)
    const changed=(fn:(doc:Document)=>void)=>{const p=unzipSync(review.bytes),doc=new DOMParser().parseFromString(strFromU8(p["word/settings.xml"]!),"application/xml");fn(doc);p["word/settings.xml"]=strToU8(new XMLSerializer().serializeToString(doc));return zipSync(p)}
    for(const mode of ["unrelated","disabled","duplicate","namespace"] as const){
      const malformed=changed(doc=>{const mark=doc.getElementsByTagNameNS(W,"trackRevisions")[0]!
        if(mode==="unrelated")doc.documentElement!.appendChild(doc.createElementNS(W,"w:evenAndOddHeaders"))
        if(mode==="disabled")mark.setAttributeNS(W,"w:val","off")
        if(mode==="duplicate")doc.documentElement!.appendChild(mark.cloneNode(true))
        if(mode==="namespace")doc.documentElement!.setAttributeNS("http://www.w3.org/2000/xmlns/","xmlns:x","urn:changed")
      })
      await expect(normalizeParagraphInsertionDecisionSettings(source,core.bytes,malformed,request)).rejects.toThrow()
      await expect(verifyParagraphInsertionReview(source,malformed,request,reason)).rejects.toThrow()
    }
    expect(source).toEqual(snapshot)
  }
})

test("a failed later review step rolls back the owned reason and native insertion group",async()=>{
  const source=await fixture(),packet=await paragraphInsertionTargets(source),target=packet.targets.find(t=>t.text==="Reference text.")!
  const request={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"after" as const,text:"New obligation.",author:"Reviewer"},reason="Explain the obligation."
  const artifact=await authorParagraphInsertionReview(source,request,reason),snapshot=new Uint8Array(artifact.bytes),dx=await docxodus()
  const session=dx.openDocxSession(artifact.bytes,{persistAnchorIds:false}),proto=Object.getPrototypeOf(session) as typeof session;session.close()
  const batch=proto.executeBatch
  for(const decision of ["accept","reject"] as const){
    let changed=false,restored=false
    const spy=spyOn(proto,"executeBatch").mockImplementation(function(this:typeof session,...args:Parameters<typeof batch>){
      if(!this.listComments().some(comment=>String(comment.id)===artifact.commentId))return batch.apply(this,args)
      const before=this.getPackageManifest().rawPackageBytesDigest,version=this.getVersion(),revisions=this.listRevisions().length
      const fail:Parameters<typeof batch>[0][number]={tool:"paragraph",action:"injected_failure",mutation:()=>{
        changed=this.listRevisions().length===revisions-1&&this.listComments().some(comment=>String(comment.id)===artifact.commentId)===(decision==="accept")
        return {success:false,created:[],removed:[],modified:[],error:{code:"internal_error",message:"Synthetic later failure"}}
      }}
      const outcome=batch.call(this,[...args[0],fail],args[1]),after=this.getPackageManifest().rawPackageBytesDigest
      restored=!outcome.success&&this.getVersion()===version&&after.value===before.value&&this.listRevisions().length===revisions&&this.listComments().some(comment=>String(comment.id)===artifact.commentId)
      return outcome
    })
    try{await expect(resolveParagraphInsertionReview(source,artifact.bytes,request,reason,decision)).rejects.toThrow("review_failed_rolled_back");expect(changed).toBe(true);expect(restored).toBe(true)}
    finally{spy.mockRestore()}
  }
  expect(artifact.bytes).toEqual(snapshot)
})

async function sameTextReviewFixture() {
  const source=await fixture(), packet=await paragraphInsertionTargets(source), target=packet.targets.filter(t=>t.text==="Reference text.")[1]!
  const request={source_sha256:packet.source_sha256,target_handle:target.target_handle,position:"before" as const,text:"Reference text.",author:"Reviewer"},reason="Explain this inserted paragraph."
  const artifact=await authorParagraphInsertionReview(source,request,reason)
  expect((await verifyParagraphInsertionReview(source,artifact.bytes,request,reason)).commentId).toBe(artifact.commentId)
  return {source,request,reason,artifact}
}
function alteredReview(bytes:Uint8Array,part:string,change:(doc:Document)=>void) {
  const parts=unzipSync(bytes),doc=new DOMParser().parseFromString(strFromU8(parts[part]!),"application/xml")
  change(doc);parts[part]=strToU8(new XMLSerializer().serializeToString(doc));return zipSync(parts)
}
async function duplicateReviewEntry(bytes:Uint8Array) {
  const parts=unzipSync(bytes)
  return await new Promise<Uint8Array>((resolve,reject)=>{
    const chunks:Uint8Array[]=[],zip=new Zip((error,data,final)=>{
      if(error)reject(error)
      else {chunks.push(data);if(final){const result=new Uint8Array(chunks.reduce((total,chunk)=>total+chunk.length,0));let offset=0;for(const chunk of chunks){result.set(chunk,offset);offset+=chunk.length}resolve(result)}}
    })
    for(const [name,part] of [...Object.entries(parts),["word/document.xml",parts["word/document.xml"]!] as [string,Uint8Array]]){
      const entry=new ZipPassThrough(name);zip.add(entry);entry.push(part,true)
    }
    zip.end()
  })
}
for(const mode of ["identical_paragraph","shift_start","shorten_end"] as const) test(`review reason rejects ${mode} anchor tampering`,async()=>{
  const {source,request,reason,artifact}=await sameTextReviewFixture(),snapshot=new Uint8Array(artifact.bytes), sourceSnapshot=new Uint8Array(source)
  const altered=alteredReview(artifact.bytes,"word/document.xml",doc=>{
    const find=(name:string)=>Array.from(doc.getElementsByTagNameNS(W,name)).filter(n=>n.getAttributeNS(W,"id")===artifact.commentId)
    const starts=find("commentRangeStart"),ends=find("commentRangeEnd"),refs=find("commentReference")
    expect(starts).toHaveLength(1);expect(ends).toHaveLength(1);expect(refs).toHaveLength(1)
    const start=starts[0]!,end=ends[0]!,refRun=refs[0]!.parentNode!,owner=start.parentNode as Element
    if(mode==="identical_paragraph") {
      const targets=Array.from(doc.getElementsByTagNameNS(W,"p")).filter(p=>p!==owner&&Array.from(p.getElementsByTagNameNS(W,"t")).map(t=>t.textContent??"").join("")===request.text)
      expect(targets.length).toBeGreaterThan(0);const target=targets[0]!
      const first=Array.from(target.childNodes).find(n=>n.nodeType===1&&(n as Element).localName!=="pPr")??null
      target.insertBefore(start,first);target.appendChild(end);target.appendChild(refRun)
      expect(start.parentNode).toBe(target);expect(start.parentNode).not.toBe(owner)
    } else {
      const content=Array.from(owner.childNodes).find(n=>n.nodeType===1&&(n as Element).localName==="ins")!
      expect(content).toBeTruthy()
      if(mode==="shift_start")owner.insertBefore(start,content.nextSibling)
      else owner.insertBefore(end,content)
    }
  })
  const alteredSnapshot=new Uint8Array(altered)
  await expect(verifyParagraphInsertionReview(source,altered,request,reason)).rejects.toThrow()
  expect(altered).toEqual(alteredSnapshot);expect(artifact.bytes).toEqual(snapshot);expect(source).toEqual(sourceSnapshot)
})
for(const mode of ["extra_comment","reply","duplicate_id"] as const) test(`review reason rejects ${mode} comment tampering`,async()=>{
  const {source,request,reason,artifact}=await sameTextReviewFixture(),snapshot=new Uint8Array(artifact.bytes)
  const parts=unzipSync(artifact.bytes),doc=new DOMParser().parseFromString(strFromU8(parts["word/comments.xml"]!),"application/xml")
  const owned=Array.from(doc.getElementsByTagNameNS(W,"comment")).find(row=>row.getAttributeNS(W,"id")===artifact.commentId)!
  expect(owned).toBeTruthy();const extra=owned.cloneNode(true) as Element
  if(mode!=="duplicate_id")extra.getAttributeNodeNS(W,"id")!.value="900"
  if(mode==="reply") {
    const W14="http://schemas.microsoft.com/office/word/2010/wordml",W15="http://schemas.microsoft.com/office/word/2012/wordml",REL="http://schemas.openxmlformats.org/package/2006/relationships",CT="http://schemas.openxmlformats.org/package/2006/content-types"
    const parent=owned.getElementsByTagNameNS(W,"p")[0]!,child=extra.getElementsByTagNameNS(W,"p")[0]!
    parent.setAttributeNS(W14,"w14:paraId","ABCDEF01");child.setAttributeNS(W14,"w14:paraId","ABCDEF02")
    parts["word/commentsExtended.xml"]=strToU8(`<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="ABCDEF01"/><w15:commentEx w15:paraId="ABCDEF02" w15:paraIdParent="ABCDEF01"/></w15:commentsEx>`)
    const rels=new DOMParser().parseFromString(strFromU8(parts["word/_rels/document.xml.rels"]!),"application/xml"),rel=rels.createElementNS(REL,"Relationship")
    rel.setAttribute("Id","reply-test");rel.setAttribute("Type","http://schemas.microsoft.com/office/2011/relationships/commentsExtended");rel.setAttribute("Target","commentsExtended.xml");rels.documentElement!.appendChild(rel)
    parts["word/_rels/document.xml.rels"]=strToU8(new XMLSerializer().serializeToString(rels))
    const types=new DOMParser().parseFromString(strFromU8(parts["[Content_Types].xml"]!),"application/xml"),override=types.createElementNS(CT,"Override")
    override.setAttribute("PartName","/word/commentsExtended.xml");override.setAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml");types.documentElement!.appendChild(override)
    parts["[Content_Types].xml"]=strToU8(new XMLSerializer().serializeToString(types))
  }
  doc.documentElement!.appendChild(extra);parts["word/comments.xml"]=strToU8(new XMLSerializer().serializeToString(doc))
  const altered=zipSync(parts),alteredSnapshot=new Uint8Array(altered)
  await expect(verifyParagraphInsertionReview(source,altered,request,reason)).rejects.toThrow()
  expect(altered).toEqual(alteredSnapshot);expect(artifact.bytes).toEqual(snapshot)
})
test("duplicate ZIP entries refuse at paragraph review entry before native decisions",async()=>{
  const {source,request,reason,artifact}=await sameTextReviewFixture(),badSource=await duplicateReviewEntry(source),badArtifact=await duplicateReviewEntry(artifact.bytes)
  const dx=await docxodus(),session=dx.openDocxSession(artifact.bytes,{persistAnchorIds:false}),proto=Object.getPrototypeOf(session) as typeof session;session.close()
  const accept=spyOn(proto,"acceptRevision"),reject=spyOn(proto,"rejectRevision"),insert=spyOn(proto,"insertParagraph")
  try {
    await expect(verifyParagraphInsertionReview(source,badArtifact,request,reason)).rejects.toThrow("docx_exact_spacing:duplicate_part")
    await expect(resolveParagraphInsertionReview(source,badArtifact,request,reason,"reject")).rejects.toThrow("docx_exact_spacing:duplicate_part")
    await expect(verifyParagraphInsertionReview(badSource,artifact.bytes,request,reason)).rejects.toThrow("docx_exact_spacing:duplicate_part")
    expect(accept).not.toHaveBeenCalled();expect(reject).not.toHaveBeenCalled();expect(insert).not.toHaveBeenCalled()
  }finally{accept.mockRestore();reject.mockRestore();insert.mockRestore()}
})
