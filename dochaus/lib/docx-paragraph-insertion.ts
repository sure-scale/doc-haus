import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { docxodus } from "./docxodus"
import { openRevision } from "./docx-amendments"
import { assertExactSpacingPackage } from "./docx-exact-spacing"

// Offline full-package proof only; no route or Word transport uses this helper.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", BODY = "word/document.xml", SETTINGS = "word/settings.xml"
const hash = (v: Uint8Array | string) => createHash("sha256").update(v).digest("hex")
const schema = z.object({ source_sha256: z.string().regex(/^[a-f0-9]{64}$/), target_handle: z.string().min(1), position: z.enum(["before", "after"]),
  text: z.string().min(1).max(4000), author: z.string().trim().min(1).max(100) }).strict()
export type DocxParagraphInsertion = z.infer<typeof schema>
type Opened = Awaited<ReturnType<typeof openRevision>>
type Revision = ReturnType<Opened["session"]["listRevisions"]>[number]
function require(value: unknown, reason: string): asserts value { if (!value) throw new Error(`docx_paragraph_insertion:${reason}`) }
const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
const kids = (node: Node) => Array.from(node.childNodes).filter((n): n is Element => n.nodeType === 1)
const child = (node: Node, name: string) => kids(node).filter(n => n.namespaceURI === W && n.localName === name)
const xml = (node: Node) => new XMLSerializer().serializeToString(node)
function parse(text: string) {
  require(!/<!DOCTYPE|<!ENTITY/i.test(text), "unsupported_xml")
  const fail = () => { throw new Error("docx_paragraph_insertion:invalid_xml") }
  return new DOMParser({errorHandler:{warning:fail,error:fail,fatalError:fail}}).parseFromString(text, "application/xml")
}
function unpack(bytes: Uint8Array) {
  assertExactSpacingPackage(bytes)
  const parts = unzipSync(bytes), docs = Object.fromEntries(Object.entries(parts).filter(([name]) => /\.(xml|rels)$/.test(name)).map(([name, bytes]) => [name, parse(strFromU8(bytes))]))
  return {parts, docs, body:docs[BODY]!, settings:docs[SETTINGS]!}
}
type Package = ReturnType<typeof unpack>
function canonical(node: Node): string {
  if (node.nodeType === 9) return JSON.stringify(Array.from(node.childNodes).filter(n => !(n.nodeType === 7 && n.nodeName === "xml")).map(canonical))
  if (node.nodeType !== 1) return JSON.stringify([node.nodeType,node.nodeName,node.nodeValue])
  const el = node as Element
  return JSON.stringify([el.namespaceURI,el.nodeName,Array.from(el.attributes).map(a => [a.namespaceURI,a.name,a.value]).sort(),Array.from(el.childNodes).map(canonical)])
}
const same = (a:Node,b:Node) => canonical(a) === canonical(b)
function settingsBase(doc:Document) {
  const copy=doc.cloneNode(true) as Document, marks=all(copy,"trackRevisions")
  require(marks.length <= 1 && marks.every(mark => mark.parentNode === copy.documentElement && !mark.childNodes.length && Array.from(mark.attributes).every(a => a.namespaceURI === W && a.localName === "val" && ["1","true","on"].includes(a.value))),"unsupported_settings")
  marks.forEach(mark => mark.parentNode!.removeChild(mark)); return copy
}
function deliver(source: Package, bytes: Uint8Array, tracked: boolean) {
  const saved = unpack(bytes)
  require(JSON.stringify(Object.keys(source.parts).sort()) === JSON.stringify(Object.keys(saved.parts).sort()), "package_inventory_changed")
  for (const name of Object.keys(source.parts)) if (name !== BODY) {
    if (name === SETTINGS) require(same(settingsBase(source.settings),settingsBase(saved.settings)),"settings_changed")
    else require(source.docs[name] ? same(source.docs[name]!,saved.docs[name]!) : hash(source.parts[name]!) === hash(saved.parts[name]!),`unrelated_part_changed:${name}`)
  }
  if (tracked) require(all(saved.settings,"trackRevisions").length === 1,"tracking_not_enabled")
  const parts = {...source.parts,[BODY]:saved.parts[BODY]!,...(tracked ? {[SETTINGS]:saved.parts[SETTINGS]!} : {})}
  return zipSync(parts,{mtime:new Date(1980,0,1)})
}
// The new plain paragraph has no explicit style: only the default paragraph
// style chain and document defaults can supply a heading or list membership.
function safeDefaultStyle(pkg:Package) {
  require(!all(pkg.settings,"documentProtection").length && !all(pkg.settings,"writeProtection").length,"protected_document")
  require(!Object.keys(pkg.parts).some(name=>/stylesWithEffects\.xml$/i.test(name)) && !Object.entries(pkg.docs).some(([name,doc])=>name.endsWith(".rels") && Array.from(doc.getElementsByTagName("*")).some(el=>/stylesWithEffects$/i.test(el.getAttribute("Type")??""))),"unresolved_default_style")
  const styles=pkg.docs["word/styles.xml"],root=styles?.documentElement
  require(root?.namespaceURI===W&&root.localName==="styles","unresolved_default_style")
  const inspectPPr=(parent:Element) => {
    const properties=child(parent,"pPr");require(properties.length<=1,"unresolved_default_style")
    const allowed=new Set(["pStyle","keepNext","keepLines","pageBreakBefore","framePr","widowControl","suppressLineNumbers","pBdr","shd","tabs","suppressAutoHyphens","kinsoku","wordWrap","overflowPunct","topLinePunct","autoSpaceDE","autoSpaceDN","bidi","adjustRightInd","snapToGrid","spacing","ind","contextualSpacing","mirrorIndents","suppressOverlap","jc","textDirection","textAlignment","textboxTightWrap","outlineLvl","divId","cnfStyle","rPr","sectPr"])
    for(const pp of properties) {
      require(!["numPr","outlineLvl","pStyle","sectPr","pPrChange","customXml","sdt"].some(name=>all(pp,name).length) &&
        kids(pp).every(el=>el.namespaceURI===W&&allowed.has(el.localName)) &&
        Array.from(pp.getElementsByTagName("*")).every(el=>el.namespaceURI===W && !["pPr","customXml","sdt"].includes(el.localName)),"unresolved_default_style")
    }
  }
  const defaults=child(root,"docDefaults");require(defaults.length<=1,"unresolved_default_style")
  if(defaults[0]) {
    require(kids(defaults[0]).every(el=>el.namespaceURI===W&&["rPrDefault","pPrDefault"].includes(el.localName)),"unresolved_default_style")
    const ps=child(defaults[0],"pPrDefault"),rs=child(defaults[0],"rPrDefault")
    require(ps.length<=1&&rs.length<=1,"unresolved_default_style")
    if(ps[0]) {require(kids(ps[0]).every(el=>el.namespaceURI===W&&el.localName==="pPr"),"unresolved_default_style");inspectPPr(ps[0])}
  }
  const stylesById=child(root,"style")
  const defaultParagraph=stylesById.filter(style=>{
    if(style.getAttributeNS(W,"type")!=="paragraph"||!style.hasAttributeNS(W,"default"))return false
    const value=style.getAttributeNS(W,"default");require(value&&["0","1","true","false","on","off"].includes(value),"unresolved_default_style")
    return ["1","true","on"].includes(value)
  })
  require(defaultParagraph.length===1,"unresolved_default_style")
  let id=defaultParagraph[0]!.getAttributeNS(W,"styleId");const seen=new Set<string>()
  while(id) {
    require(!seen.has(id),"unresolved_default_style");seen.add(id)
    const found=stylesById.filter(style=>style.getAttributeNS(W,"styleId")===id)
    require(found.length===1&&found[0]!.getAttributeNS(W,"type")==="paragraph","unresolved_default_style")
    const style=found[0]!
    require(kids(style).every(el=>el.namespaceURI===W&&["name","aliases","basedOn","next","link","autoRedefine","hidden","uiPriority","semiHidden","unhideWhenUsed","qFormat","locked","personal","personalCompose","personalReply","rsid","pPr","rPr"].includes(el.localName)),"unresolved_default_style")
    inspectPPr(style)
    const bases=child(style,"basedOn");require(bases.length<=1,"unresolved_default_style")
    if(!bases.length)break
    require(!kids(bases[0]!).length && !!bases[0]!.getAttributeNS(W,"val"),"unresolved_default_style")
    id=bases[0]!.getAttributeNS(W,"val")
  }
  require(seen.size>0,"unresolved_default_style")
}
function isBoundarySafe(body:Element,target:Element,position:"before"|"after") {
  const names = new Map([["bookmarkStart","bookmark"],["bookmarkEnd","bookmark"],["commentRangeStart","comment"],["commentRangeEnd","comment"],
    ["moveFromRangeStart","moveFrom"],["moveFromRangeEnd","moveFrom"],["moveToRangeStart","moveTo"],["moveToRangeEnd","moveTo"],["permStart","perm"],["permEnd","perm"]])
  const active=new Set<string>(),seen=new Set<string>(),bookmarkNames=new Set<string>(); let safe=false
  const walk=(node:Element) => {
    if (node.namespaceURI === W && /Range(Start|End)$/.test(node.localName) && !names.has(node.localName)) throw new Error("docx_paragraph_insertion:unsupported_range")
    const kind=node.namespaceURI === W ? names.get(node.localName) : undefined
    if (kind) {
      const id=node.getAttributeNS(W,"id"), start=node.localName.endsWith("Start")
      const allowed=kind==="bookmark" && start ? ["id","name"] : kind==="perm" && start ? ["id","ed","edGrp"] :
        (kind==="moveFrom"||kind==="moveTo") && start ? ["id","name","author","date"] : ["id"]
      require(id && /^[+-]?\d+$/.test(id) && !node.childNodes.length &&
        Array.from(node.attributes).every(attr=>attr.namespaceURI===W&&allowed.includes(attr.localName)),"invalid_range")
      if(kind==="bookmark" && start) {
        const name=node.getAttributeNS(W,"name");require(name && !bookmarkNames.has(name.toLowerCase()),"invalid_range")
        bookmarkNames.add(name.toLowerCase())
      }
      const key=`${kind}:${BigInt(id).toString()}`
      if(start){require(!seen.has(key),"ambiguous_range");seen.add(key);active.add(key)}
      else {require(active.has(key),"unpaired_range");active.delete(key)}
    }
    for(const next of kids(node)) walk(next)
  }
  for(const node of kids(body)) {
    if(node === target && position === "before") safe=active.size===0
    walk(node)
    if(node === target && position === "after") safe=active.size===0
  }
  require(active.size===0,"unpaired_range")
  return safe
}
function plainReference(node:Element,text:string) {
  const clean=(element:Element)=>Array.from(element.childNodes).every(n=>n.nodeType===1||n.nodeType===3&&!n.nodeValue?.trim())
  if (!clean(node) || all(node,"p").length || /[\r\n\t]/.test(text) || child(node,"pPr").length>1) return false
  const direct=kids(node)
  if(direct.some(n=>n.namespaceURI!==W || !["pPr","r","bookmarkStart","bookmarkEnd"].includes(n.localName)))return false
  for(const run of child(node,"r")) {
    if(!clean(run) || child(run,"rPr").length>1 || kids(run).some(n=>n.namespaceURI!==W || !["rPr","t"].includes(n.localName)))return false
    for(const t of child(run,"t")) if(Array.from(t.childNodes).some(n=>n.nodeType!==3&&n.nodeType!==4) ||
      Array.from(t.attributes).some(a=>a.namespaceURI!=="http://www.w3.org/XML/1998/namespace" || a.localName!=="space" || !["preserve","default"].includes(a.value))) return false
  }
  return all(node,"t").map(t=>t.textContent??"").join("")===text
}
function candidates(opened:Opened) {
  const body=child(opened.documents[BODY]!.documentElement!,"body")
  require(body.length===1,"invalid_body")
  const siblings=kids(body[0]!), sourceHash=opened.revision.source_sha256
  return opened.paragraphs.flatMap((paragraph,ordinal) => {
    const node=paragraph.node, index=siblings.indexOf(node)
    if(index<0 || node.parentNode!==body[0] || !plainReference(node,paragraph.match.text) || paragraph.match.enclosingAnchor.kind!=="p" || paragraph.metadata.list || paragraph.metadata.outlineLevel !== undefined ||
      ["sectPr","sdt","fldChar","fldSimple","instrText","footnoteReference","endnoteReference","hyperlink","ins","del","moveFrom","moveTo","pPrChange","rPrChange"].some(name=>all(node,name).length)) return []
    const neighbors=[siblings[index-1],node,siblings[index+1]].filter((n):n is Element => !!n)
    if(neighbors.some(n => all(n,"sectPr").length)) return []
    const before=isBoundarySafe(body[0]!,node,"before"), after=isBoundarySafe(body[0]!,node,"after")
    if(!before&&!after)return []
    const neighbor=(element:Element|undefined)=>opened.paragraphs.find(row=>row.node===element)?.match.text ?? null
    return [{target_handle:hash(JSON.stringify([sourceHash,ordinal,index,canonical(node)])),text:paragraph.match.text,before,after,ordinal,index,anchor:paragraph.match.enclosingAnchor.id,
      headings:opened.revision.paragraphs[ordinal]!.headings,previous_paragraph:neighbor(siblings[index-1]),next_paragraph:neighbor(siblings[index+1]),body_paragraph_number:siblings.slice(0,index+1).filter(n=>n.namespaceURI===W&&n.localName==="p").length}]
  })
}
function input(source:Uint8Array,value:DocxParagraphInsertion) {
  const request=schema.parse(value)
  require(request.source_sha256===hash(source),"stale_source")
  require(request.text===request.text.trim() && request.text.isWellFormed() && !/[\u0000-\u001f\u007f]/.test(request.text),"unsupported_literal")
  return request
}
async function open(source:Uint8Array) { unpack(source);return openRevision(source,{pendingRevisions:"layer"}) }
export async function paragraphInsertionTargets(bytes:Uint8Array) {
  const source=new Uint8Array(bytes),opened=await open(source)
  try {
    let supported=true
    try {safeDefaultStyle(unpack(source))} catch(error) {
      if (!(error instanceof Error) || !["docx_paragraph_insertion:protected_document","docx_paragraph_insertion:unresolved_default_style"].includes(error.message)) throw error
      supported=false
    }
    return {source_sha256:hash(source),targets:candidates(opened).map(({target_handle,text,before,after,headings,previous_paragraph,next_paragraph,body_paragraph_number})=>({target_handle,text,before,after,supported,headings,previous_paragraph,next_paragraph,body_paragraph_number}))}
  }
  finally {opened.session.close()}
}
function selected(opened:Opened,request:DocxParagraphInsertion,pkg:Package) {
  safeDefaultStyle(pkg)
  const match=candidates(opened).find(c=>c.target_handle===request.target_handle)
  require(match && match[request.position],"target_unavailable")
  require(opened.session.listRevisions().every(revision=>revision.resolutionStatus==="supported"),"unsupported_prior_revision")
  return match
}
function literal(text:string) {return text.replace(/[\\`*_{}\[\]()#+.!<>|~\-]/g,"\\$&")}
function revisionSignature(rows:Revision[],pkg:Package,body=pkg.body) {
  return JSON.stringify(rows.toSorted((a,b)=>a.id.localeCompare(b.id)).map(revision=>{
    const {anchorId:_anchor,affectedAnchors:_affected,...identity}=revision
    const part=revision.partUri==="/word/document.xml" ? body : pkg.docs[revision.partUri.replace(/^\//,"")]
    require(part && revision.constituentKeys.length>0 && new Set(revision.constituentKeys).size===revision.constituentKeys.length,"prior_revision_unavailable")
    const extents=revision.constituentKeys.map(key=>{
      const match=/^kind=([^:]+):(.+):([^:]+):([^:]+)$/.exec(key);require(match,"prior_revision_unavailable")
      const [, ,namespace,name,id]=match
      const nodes=Array.from(part.getElementsByTagNameNS(namespace!,name!)).filter(node=>node.getAttributeNS(W,"id")===id)
      require(nodes.length===1,"prior_revision_ambiguous")
      const path:number[]=[];let node:Node=nodes[0]!
      while(node.parentNode&&node.parentNode.nodeType!==9){path.unshift(Array.from(node.parentNode.childNodes).indexOf(node as ChildNode));node=node.parentNode}
      require(node===part.documentElement,"prior_revision_unavailable")
      return {key,path,subtree:canonical(nodes[0]!)}
    })
    return {identity,extents}
  }))
}
function inserted(source:Package,output:Package,index:number,request:DocxParagraphInsertion) {
  const clone=output.body.cloneNode(true) as Document, body=child(clone.documentElement!,"body")[0]!, rows=kids(body)
  const at=index+Number(request.position==="after"), paragraph=rows[at]
  require(paragraph?.namespaceURI===W&&paragraph.localName==="p","inserted_paragraph_missing")
  body.removeChild(paragraph)
  require(same(source.body,clone),"original_document_changed")
  return {paragraph,remainder:clone,at}
}
function plainParagraph(paragraph:Element,text:string,tracked:boolean,author:string) {
  const only=(node:Element,names:string[])=>require(Array.from(node.childNodes).every(n=>n.nodeType===1&&(n as Element).namespaceURI===W&&names.includes((n as Element).localName)),"unexpected_new_paragraph")
  require(!paragraph.attributes.length,"unexpected_new_paragraph")
  const p=kids(paragraph), carriers:{id:string;key:string;date:string}[]=[]
  const stamp=(node:Element,kind:"Content"|"ParaMark")=>{
    const id=node.getAttributeNS(W,"id"),date=node.getAttributeNS(W,"date")
    const time=date?.match(/^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.(\d{1,7}))?Z$/)
    const parsed=time ? Date.parse(date!) : NaN
    const valid=!!time && Number.isFinite(parsed) && new Date(parsed).toISOString() === `${time[1]}.${(time[2] ?? "").slice(0,3).padEnd(3,"0")}Z`
    require(id && /^(0|[1-9]\d*)$/.test(id) && node.getAttributeNS(W,"author")===author && valid &&
      Array.from(node.attributes).every(a=>a.namespaceURI===W&&["id","author","date"].includes(a.localName)),"invalid_revision_carrier")
    carriers.push({id,key:`kind=${kind}:${W}:ins:${id}`,date:date!})
  }
  let content:Element
  if(tracked) {
    only(paragraph,["pPr","ins"]);require(p.length===2&&p[0]!.localName==="pPr"&&p[1]!.localName==="ins","unexpected_new_paragraph")
    const pp=p[0]!,wrap=p[1]!;require(!pp.attributes.length,"unexpected_new_paragraph");only(pp,["rPr"])
    const markProps=child(pp,"rPr");require(markProps.length===1&&!markProps[0]!.attributes.length,"invalid_paragraph_mark");only(markProps[0]!,["ins"])
    const marks=child(markProps[0]!,"ins");require(marks.length===1&&!marks[0]!.childNodes.length,"invalid_paragraph_mark")
    stamp(marks[0]!,"ParaMark");stamp(wrap,"Content");content=wrap
  } else {only(paragraph,["r"]);content=paragraph}
  only(content,["r"]);const runs=child(content,"r");require(runs.length===1&&!runs[0]!.attributes.length,"unexpected_literal_run")
  only(runs[0]!,["t"]);const texts=child(runs[0]!,"t")
  require(texts.length===1&&texts[0]!.textContent===text&&Array.from(texts[0]!.childNodes).every(n=>n.nodeType===3)&&
    Array.from(texts[0]!.attributes).every(a=>a.namespaceURI==="http://www.w3.org/XML/1998/namespace"&&a.localName==="space"&&a.value==="preserve"),"unexpected_literal_text")
  require(new Set(carriers.map(c=>c.id)).size===carriers.length,"duplicate_revision_id")
  return carriers
}
function sourceIds(pkg:Package) {
  const result=new Set<string>()
  for(const doc of Object.values(pkg.docs))for(const node of all(doc,"*"))if(node.hasAttributeNS(W,"id")){
    const id=node.getAttributeNS(W,"id");require(id && /^[+-]?\d+$/.test(id),"unsupported_source_id");result.add(BigInt(id).toString())
  }
  return result
}
function verifyPlainMetadata(opened:Opened,physicalIndex:number) {
  const paragraph=all(opened.documents[BODY]!,"p")[physicalIndex]
  const ordinal=opened.paragraphs.findIndex(row=>row.node===paragraph)
  require(ordinal>=0,"inserted_paragraph_unavailable")
  const row=opened.paragraphs[ordinal]!,view=opened.revision.paragraphs[ordinal]!
  require(row.match.enclosingAnchor.kind==="p" && !row.metadata.list && row.metadata.outlineLevel===undefined && !view.label,"inserted_paragraph_not_plain")
  return row
}
async function inspect(source:Uint8Array,bytes:Uint8Array,request:DocxParagraphInsertion) {
  const inputPkg=unpack(source),outputPkg=unpack(bytes),baseline=await open(source)
  let current:Opened | undefined
  try {
    current=await open(bytes)
    const target=selected(baseline,request,inputPkg),before=baseline.session.listRevisions()
    deliver(inputPkg,bytes,true)
    const placed=inserted(inputPkg,outputPkg,target.index,request), physical=plainParagraph(placed.paragraph,request.text,true,request.author), ids=sourceIds(inputPkg)
    require(physical.length===2&&physical.every(c=>!ids.has(BigInt(c.id).toString())),"source_revision_id_collision")
    require(physical[0]!.date===physical[1]!.date,"revision_date_changed")
    const rows=current.session.listRevisions(),old=new Set(before.map(r=>r.id)),existing=rows.filter(r=>old.has(r.id)),added=rows.filter(r=>!old.has(r.id))
    require(new Set(rows.map(r=>r.id)).size===rows.length&&revisionSignature(existing,outputPkg,placed.remainder)===revisionSignature(before,inputPkg),"prior_revision_changed")
    require(added.length===1,"unexpected_revision_group")
    const group=added[0]!, expected=physical.map(c=>c.key).sort()
    require(group.resolutionStatus==="supported"&&group.type==="insert"&&group.family==="content_insert"&&group.author===request.author&&group.date===physical[0]!.date&&group.partUri==="/word/document.xml"&&group.scope==="body"&&group.text===`${request.text}¶`&&
      JSON.stringify([...group.constituentKeys].sort())===JSON.stringify(expected)&&JSON.stringify([...group.constituentIds].sort())===JSON.stringify(physical.map(c=>c.id).sort()),"revision_membership_changed")
    const actual=kids(child(outputPkg.body.documentElement!,"body")[0]!)[placed.at]
    const paragraphIndex=actual ? all(outputPkg.body,"p").indexOf(actual) : -1
    require(paragraphIndex>=0,"inserted_paragraph_missing")
    const projected=verifyPlainMetadata(current,paragraphIndex)
    require(projected.match.text===request.text&&group.anchorId===projected.match.enclosingAnchor.id&&group.affectedAnchors.length===1&&group.affectedAnchors[0]!.id===group.anchorId,"revision_extent_changed")
    return {inputPkg,outputPkg,baseline,current,target,before,group,paragraphIndex}
  }catch(error){baseline.session.close();current?.session.close();throw error}
}
export async function verifyParagraphInsertion(sourceBytes:Uint8Array,artifact:Uint8Array,value:DocxParagraphInsertion) {
  const source=new Uint8Array(sourceBytes),request=input(source,value),checked=await inspect(source,new Uint8Array(artifact),request)
  try{return {revisions:[checked.group],paragraph:{xml_paragraph_index:checked.paragraphIndex,text:request.text}}}
  finally{checked.baseline.session.close();checked.current.session.close()}
}
export async function resolveParagraphInsertion(sourceBytes:Uint8Array,artifact:Uint8Array,value:DocxParagraphInsertion,decision:"accept"|"reject") {
  const source=new Uint8Array(sourceBytes),request=input(source,value);require(decision==="accept"||decision==="reject","invalid_decision")
  const checked=await inspect(source,new Uint8Array(artifact),request)
  try{
    const {current,group}=checked
    const matches=current.session.listRevisions().filter(row=>row.id===group.id&&JSON.stringify([...row.constituentKeys].sort())===JSON.stringify([...group.constituentKeys].sort()))
    require(matches.length===1,"revision_group_changed")
    const result=decision==="accept"?current.session.acceptRevision(group.id):current.session.rejectRevision(group.id)
    require(result.success,"resolution_failed")
    const saved=deliver(checked.inputPkg,current.session.save(),false),pkg=unpack(saved)
    const remainder=decision==="reject"?pkg.body:inserted(checked.inputPkg,pkg,checked.target.index,request).remainder
    if(decision==="reject") require(same(pkg.body,checked.inputPkg.body),"rejection_did_not_restore_source")
    else plainParagraph(inserted(checked.inputPkg,pkg,checked.target.index,request).paragraph,request.text,false,request.author)
    require(revisionSignature(current.session.listRevisions(),pkg,remainder)===revisionSignature(checked.before,checked.inputPkg),"resolution_changed_prior_revisions")
    const reopened=await open(saved)
    try {
      require(revisionSignature(reopened.session.listRevisions(),pkg,remainder)===revisionSignature(checked.before,checked.inputPkg),"saved_prior_revisions_changed")
      if(decision==="accept") {
        const element=kids(child(pkg.body.documentElement!,"body")[0]!)[checked.target.index+Number(request.position==="after")]!
        const index=all(pkg.body,"p").indexOf(element)
        require(index>=0 && verifyPlainMetadata(reopened,index).match.text===request.text,"inserted_paragraph_not_plain")
      }
    }
    finally {reopened.session.close()}
    return decision==="reject"?source:saved
  }finally{checked.baseline.session.close();checked.current.session.close()}
}
// Review decisions may leave Track Changes enabled. Accept only the exact
// already-verified tracked settings or exact source settings, then restore the
// source setting; every other part is left for the review wrapper to prove.
export async function normalizeParagraphInsertionDecisionSettings(sourceBytes:Uint8Array,trackedCore:Uint8Array,actualBytes:Uint8Array,value:DocxParagraphInsertion) {
  const source=new Uint8Array(sourceBytes),tracked=new Uint8Array(trackedCore),actual=new Uint8Array(actualBytes)
  await verifyParagraphInsertion(source,tracked,value)
  const original=unpack(source), verified=unpack(tracked), saved=unpack(actual)
  require(same(saved.settings,original.settings)||same(saved.settings,verified.settings),"unexpected_decision_settings")
  const parts={...saved.parts,[SETTINGS]:original.parts[SETTINGS]!}
  const result=zipSync(parts,{mtime:new Date(1980,0,1)});unpack(result)
  return result
}
export async function authorParagraphInsertion(sourceBytes:Uint8Array,value:DocxParagraphInsertion) {
  const source=new Uint8Array(sourceBytes),request=input(source,value),pkg=unpack(source),opened=await open(source)
  let bytes:Uint8Array
  try{
    const target=selected(opened,request,pkg)
    opened.session.setRevisionAuthor(request.author);opened.session.setTrackedChanges((await docxodus()).TrackedChangeMode.RenderInline)
    const result=opened.session.insertParagraph(target.anchor,request.position,literal(request.text))
    require(result.success,"authoring_failed")
    bytes=deliver(pkg,opened.session.save(),true)
  }finally{opened.session.close()}
  const verified=await verifyParagraphInsertion(source,bytes,request)
  await resolveParagraphInsertion(source,bytes,request,"reject");await resolveParagraphInsertion(source,bytes,request,"accept")
  return {bytes,revisions:verified.revisions,paragraph:verified.paragraph}
}
