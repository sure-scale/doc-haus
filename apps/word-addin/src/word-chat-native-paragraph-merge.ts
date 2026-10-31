import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { assertNativeWordChatInlineDependencies, projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { normalizeNativeWordChatCommentIds } from "./word-chat-native-comments"

type Projection = ReturnType<typeof projectNativeWordChatBody>
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", W14 = "http://schemas.microsoft.com/office/word/2010/wordml", W16 = "http://schemas.microsoft.com/office/word/2023/wordml/word16du", PKG = "http://schemas.microsoft.com/office/2006/xmlPackage", XMLNS = "http://www.w3.org/2000/xmlns/", XML = "http://www.w3.org/XML/1998/namespace"
const elements = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const named = (n: Node, name: string) => elements(n).filter(e => e.namespaceURI === W && e.localName === name)
const fail = (reason: string): never => { throw Error(`Native paragraph merge could not be verified: ${reason}`) }
const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const numeric = (s: string) => s.replace(/^0+(?=\d)/, "")
function parse(xml: string) {
  if (xml.length > 16_000_000 || /<!DOCTYPE/i.test(xml)) return fail("invalid package")
  const reject = () => { throw Error("Invalid Word XML") }
  const doc = new DOMParser({errorHandler:{warning:reject,error:reject,fatalError:reject}}).parseFromString(xml,"application/xml")
  if (doc.doctype || doc.documentElement.namespaceURI !== PKG || doc.documentElement.localName !== "package") return fail("invalid package")
  const parts = new Map<string,Element>()
  for (const part of elements(doc.documentElement)) { const name=part.getAttributeNS(PKG,"name"); if (part.namespaceURI!==PKG || part.localName!=="part" || !name || parts.has(name) || parts.size>=64) return fail("invalid package parts"); parts.set(name,part) }
  const main=parts.get("/word/document.xml"), bodies=main ? Array.from(main.getElementsByTagNameNS(W,"body")) : []
  if (bodies.length!==1) return fail("invalid main story")
  return {doc,body:bodies[0]!,parts}
}
function canon(n: Node, strict = false): unknown {
  if (n.nodeType===1) { const e=n as Element; return [e.namespaceURI,e.localName,Array.from(e.attributes).filter(a=>a.namespaceURI!==XMLNS && !(!strict && a.namespaceURI===W && ["rsidR","rsidRDefault"].includes(a.localName))).map(a=>[a.namespaceURI,a.localName,a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))),Array.from(e.childNodes).map(c=>canon(c,strict)).filter(x=>x!==undefined)] }
  if (n.nodeType===3 || n.nodeType===4) return n.nodeValue?.trim() || ["t","delText"].includes((n.parentNode as Element)?.localName) ? ["text",n.nodeValue] : undefined
  if (n.nodeType===7 || n.nodeType===8) return [n.nodeType,n.nodeName,n.nodeValue]
  return undefined
}
function validMeta(e: Element) { return Array.from(e.attributes).every(a=>a.namespaceURI===XMLNS || a.namespaceURI===W && ["rsidR","rsidRDefault"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value) || a.namespaceURI===W14 && ["paraId","textId"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)) }
function trimTerminal(body: Element, projection: Projection) {
  const children=elements(body), section=children.at(-1), tail=children.at(-2)
  if (section?.namespaceURI!==W || section.localName!=="sectPr" || tail?.namespaceURI!==W || tail.localName!=="p" || projection.paragraphs.some(p=>p.path===`body/p[${children.length-2}]`)) return
  if (!validMeta(tail) || Array.from(tail.childNodes).some(n=>n.nodeType!==3 || Boolean(n.nodeValue?.trim()))) return fail("terminal carrier")
  body.removeChild(tail)
}
function qualified(p: Element, text: string) {
  const children=elements(p), pp=named(p,"pPr"), runs=named(p,"r")
  if (!validMeta(p) || pp.length>1 || !runs.length || children.some(c=>c.namespaceURI!==W || !["pPr","r"].includes(c.localName)) || pp.length && children[0]!==pp[0]) return fail("target paragraph structure")
  if (Array.from(p.getElementsByTagNameNS(W,"*")).some(e=>["ins","del","moveFrom","moveTo","pPrChange","rPrChange","numPr","outlineLvl","sectPr","fldChar","instrText","vanish","webHidden","commentRangeStart","commentRangeEnd","commentReference","bookmarkStart","bookmarkEnd","permStart","permEnd"].includes(e.localName))) return fail("target unsupported content")
  for (const r of runs) {
    const ts=named(r,"t"), props=named(r,"rPr")
    if (ts.length!==1 || props.length>1 || elements(r).some(c=>c.namespaceURI!==W || !["rPr","t"].includes(c.localName)) || elements(ts[0]!).length || Array.from(ts[0]!.childNodes).some(n=>n.nodeType!==3)) return fail("target run structure")
  }
  if (runs.map(r=>named(r,"t")[0]!.textContent??"").join("")!==text) return fail("target text")
}
/** The host also checks native Normal style, non-list status and compatible effective metrics. */
export function assertNativeWordChatParagraphMergeTarget(projection: Projection, target: WordChatTarget): [number,number] {
  if (!target?.id || target.paragraphs?.length!==2 || !target.paragraphs[0]!.text || !target.paragraphs[1]!.text || target.startOffset!==0 || target.endOffset!==target.paragraphs[1]!.text.length || target.text!==`${target.paragraphs[0]!.text}\r${target.paragraphs[1]!.text}`) return fail("target binding")
  const i=projection.paragraphs.findIndex(p=>p.id===target.paragraphs[0]!.id), left=projection.paragraphs[i],right=projection.paragraphs[i+1]
  if (!left || !right || right.id!==target.paragraphs[1]!.id || left.id===right.id || left.text!==target.paragraphs[0]!.text || right.text!==target.paragraphs[1]!.text || left.nesting!==0 || right.nesting!==0 || !/^body\/p\[\d+\]$/.test(left.path) || !/^body\/p\[\d+\]$/.test(right.path)) return fail("target binding")
  const parsed=parse(projection.ooxml), children=elements(parsed.body), l=children[Number(/^body\/p\[(\d+)\]$/.exec(left.path)![1])], r=children[Number(/^body\/p\[(\d+)\]$/.exec(right.path)![1])]
  if (!l || !r || children.indexOf(r)!==children.indexOf(l)+1 || l.localName!=="p" || r.localName!=="p") return fail("target adjacency")
  qualified(l,left.text);qualified(r,right.text)
  if (!equal(named(l,"pPr").map(x=>canon(x)),named(r,"pPr").map(x=>canon(x)))) return fail("paragraph properties")
  const active=new Set<string>();let fields=0
  const walk=(e:Element)=>{
    if ((e===l || e===r) && (active.size || fields)) return fail("crossing annotation")
    const starts:Record<string,string>={commentRangeStart:"comment",bookmarkStart:"bookmark",permStart:"perm",moveFromRangeStart:"moveFrom",moveToRangeStart:"moveTo"}, ends:Record<string,string>={commentRangeEnd:"comment",bookmarkEnd:"bookmark",permEnd:"perm",moveFromRangeEnd:"moveFrom",moveToRangeEnd:"moveTo"}
    if(e.namespaceURI===W) { const family=starts[e.localName]??ends[e.localName];if(family){const id=e.getAttributeNS(W,"id"),key=`${family}:${id}`;if(!id || starts[e.localName] && active.has(key) || ends[e.localName] && !active.has(key))return fail("annotation graph");if(starts[e.localName])active.add(key);else active.delete(key)} if(e.localName==="fldChar"){const kind=e.getAttributeNS(W,"fldCharType");if(kind==="begin")fields++;else if(kind==="end"){if(!fields)return fail("field graph");fields--}} }
    for(const c of elements(e))walk(c)
    if((e===l || e===r) && (active.size || fields))return fail("crossing annotation")
  }
  walk(parsed.body);if(active.size || fields)return fail("annotation graph")
  return [i,i+1]
}
/** Prove both owned carriers, then inverse only them against the complete fresh story. */
export function verifyNativeWordChatParagraphMerge(before:Projection, after:Projection, target:WordChatTarget):string[] {
  const [li,ri]=assertNativeWordChatParagraphMergeTarget(before,target)
  if(before.paragraphs.length!==after.paragraphs.length || before.paragraphs.some((p,i)=>{const q=after.paragraphs[i];return !q || p.id!==q.id || p.path!==q.path || p.nesting!==q.nesting || q.text!==(i===ri?` ${p.text}`:p.text)})) return fail("native paragraph inventory")
  assertNativeWordChatInlineDependencies(before.ooxml,after.ooxml)
  const old=parse(before.ooxml), current=parse(after.ooxml)
  trimTerminal(old.body,before);trimTerminal(current.body,after)
  const pos=Number(/^body\/p\[(\d+)\]$/.exec(before.paragraphs[li]!.path)![1]), oldLeft=elements(old.body)[pos]!, oldRight=elements(old.body)[pos+1]!, left=elements(current.body)[pos],right=elements(current.body)[pos+1]
  if(!left || !right || left.namespaceURI!==W || right.namespaceURI!==W || left.localName!=="p" || right.localName!=="p")return fail("target adjacency")
  const lp=named(left,"pPr"), rp=lp.length===1?named(lp[0]!,"rPr"):[], marks=rp.length===1?named(rp[0]!,"del"):[], additions=named(right,"ins")
  if(lp.length!==1 || rp.length!==1 || marks.length!==1 || additions.length!==1 || elements(right)[named(right,"pPr").length]!==additions[0] || marks[0]!.childNodes.length) return fail("revision carriers")
  const mark=marks[0]!, ins=additions[0]!, runs=named(ins,"r"), t=runs.length===1?named(runs[0]!,"t"):[]
  if(Array.from(ins.childNodes).some(n=>n.nodeType!==1 && (n.nodeType!==3 || Boolean(n.nodeValue?.trim()))) || elements(ins).length!==1 || runs.length!==1 || Array.from(runs[0]!.childNodes).some(n=>n.nodeType!==1 && (n.nodeType!==3 || Boolean(n.nodeValue?.trim()))) || elements(runs[0]!).length!==1 || t.length!==1 || t[0]!.textContent!==" " || elements(t[0]!).length || Array.from(t[0]!.childNodes).some(n=>n.nodeType!==3) || Array.from(runs[0]!.attributes).some(a=>a.namespaceURI!==XMLNS) || Array.from(t[0]!.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===XML && a.localName==="space" && a.value==="preserve")) || t[0]!.getAttributeNS(XML,"space")!=="preserve") return fail("inserted space")
  const author=mark.getAttributeNS(W,"author"),date=mark.getAttributeNS(W,"date"),ids:string[]=[]
  for(const e of [mark,ins]) {
    const id=e.getAttributeNS(W,"id"),utc=e.getAttributeNS(W16,"dateUtc")
    if(!id || !/^\d+$/.test(id) || !author?.trim() || !date || Number.isNaN(Date.parse(date)) || e.getAttributeNS(W,"author")!==author || e.getAttributeNS(W,"date")!==date || utc && utc!==date || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && ["id","author","date"].includes(a.localName)) && !(a.namespaceURI===W16 && a.localName==="dateUtc"))) return fail("revision ownership")
    ids.push(numeric(id))
  }
  if(ids[0]===ids[1])return fail("revision IDs")
  for(const [prior,next] of [[oldLeft,left],[oldRight,right]] as const) {
    if(prior.getAttributeNS(W14,"paraId")!==next.getAttributeNS(W14,"paraId") || Boolean(prior.getAttributeNS(W14,"textId"))!==Boolean(next.getAttributeNS(W14,"textId")))return fail("paragraph identity")
    for(const a of Array.from(next.attributes)) if(a.namespaceURI!==XMLNS && !(a.namespaceURI===W && ["rsidR","rsidRDefault","rsidDel"].includes(a.localName) && /^[0-9A-Fa-f]{8}$/.test(a.value)) && !(a.namespaceURI===W14 && ["paraId","textId"].includes(a.localName) && /^[0-9A-Fa-f]{8}$/.test(a.value)))return fail("paragraph metadata")
    next.removeAttributeNS(W,"rsidDel")
    if(prior.hasAttributeNS(W14,"textId")) next.setAttributeNS(W14,"w14:textId",prior.getAttributeNS(W14,"textId")!)
  }
  rp[0]!.removeChild(mark)
  const priorProps=named(oldLeft,"pPr"), priorRunProps=priorProps.length===1?named(priorProps[0]!,"rPr"):[]
  const empty=(e:Element)=>!elements(e).length && Array.from(e.attributes).every(a=>a.namespaceURI===XMLNS) && Array.from(e.childNodes).every(n=>n.nodeType===3 && !n.nodeValue?.trim())
  if(!priorRunProps.length && empty(rp[0]!))lp[0]!.removeChild(rp[0]!)
  if(!priorProps.length && empty(lp[0]!))left.removeChild(lp[0]!)
  right.removeChild(ins)
  assertNativeWordChatStoryInverse(before,after,new XMLSerializer().serializeToString(current.doc),author!,ids)
  return [author!]
}

/** Compare the reconstructed whole story and dependencies, preserving prior review provenance. */
export function assertNativeWordChatStoryInverse(before:Projection, after:Projection, inverseAfterXml:string, author:string, ids:string[]):void {
  assertNativeWordChatInlineDependencies(before.ooxml,after.ooxml)
  const old=parse(before.ooxml)
  trimTerminal(old.body,before)
  const normalized=parse(normalizeNativeWordChatCommentIds(before.ooxml,inverseAfterXml))
  // Preserve every available dependency, permitting only Word's qualified
  // rsid inventory and a newly introduced people relationship/part.
  const extra=[...normalized.parts.keys()].filter(name=>!old.parts.has(name))
  if ([...old.parts.keys()].some(name=>!normalized.parts.has(name)) || extra.some(name=>name!=="/word/people.xml") || extra.length>1) return fail("package dependencies")
  if(extra.length) {
    if(!old.parts.has("/word/_rels/document.xml.rels")) return fail("people relationship")
    const part=normalized.parts.get(extra[0]!)!, ns="http://schemas.microsoft.com/office/word/2012/wordml", data=elements(part), people=Array.from(part.getElementsByTagNameNS(ns,"people")), persons=Array.from(part.getElementsByTagNameNS(ns,"person"))
    if(data.length!==1 || data[0]!.namespaceURI!==PKG || data[0]!.localName!=="xmlData" || elements(data[0]!).length!==1 || people.length!==1 || elements(data[0]!)[0]!==people[0] || elements(people[0]!).length!==1 || Array.from(people[0]!.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI==="http://schemas.openxmlformats.org/markup-compatibility/2006" && a.localName==="Ignorable")) || persons.length!==1 || persons[0]!.parentNode!==people[0] || persons[0]!.getAttributeNS(ns,"author")!==author) return fail("people metadata")
    const c=elements(persons[0]!)
    if(c.length!==1 || c[0]!.namespaceURI!==ns || c[0]!.localName!=="presenceInfo" || elements(c[0]!).length || !c[0]!.getAttributeNS(ns,"providerId") || !c[0]!.getAttributeNS(ns,"userId") || Array.from(persons[0]!.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===ns && a.localName==="author")) || Array.from(c[0]!.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===ns && ["providerId","userId"].includes(a.localName)))) return fail("people metadata")
  }
  for(const [name,prior] of old.parts) {
    if(name==="/word/document.xml")continue
    const next=normalized.parts.get(name)!
    if(name==="/word/_rels/document.xml.rels" && extra.length) {
      const rel="http://schemas.openxmlformats.org/package/2006/relationships",type="http://schemas.microsoft.com/office/2011/relationships/people",theme="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", ref="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      const a=Array.from(prior.getElementsByTagNameNS(rel,"Relationship")),b=Array.from(next.getElementsByTagNameNS(rel,"Relationship")),added=b.filter(e=>e.getAttribute("Type")===type)
      const valid=(list:Element[])=>list.every(e=>e.getAttribute("Id") && !elements(e).length && Array.from(e.attributes).every(x=>x.namespaceURI===XMLNS || !x.namespaceURI && ["Id","Type","Target"].includes(x.name))) && new Set(list.map(e=>e.getAttribute("Id"))).size===list.length
      if(!valid(a) || !valid(b) || a.some(e=>e.getAttribute("Type")===type) || added.length!==1 || added[0]!.getAttribute("Target")!=="people.xml")return fail("people relationship")
      const shell=(p:Element)=>{const copy=p.cloneNode(true) as Element;for(const e of Array.from(copy.getElementsByTagNameNS(rel,"Relationship")))e.parentNode!.removeChild(e);return canon(copy,true)}
      if(!equal(shell(prior),shell(next)) || a.length!==b.length-1)return fail("relationship dependencies")
      const used=new Set<Element>()
      for(const x of a){const sig=(e:Element)=>{const c=e.cloneNode(true) as Element;c.removeAttribute("Id");return canon(c)};const candidates=b.filter(y=>y!==added[0] && !used.has(y) && equal(sig(x),sig(y)));if(candidates.length!==1)return fail("relationship mapping");const y=candidates[0]!;if(x.getAttribute("Id")!==y.getAttribute("Id")){if(x.getAttribute("Type")!==theme)return fail("relationship mapping");const referenced=(doc:Document,id:string)=>Array.from(doc.getElementsByTagName("*")).some(e=>Array.from(e.attributes).some(a=>a.namespaceURI===ref && a.value===id));if(referenced(old.doc,x.getAttribute("Id")!) || referenced(normalized.doc,y.getAttribute("Id")!))return fail("relationship mapping")}used.add(y)}
      continue
    }
    const clean=(part:Element)=>{const c=part.cloneNode(true) as Element;if(name==="/word/settings.xml"){const lists=Array.from(c.getElementsByTagNameNS(W,"rsids"));if(lists.length>1)return fail("settings rsids");for(const list of lists){if(Array.from(list.attributes).some(a=>a.namespaceURI!==XMLNS) || elements(list).some(e=>e.namespaceURI!==W || !["rsid","rsidRoot"].includes(e.localName) || elements(e).length || !/^[0-9A-Fa-f]{8}$/.test(e.getAttributeNS(W,"val")??"") || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && a.localName==="val"))))return fail("settings rsids");list.parentNode!.removeChild(list)}}return canon(c,true)}
    if(name==="/word/settings.xml" && prior.getElementsByTagNameNS(W,"rsids").length!==next.getElementsByTagNameNS(W,"rsids").length || !equal(clean(prior),clean(next)))return fail("package dependencies")
  }
  trimTerminal(normalized.body,after)
  const carriers=(body:Element)=>{const out:Array<{path:string;node:Element}>=[];const visit=(e:Element,path:string)=>{if(e.namespaceURI===W && ["ins","del","pPrChange","rPrChange"].includes(e.localName))out.push({path,node:e});elements(e).forEach((c,i)=>visit(c,`${path}/${c.localName}[${i}]`))};visit(body,"body");return out}
  const oldRevs=carriers(old.body), nowRevs=carriers(normalized.body), oldIds=oldRevs.map(x=>x.node.getAttributeNS(W,"id")), nowIds=nowRevs.map(x=>x.node.getAttributeNS(W,"id"))
  if(oldRevs.length!==nowRevs.length || [...oldIds,...nowIds].some(id=>!id || !/^\d+$/.test(id)) || new Set(oldIds.map(id=>numeric(id!))).size!==oldIds.length || new Set(nowIds.map(id=>numeric(id!))).size!==nowIds.length || nowIds.some(id=>ids.includes(numeric(id!))))return fail("prior revision IDs")
  for(let i=0;i<oldRevs.length;i++){const a=oldRevs[i]!,b=nowRevs[i]!;if(a.path!==b.path || a.node.namespaceURI!==b.node.namespaceURI || a.node.localName!==b.node.localName)return fail("prior revision path");b.node.setAttributeNS(W,b.node.getAttributeNodeNS(W,"id")!.name,oldIds[i]!)}
  for(const body of [old.body,normalized.body]) {
    const visit=(e:Element)=>{for(const key of ["rsidR","rsidRDefault"])if(e.hasAttributeNS(W,key) && !/^[0-9A-Fa-f]{8}$/.test(e.getAttributeNS(W,key)!))return fail("invalid serialization metadata");elements(e).forEach(visit)}
    visit(body)
    const last=elements(body).at(-1);if(last?.namespaceURI===W && last.localName==="sectPr"){if(last.hasAttributeNS(W,"rsidSect") && !/^[0-9A-Fa-f]{8}$/.test(last.getAttributeNS(W,"rsidSect")!))return fail("invalid section metadata");last.removeAttributeNS(W,"rsidSect")}
  }
  if(!equal(canon(old.body),canon(normalized.body)))return fail("main story inverse")
}
