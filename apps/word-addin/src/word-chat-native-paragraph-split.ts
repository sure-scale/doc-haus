import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import type { WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { assertNativeWordChatStoryInverse } from "./word-chat-native-paragraph-merge"

type Projection = ReturnType<typeof projectNativeWordChatBody>
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", W14 = "http://schemas.microsoft.com/office/word/2010/wordml", W16 = "http://schemas.microsoft.com/office/word/2023/wordml/word16du", PKG = "http://schemas.microsoft.com/office/2006/xmlPackage", XMLNS = "http://www.w3.org/2000/xmlns/", XML = "http://www.w3.org/XML/1998/namespace"
const elements = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const named = (n: Node, name: string) => elements(n).filter(e => e.namespaceURI === W && e.localName === name)
const fail = (reason: string): never => { throw Error(`Native paragraph split could not be verified: ${reason}`) }
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
// Projection paths are element indices within body/table/row/cell, not text searches.
function atPath(body:Element,path:string):Element {
  const pieces=path.split("/");if(pieces.shift()!=="body" || !pieces.length)return fail("paragraph path")
  let node=body
  for(const piece of pieces){const m=/^(p|tbl|tr|tc)\[(0|[1-9]\d*)\]$/.exec(piece);if(!m)return fail("paragraph path");const child=elements(node)[Number(m[2])];if(!child || child.namespaceURI!==W || child.localName!==m[1])return fail("paragraph path");node=child}
  return node
}
const cleanContainer=(e:Element)=>Array.from(e.childNodes).every(n=>n.nodeType===1 || n.nodeType===3 && !n.nodeValue?.trim())
function runChars(run:Element, textName:"t"|"delText", deleted=false):Array<{text:string;props:unknown}> {
  if(run.namespaceURI!==W || run.localName!=="r" || !cleanContainer(run))return fail("run structure")
  if(Array.from(run.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && (["rsidR","rsidRPr"].includes(a.localName) || deleted && a.localName==="rsidDel") && /^[0-9a-f]{8}$/i.test(a.value))))return fail("run attributes")
  const rs=elements(run),props=named(run,"rPr"),texts=named(run,textName)
  if(props.length>1 || texts.length!==1 || rs.length!==1+props.length || rs.at(-1)!==texts[0] || props.length && rs[0]!==props[0])return fail("run content")
  const t=texts[0]!;if(elements(t).length || Array.from(t.childNodes).some(n=>n.nodeType!==3) || Array.from(t.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===XML && a.localName==="space" && a.value==="preserve")))return fail("run text")
  const value=t.textContent??"";if(!value || (/^\s|\s$/.test(value) && t.getAttributeNS(XML,"space")!=="preserve"))return fail("run whitespace")
  const property=props.length?canon(props[0]!,true):null
  return Array.from({length:value.length},(_,i)=>({text:value[i]!,props:property}))
}
/** Only a complete, qualified ordinary paragraph can be split. The host checks effective style and metrics. */
export function assertNativeWordChatParagraphSplitTarget(projection: Projection, target: WordChatTarget): number {
  if (!target?.id || target.paragraphs?.length!==1 || !target.text || target.startOffset!==0 || target.endOffset!==target.text.length || target.paragraphs[0]!.text!==target.text) return fail("target binding")
  const i=projection.paragraphs.findIndex(p=>p.id===target.paragraphs[0]!.id), item=projection.paragraphs[i]
  if (!item || item.text!==target.text || item.nesting>1 || item.nesting<0) return fail("target binding")
  const {body}=parse(projection.ooxml), p=atPath(body,item.path)
  if(!p || p.namespaceURI!==W || p.localName!=="p" || !validMeta(p)) return fail("target structure")
  const children=elements(p), runs=named(p,"r"), pp=named(p,"pPr")
  if(!cleanContainer(p) || !runs.length || pp.length>1 || children.length!==runs.length+pp.length || pp.length && children[0]!==pp[0] || pp.length && named(pp[0]!,"rPr").length) return fail("unsupported paragraph properties")
  const chars=runs.flatMap(r=>runChars(r,"t"));if(chars.map(c=>c.text).join("")!==target.text)return fail("target text")
  if(Array.from(p.getElementsByTagNameNS(W,"*")).some(e=>["ins","del","moveFrom","moveTo","pPrChange","rPrChange","numPr","outlineLvl","sectPr","fldChar","instrText","vanish","webHidden","commentRangeStart","commentRangeEnd","commentReference","bookmarkStart","bookmarkEnd","permStart","permEnd"].includes(e.localName)))return fail("target unsupported content")
  const active=new Set<string>();let fields=0
  const walk=(e:Element)=>{
    if(e===p && (active.size || fields))return fail("crossing annotation")
    const starts:Record<string,string>={commentRangeStart:"comment",bookmarkStart:"bookmark",permStart:"perm",moveFromRangeStart:"moveFrom",moveToRangeStart:"moveTo"}, ends:Record<string,string>={commentRangeEnd:"comment",bookmarkEnd:"bookmark",permEnd:"perm",moveFromRangeEnd:"moveFrom",moveToRangeEnd:"moveTo"}
    if(e.namespaceURI===W){const family=starts[e.localName]??ends[e.localName];if(family){const id=e.getAttributeNS(W,"id"),key=`${family}:${id}`;if(!id || starts[e.localName] && active.has(key) || ends[e.localName] && !active.has(key))return fail("annotation graph");if(starts[e.localName])active.add(key);else active.delete(key)}if(e.localName==="fldChar"){const kind=e.getAttributeNS(W,"fldCharType");if(kind==="begin")fields++;else if(kind==="end"){if(!fields)return fail("field graph");fields--}}}
    for(const c of elements(e))walk(c)
    if(e===p && (active.size || fields))return fail("crossing annotation")
  };walk(body);if(active.size || fields)return fail("annotation graph")
  return i
}
/** Verify the observed split carriers and reconstruct the source before comparing the whole story. */
export function verifyNativeWordChatParagraphSplit(before:Projection, after:Projection, target:WordChatTarget, replacement:string):string[] {
  const i=assertNativeWordChatParagraphSplitTarget(before,target), fragments=replacement.split("\r")
  if(fragments.length!==2 || !fragments[0] || !fragments[1] || target.text!==`${fragments[0]} ${fragments[1]}`)return fail("replacement geometry")
  const a=after.paragraphs[i], b=after.paragraphs[i+1], original=before.paragraphs[i]!
  const ids=after.paragraphs.map(p=>p.id)
  if(after.paragraphs.length!==before.paragraphs.length+1 || !a || !b || !a.id || a.id===b.id || before.paragraphs.some(p=>p.id===a.id) || new Set(ids).size!==ids.length || a.text!==fragments[0] || b.text!==fragments[1] || b.id!==original.id || a.nesting!==original.nesting || b.nesting!==original.nesting || before.paragraphs.some((p,j)=>{if(j===i)return false;const q=after.paragraphs[j+(j>i?1:0)];return !q || p.id!==q.id || p.text!==q.text || p.nesting!==q.nesting}))return fail("native paragraph inventory")
  const old=parse(before.ooxml), current=parse(after.ooxml)
  trimTerminal(old.body,before);trimTerminal(current.body,after)
  const prior=atPath(old.body,original.path), left=atPath(current.body,a.path),right=atPath(current.body,b.path)
  if(left.parentNode!==right.parentNode || !left.parentNode || left.namespaceURI!==W || right.namespaceURI!==W || left.localName!=="p" || right.localName!=="p" || elements(left.parentNode).indexOf(right)!==elements(left.parentNode).indexOf(left)+1 || !validMeta(left) || !validMeta(right))return fail("target paragraphs")
  const parent=left.parentNode, sourceRuns=named(prior,"r"), lc=elements(left),rc=elements(right), pp=named(left,"pPr"), pr=pp.length===1?named(pp[0]!,"rPr"):[], marks=pr.length===1?named(pr[0]!,"ins"):[], dels=named(right,"del"), lr=named(left,"r"),rr=named(right,"r")
  if(!cleanContainer(left) || !cleanContainer(right) || pp.length!==1 || lc[0]!==pp[0] || pr.length!==1 || marks.length!==1 || !cleanContainer(pp[0]!) || !cleanContainer(pr[0]!) || marks[0]!.childNodes.length || dels.length!==1 || rc[named(right,"pPr").length]!==dels[0] || lc.length!==1+lr.length || rc.length!==named(right,"pPr").length+1+rr.length)return fail("revision carriers")
  const mark=marks[0]!, del=dels[0]!, dr=named(del,"r")
  if(!cleanContainer(del) || elements(del).length!==1 || dr.length!==1)return fail("deleted space")
  const deleted=runChars(dr[0]!,"delText",true), source=sourceRuns.flatMap(r=>runChars(r,"t")), leftChars=lr.flatMap(r=>runChars(r,"t")), rightChars=rr.flatMap(r=>runChars(r,"t")), reconstructed=[...leftChars,...deleted,...rightChars]
  if(deleted.length!==1 || deleted[0]!.text!==" " || !equal(source,reconstructed))return fail("run provenance")
  const boundary=source[leftChars.length];if(!boundary || boundary.text!==" ")return fail("boundary provenance")
  // The added mark may carry only the original boundary run properties.
  const markProps=pr[0]!.cloneNode(true) as Element, markCopy=named(markProps,"ins")[0]!;markProps.removeChild(markCopy)
  if(Array.from(markProps.attributes).some(a=>a.namespaceURI!==XMLNS))return fail("paragraph mark attributes")
  const markCanon=elements(markProps).length ? canon(markProps,true) : null
  if(!equal(markCanon,boundary.props))return fail("paragraph mark formatting")
  const leftProps=pp[0]!.cloneNode(true) as Element;leftProps.removeChild(named(leftProps,"rPr")[0]!)
  const priorProps=named(prior,"pPr"),rightProps=named(right,"pPr")
  const empty=(e:Element)=>!elements(e).length && Array.from(e.attributes).every(x=>x.namespaceURI===XMLNS) && cleanContainer(e)
  if(priorProps.length===0 ? !empty(leftProps) || rightProps.length!==0 : rightProps.length!==1 || !equal(canon(leftProps,true),canon(priorProps[0]!,true)) || !equal(canon(rightProps[0]!,true),canon(priorProps[0]!,true)))return fail("paragraph properties")
  const author=mark.getAttributeNS(W,"author"),date=mark.getAttributeNS(W,"date"),owned:string[]=[]
  for(const e of [mark,del]){const id=e.getAttributeNS(W,"id"),utc=e.getAttributeNS(W16,"dateUtc");if(!id || !/^\d+$/.test(id) || !author?.trim() || !date || Number.isNaN(Date.parse(date)) || e.getAttributeNS(W,"author")!==author || e.getAttributeNS(W,"date")!==date || utc && utc!==date || Array.from(e.attributes).some(x=>x.namespaceURI!==XMLNS && !(x.namespaceURI===W && ["id","author","date"].includes(x.localName)) && !(x.namespaceURI===W16 && x.localName==="dateUtc")))return fail("revision ownership");owned.push(numeric(id))}
  if(owned[0]===owned[1])return fail("revision IDs")
  if(!prior.getAttributeNS(W14,"paraId") || right.getAttributeNS(W14,"paraId")!==prior.getAttributeNS(W14,"paraId") || !left.getAttributeNS(W14,"paraId") || left.getAttributeNS(W14,"paraId")===right.getAttributeNS(W14,"paraId") || Boolean(prior.getAttributeNS(W14,"textId"))!==Boolean(right.getAttributeNS(W14,"textId")) || left.getAttributeNS(W14,"textId")!==prior.getAttributeNS(W14,"textId"))return fail("paragraph identity")
  const allParaIds=Array.from(current.body.getElementsByTagNameNS(W,"p")).map(p=>p.getAttributeNS(W14,"paraId")).filter(Boolean)
  if(new Set(allParaIds).size!==allParaIds.length)return fail("duplicate paragraph identity")
  // Target properties and characters are proven above; replace only that pair for the full-story inverse.
  const restored=prior.cloneNode(true) as Element
  parent.replaceChild(restored,right);parent.removeChild(left)
  assertNativeWordChatStoryInverse(before,after,new XMLSerializer().serializeToString(current.doc),author!,owned)
  return [author!]
}
