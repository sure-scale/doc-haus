import { DOMParser, XMLSerializer } from "@xmldom/xmldom"

const PKG="http://schemas.microsoft.com/office/2006/xmlPackage"
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XMLNS="http://www.w3.org/2000/xmlns/"
const MAIN="/word/document.xml", COMMENTS="/word/comments.xml"
const kinds=new Set(["commentRangeStart","commentRangeEnd","commentReference"])
const bookmarkKinds=new Set(["bookmarkStart","bookmarkEnd"])
const allKinds=new Set([...kinds,...bookmarkKinds])
const elems=(n:Node)=>Array.from(n.childNodes).filter((c):c is Element=>c.nodeType===1)
const fail=():never=>{throw Error("Word comment anchors could not be verified")}
function canonical(node:Node):unknown {
  if(node.nodeType===1){const e=node as Element;return [e.namespaceURI,e.localName,Array.from(e.attributes).filter(a=>a.namespaceURI!==XMLNS).map(a=>[a.namespaceURI,a.localName,a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))),Array.from(e.childNodes).map(canonical).filter(x=>x!==undefined)]}
  if(node.nodeType===3||node.nodeType===4)return node.nodeValue?.trim() || (node.parentNode as Element)?.localName==="t"?["text",node.nodeValue]:undefined
  if(node.nodeType===7||node.nodeType===8)return [node.nodeType,node.nodeName,node.nodeValue]
  return undefined
}
function parse(xml:string) {
  if(typeof xml!=="string" || xml.length>16_000_000 || /<!DOCTYPE/i.test(xml))return fail()
  const reject=()=>{throw Error("Invalid Word XML")};const doc=new DOMParser({errorHandler:{warning:reject,error:reject,fatalError:reject}}).parseFromString(xml,"application/xml")
  if(doc.doctype || doc.documentElement.namespaceURI!==PKG || doc.documentElement.localName!=="package")return fail()
  const parts=new Map<string,Element>()
  for(const part of elems(doc.documentElement)){const name=part.getAttributeNS(PKG,"name");if(part.namespaceURI!==PKG||part.localName!=="part"||!name||parts.has(name)||parts.size>=64)return fail();parts.set(name,part)}
  if(!parts.has(MAIN))return fail();return {doc,parts}
}
function normalizedComment(comment:Element) {
  const copy=comment.cloneNode(true) as Element;copy.removeAttributeNS(W,"id")
  for(const p of Array.from(copy.getElementsByTagNameNS(W,"p")))for(const name of ["rsidR","rsidRDefault"]){const value=p.getAttributeNS(W,name);if(p.hasAttributeNS(W,name)){if(typeof value!=="string"||!/^[0-9a-fA-F]{8}$/.test(value))return fail();p.removeAttributeNS(W,name)}}
  return JSON.stringify(canonical(copy))
}
function graph(parts:Map<string,Element>) {
  const body=Array.from(parts.get(MAIN)!.getElementsByTagNameNS(W,"body"));if(body.length!==1)return fail()
  for(const [name,part] of parts)if(name!==MAIN&&name!==COMMENTS)for(const kind of allKinds)if(part.getElementsByTagNameNS(W,kind).length)return fail()
  const anchors=new Map<string,Map<string,{path:string;node:Element}>>()
  const bookmarks=new Map<string,Map<string,{path:string;node:Element}>>()
  let paragraph=-1, offset=0, active=false, ordinal=0, lastPosition=""
  const unsupportedParagraphs=new Set<number>()
  const walk=(node:Element, deleted=false)=>{
    const word=node.namespaceURI===W, kind=node.localName
    if(word&&kind==="p"){if(active)return fail();paragraph++;offset=0;active=true}
    const hidden=deleted || word&&["del","moveFrom"].includes(kind)
    if(active&&word&&["instrText","fldChar","drawing","pict","object","sym","footnoteReference","endnoteReference","noBreakHyphen","softHyphen"].includes(kind))unsupportedParagraphs.add(paragraph)
    if(word&&allKinds.has(kind)){
      if(!active||hidden)return fail()
      const id=node.getAttributeNS(W,"id");if(!id||!/^\d+$/.test(id))return fail()
      const position=`${paragraph}:${offset}`;if(position!==lastPosition){lastPosition=position;ordinal=0}const path=`${position}:${ordinal++}`
      const owner=kinds.has(kind)?anchors:bookmarks, map=owner.get(id)??new Map();if(map.has(kind))return fail();map.set(kind,{path,node});owner.set(id,map)
    }
    if(word&&!hidden&&kind==="t")offset+=(node.textContent??"").length
    else if(word&&!hidden&&["tab","br","cr"].includes(kind))offset++
    else for(const child of elems(node))walk(child,hidden)
    if(word&&kind==="p")active=false
  }
  walk(body[0]!)
  for(const kind of allKinds)if(parts.get(MAIN)!.getElementsByTagNameNS(W,kind).length!==body[0]!.getElementsByTagNameNS(W,kind).length)return fail()
  for(const map of [anchors,bookmarks])for(const pair of map.values())for(const item of pair.values())if(unsupportedParagraphs.has(Number(item.path.split(":")[0])))return fail()
  const commentsPart=parts.get(COMMENTS),comments=commentsPart?Array.from(commentsPart.getElementsByTagNameNS(W,"comment")):[]
  const result=new Map<string,{comment:Element;anchors:Map<string,{path:string;node:Element}>;payload:string}>()
  for(const comment of comments){const id=comment.getAttributeNS(W,"id");if(!id||!/^\d+$/.test(id)||result.has(id))return fail();const references=anchors.get(id);if(!references||references.size!==3||[...kinds].some(k=>!references.has(k)))return fail();result.set(id,{comment,anchors:references,payload:normalizedComment(comment)})}
  if(anchors.size!==result.size)return fail()
  for(const pair of bookmarks.values())if(pair.size!==2 || [...bookmarkKinds].some(k=>!pair.has(k)) || !pair.get("bookmarkStart")!.node.getAttributeNS(W,"name"))return fail()
  return {comments:result,bookmarks}
}
/** Map only fully matched comment and bookmark identities and their body anchors. Other
 * package parts are not modified and must be verified by the caller. */
export function normalizeNativeWordChatCommentIds(beforeXml:string,afterXml:string):string {
  const before=parse(beforeXml),after=parse(afterXml),left=graph(before.parts),right=graph(after.parts),a=left.comments,b=right.comments
  if(a.size!==b.size || before.parts.has(COMMENTS)!==after.parts.has(COMMENTS))return fail()
  const used=new Set<string>()
  for(const [oldId,old] of a){
    const matches=[...b.entries()].filter(([,candidate])=>candidate.payload===old.payload && [...kinds].every(kind=>{
      const left=old.anchors.get(kind)!,right=candidate.anchors.get(kind)!;if(left.path!==right.path)return false
      const l=left.node.cloneNode(true) as Element,r=right.node.cloneNode(true) as Element;l.removeAttributeNS(W,"id");r.removeAttributeNS(W,"id");return JSON.stringify(canonical(l))===JSON.stringify(canonical(r))
    }))
    if(matches.length!==1||used.has(matches[0]![0]))return fail()
    const [newId,target]=matches[0]!;used.add(newId)
    const oldParas=Array.from(old.comment.getElementsByTagNameNS(W,"p")),newParas=Array.from(target.comment.getElementsByTagNameNS(W,"p"))
    if(oldParas.length!==newParas.length)return fail()
    for(let i=0;i<oldParas.length;i++) for(const key of ["rsidR","rsidRDefault"]) {
      const oldAttr=oldParas[i]!.getAttributeNodeNS(W,key), newAttr=newParas[i]!.getAttributeNodeNS(W,key)
      if(Boolean(oldAttr)!==Boolean(newAttr))return fail()
      if(oldAttr&&newAttr){const name=newAttr.name;newParas[i]!.removeAttributeNS(W,key);newParas[i]!.setAttributeNS(W,name,oldAttr.value)}
    }
    for(const node of [target.comment,...[...target.anchors.values()].map(x=>x.node)]){const name=node.getAttributeNodeNS(W,"id")!.name;node.removeAttributeNS(W,"id");node.setAttributeNS(W,name,oldId)}
  }
  if(left.bookmarks.size!==right.bookmarks.size)return fail()
  const usedBookmarks=new Set<string>()
  for(const [oldId,old] of left.bookmarks) {
    const matches=[...right.bookmarks.entries()].filter(([,candidate])=>[...bookmarkKinds].every(kind=>{
      const l=old.get(kind)!,r=candidate.get(kind)!;if(l.path!==r.path)return false
      const x=l.node.cloneNode(true) as Element,y=r.node.cloneNode(true) as Element;x.removeAttributeNS(W,"id");y.removeAttributeNS(W,"id");return JSON.stringify(canonical(x))===JSON.stringify(canonical(y))
    }))
    if(matches.length!==1 || usedBookmarks.has(matches[0]![0]))return fail()
    const [newId,pair]=matches[0]!;usedBookmarks.add(newId)
    for(const {node} of pair.values()){const name=node.getAttributeNodeNS(W,"id")!.name;node.removeAttributeNS(W,"id");node.setAttributeNS(W,name,oldId)}
  }
  if(before.parts.has(COMMENTS) && JSON.stringify(canonical(before.parts.get(COMMENTS)!))!==JSON.stringify(canonical(after.parts.get(COMMENTS)!)))return fail()
  return new XMLSerializer().serializeToString(after.doc)
}
