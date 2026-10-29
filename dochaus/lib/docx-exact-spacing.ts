import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { openRevision, type DocxTableContext } from "./docx-amendments"

// Offline proof only: no route or Word transport calls this function.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const BODY = "word/document.xml"
const hash = (v: Uint8Array) => createHash("sha256").update(v).digest("hex")
const requireProof = (v: unknown, code: string): asserts v => { if (!v) throw new Error(`docx_exact_spacing:${code}`) }
const elements = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const all = (n: Document | Element, name: string) => Array.from(n.getElementsByTagNameNS(W, name))
const child = (n: Element, name: string) => elements(n).filter(c => c.namespaceURI === W && c.localName === name)
function parse(bytes: Uint8Array) {
  const xml = strFromU8(bytes)
  requireProof(!/<!DOCTYPE|<!ENTITY/i.test(xml), "unsafe_xml")
  const fail = () => { throw new Error("docx_exact_spacing:invalid_xml") }
  const doc = new DOMParser({errorHandler:{warning:fail,error:fail,fatalError:fail}}).parseFromString(xml,"application/xml")
  for (const el of Array.from(doc.getElementsByTagName("*"))) for (const attr of Array.from(el.attributes)) {
    const qualified = attr.namespaceURI === MC && ["Ignorable","MustUnderstand","ProcessContent","PreserveElements","PreserveAttributes"].includes(attr.localName)
    const choiceRequires = el.namespaceURI === MC && el.localName === "Choice" && !attr.namespaceURI && attr.localName === "Requires"
    if (!qualified && !choiceRequires) continue
    const prefixList = choiceRequires || ["Ignorable","MustUnderstand"].includes(attr.localName)
    const tokens=attr.value.trim().split(/\s+/).filter(Boolean)
    requireProof(tokens.length > 0,"invalid_mc_token")
    for (const token of tokens) {
      const nc = "[A-Za-z_][A-Za-z0-9_.-]*"
      const valid = prefixList ? new RegExp(`^${nc}$`).test(token) : new RegExp(`^${nc}:${nc}$`).test(token)
      requireProof(valid,"invalid_mc_token")
      const prefix = token.split(":")[0]!
      requireProof(!!el.lookupNamespaceURI(prefix), "unbound_mc_prefix")
    }
  }
  return doc
}
function canonical(node: Node): string {
  if (node.nodeType === 9) return JSON.stringify(Array.from(node.childNodes).filter(n => !(n.nodeType === 7 && n.nodeName === "xml")).map(canonical))
  if (node.nodeType !== 1) return JSON.stringify([node.nodeType,node.nodeName,node.nodeValue])
  const el = node as Element
  return JSON.stringify([el.namespaceURI,el.nodeName,Array.from(el.attributes).map(a => [a.namespaceURI,a.name,a.value]).sort(),Array.from(el.childNodes).map(canonical)])
}
const same = (a: Node,b: Node) => canonical(a) === canonical(b)
const serialize = (doc: Document) => strToU8(new XMLSerializer().serializeToString(doc))
type Package = ReturnType<typeof unpack>
function unpack(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  requireProof(parts[BODY] && parts["word/settings.xml"],"missing_parts")
  const docs = Object.fromEntries(Object.entries(parts).filter(([n]) => /\.(xml|rels)$/.test(n)).map(([n,b]) => [n,parse(b)]))
  return {parts,docs,body:docs[BODY]!}
}
function samePackage(a: Package,b: Package) {
  requireProof(JSON.stringify(Object.keys(a.parts).sort()) === JSON.stringify(Object.keys(b.parts).sort()),"inventory_changed")
  for (const name of Object.keys(a.parts)) if (name !== BODY) {
    requireProof(a.docs[name] ? !!b.docs[name] && same(a.docs[name]!,b.docs[name]!) : hash(a.parts[name]!) === hash(b.parts[name]!),"other_part_changed")
  }
}
const ppr = (p: Element) => { const pps=child(p,"pPr"); requireProof(pps.length <= 1,"ambiguous_ppr"); return pps[0] }
function spacing(p: Element) { const pp=ppr(p); const rows=pp ? child(pp,"spacing") : []; requireProof(rows.length <= 1,"ambiguous_spacing"); return rows[0] }
const ranges = new Map([["bookmarkStart","bookmark"],["bookmarkEnd","bookmark"],["commentRangeStart","comment"],["commentRangeEnd","comment"],["permStart","perm"],["permEnd","perm"]])
function protect(body: Document, target: Element) {
  const active = new Set<string>(), seen = new Set<string>(), bookmarkNames = new Set<string>(); let inside = false
  function walk(node: Node) {
    if (node === target) { requireProof(active.size === 0,"protected_range"); inside = true }
    if (node.nodeType === 1) {
      const e = node as Element
      if (e.namespaceURI === W && /Range(?:Start|End)$/.test(e.localName))
        requireProof(ranges.has(e.localName),"unsupported_range")
      if (e.namespaceURI === W && ranges.has(e.localName)) {
        if (e.localName === "bookmarkStart") {
          const name=e.getAttributeNS(W,"name")
          requireProof(typeof name === "string" && name.length > 0,"invalid_bookmark_name")
          const normalizedName=name.toLowerCase()
          requireProof(!bookmarkNames.has(normalizedName),"invalid_bookmark_name")
          bookmarkNames.add(normalizedName)
        }
        const id=e.getAttributeNS(W,"id")
        requireProof(/^\d+$/.test(id),"invalid_range")
        const key=`${ranges.get(e.localName)}:${BigInt(id).toString()}`
        if (inside) {
          requireProof(ranges.get(e.localName) === "bookmark" && e.parentNode === target && !e.childNodes?.length &&
            Array.from(e.attributes).every(a => a.namespaceURI === W && (a.localName === "id" || e.localName === "bookmarkStart" && a.localName === "name")) &&
            (e.localName !== "bookmarkStart" || !!e.getAttributeNS(W,"name")),"protected_range")
        }
        if (e.localName.endsWith("Start")) { requireProof(!active.has(key) && !seen.has(key),"invalid_range"); active.add(key); seen.add(key) }
        else { requireProof(active.delete(key),"invalid_range") }
      }
    }
    for (const c of Array.from(node.childNodes ?? [])) walk(c)
    if (node === target) { requireProof(active.size === 0,"protected_range"); inside = false }
  }
  walk(body); requireProof(active.size === 0,"invalid_range")
  for (let n: Node | null = target; n && n.nodeType !== 9; n=n.parentNode)
    requireProof(!(n.nodeType === 1 && (n as Element).namespaceURI === W && ["sdt","customXml","ins","del","moveFrom","moveTo"].includes((n as Element).localName)),"protected_ancestor")
  for (const name of ["sdt","fldSimple","fldChar","instrText","pPrChange","ins","del","moveFrom","moveTo","commentReference"]) requireProof(!all(target,name).length,"protected_target")
}
function maskAfter(p: Element) {
  const copy=p.cloneNode(true) as Element, s=spacing(copy)
  if (s) s.removeAttributeNS(W,"after")
  return copy
}
function verifyBody(source: Document,candidate: Document,ordinal: number,value: number) {
  const a=all(source,"p")[ordinal], b=all(candidate,"p")[ordinal]
  requireProof(a && b,"target_missing")
  requireProof(spacing(b)?.getAttributeNS(W,"after") === String(value),"wrong_after")
  requireProof(same(maskAfter(a),maskAfter(b)),"unapproved_paragraph_change")
  const clone=candidate.cloneNode(true) as Document, selected=all(clone,"p")[ordinal]!
  selected.parentNode!.replaceChild(clone.importNode(a,true),selected)
  requireProof(same(source,clone),"other_body_changed")
}
function allocate(pkg: Package) {
  const used = new Set<string>()
  for (const doc of Object.values(pkg.docs)) for (const e of Array.from(doc.getElementsByTagName("*"))) if (e.hasAttributeNS(W,"id")) {
    const id=e.getAttributeNS(W,"id"); if (/^\+?\d+$/.test(id)) used.add(BigInt(id).toString())
    else if (["ins","del","moveFrom","moveTo","pPrChange","rPrChange","tblPrChange","tcPrChange","trPrChange"].includes(e.localName) && e.namespaceURI === W) throw new Error("docx_exact_spacing:invalid_revision_id")
  }
  for (let n=0;n<=used.size && n<=2147483647;n++) if (!used.has(String(n))) return String(n)
  throw new Error("docx_exact_spacing:no_revision_id")
}
export type ExactSpacingRequest = { sourceSha256: string; paragraphHandle: string; paragraphText: string; tableContext: DocxTableContext | null; spacingAfterPoints: number; author: string }
export async function authorExactParagraphSpacing(input: Uint8Array, request: ExactSpacingRequest) {
  const source = new Uint8Array(input), pkg=unpack(source)
  requireProof(/^[a-f0-9]{64}$/.test(request.sourceSha256) && hash(source) === request.sourceSha256,"stale_source")
  requireProof(request.author.trim() && Number.isFinite(request.spacingAfterPoints) && request.spacingAfterPoints >= 0 && request.spacingAfterPoints <= 48 && Number.isInteger(request.spacingAfterPoints*20),"invalid_request")
  requireProof(!all(pkg.docs["word/settings.xml"]!,"documentProtection").length,"protected_document")
  for (const doc of Object.values(pkg.docs)) for (const name of ["contextualSpacing","snapToGrid","docGrid"]) requireProof(!all(doc,name).length,"unresolved_spacing")
  for (const doc of Object.values(pkg.docs)) for (const s of all(doc,"spacing")) requireProof(!s.hasAttributeNS(W,"afterAutospacing"),"unresolved_spacing")
  const opened=await openRevision(source,{pendingRevisions:"layer"})
  try {
    const indices=opened.revision.paragraphs.flatMap((p,i) => p.paragraph_handle === request.paragraphHandle ? [i] : [])
    requireProof(indices.length === 1,"unknown_handle")
    const index=indices[0]!, view=opened.revision.paragraphs[index]!, target=opened.paragraphs[index]!
    requireProof(view.text === request.paragraphText && JSON.stringify(view.table) === JSON.stringify(request.tableContext),"target_changed")
    requireProof(target.part === BODY && !target.pending && (!view.table || !view.table.unavailable_reason && view.table.row_span === 1 && view.table.column_span === 1 && view.table.vertical_merge === "none"),"unsupported_target")
    const ordinal=all(opened.documents[BODY]!,"p").indexOf(target.node), para=all(pkg.body,"p")[ordinal]
    requireProof(ordinal>=0 && para,"target_missing"); protect(pkg.body,para)
    const table=para.parentNode?.parentNode?.parentNode as Element | undefined
    if (view.table) requireProof(table?.namespaceURI === W && table.localName === "tbl" && !all(table,"tblStyle").length,"unsupported_table")
    const old=ppr(para), before=spacing(para)?.getAttributeNS(W,"after") ?? null
    requireProof(old && spacing(para),"unsupported_empty_properties")
    const desired=request.spacingAfterPoints*20
    requireProof(before !== String(desired),"no_change")
    const prior=opened.session.listRevisions(), priorIds=new Set(prior.map(x=>x.id)); requireProof(priorIds.size === prior.length,"ambiguous_revisions")
    const set=opened.session.setParagraphFormat(target.match.enclosingAnchor.id,{spacingAfter:desired}); requireProof(set.success,"sdk_refused")
    const bakedCandidate=unpack(opened.session.save()); samePackage(pkg,bakedCandidate); verifyBody(pkg.body,bakedCandidate.body,ordinal,desired)
    const baked=zipSync({...pkg.parts,[BODY]:bakedCandidate.parts[BODY]!})
    const trackedDoc=parse(bakedCandidate.parts[BODY]!), trackedPara=all(trackedDoc,"p")[ordinal]!, current=ppr(trackedPara)!, id=allocate(pkg)
    const mark=trackedDoc.createElementNS(W,"w:pPrChange"); mark.setAttributeNS(W,"w:id",id); mark.setAttributeNS(W,"w:author",request.author); mark.setAttributeNS(W,"w:date",new Date().toISOString()); mark.appendChild(trackedDoc.importNode(old,true)); current.appendChild(mark)
    const tracked=zipSync({...pkg.parts,[BODY]:serialize(trackedDoc)}), dx=await docxodus(), probe=dx.openDocxSession(tracked,{persistAnchorIds:false})
    let revisionId:string
    try {
      const rows=probe.listRevisions(), added=rows.filter(x=>!priorIds.has(x.id)); requireProof(rows.length === prior.length+1 && added.length === 1,"revision_membership")
      const rev=added[0]!; requireProof(rev.type === "format" && rev.author === request.author && rev.partUri === "/word/document.xml" && rev.constituentIds.includes(id) && rev.constituentIds.length === 1 && rev.constituentKeys.length === 1,"revision_membership")
      const oldNow=rows.filter(x=>priorIds.has(x.id)); const clean=(rs:typeof rows)=>rs.map(({anchorId:_a,affectedAnchors:_b,...r})=>r).sort((a,b)=>a.id.localeCompare(b.id)); requireProof(JSON.stringify(clean(prior))===JSON.stringify(clean(oldNow)),"prior_revision_changed")
      revisionId=rev.id
    } finally {probe.close()}
    for (const accept of [false,true]) {
      const session=dx.openDocxSession(tracked,{persistAnchorIds:false})
      try {
        const result=accept?session.acceptRevision(revisionId):session.rejectRevision(revisionId); requireProof(result.success,"resolution_failed")
        const resolved=unpack(session.save()), expected=accept?bakedCandidate:pkg; samePackage(expected,resolved); requireProof(same(expected.body,resolved.body),"resolution_changed")
        const rows=session.listRevisions(); const clean=(rs:typeof rows)=>rs.map(({anchorId:_a,affectedAnchors:_b,...r})=>r).sort((a,b)=>a.id.localeCompare(b.id)); requireProof(JSON.stringify(clean(prior))===JSON.stringify(clean(rows)),"prior_revision_changed")
      } finally {session.close()}
    }
    return {trackedBytes:tracked,bakedBytes:baked,revisionId,beforeDirect:before,afterDirect:String(desired)}
  } finally {opened.session.close()}
}
