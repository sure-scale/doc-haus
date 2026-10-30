import { DOMParser } from "@xmldom/xmldom"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const W16DU = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const XML = "http://www.w3.org/XML/1998/namespace"
const MAIN = "/word/document.xml"
const elements = (n:Node) => Array.from(n.childNodes).filter((c):c is Element => c.nodeType === 1)
const named = (n:Node, name:string) => elements(n).filter(e => e.namespaceURI === W && e.localName === name)
const only = (n:Node, names:string[]) => elements(n).every(e => e.namespaceURI === W && names.includes(e.localName))
const fail = ():never => { throw new Error("Native inserted row could not be verified") }
const canon = (node:Node):unknown => {
  if (node.nodeType === 1) {
    const e=node as Element
    return [e.namespaceURI,e.localName,Array.from(e.attributes).filter(a=>a.namespaceURI!==XMLNS).map(a=>[a.namespaceURI,a.localName,a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b))),Array.from(e.childNodes).map(canon).filter(x=>x!==undefined)]
  }
  if (node.nodeType===3 || node.nodeType===4) return node.nodeValue?.trim() || (node.parentNode as Element)?.localName === "t" ? ["text",node.nodeValue] : undefined
  if (node.nodeType===7 || node.nodeType===8) return [node.nodeType,node.nodeName,node.nodeValue]
  return undefined
}
function parts(xml:string) {
  if (xml.length>16_000_000 || /<!DOCTYPE/i.test(xml)) return fail()
  const reject=()=>{throw Error("Invalid native XML")}
  const doc=new DOMParser({errorHandler:{warning:reject,error:reject,fatalError:reject}}).parseFromString(xml,"application/xml")
  if (doc.doctype || doc.documentElement.namespaceURI!==PKG || doc.documentElement.localName!=="package") return fail()
  const result=new Map<string,Element>()
  for(const p of elements(doc.documentElement)) { const name=p.getAttributeNS(PKG,"name"); if(p.namespaceURI!==PKG || p.localName!=="part" || !name || result.has(name) || result.size>=64) return fail(); result.set(name,p) }
  if(!result.has(MAIN)) return fail()
  return result
}
/** Proves the new row's native carrier ownership and literal plain runs, and
 * exact available non-main Flat OPC dependencies. Caller separately proves
 * old-body remainder, fresh binding, current native fonts and full DOCX. */
export function verifyNativeWordChatInsertedRow(beforeXml:string, afterXml:string, tableIndex:number, priorRows:number, cells:string[]): {authors:string[]} {
  const before=parts(beforeXml), after=parts(afterXml)
  if(JSON.stringify([...before.keys()].sort())!==JSON.stringify([...after.keys()].sort())) return fail()
  const dependency = (name:string, part:Element) => {
    const copy=part.cloneNode(true) as Element
    if(name === "/word/comments.xml") for(const paragraph of Array.from(copy.getElementsByTagNameNS(W,"p"))) for(const attr of ["rsidR","rsidRDefault"]) {
      const value=paragraph.getAttributeNS(W,attr); if(value!==null) { if(!/^[0-9A-Fa-f]{8}$/.test(value)) return fail();paragraph.removeAttributeNS(W,attr) }
    }
    if(name === "/word/settings.xml") {
      const inventories=Array.from(copy.getElementsByTagNameNS(W,"rsids")); if(inventories.length>1) return fail()
      for(const inventory of inventories) {
        if(Array.from(inventory.attributes).some(a=>a.namespaceURI!==XMLNS) || !only(inventory,["rsid","rsidRoot"]) || elements(inventory).some(e=>elements(e).length || !/^[0-9A-Fa-f]{8}$/.test(e.getAttributeNS(W,"val")??"") || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && a.localName==="val")))) return fail()
        inventory.parentNode!.removeChild(inventory)
      }
    }
    return canon(copy)
  }
  for(const [name,part] of before) if(name!==MAIN) {
    const other=after.get(name)!
    if(name === "/word/comments.xml") {
      const keys=(p:Element)=>Array.from(p.getElementsByTagNameNS(W,"p")).map(e=>["rsidR","rsidRDefault"].filter(k=>e.hasAttributeNS(W,k)))
      if(JSON.stringify(keys(part))!==JSON.stringify(keys(other))) return fail()
    }
    if(name === "/word/settings.xml" && part.getElementsByTagNameNS(W,"rsids").length!==other.getElementsByTagNameNS(W,"rsids").length) return fail()
    if(JSON.stringify(dependency(name,part))!==JSON.stringify(dependency(name,other))) return fail()
  }
  const main=after.get(MAIN)!, bodies=Array.from(main.getElementsByTagNameNS(W,"body")), oldBodies=Array.from(before.get(MAIN)!.getElementsByTagNameNS(W,"body"))
  if(bodies.length!==1 || oldBodies.length!==1 || !Number.isSafeInteger(tableIndex) || tableIndex<0 || !Number.isSafeInteger(priorRows) || priorRows<1 || !cells.length || cells.length>32) return fail()
  const tables=named(bodies[0]!,"tbl"), oldTables=named(oldBodies[0]!,"tbl"); if(tableIndex>=tables.length || oldTables.length!==tables.length) return fail()
  const rows=named(tables[tableIndex]!,"tr"), oldRows=named(oldTables[tableIndex]!,"tr"); if(rows.length!==priorRows+1 || oldRows.length!==priorRows) return fail()
  const referenceCells=named(oldRows.at(-1)!,"tc"); if(referenceCells.length!==cells.length) return fail()
  const metadata = (element:Element) => Array.from(element.attributes).every(a => a.namespaceURI===XMLNS ||
    a.namespaceURI===W && ["rsidR","rsidRPr","rsidRDefault"].includes(a.localName) && /^[0-9A-Fa-f]{8}$/.test(a.value) ||
    a.namespaceURI===W14 && ["paraId","textId"].includes(a.localName) && /^[0-9A-Fa-f]{8}$/.test(a.value))
  const row=rows.at(-1)!, rowCells=named(row,"tc"), rowPr=named(row,"trPr"), marks=rowPr.length===1?named(rowPr[0]!,"ins"):[]
  if(!metadata(row) || !only(row,["trPr","tc"]) || rowCells.length!==cells.length || rowPr.length!==1 || !only(rowPr[0]!,["ins"]) || marks.length!==1 || marks[0]!.childNodes.length) return fail()
  const ids=new Set<string>(), authors=new Set<string>(); let stamp:string|undefined
  const mark=(e:Element) => {
    const id=e.getAttributeNS(W,"id"), author=e.getAttributeNS(W,"author"), date=e.getAttributeNS(W,"date"), utc=e.getAttributeNS(W16DU,"dateUtc")
    if(!id || !/^\d+$/.test(id) || ids.has(id) || !author?.trim() || !date || Number.isNaN(Date.parse(date)) || utc && utc!==date || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && ["id","author","date"].includes(a.localName)) && !(a.namespaceURI===W16DU && a.localName==="dateUtc"))) return fail()
    ids.add(id);authors.add(author);const key=JSON.stringify([author,date,utc??null]);if(stamp && stamp!==key) return fail();stamp=key
  }
  mark(marks[0]!)
  const noAttrs=(p:Element) => Array.from(p.attributes).every(a=>a.namespaceURI===XMLNS)
  const basic=(p:Element) => noAttrs(p) && only(p,["b","bCs","i","iCs","u"]) && elements(p).every(e=>!elements(e).length && Array.from(e.attributes).every(a=>a.namespaceURI===XMLNS || a.namespaceURI===W && a.localName==="val"))
  for(let i=0;i<cells.length;i++) {
    const cell=rowCells[i]!, ps=named(cell,"p"), cellProps=named(cell,"tcPr"), referenceProps=named(referenceCells[i]!,"tcPr")
    if(cellProps.length>1 || cellProps.length!==referenceProps.length || cellProps.length && JSON.stringify(canon(cellProps[0]!))!==JSON.stringify(canon(referenceProps[0]!)) ||
       [cell,...referenceCells.slice(i,i+1)].some(c=>["gridSpan","vMerge","hMerge"].some(n=>c.getElementsByTagNameNS(W,n).length))) return fail()
    if(!metadata(cell) || !only(cell,["tcPr","p"]) || ps.length!==1 || !metadata(ps[0]!) || !only(ps[0]!,["pPr","ins"])) return fail()
    const p=ps[0]!, pp=named(p,"pPr"), pr=pp.length===1?named(pp[0]!,"rPr"):[], pm=pr.length===1?named(pr[0]!,"ins"):[]
    if(pp.length!==1 || !noAttrs(pp[0]!) || pr.length!==1 || !noAttrs(pr[0]!) || pm.length!==1 || pm[0]!.childNodes.length || !only(pp[0]!,["rPr"]) || !only(pr[0]!,["ins","b","bCs","i","iCs","u"])) return fail()
    if(elements(pr[0]!).some(e=>e.localName!=="ins" && (!only(e,[]) || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && a.localName==="val"))))) return fail()
    mark(pm[0]!)
    let value=""
    for(const ins of named(p,"ins")) {
      mark(ins);if(!only(ins,["r"])) return fail()
      for(const run of named(ins,"r")) {
        if(!metadata(run) || !only(run,["rPr","t"])) return fail()
        const props=named(run,"rPr"), ts=named(run,"t"); if(props.length>1 || ts.length!==1) return fail()
        if(props.length) {
          const r=props[0]!, changes=named(r,"rPrChange")
          if(!noAttrs(r) || !only(r,["b","bCs","i","iCs","u","rPrChange"]) || changes.length>1 || elements(r).some(e=>e.localName!=="rPrChange" && (!(["0","false","off","none"].includes(e.getAttributeNS(W,"val")??"")) || !only(e,[]) || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && a.localName==="val"))))) return fail()
          for(const change of changes) { mark(change);const old=named(change,"rPr");if(old.length!==1 || !only(change,["rPr"]) || !basic(old[0]!)) return fail() }
        }
        const t=ts[0]!;if(elements(t).length || Array.from(t.childNodes).some(n=>n.nodeType!==3) || Array.from(t.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===XML && a.localName==="space" && a.value==="preserve"))) return fail();value+=t.textContent??""
      }
    }
    if(value!==cells[i] || cells[i] && !named(p,"ins").length) return fail()
  }
  // Reject comments, PIs or non-whitespace text outside literal text runs.
  const clean=(n:Element):boolean=>Array.from(n.childNodes).every(c=>c.nodeType===1?clean(c as Element):c.nodeType===3 && (n.namespaceURI===W && n.localName==="t" || !c.nodeValue?.trim()))
  if(!clean(row)) return fail()
  return {authors:[...authors]}
}

export function verifyNativeWordChatAddedRow(beforeXml:string, afterXml:string, binding:{tableIndex:number;rows:unknown[]}, cells:string[]): string[] {
  return verifyNativeWordChatInsertedRow(beforeXml, afterXml, binding.tableIndex, binding.rows.length, cells).authors
}
