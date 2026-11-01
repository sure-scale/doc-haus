import { DOMParser } from "@xmldom/xmldom"
import type { WordChatCellRun } from "../../../dochaus/lib/word-chat-amendment-contract"
import { nativeWordChatPeopleGrowth } from "./word-chat-native-people"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const W16DU = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const XML = "http://www.w3.org/XML/1998/namespace"
const MAIN = "/word/document.xml"
const RELS = "/word/_rels/document.xml.rels"
const PEOPLE = "/word/people.xml"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const PEOPLE_TYPE = "http://schemas.microsoft.com/office/2011/relationships/people"
const THEME_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
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
const RECORDED: Record<string,string> = {tblPrChange:"tblPr", tblGridChange:"tblGrid", trPrChange:"trPr", tcPrChange:"tcPr"}
/** With Track Formatting on, Word records the table's, grid's, rows' and
 * cells' properties when it appends a row, even where none changed, and may
 * fix an automatic table width at the grid total. Remove records whose old
 * properties equal the current ones and restore that one width; any other
 * recorded change stays for the caller's comparison to reject. */
export function settleNativeWordChatTableRecords(table:Element) {
  const holders=[...named(table,"tblPr"),...named(table,"tblGrid")]
  for(const row of named(table,"tr")) holders.push(...named(row,"trPr"),...named(row,"tc").flatMap(c=>named(c,"tcPr")))
  const props=(e:Element)=>JSON.stringify(elements(e).filter(c=>!(c.namespaceURI===W && ["ins","del",...Object.keys(RECORDED)].includes(c.localName))).map(canon))
  for(const holder of holders) {
    const records=elements(holder).filter(e=>e.namespaceURI===W && RECORDED[e.localName]===holder.localName)
    if(records.length!==1) continue
    const record=records[0]!, recorded=elements(record)
    if(recorded.length!==1 || recorded[0]!.namespaceURI!==W || recorded[0]!.localName!==holder.localName || Array.from(recorded[0]!.attributes).some(a=>a.namespaceURI!==XMLNS) ||
       Array.from(record.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && ["id","author","date"].includes(a.localName)) && !(a.namespaceURI===W16DU && a.localName==="dateUtc"))) continue
    const old=recorded[0]!
    if(props(holder)===props(old)) { holder.removeChild(record); continue }
    if(holder.localName!=="tblPr") continue
    const width=named(holder,"tblW"), oldWidth=named(old,"tblW"), grid=named(table,"tblGrid")
    if(width.length!==1 || oldWidth.length!==1 || grid.length!==1 || oldWidth[0]!.getAttributeNS(W,"type")!=="auto" || width[0]!.getAttributeNS(W,"type")!=="dxa" ||
       Array.from(width[0]!.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && ["w","type"].includes(a.localName)))) continue
    const total=named(grid[0]!,"gridCol").reduce((sum,col)=>sum+Number(col.getAttributeNS(W,"w")),0)
    if(!Number.isSafeInteger(total) || String(total)!==width[0]!.getAttributeNS(W,"w")) continue
    const restored=oldWidth[0]!.cloneNode(true) as Element
    holder.replaceChild(restored,width[0]!)
    if(props(holder)===props(old)) holder.removeChild(record)
    else holder.replaceChild(width[0]!,restored)
  }
}
/** Direct run flags: only b/bCs/i/iCs/u with a bare val. A toggle is on when
 * present without val or with an on value; a complex-script counterpart must be
 * absent or agree; underline is exactly single or none. Anything else fails. */
function directFlags(props:Element|undefined): Omit<WordChatCellRun,"text"> {
  const direct=props ? elements(props).filter(e=>e.localName!=="rPrChange") : []
  if(new Set(direct.map(e=>e.localName)).size!==direct.length || direct.some(e=>!only(e,[]) || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && a.localName==="val")))) return fail()
  const find=(name:string)=>direct.find(e=>e.localName===name)
  const toggle=(e:Element|undefined)=>{
    if(!e) return undefined
    if(!e.hasAttributeNS(W,"val")) return true
    const value=e.getAttributeNS(W,"val")??""
    return ["1","true","on"].includes(value) ? true : ["0","false","off"].includes(value) ? false : fail()
  }
  const bold=toggle(find("b"))??false, boldCs=toggle(find("bCs")), italic=toggle(find("i"))??false, italicCs=toggle(find("iCs")), u=find("u")
  if(boldCs!==undefined && boldCs!==bold || italicCs!==undefined && italicCs!==italic || u && !u.hasAttributeNS(W,"val")) return fail()
  const underline=!u || u.getAttributeNS(W,"val")==="none" ? false : u.getAttributeNS(W,"val")==="single" ? true : fail()
  return {bold,italic,underline}
}
/** Proves the new row's native carrier ownership, literal runs whose direct
 * flags group into exactly the reviewed spans (plain rows: one off span per
 * cell), and exact available non-main Flat OPC dependencies. Word may split a
 * span across runs. Caller separately proves old-body remainder, fresh
 * binding, current native fonts and full DOCX. */
export function verifyNativeWordChatInsertedRow(beforeXml:string, afterXml:string, tableIndex:number, priorRows:number, cells:string[], runs?:WordChatCellRun[][]): {authors:string[]} {
  const before=parts(beforeXml), after=parts(afterXml)
  // A first revision by this Word user adds its people part and relationship.
  const extra=[...after.keys()].filter(name=>!before.has(name)), people=extra.length===1 && extra[0]===PEOPLE ? after.get(PEOPLE)! : undefined
  if([...before.keys()].some(name=>!after.has(name)) || extra.length && !people) return fail()
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
    if(people && name === RELS) { peopleRelationships(part,other); continue }
    if(name === PEOPLE) continue
    if(JSON.stringify(dependency(name,part))!==JSON.stringify(dependency(name,other))) return fail()
  }
  /** Word adds exactly one people relationship and may renumber only the
   * unreferenced theme relationship to make room for it. */
  function peopleRelationships(left:Element, right:Element) {
    const referenced=(map:Map<string,Element>, id:string) => [...map.values()].some(part=>Array.from(part.getElementsByTagName("*")).some(e=>Array.from(e.attributes).some(a=>a.namespaceURI===R && a.value===id)))
    const all=(part:Element)=>Array.from(part.getElementsByTagNameNS(REL,"Relationship")), l=all(left), added=all(right).filter(e=>e.getAttribute("Type")===PEOPLE_TYPE), r=all(right).filter(e=>!added.includes(e))
    const ids=(items:Element[])=>items.every(e=>Boolean(e.getAttribute("Id"))) && new Set(items.map(e=>e.getAttribute("Id"))).size===items.length
    if(l.some(e=>e.getAttribute("Type")===PEOPLE_TYPE) || added.length!==1 || added[0]!.getAttribute("Target")!=="people.xml" || elements(added[0]!).length ||
       Array.from(added[0]!.attributes).some(a=>a.namespaceURI!==XMLNS && (Boolean(a.namespaceURI) || !["Id","Type","Target"].includes(a.name))) || !ids(l) || !ids(all(right)) || l.length!==r.length) return fail()
    const shell=(part:Element)=>{ const c=part.cloneNode(true) as Element; for(const e of all(c)) e.parentNode!.removeChild(e); return JSON.stringify(canon(c)) }
    const sig=(e:Element)=>{ const c=e.cloneNode(true) as Element; c.removeAttribute("Id"); return JSON.stringify(canon(c)) }
    if(shell(left)!==shell(right)) return fail()
    const used=new Set<Element>()
    for(const x of l) {
      const matches=r.filter(y=>!used.has(y) && sig(x)===sig(y)); if(matches.length!==1) return fail()
      const y=matches[0]!, oldId=x.getAttribute("Id")!, newId=y.getAttribute("Id")!
      if(oldId!==newId && (x.getAttribute("Type")!==THEME_TYPE || referenced(before,oldId) || referenced(after,newId))) return fail()
      used.add(y)
    }
  }
  const main=after.get(MAIN)!, bodies=Array.from(main.getElementsByTagNameNS(W,"body")), oldBodies=Array.from(before.get(MAIN)!.getElementsByTagNameNS(W,"body"))
  if(bodies.length!==1 || oldBodies.length!==1 || !Number.isSafeInteger(tableIndex) || tableIndex<0 || !Number.isSafeInteger(priorRows) || priorRows<1 || !cells.length || cells.length>32 ||
     runs!==undefined && (!Array.isArray(runs) || runs.length!==cells.length || runs.some(cell=>!Array.isArray(cell)))) return fail()
  const tables=named(bodies[0]!,"tbl"), oldTables=named(oldBodies[0]!,"tbl"); if(tableIndex>=tables.length || oldTables.length!==tables.length) return fail()
  settleNativeWordChatTableRecords(tables[tableIndex]!); settleNativeWordChatTableRecords(oldTables[tableIndex]!)
  const rows=named(tables[tableIndex]!,"tr"), oldRows=named(oldTables[tableIndex]!,"tr"); if(rows.length!==priorRows+1 || oldRows.length!==priorRows) return fail()
  const referenceCells=named(oldRows.at(-1)!,"tc"); if(referenceCells.length!==cells.length) return fail()
  const metadata = (element:Element) => Array.from(element.attributes).every(a => a.namespaceURI===XMLNS ||
    a.namespaceURI===W && ["rsidR","rsidRPr","rsidRDefault","rsidTr"].includes(a.localName) && /^[0-9A-Fa-f]{8}$/.test(a.value) ||
    a.namespaceURI===W14 && ["paraId","textId"].includes(a.localName) && /^[0-9A-Fa-f]{8}$/.test(a.value))
  const row=rows.at(-1)!, rowCells=named(row,"tc"), rowPr=named(row,"trPr"), marks=rowPr.length===1?named(rowPr[0]!,"ins"):[]
  // Word gives an appended row the old last row's own row properties.
  const referenceRowPr=named(oldRows.at(-1)!,"trPr"), rowRest=(props:Element|undefined)=>props ? elements(props).filter(e=>!(e.namespaceURI===W && e.localName==="ins")).map(canon) : []
  if(!metadata(row) || !only(row,["trPr","tc"]) || rowCells.length!==cells.length || rowPr.length!==1 || referenceRowPr.length>1 || marks.length!==1 || marks[0]!.childNodes.length ||
     referenceRowPr.some(p=>["ins","del","trPrChange"].some(n=>p.getElementsByTagNameNS(W,n).length)) || JSON.stringify(rowRest(rowPr[0]))!==JSON.stringify(rowRest(referenceRowPr[0]))) return fail()
  const ids=new Set<string>(), authors=new Set<string>(); let stamp:string|undefined
  const mark=(e:Element) => {
    const id=e.getAttributeNS(W,"id"), author=e.getAttributeNS(W,"author"), date=e.getAttributeNS(W,"date"), utc=e.getAttributeNS(W16DU,"dateUtc")
    if(!id || !/^\d+$/.test(id) || ids.has(id) || !author?.trim() || !date || Number.isNaN(Date.parse(date)) || utc && utc!==date || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && ["id","author","date"].includes(a.localName)) && !(a.namespaceURI===W16DU && a.localName==="dateUtc"))) return fail()
    ids.add(id);authors.add(author);const key=JSON.stringify([author,date,utc??null]);if(stamp && stamp!==key) return fail();stamp=key
  }
  mark(marks[0]!)
  const noAttrs=(p:Element) => Array.from(p.attributes).every(a=>a.namespaceURI===XMLNS)
  /** Word gives an added row's paragraphs the reference row's paragraph
   * properties (spacing, alignment, style). Apart from the mark's rPr they
   * must equal that column's paragraph in the old last row and carry no
   * revision of their own. */
  const inherited=(pPr:Element, referenceCell:Element) => {
    const reference=named(referenceCell,"p"), referenceProps=reference.length===1 ? named(reference[0]!,"pPr") : []
    if(reference.length!==1 || referenceProps.length>1) return false
    const rest=(props:Element|undefined)=>{
      if(!props) return []
      const copy=props.cloneNode(true) as Element
      for(const r of named(copy,"rPr")) copy.removeChild(r)
      return ["ins","del","pPrChange","rPrChange"].some(n=>copy.getElementsByTagNameNS(W,n).length) ? undefined : (canon(copy) as unknown[])[3]
    }
    const own=rest(pPr)
    return own!==undefined && JSON.stringify(own)===JSON.stringify(rest(referenceProps[0]))
  }
  const basic=(p:Element) => noAttrs(p) && only(p,["b","bCs","i","iCs","u"]) && elements(p).every(e=>!elements(e).length && Array.from(e.attributes).every(a=>a.namespaceURI===XMLNS || a.namespaceURI===W && a.localName==="val"))
  for(let i=0;i<cells.length;i++) {
    const cell=rowCells[i]!, ps=named(cell,"p"), cellProps=named(cell,"tcPr"), referenceProps=named(referenceCells[i]!,"tcPr")
    if(cellProps.length>1 || cellProps.length!==referenceProps.length || cellProps.length && JSON.stringify(canon(cellProps[0]!))!==JSON.stringify(canon(referenceProps[0]!)) ||
       [cell,...referenceCells.slice(i,i+1)].some(c=>["gridSpan","vMerge","hMerge"].some(n=>c.getElementsByTagNameNS(W,n).length))) return fail()
    if(!metadata(cell) || !only(cell,["tcPr","p"]) || ps.length!==1 || !metadata(ps[0]!) || !only(ps[0]!,["pPr","ins"])) return fail()
    const p=ps[0]!, pp=named(p,"pPr"), pr=pp.length===1?named(pp[0]!,"rPr"):[], pm=pr.length===1?named(pr[0]!,"ins"):[]
    if(pp.length!==1 || !noAttrs(pp[0]!) || pr.length!==1 || !noAttrs(pr[0]!) || pm.length!==1 || pm[0]!.childNodes.length || !inherited(pp[0]!,referenceCells[i]!) || !only(pr[0]!,["ins","b","bCs","i","iCs","u"])) return fail()
    if(elements(pr[0]!).some(e=>e.localName!=="ins" && (!only(e,[]) || Array.from(e.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===W && a.localName==="val"))))) return fail()
    mark(pm[0]!)
    let value=""
    const groups:WordChatCellRun[]=[]
    for(const ins of named(p,"ins")) {
      mark(ins);if(!only(ins,["r"])) return fail()
      for(const run of named(ins,"r")) {
        if(!metadata(run) || !only(run,["rPr","t"])) return fail()
        const props=named(run,"rPr"), ts=named(run,"t"); if(props.length>1 || ts.length!==1) return fail()
        if(props.length) {
          const r=props[0]!, changes=named(r,"rPrChange")
          if(!noAttrs(r) || !only(r,["b","bCs","i","iCs","u","rPrChange"]) || changes.length>1) return fail()
          for(const change of changes) { mark(change);const old=named(change,"rPr");if(old.length!==1 || !only(change,["rPr"]) || !basic(old[0]!)) return fail() }
        }
        const flags=directFlags(props[0])
        const t=ts[0]!;if(elements(t).length || Array.from(t.childNodes).some(n=>n.nodeType!==3) || Array.from(t.attributes).some(a=>a.namespaceURI!==XMLNS && !(a.namespaceURI===XML && a.localName==="space" && a.value==="preserve"))) return fail()
        const text=t.textContent??"", last=groups.at(-1);value+=text
        if(!text) continue
        if(last && last.bold===flags.bold && last.italic===flags.italic && last.underline===flags.underline) last.text+=text
        else groups.push({text,...flags})
      }
    }
    const planned=runs?.[i]?.length ? runs[i]! : cells[i] ? [{text:cells[i]!,bold:false,italic:false,underline:false}] : []
    if(value!==cells[i] || cells[i] && !named(p,"ins").length || groups.length!==planned.length ||
       groups.some((g,k)=>g.text!==planned[k]!.text || g.bold!==planned[k]!.bold || g.italic!==planned[k]!.italic || g.underline!==planned[k]!.underline)) return fail()
  }
  // Reject comments, PIs or non-whitespace text outside literal text runs.
  const clean=(n:Element):boolean=>Array.from(n.childNodes).every(c=>c.nodeType===1?clean(c as Element):c.nodeType===3 && (n.namespaceURI===W && n.localName==="t" || !c.nodeValue?.trim()))
  if(!clean(row)) return fail()
  // A new or grown people part lists only this row's revision author.
  if(after.has(PEOPLE) && !nativeWordChatPeopleGrowth(before.get(PEOPLE), after.get(PEOPLE)!, authors)) return fail()
  return {authors:[...authors]}
}

export function verifyNativeWordChatAddedRow(beforeXml:string, afterXml:string, binding:{tableIndex:number;rows:unknown[]}, cells:string[], runs?:WordChatCellRun[][]): string[] {
  return verifyNativeWordChatInsertedRow(beforeXml, afterXml, binding.tableIndex, binding.rows.length, cells, runs).authors
}

/** Whether Word, while adding a row, also re-fitted the bound table's column
 * widths as a tracked table-format change. Word for Windows can do this after
 * earlier tracked edits; the row is then present but cannot be confirmed, and
 * the lawyer needs to know why. Diagnostic only: never admits the change. */
export function nativeWordChatRowRefitWidths(beforeXml:string, afterXml:string, tableIndex:number): boolean {
  try {
    const table=(xml:string)=>{ const t=named(named(parts(xml).get(MAIN)!.getElementsByTagNameNS(W,"document")[0]!,"body")[0]!,"tbl")[tableIndex]; if(!t) return fail(); settleNativeWordChatTableRecords(t); return t }
    const widths=(t:Element)=>named(t,"tblGrid").flatMap(g=>named(g,"gridCol")).map(c=>c.getAttributeNS(W,"w"))
    const before=table(beforeXml), after=table(afterXml)
    return JSON.stringify(widths(before))!==JSON.stringify(widths(after)) && named(named(after,"tblGrid")[0] ?? after,"tblGridChange").length===1
  } catch { return false }
}
