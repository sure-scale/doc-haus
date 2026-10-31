import { unzipSync } from "fflate"
import type { WordChatAuthoredTableRow } from "../../../dochaus/lib/word-chat-contract"
import { canonicalWordChatCellRuns, type WordChatCellRun, type WordChatTableRowInsertionProposal } from "../../../dochaus/lib/word-chat-amendment-contract"
import { docxRowFormatCascade } from "../../../dochaus/lib/docx-row-format-cascade"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { validWordChatFormatting, type WordChatAuthoredAmendment, type WordChatInlineFormat, type WordChatParagraphFormat, type WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { reviewedWordChatParagraphFormatting, wordChatParagraphFormatChanged } from "../../../dochaus/lib/word-chat-formatting-ooxml"
import { assertWordChatInlineFormattingTarget, verifyWordChatInlineFormatting } from "../../../dochaus/lib/word-chat-inline-formatting-ooxml"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const MAIN = "/word/document.xml"
const elements = (node: Node) => Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1)
function parse(xml: string, limit = 2_000_000) {
  if (typeof xml !== "string" || xml.length > limit || /<!DOCTYPE/i.test(xml)) throw new Error("Invalid clause OOXML")
  const fail = () => { throw new Error("Invalid clause OOXML") }
  const document = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
  if (document.doctype || document.documentElement?.namespaceURI !== PKG || document.documentElement.localName !== "package") throw new Error("Expected a scoped Word package")
  const parts = new Map<string, Element>()
  for (const part of elements(document.documentElement)) {
    const name = part.getAttributeNS(PKG, "name")
    if (part.namespaceURI !== PKG || part.localName !== "part" || !name || parts.has(name) || parts.size >= 64) throw new Error("Invalid Word package part")
    parts.set(name, part)
  }
  if (!parts.has(MAIN)) throw new Error("Missing clause content")
  return parts
}

/** Bind faithful main-story paragraphs to the same-read native inventory.
 * Table geometry is represented by structural paths, not text search. */
export function projectNativeWordChatBody(xml: string, inventory: Array<{ id: string; text: string; tableNestingLevel?: number }>) {
  if (!inventory.length || inventory.some(p => !p.id || typeof p.text !== "string") || new Set(inventory.map(p => p.id)).size !== inventory.length)
    throw new Error("Invalid native paragraph inventory")
  const parts = parse(xml, 16_000_000), data = elements(parts.get(MAIN)!).filter(n => n.namespaceURI === PKG && n.localName === "xmlData")
  if (data.length !== 1 || elements(data[0]!).length !== 1) throw new Error("Invalid main story")
  const doc = elements(data[0]!)[0]!, bodies = elements(doc).filter(n => n.namespaceURI === W && n.localName === "body")
  if (doc.namespaceURI !== W || doc.localName !== "document" || bodies.length !== 1) throw new Error("Invalid main story")
  const body = bodies[0]!, children = elements(body), nodes: Array<{ node: Element; path: string; nesting: number }> = []
  const walk = (parent: Element, path: string, container: "body" | "cell", nesting = 0) => {
    elements(parent).forEach((node, i) => {
      if (node.namespaceURI !== W) throw new Error("Unsupported main-story structure")
      const loc = `${path}/${node.localName}[${i}]`
      if (node.localName === "p") nodes.push({ node, path: loc, nesting })
      else if (node.localName === "tbl") {
        elements(node).forEach((row, ri) => {
          if (row.namespaceURI !== W) throw new Error("Unsupported table structure")
          if (["tblPr", "tblGrid"].includes(row.localName)) return
          if (row.localName !== "tr") throw new Error("Unsupported table structure")
          elements(row).forEach((cell, ci) => {
            if (cell.namespaceURI !== W) throw new Error("Unsupported table row")
            if (cell.localName === "trPr") return
            if (cell.localName !== "tc") throw new Error("Unsupported table row")
            walk(cell, `${loc}/tr[${ri}]/tc[${ci}]`, "cell", nesting + 1)
          })
        })
      } else if (!(container === "body" && node.localName === "sectPr" && i === elements(parent).length - 1) && !(container === "cell" && node.localName === "tcPr"))
        throw new Error("Unsupported main-story structure")
    })
  }
  walk(body, "body", "body")
  let terminalCarrier = false
  if (nodes.length === inventory.length + 1) {
    const last = nodes.at(-1)!.node, bodyChildren = children.filter(n => !(n.namespaceURI === W && n.localName === "sectPr"))
    const permitted = (a: Attr) => a.namespaceURI === XMLNS || a.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(a.localName) || a.namespaceURI === W14 && ["paraId", "textId"].includes(a.localName)
    if (children.at(-1)?.namespaceURI !== W || children.at(-1)?.localName !== "sectPr" || last !== bodyChildren.at(-1) || last.parentNode !== body || Array.from(last.childNodes).some(n => n.nodeType !== 3 || Boolean(n.nodeValue?.trim())) || Array.from(last.attributes).some(a => !permitted(a)))
      throw new Error("Unqualified terminal carrier")
    nodes.pop()
    terminalCarrier = true
  }
  if (nodes.length !== inventory.length || nodes.some((n, i) => text(n.node, "Current") !== inventory[i]!.text || inventory[i]!.tableNestingLevel !== undefined && n.nesting !== inventory[i]!.tableNestingLevel)) throw new Error("Native paragraph projection mismatch")
  const serializer = new XMLSerializer()
  const skeleton = body.cloneNode(true) as Element
  // Whole-story exports assign the terminal section a session rsid even when
  // the section and all layout properties remain unchanged.
  const terminalSection = elements(skeleton).at(-1)
  if (terminalSection?.namespaceURI === W && terminalSection.localName === "sectPr") terminalSection.removeAttributeNS(W, "rsidSect")
  if (terminalCarrier) {
    const last = elements(skeleton).filter(n => !(n.namespaceURI === W && n.localName === "sectPr")).at(-1)!
    skeleton.removeChild(last)
  }
  const replaceParagraphs = (node: Element) => {
    for (const child of elements(node)) {
      if (child.namespaceURI === W && child.localName === "p") node.replaceChild(child.ownerDocument!.createElementNS(W, "w:p"), child)
      else replaceParagraphs(child)
    }
  }
  replaceParagraphs(skeleton)
  // Outside paragraphs only table and section revision marks carry w:id, and
  // Word renumbers them in story order on every export. Kind, author and date
  // stay in the comparison.
  for (const node of Array.from(skeleton.getElementsByTagNameNS(W, "*"))) node.removeAttributeNS(W, "id")
  return {
    ooxml: xml,
    paragraphs: nodes.map((n, i) => ({ id: inventory[i]!.id, text: inventory[i]!.text, path: n.path, nesting: n.nesting, xml: serializer.serializeToString(n.node) })),
    geometry: JSON.stringify(canonical(skeleton, true)),
    /** Preserve every package dependency while scoping the main body to bound paragraphs. */
    scope(indices: number[]) {
      if (!indices.length || indices.some((n, i) => !Number.isSafeInteger(n) || n < 0 || n >= nodes.length || i > 0 && n !== indices[i - 1]! + 1)) throw new Error("Invalid paragraph scope")
      const clone = body.ownerDocument!.cloneNode(true) as Document
      const main = Array.from(clone.getElementsByTagNameNS(PKG, "part")).find(part => part.getAttributeNS(PKG, "name") === MAIN)
      const cloneBody = main?.getElementsByTagNameNS(W, "body")[0]
      if (!cloneBody) throw new Error("Missing main story")
      while (cloneBody.firstChild) cloneBody.removeChild(cloneBody.firstChild)
      for (const i of indices) cloneBody.appendChild(nodes[i]!.node.cloneNode(true))
      return serializer.serializeToString(clone)
    },
  }
}
/** Verify only the protected main-story remainder after an appended row.
 * The caller must separately verify new-row formatting, carrier ownership and
 * compressed-package dependencies; this function does not establish them. */
export function verifyNativeWordChatRowRemainder(
  beforeXml: string, afterXml: string,
  beforeInventory: Array<{id:string; text:string; tableNestingLevel?:number}>,
  afterInventory: Array<{id:string; text:string; tableNestingLevel?:number}>,
  binding: {tableIndex:number; rows:Array<Array<{id:string; text:string}>>}, cells: string[],
) {
  const before = projectNativeWordChatBody(beforeXml, beforeInventory)
  const after = projectNativeWordChatBody(afterXml, afterInventory)
  const fail = (): never => { throw new Error("Native row remainder mismatch") }
  const bodyOf = (xml: string) => {
    const part = parse(xml, 16_000_000).get(MAIN)!
    const bodies = Array.from(part.getElementsByTagNameNS(W, "body"))
    if (bodies.length !== 1) return fail()
    return bodies[0]!
  }
  const oldBody = bodyOf(beforeXml), newBody = bodyOf(afterXml)
  const tables = (body: Element) => elements(body).filter(e => e.namespaceURI === W && e.localName === "tbl")
  if (!Number.isSafeInteger(binding.tableIndex) || binding.tableIndex < 0 || !binding.rows.length || !cells.length || cells.length > 32 ||
      binding.rows.some(row => row.length !== cells.length) || cells.some(cell => typeof cell !== "string")) return fail()
  const oldTables = tables(oldBody), newTables = tables(newBody)
  if (oldTables.length !== newTables.length || binding.tableIndex >= oldTables.length) return fail()
  const oldTable = oldTables[binding.tableIndex]!, newTable = newTables[binding.tableIndex]!
  const rows = (table: Element) => elements(table).filter(e => e.namespaceURI === W && e.localName === "tr")
  const oldRows = rows(oldTable), newRows = rows(newTable)
  if (oldRows.length !== binding.rows.length || newRows.length !== oldRows.length + 1) return fail()
  const newRow = newRows.at(-1)!
  const rowCells = (row: Element) => elements(row).filter(e => e.namespaceURI === W && e.localName === "tc")
  const addedCells = rowCells(newRow)
  if (addedCells.length !== cells.length || elements(newRow).some(e => e.namespaceURI !== W || !["trPr", "tc"].includes(e.localName))) return fail()
  const addedIds: string[] = []
  for (let rowIndex = 0; rowIndex < newRows.length; rowIndex++) {
    const currentCells = rowCells(newRows[rowIndex]!)
    if (currentCells.length !== cells.length) return fail()
    for (let col = 0; col < currentCells.length; col++) {
      const cell = currentCells[col]!, paragraphs = elements(cell).filter(e => e.namespaceURI === W && e.localName === "p")
      if (paragraphs.length !== 1 || elements(cell).some(e => e.namespaceURI !== W || !["tcPr", "p"].includes(e.localName))) return fail()
      const tablePos = elements(newBody).indexOf(newTable), rowPos = elements(newTable).indexOf(newRows[rowIndex]!), cellPos = elements(newRows[rowIndex]!).indexOf(cell), paraPos = elements(cell).indexOf(paragraphs[0]!)
      const path = `body/tbl[${tablePos}]/tr[${rowPos}]/tc[${cellPos}]/p[${paraPos}]`
      const found = after.paragraphs.filter(p => p.path === path)
      if (found.length !== 1) return fail()
      if (rowIndex === oldRows.length) {
        if (found[0]!.text !== cells[col]) return fail()
        addedIds.push(found[0]!.id)
      } else if (found[0]!.id !== binding.rows[rowIndex]![col]!.id || found[0]!.text !== binding.rows[rowIndex]![col]!.text) return fail()
    }
  }
  const oldIds = binding.rows.flat().map(cell => cell.id)
  if (new Set(oldIds).size !== oldIds.length || oldIds.some(id => !before.paragraphs.some(p => p.id === id))) return fail()
  const filtered = after.paragraphs.filter(p => !addedIds.includes(p.id))
  if (filtered.length !== before.paragraphs.length || filtered.some((p, i) => p.id !== before.paragraphs[i]!.id || p.text !== before.paragraphs[i]!.text || p.path !== before.paragraphs[i]!.path)) return fail()
  const carrierNames = new Set(["ins", "del", "pPrChange", "rPrChange"])
  const collect = (root: Element) => {
    const result: Array<{path:string; node:Element}> = []
    const visit = (node:Element, path:string) => {
      if (node.namespaceURI === W && carrierNames.has(node.localName)) result.push({path, node})
      for (const [i, child] of elements(node).entries()) visit(child, `${path}/${child.localName}[${i}]`)
    }
    visit(root,"body"); return result
  }
  const allAfter = collect(newBody)
  const afterIds = allAfter.map(item => item.node.getAttributeNS(W,"id"))
  if (afterIds.some(id => !id || !/^\d+$/.test(id)) || new Set(afterIds).size !== afterIds.length) return fail()
  newRow.parentNode!.removeChild(newRow)
  const old = collect(oldBody), current = collect(newBody)
  const ids = (items: typeof old) => items.map(item => item.node.getAttributeNS(W,"id"))
  const oldIdsAll = ids(old), currentIds = ids(current)
  if (old.length !== current.length || oldIdsAll.some(id => !id || !/^\d+$/.test(id)) || currentIds.some(id => !id || !/^\d+$/.test(id)) ||
      new Set(oldIdsAll).size !== old.length || new Set(currentIds).size !== current.length) return fail()
  for (let i = 0; i < old.length; i++) {
    const a = old[i]!, b = current[i]!
    if (a.path !== b.path || a.node.namespaceURI !== b.node.namespaceURI || a.node.localName !== b.node.localName) return fail()
    const name = b.node.getAttributeNodeNS(W,"id")!.name
    b.node.removeAttributeNS(W,"id"); b.node.setAttributeNS(W,name,a.node.getAttributeNS(W,"id")!)
  }
  for (const body of [oldBody, newBody]) {
    const section = elements(body).at(-1)
    if (section?.namespaceURI === W && section.localName === "sectPr") section.removeAttributeNS(W,"rsidSect")
  }
  if (JSON.stringify(canonical(oldBody,true)) !== JSON.stringify(canonical(newBody,true))) return fail()
  return { addedParagraphIds: addedIds, carrierMap: old.map((item,i) => ({path:item.path, type:item.node.localName, before:oldIdsAll[i]!, after:currentIds[i]!})) }
}

/** A mixed-format row states only its "on" spans, so it needs a table whose styles apply none. */
export class WordChatRowFormatRefusal extends Error {
  constructor() { super("This table's style already applies bold, italic or underline, so a mixed-format row can't be added safely. Your document is unchanged."); this.name = "WordChatRowFormatRefusal" }
}
export function assertWordChatRowFormatNeutral(source: Uint8Array, proposal: WordChatTableRowInsertionProposal) {
  if (proposal.runs && !docxRowFormatCascade(source, proposal.table.tableIndex).neutral) throw new WordChatRowFormatRefusal()
}
/** Verify the reviewed SDK oracle only; never import its bytes into Word. */
export async function verifyReviewedWordChatRowOracle(source: Uint8Array, oracle: WordChatAuthoredTableRow, proposal: WordChatTableRowInsertionProposal) {
  const fail = (): never => { throw new Error("Reviewed row oracle mismatch") }
  const max = 16 * 1024 * 1024, aggregate = 64 * 1024 * 1024
  if (!(source instanceof Uint8Array) || !source.length || source.length > max || !oracle || !proposal || oracle.kind !== "table_row_insert" || proposal.kind !== "table_row_insert" || proposal.position !== "end" || oracle.proposalId !== proposal.id || oracle.author !== "doc.haus" || JSON.stringify(oracle.table) !== JSON.stringify(proposal.table) || JSON.stringify(oracle.cells) !== JSON.stringify(proposal.cells)) return fail()
  const digest = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", source))).map(x => x.toString(16).padStart(2,"0")).join("")
  const req = oracle.writerRequest
  if (oracle.sourceSha256 !== digest || !req || req.source_sha256 !== digest || req.position !== "after" || req.author !== oracle.author || !req.target_handle || JSON.stringify(req.cells) !== JSON.stringify(proposal.cells)) return fail()
  // Spans are the proposal's canonical runs, carried unchanged; absent for a plain row.
  const runs = proposal.runs
  if (runs !== undefined && canonicalWordChatCellRuns(proposal.cells, runs) !== runs || JSON.stringify(oracle.runs) !== JSON.stringify(runs) || JSON.stringify(req.runs) !== JSON.stringify(runs)) return fail()
  const encoded = oracle.authoredBase64
  if (typeof encoded !== "string" || !encoded.length || encoded.length > Math.ceil(max/3)*4 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(encoded)) return fail()
  let decoded: string
  try { decoded = atob(encoded) } catch { return fail() }
  if (!decoded.length || decoded.length > max || btoa(decoded) !== encoded) return fail()
  const unpack = (bytes: Uint8Array) => {
    let size=0, count=0; const names=new Set<string>()
    try { return unzipSync(bytes,{filter(entry) { if (++count > 4096 || names.has(entry.name) || entry.originalSize > max || (size += entry.originalSize) > aggregate) throw Error("archive limit"); names.add(entry.name); return true }}) } catch { return fail() }
  }
  const a=unpack(source), b=unpack(Uint8Array.from(decoded,c=>c.charCodeAt(0))), keys=Object.keys(a).sort()
  if (JSON.stringify(keys)!==JSON.stringify(Object.keys(b).sort()) || !a["word/document.xml"] || !a["word/settings.xml"]) return fail()
  for (const k of keys) if (k!=="word/document.xml" && k!=="word/settings.xml" && (a[k]!.length!==b[k]!.length || a[k]!.some((v,i)=>v!==b[k]![i]))) return fail()
  const parsePart = (bytes:Uint8Array) => {
    const reject = () => { throw Error("invalid XML") }
    const doc=new DOMParser({errorHandler:{warning:reject,error:reject,fatalError:reject}}).parseFromString(new TextDecoder("utf-8",{fatal:true}).decode(bytes),"application/xml")
    if (doc.doctype || !doc.documentElement) return fail()
    return doc
  }
  const sa=parsePart(a["word/settings.xml"]!), sb=parsePart(b["word/settings.xml"]!)
  for (const [doc,required] of [[sa,false],[sb,true]] as const) {
    const marks=Array.from(doc.getElementsByTagNameNS(W,"trackRevisions"))
    if (marks.length>1 || required && marks.length!==1 || marks.some(m=>m.parentNode!==doc.documentElement || m.childNodes.length || Array.from(m.attributes).some(x=>x.namespaceURI!==W || x.localName!=="val" || !["1","true","on"].includes(x.value)))) return fail()
    for (const m of marks) m.parentNode!.removeChild(m)
  }
  if (JSON.stringify(canonical(sa.documentElement))!==JSON.stringify(canonical(sb.documentElement))) return fail()
  const da=parsePart(a["word/document.xml"]!), db=parsePart(b["word/document.xml"]!)
  const direct=(e:Element,n:string)=>elements(e).filter(x=>x.namespaceURI===W && x.localName===n)
  const ab=Array.from(da.getElementsByTagNameNS(W,"body")), bb=Array.from(db.getElementsByTagNameNS(W,"body"))
  if (ab.length!==1 || bb.length!==1) return fail()
  const at=direct(ab[0]!,"tbl"), bt=direct(bb[0]!,"tbl"), table=proposal.table
  if (!table || !Number.isSafeInteger(table.tableIndex) || table.tableIndex<0 || at.length!==bt.length || table.tableIndex>=at.length || !table.rows.length || !proposal.cells.length || proposal.cells.length>32) return fail()
  const old=direct(at[table.tableIndex]!,"tr"), target=bt[table.tableIndex]!, next=direct(target,"tr")
  if (old.length!==table.rows.length || next.length!==old.length+1) return fail()
  if (runs && !docxRowFormatCascade(source, table.tableIndex).neutral) throw new WordChatRowFormatRefusal()
  for (let i=0;i<old.length;i++) { const cells=direct(old[i]!,"tc"), expected=table.rows[i]!; if (cells.length!==proposal.cells.length || expected.length!==cells.length || cells.some((c,j)=>text(c,"Current")!==expected[j]!.text)) return fail() }
  const row=next.at(-1)!, cells=direct(row,"tc"), ref=direct(old.at(-1)!,"tc"), carriers:Element[]=[]
  const only=(e:Element,names:string[]) => elements(e).every(x=>x.namespaceURI===W && names.includes(x.localName))
  const mark=(e:Element) => { if (e.namespaceURI!==W || e.localName!=="ins" || !/^(0|[1-9]\d*)$/.test(e.getAttributeNS(W,"id")??"") || e.getAttributeNS(W,"author")!=="doc.haus" || !e.getAttributeNS(W,"date") || Array.from(e.attributes).some(a=>a.namespaceURI!==W || !["id","author","date"].includes(a.localName))) return fail(); carriers.push(e) }
  const rp=direct(row,"trPr"), rm=rp.length===1?direct(rp[0]!,"ins"):[]
  if (row.attributes.length || !only(row,["trPr","tc"]) || cells.length!==proposal.cells.length || rp.length!==1 || rp[0]!.attributes.length || elements(rp[0]!).length!==1 || rm.length!==1 || rm[0]!.childNodes.length) return fail()
  mark(rm[0]!)
  // The SDK's tracked span format: <w:rPr>[b][i][u]<w:rPrChange …><w:rPr/></w:rPrChange></w:rPr>, nothing else.
  const spanFlags = (rPr: Element) => {
    const props = elements(rPr), change = props.at(-1), flags = props.slice(0,-1), names = flags.map(e=>e.localName)
    if (rPr.attributes.length || props.some(e=>e.namespaceURI!==W) || change?.localName!=="rPrChange" || !flags.length ||
      JSON.stringify(names)!==JSON.stringify(["b","i","u"].filter(n=>names.includes(n))) ||
      flags.some(e=>elements(e).length || (e.localName==="u" ? e.attributes.length!==1 || e.getAttributeNS(W,"val")!=="single" : e.attributes.length))) return fail()
    const prior = elements(change)
    if (!/^(0|[1-9]\d*)$/.test(change.getAttributeNS(W,"id")??"") || change.getAttributeNS(W,"author")!=="doc.haus" || !change.getAttributeNS(W,"date") ||
      Array.from(change.attributes).some(a=>a.namespaceURI!==W || !["id","author","date"].includes(a.localName)) ||
      prior.length!==1 || prior[0]!.namespaceURI!==W || prior[0]!.localName!=="rPr" || prior[0]!.attributes.length || elements(prior[0]!).length) return fail()
    carriers.push(change)
    return { bold:names.includes("b"), italic:names.includes("i"), underline:names.includes("u") }
  }
  const cleanNodes = (node: Element): boolean => Array.from(node.childNodes).every(child => child.nodeType === 1 ? cleanNodes(child as Element) : child.nodeType === 3 && (node.namespaceURI === W && node.localName === "t" || !child.nodeValue?.trim()))
  if (!cleanNodes(row)) return fail()
  for (let i=0;i<cells.length;i++) {
    const c=cells[i]!, p=direct(c,"p"), cp=direct(c,"tcPr"), op=direct(ref[i]!,"tcPr")
    if (c.attributes.length || !only(c,["tcPr","p"]) || p.length!==1 || p[0]!.attributes.length || cp.length!==op.length || cp.length>1 || cp.length && JSON.stringify(canonical(cp[0]!))!==JSON.stringify(canonical(op[0]!)) || text(c,"Current")!==proposal.cells[i]) return fail()
    const pp=direct(p[0]!,"pPr"), pr=pp.length===1?direct(pp[0]!,"rPr"):[], pm=pr.length===1?direct(pr[0]!,"ins"):[]
    if (!only(p[0]!,["pPr","ins"]) || pp.length!==1 || pp[0]!.attributes.length || elements(pp[0]!).length!==1 || pr.length!==1 || pr[0]!.attributes.length || elements(pr[0]!).length!==1 || pm.length!==1 || pm[0]!.childNodes.length) return fail()
    mark(pm[0]!)
    const spans: WordChatCellRun[] = []
    for (const ins of direct(p[0]!,"ins")) { mark(ins); if (!only(ins,["r"])) return fail(); for (const run of direct(ins,"r")) {
      const ts=direct(run,"t"), rPr=direct(run,"rPr")
      if (run.attributes.length || ts.length!==1 || rPr.length>1 || rPr.length && (!runs || elements(run)[0]!==rPr[0]) || elements(run).length!==1+rPr.length || elements(ts[0]!).length || Array.from(ts[0]!.childNodes).some(n=>n.nodeType!==3) || Array.from(ts[0]!.attributes).some(a=>a.namespaceURI!=="http://www.w3.org/XML/1998/namespace" || a.localName!=="space" || a.value!=="preserve")) return fail()
      const flags = rPr.length ? spanFlags(rPr[0]!) : { bold:false, italic:false, underline:false }, last = spans.at(-1), value = ts[0]!.textContent ?? ""
      if (last && last.bold===flags.bold && last.italic===flags.italic && last.underline===flags.underline) last.text += value
      else spans.push({ text:value, ...flags })
    } }
    if (runs && JSON.stringify(spans)!==JSON.stringify(runs[i]!.map(r=>({text:r.text,bold:r.bold,italic:r.italic,underline:r.underline})))) return fail()
  }
  const ids=carriers.map(c=>c.getAttributeNS(W,"id")!), oldIds=new Set(Array.from(da.getElementsByTagNameNS(W,"*")).flatMap(e=>e.hasAttributeNS(W,"id")?[e.getAttributeNS(W,"id")!.replace(/^\+?0+(?=\d)/, "")]:[]))
  if (new Set(ids).size!==ids.length || ids.some(id=>oldIds.has(id))) return fail()
  target.removeChild(row)
  if (JSON.stringify(canonical(da.documentElement))!==JSON.stringify(canonical(db.documentElement))) return fail()
  return {sourceSha256:digest,cells:[...proposal.cells]}
}

function content(parts: Map<string, Element>, table: boolean, section: boolean) {
  const data = elements(parts.get(MAIN)!).filter(node => node.namespaceURI === PKG && node.localName === "xmlData")
  if (data.length !== 1 || elements(data[0]!).length !== 1) throw new Error("Invalid clause document")
  const document = elements(data[0]!)[0]!
  if (document.namespaceURI !== W || document.localName !== "document") throw new Error("Invalid clause document")
  const bodies = elements(document).filter(node => node.namespaceURI === W && node.localName === "body")
  if (bodies.length !== 1) throw new Error("Invalid clause body")
  const children = elements(bodies[0]!)
  if (section && children.at(-1)?.namespaceURI === W && children.at(-1)?.localName === "sectPr") children.pop()
  if (children.length && children.every(node => node.namespaceURI === W && node.localName === "p")) return { nodes: children, shape: children }
  if (table && children.length === 1 && children[0]!.namespaceURI === W && children[0]!.localName === "tbl") {
    const wrapper = children[0]!, cells = Array.from(wrapper.getElementsByTagNameNS(W, "tc"))
    // A paragraph fragment may carry its existing cell wrapper. Never import
    // another cell or guess among additional empty/nonempty table paragraphs.
    if (cells.length !== 1 || wrapper.getElementsByTagNameNS(W, "tbl").length) throw new Error("Choose paragraphs within one ordinary table cell")
    const nodes = elements(cells[0]!).filter(node => node.namespaceURI === W && node.localName === "p")
    if (!nodes.length || nodes.length !== wrapper.getElementsByTagNameNS(W, "p").length) throw new Error("Unsupported table clause content")
    return { nodes, shape: children }
  }
  throw new Error("Choose ordinary paragraphs within one document passage")
}
function text(node: Node, version: "Current" | "Original" = "Current"): string {
  if (node.nodeType !== 1) return ""
  const element = node as Element
  if (element.namespaceURI === W) {
    if (["pPr", "rPr"].includes(element.localName)) return ""
    if (element.localName === (version === "Current" ? "del" : "ins")) return ""
    if (["t", "delText"].includes(element.localName)) return element.textContent || ""
    if (element.localName === "tab") return "\t"
    if (["br", "cr"].includes(element.localName)) return "\u000b"
  }
  return elements(element).map(child => text(child, version)).join("")
}
function canonical(node: Node, paragraph = false): unknown {
  if (node.nodeType === 1) {
    const element = node as Element
    const attributes = Array.from(element.attributes).filter(attr => attr.namespaceURI !== XMLNS && !(
      paragraph && (attr.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(attr.localName)
        || element.namespaceURI === W && element.localName === "p" && attr.namespaceURI === W14 && attr.localName === "paraId")
    )).map(attr => [attr.namespaceURI || "", attr.localName, attr.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)))
    return [element.namespaceURI || "", element.localName, attributes, Array.from(element.childNodes).map(child => canonical(child, paragraph)).filter(child => child !== undefined)]
  }
  if (node.nodeType === 3 || node.nodeType === 4) {
    const parent = node.parentNode as Element | null
    if (!node.nodeValue?.trim() && !(parent?.namespaceURI === W && ["t", "delText", "instrText"].includes(parent.localName))) return undefined
    return ["text", node.nodeValue]
  }
  if (node.nodeType === 7 || node.nodeType === 8) return [node.nodeType, node.nodeName, node.nodeValue]
  return undefined
}
export function capturedWordChatParagraphs(xml: string): string[] {
  return content(parse(xml), true, true).nodes.map(node => text(node))
}

// Compare accepted character formatting, paragraph properties and named bookmark
// anchors rather than run boundaries: Word may split runs during OOXML import.
function paragraphView(paragraph: Element, version: "Current" | "Original") {
  const runs: Array<{ text: string; format: string }> = []
  const bookmarks = new Map<string, { name: string; start: number; end?: number }>()
  let offset = 0
  const append = (value: string, format: string) => {
    if (!value) return
    const last = runs.at(-1)
    if (last?.format === format) last.text += value
    else runs.push({ text: value, format })
    offset += value.length
  }
  const visit = (node: Element, format = "null") => {
    if (node.namespaceURI === W) {
      if (node.localName === (version === "Current" ? "del" : "ins") || ["pPr", "rPr"].includes(node.localName)) return
      if (node.localName === "r") format = JSON.stringify(canonical(elements(node).find(child => child.namespaceURI === W && child.localName === "rPr") ?? nullNode) ?? null)
      if (["t", "delText", "tab", "br", "cr"].includes(node.localName)) { append(text(node, version), format); return }
      if (node.localName === "bookmarkStart") {
        const id = node.getAttributeNS(W, "id"), name = node.getAttributeNS(W, "name")
        if (!id || !name || bookmarks.has(id)) throw new Error("Invalid clause bookmark")
        bookmarks.set(id, { name, start: offset }); return
      }
      if (node.localName === "bookmarkEnd") {
        const mark = bookmarks.get(node.getAttributeNS(W, "id") || "")
        if (!mark || mark.end !== undefined) throw new Error("Invalid clause bookmark")
        mark.end = offset; return
      }
    }
    for (const child of elements(node)) visit(child, format)
  }
  visit(paragraph)
  const marks = [...bookmarks.values()]
  if (marks.some(mark => mark.end === undefined) || new Set(marks.map(mark => mark.name)).size !== marks.length) throw new Error("Incomplete clause bookmark")
  const currentProperties = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")
  const changes = currentProperties ? elements(currentProperties).filter(node => node.namespaceURI === W && node.localName === "pPrChange") : []
  if (changes.length > 1) throw new Error("Ambiguous paragraph formatting")
  const old = changes[0] && elements(changes[0]).find(node => node.namespaceURI === W && node.localName === "pPr")
  if (changes[0] && !old) throw new Error("Missing original paragraph properties")
  const properties = (version === "Original" && old ? old : currentProperties)?.cloneNode(true) as Element | undefined
  if (properties) {
    for (const change of elements(properties)) if (change.namespaceURI === W && change.localName === "pPrChange") properties.removeChild(change)
    const run = elements(properties).find(node => node.namespaceURI === W && node.localName === "rPr")
    if (run) {
      for (const mark of elements(run)) if (mark.namespaceURI === W && ["ins", "del"].includes(mark.localName)) run.removeChild(mark)
      if (!elements(run).length && !run.attributes.length) properties.removeChild(run)
    }
  }
  return { properties: canonical(properties && (elements(properties).length || properties.attributes.length) ? properties : nullNode) ?? null, runs,
    bookmarks: marks.sort((a, b) => a.name.localeCompare(b.name)) }
}
const nullNode = { nodeType: 0 } as Node
function paragraphMark(paragraph: Element) {
  const properties = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")
  const run = properties && elements(properties).find(node => node.namespaceURI === W && node.localName === "rPr")
  const marks = run ? elements(run).filter(node => node.namespaceURI === W && ["ins", "del"].includes(node.localName)) : []
  if (marks.length > 1) throw new Error("Ambiguous paragraph revision")
  return marks[0]
}
function clauseView(xml: string, version: "Current" | "Original") {
  return content(parse(xml), true, true).nodes.filter(node => {
    const mark = paragraphMark(node)
    if (mark?.localName !== (version === "Current" ? "del" : "ins")) return true
    if (text(node, version)) throw new Error("Only complete paragraph insertions and deletions are supported")
    return false
  }).map(node => paragraphView(node, version))
}
export function importedWordChatDifference(authored: string, observed: string): string | undefined {
  const marks = (xml: string) => content(parse(xml), true, true).nodes.map(paragraphMark).filter((mark): mark is Element => Boolean(mark))
    .map(mark => [mark.localName, mark.getAttributeNS(W, "author")])
  const formats = (xml: string) => content(parse(xml), true, true).nodes.map(paragraph => {
    const pPr = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")
    return pPr ? elements(pPr).filter(node => node.namespaceURI === W && node.localName === "pPrChange").map(node => node.getAttributeNS(W, "author")) : []
  })
  for (const version of ["Current", "Original"] as const) {
    const expected = clauseView(authored, version), actual = clauseView(observed, version)
    if (expected.length !== actual.length) return `${version.toLowerCase()} paragraph count`
    for (let i = 0; i < expected.length; i++) for (const key of ["properties", "runs", "bookmarks"] as const)
      if (JSON.stringify(expected[i]![key]) !== JSON.stringify(actual[i]![key])) return `${version.toLowerCase()} paragraph ${i + 1} ${key}`
  }
  if (JSON.stringify(marks(authored)) !== JSON.stringify(marks(observed))) return "paragraph revision marks"
  if (JSON.stringify(formats(authored)) !== JSON.stringify(formats(observed))) return "format revision marks"
}
export function sameImportedWordChat(authored: string, observed: string): boolean {
  return importedWordChatDifference(authored, observed) === undefined
}

const named = (node: Element, name: string) => elements(node).filter(child => child.namespaceURI === W && child.localName === name)
const unique = (nodes: Element[]) => nodes.length === 1 ? nodes[0] : undefined
function dependencyRoot(parts: Map<string, Element>, name: string) {
  const part = parts.get(name), data = part && unique(elements(part).filter(node => node.namespaceURI === PKG && node.localName === "xmlData"))
  return data && unique(elements(data))
}

/** Word may materialize a list level's inherited indentation into the old pPr
 * when recording a formatting revision. Normalize only that exact proven shape,
 * with unchanged styles/numbering and no competing inherited list or indent. */
function normalizeNativeNumberingOldProperties(captured: string, observed: string, before: Element[], after: Element[]) {
  const source = parse(captured), current = parse(observed)
  const numbering = dependencyRoot(source, "/word/numbering.xml"), observedNumbering = dependencyRoot(current, "/word/numbering.xml")
  const styles = dependencyRoot(source, "/word/styles.xml"), observedStyles = dependencyRoot(current, "/word/styles.xml")
  if (!numbering || !observedNumbering || !styles || !observedStyles ||
    JSON.stringify(canonical(numbering)) !== JSON.stringify(canonical(observedNumbering)) ||
    JSON.stringify(canonical(styles)) !== JSON.stringify(canonical(observedStyles))) return
  if (numbering.namespaceURI !== W || numbering.localName !== "numbering" || styles.namespaceURI !== W || styles.localName !== "styles") return
  const styleList = named(styles, "style"), defaults = named(styles, "docDefaults")
  if (defaults.length > 1 || defaults.some(node => node.getElementsByTagNameNS(W, "ind").length || node.getElementsByTagNameNS(W, "numPr").length)) return
  const safeStyle = (original: Element) => {
    const styleRef = unique(named(original, "pStyle"))?.getAttributeNS(W, "val")
    const defaultStyle = unique(styleList.filter(style => style.getAttributeNS(W, "type") === "paragraph" && ["1", "true"].includes(style.getAttributeNS(W, "default") ?? "")))
    const first = styleRef || defaultStyle?.getAttributeNS(W, "styleId")
    if (!first) return false
    const visited = new Set<string>()
    for (const seed of [first, defaultStyle?.getAttributeNS(W, "styleId")]) {
      let id = seed
      const chain = new Set<string>()
      while (id && !visited.has(id)) {
        if (chain.has(id)) return false
        chain.add(id)
        const style = unique(styleList.filter(item => item.getAttributeNS(W, "styleId") === id && item.getAttributeNS(W, "type") === "paragraph"))
        if (!style || named(style, "pPr").length > 1 || style.getElementsByTagNameNS(W, "ind").length || style.getElementsByTagNameNS(W, "numPr").length || named(style, "basedOn").length > 1) return false
        id = unique(named(style, "basedOn"))?.getAttributeNS(W, "val")
      }
      if (id && chain.has(id)) return false
      for (const item of chain) visited.add(item)
    }
    return true
  }
  for (const [index, paragraph] of after.entries()) {
    const original = unique(named(before[index]!, "pPr")), properties = unique(named(paragraph, "pPr"))
    const change = properties && unique(named(properties, "pPrChange")), old = change && unique(named(change, "pPr"))
    if (!original || !old || named(original, "ind").length || !safeStyle(original)) continue
    const originalNum = unique(named(original, "numPr")), oldNum = unique(named(old, "numPr")), ind = unique(named(old, "ind"))
    const numId = originalNum && unique(named(originalNum, "numId"))?.getAttributeNS(W, "val")
    const level = originalNum && unique(named(originalNum, "ilvl"))?.getAttributeNS(W, "val")
    if (!originalNum || !oldNum || !ind || !numId || numId === "0" || !level || !/^[0-8]$/.test(level)) continue
    const num = unique(named(numbering, "num").filter(node => node.getAttributeNS(W, "numId") === numId))
    const abstractId = num && unique(named(num, "abstractNumId"))?.getAttributeNS(W, "val")
    const abstract = abstractId && unique(named(numbering, "abstractNum").filter(node => node.getAttributeNS(W, "abstractNumId") === abstractId))
    if (!num || !abstract || named(abstract, "numStyleLink").length || named(abstract, "styleLink").length || abstract.getElementsByTagNameNS(W, "pStyle").length) continue
    const overrides = named(num, "lvlOverride").filter(node => node.getAttributeNS(W, "ilvl") === level)
    if (overrides.length > 1 || overrides.some(node => node.getElementsByTagNameNS(W, "pStyle").length)) continue
    const overrideLevels = overrides[0] ? named(overrides[0], "lvl") : []
    if (overrideLevels.length > 1) continue
    const lvl = overrideLevels[0] ?? unique(named(abstract, "lvl").filter(node => node.getAttributeNS(W, "ilvl") === level))
    const lvlPr = lvl && unique(named(lvl, "pPr")), lvlInd = lvlPr && unique(named(lvlPr, "ind"))
    if (!lvlInd || JSON.stringify(canonical(lvlInd)) !== JSON.stringify(canonical(ind))) continue
    const normalized = old.cloneNode(true) as Element, normalizedNum = unique(named(normalized, "numPr"))!
    normalized.removeChild(unique(named(normalized, "ind"))!)
    const oldLevel = named(normalizedNum, "ilvl")
    if (!oldLevel.length && level === "0") normalizedNum.insertBefore(originalNum.ownerDocument!.importNode(unique(named(originalNum, "ilvl"))!, true), normalizedNum.firstChild)
    if (JSON.stringify(canonical(normalized)) !== JSON.stringify(canonical(original))) continue
    change!.replaceChild(change!.ownerDocument!.importNode(original, true), old)
  }
}

/** Word can renumber pre-existing inline revision IDs during a native save.
 * Match them in paragraph order with all other markup exact and unique IDs. */
function normalizeNativeRevisionIds(before: Element[], after: Element[]) {
  const collect = (paragraph: Element) => {
    const revisions: Element[] = []
    const visit = (node: Element) => {
      if (node.namespaceURI === W && ["pPr", "rPr"].includes(node.localName)) return
      if (node.namespaceURI === W && ["ins", "del"].includes(node.localName)) { revisions.push(node); return }
      elements(node).forEach(visit)
    }
    visit(paragraph); return revisions
  }
  const originals = before.map(collect), observed = after.map(collect)
  const ids = (values: Element[][]) => values.flat().map(node => node.getAttributeNS(W, "id"))
  const originalIds = ids(originals), currentIds = ids(observed)
  if (originalIds.some(id => !id || !/^\d+$/.test(id)) || currentIds.some(id => !id || !/^\d+$/.test(id)) ||
    new Set(originalIds).size !== originalIds.length || new Set(currentIds).size !== currentIds.length ||
    originals.length !== observed.length || originals.some((marks, i) => marks.length !== observed[i]!.length)) return false
  const pairs = originals.flatMap((marks, i) => marks.map((mark, j) => [mark, observed[i]![j]!] as const))
  const withoutId = (node: Element) => { const copy = node.cloneNode(true) as Element; copy.removeAttributeNS(W, "id"); return canonical(copy) }
  if (pairs.some(([old, current]) => JSON.stringify(withoutId(old)) !== JSON.stringify(withoutId(current)))) return false
  for (const [old, current] of pairs) {
    const name = current.getAttributeNodeNS(W, "id")!.name
    current.removeAttributeNS(W, "id")
    current.setAttributeNS(W, name, old.getAttributeNS(W, "id")!)
  }
  return true
}

/** Verify native paragraph formatting against both the captured source and the
 * independently checked authored artifact. Word owns the revision author for
 * property mutations; no other author or revision content is normalized. */
export function verifiedNativeWordChatFormatting(captured: string, authored: string, observed: string, formatting: WordChatParagraphFormat[], inserted = new Set<number>()): string[] {
  const before = content(parse(captured), true, true).nodes
  const after = content(parse(observed), true, true).nodes.map(node => node.cloneNode(true) as Element)
  const authors: string[] = []
  for (const [index, paragraph] of after.entries()) {
    for (const name of ["rsidR", "rsidRDefault"]) {
      const original = before[index]
      if ([original, paragraph].some(node => node?.hasAttributeNS(W, name) && !/^[0-9a-fA-F]{8}$/.test(node.getAttributeNS(W, name) ?? "")))
        throw new Error("Invalid paragraph revision-session identifier")
      if (original?.hasAttributeNS(W, name)) paragraph.setAttributeNS(W, `w:${name}`, original.getAttributeNS(W, name)!)
      else paragraph.removeAttributeNS(W, name)
    }
    const properties = elements(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")
    const changes = properties ? elements(properties).filter(node => node.namespaceURI === W && node.localName === "pPrChange") : []
    if (changes.length > 1) throw new Error("Ambiguous native formatting revision")
    if (changes[0]) {
      const author = changes[0].getAttributeNS(W, "author")
      if (!author?.trim()) throw new Error("Missing native formatting author")
      authors.push(author)
      changes[0].setAttributeNS(W, "w:author", "doc.haus")
      // Word changes this paragraph-property editing-session ID when it
      // creates the verified pPrChange. It is not a content or style value.
      if (!before[index]?.hasAttributeNS(W, "rsidP")) paragraph.removeAttributeNS(W, "rsidP")
      else paragraph.setAttributeNS(W, "w:rsidP", before[index]!.getAttributeNS(W, "rsidP")!)
    }
  }
  normalizeNativeRevisionIds(before, after)
  normalizeNativeNumberingOldProperties(captured, observed, before, after)
  const expected = content(parse(authored), false, false).nodes
  // Inserted properties are already present in the verified wording import.
  // Only retained paragraphs may acquire native formatting revisions here.
  const retained = formatting.filter(op => !inserted.has(op.paragraph))
  const hasExpectedRevision = expected.some(paragraph => {
    const properties = unique(named(paragraph, "pPr"))
    return properties && named(properties, "pPrChange").length > 0
  })
  if (hasExpectedRevision) reviewedWordChatParagraphFormatting(before, after, retained)
  else {
    if (before.length !== after.length) throw new Error("Formatting changes paragraph boundaries")
    before.forEach((paragraph, index) => wordChatParagraphFormatChanged(paragraph, after[index]!))
  }
  if (after.length !== expected.length || after.some((paragraph, index) =>
    JSON.stringify(paragraphView(paragraph, "Current")) !== JSON.stringify(paragraphView(expected[index]!, "Current")) ||
    JSON.stringify(paragraphView(paragraph, "Original")) !== JSON.stringify(paragraphView(expected[index]!, "Original"))))
    throw new Error("Native formatting differs from the reviewed artifact")
  return [...new Set(authors)]
}

/** Both pieces are validated before Word is mutated. Inserted paragraph marks
 * are identified only from the separately reviewed wording artifact. */
export function reviewedMixedWordChatArtifact(artifact: WordChatAuthoredAmendment, proposalId: string, target: WordChatTarget, expected: string[], captured: string, formatting: WordChatParagraphFormat[]) {
  if (!artifact.wording || !validWordChatFormatting(target, formatting, expected.join("\r")) ||
    JSON.stringify(artifact.formatting) !== JSON.stringify(formatting)) throw new Error("Invalid mixed amendment")
  const wording = { proposalId, author: artifact.author, paragraphs: artifact.paragraphs, revisions: artifact.wording.revisions, ooxml: artifact.wording.ooxml } satisfies WordChatAuthoredAmendment
  reviewedWordChatArtifact(wording, proposalId, target, expected, captured)
  const delivered = content(parse(wording.ooxml), false, false).nodes
  const after = content(parse(artifact.ooxml), false, false).nodes
  if (delivered.some(node => paragraphMark(node)?.localName === "del")) throw new Error("Mixed deletion is not supported")
  const inserted = new Set(delivered.flatMap((node, index) => paragraphMark(node)?.localName === "ins" ? [index] : []))
  let before = delivered
  if (inserted.size) {
    if (!artifact.wording.unformattedOoxml) throw new Error("Missing insertion formatting baseline")
    const baseline = { ...wording, ooxml: artifact.wording.unformattedOoxml }
    reviewedWordChatArtifact(baseline, proposalId, target, expected, captured)
    unchangedFormattingDependencies(captured, baseline.ooxml)
    unchangedFormattingDependencies(captured, wording.ooxml)
    before = content(parse(baseline.ooxml), false, false).nodes
    if (before.length !== delivered.length) throw new Error("Insertion baseline differs")
    for (const [index, paragraph] of before.entries()) {
      if (inserted.has(index)) {
        const pPr = unique(named(paragraph, "pPr")), run = pPr && unique(named(pPr, "rPr"))
        // A new paragraph has no authority for arbitrary pre-existing direct
        // properties. Only its independently checked insertion mark is allowed.
        if (!pPr || !run || elements(pPr).length !== 1 || elements(run).length !== 1 ||
          paragraphMark(paragraph)?.localName !== "ins" || Array.from(pPr.attributes).some(a => a.namespaceURI !== XMLNS) ||
          Array.from(run.attributes).some(a => a.namespaceURI !== XMLNS)) throw new Error("Unreviewed insertion properties")
        // Inserted wording is literal, with only the reviewed paragraph-level
        // formatting. Do not let a server-authored baseline introduce bold or
        // other unreviewed run properties inside that insertion.
        const inlineHasProperties = (node: Element): boolean => {
          if (node.namespaceURI === W && node.localName === "pPr") return false
          if (node.namespaceURI === W && node.localName === "rPr") return true
          return elements(node).some(inlineHasProperties)
        }
        if (inlineHasProperties(paragraph)) throw new Error("Unreviewed insertion run formatting")
      }
      const op = inserted.has(index) ? formatting.find(item => item.paragraph === index) : undefined
      wordChatParagraphFormatChanged(paragraph, delivered[index]!, op)
    }
  } else if (artifact.wording.unformattedOoxml) throw new Error("Unexpected insertion baseline")
  unchangedFormattingDependencies(captured, artifact.ooxml)
  const byParagraph = reviewedWordChatParagraphFormatting(before, after, formatting, inserted)
  const formats = byParagraph.flat()
  if (artifact.proposalId !== proposalId || artifact.author !== "doc.haus" || JSON.stringify(artifact.paragraphs) !== JSON.stringify(expected) ||
    JSON.stringify(artifact.revisions) !== JSON.stringify([...wording.revisions, ...formats])) throw new Error("Mixed amendment metadata differs from markup")
  const nativeFormatCount = byParagraph.reduce((count, changes, index) => count + (inserted.has(index) ? 0 : changes.length), 0)
  return { wording, inserted, formats, nativeFormatCount }
}

export function sameWordChatParagraphContent(before: string, after: string): boolean {
  const view = (xml: string) => content(parse(xml), true, true).shape.map(node => canonical(node, true))
  return JSON.stringify(view(before)) === JSON.stringify(view(after))
}

/** After a native formatting write, Word may renumber matching inline
 * revisions in neighbouring paragraphs. All remaining markup stays exact. */
export function sameNativeWordChatParagraphContent(before: string, after: string): boolean {
  const original = content(parse(before), true, true)
  const observed = content(parse(after), true, true)
  // The parsed observed package is private to this comparison; normalize its
  // bound paragraphs in place so the complete shape retains its table wrapper.
  if (!normalizeNativeRevisionIds(original.nodes, observed.nodes)) return false
  return JSON.stringify(original.shape.map(node => canonical(node, true))) === JSON.stringify(observed.shape.map(node => canonical(node, true)))
}

/** The reviewed paragraphs (text and pPr) must be unchanged immediately before
 * a native text mutation. Word metadata such as rsids is not wording. */
export function sameWordChatTargetSnapshot(before: string, after: string): boolean {
  return sameWordChatParagraphContent(before, after)
}

/** Compare the complete native main story, including tables and section layout. */
export function sameNativeWordChatBodySnapshot(before: string, after: string): boolean {
  try {
    const body = (xml: string) => {
      const data = elements(parse(xml, 16_000_000).get(MAIN)!).filter(node => node.namespaceURI === PKG && node.localName === "xmlData")
      if (data.length !== 1 || elements(data[0]!).length !== 1) throw new Error("Invalid main story")
      const document = elements(data[0]!)[0]!
      const bodies = elements(document).filter(node => node.namespaceURI === W && node.localName === "body")
      if (document.namespaceURI !== W || document.localName !== "document" || bodies.length !== 1) throw new Error("Invalid main story")
      // Match the existing native projection's terminal section export rule.
      const terminalSection = elements(bodies[0]!).at(-1)
      if (terminalSection?.namespaceURI === W && terminalSection.localName === "sectPr") terminalSection.removeAttributeNS(W, "rsidSect")
      return canonical(bodies[0]!, true)
    }
    return JSON.stringify(body(before)) === JSON.stringify(body(after))
  } catch { return false }
}

function singleParagraphRuns(xml: string, version: "Current" | "Original") {
  const nodes = content(parse(xml), true, true).nodes
  if (nodes.length !== 1) throw new Error("Expected one reviewed paragraph")
  return paragraphView(nodes[0]!, version).runs
}

export type WordChatWordingHunk = { start: number; end: number; before: string; after: string }

/** Derive replacements from the already verified authored Word revisions, not
 * from model-provided offsets. Each group is a replacement, insertion or
 * deletion; ambiguous or structural groups fail closed. */
function safeUtf16Boundary(value: string, offset: number): boolean {
  return Number.isSafeInteger(offset) && offset >= 0 && offset <= value.length &&
    !(offset > 0 && offset < value.length && /[\uD800-\uDBFF]/.test(value[offset - 1]!) && /[\uDC00-\uDFFF]/.test(value[offset]!))
}
function wellFormedUtf16(value: string): boolean {
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i)
    if (code >= 0xD800 && code <= 0xDBFF) { if (++i >= value.length || value.charCodeAt(i) < 0xDC00 || value.charCodeAt(i) > 0xDFFF) return false }
    else if (code >= 0xDC00 && code <= 0xDFFF) return false
  }
  return true
}
export function authoredWordChatWordingHunks(xml: string, original: string, expected: string, start: number, end: number): WordChatWordingHunk[] {
  if (!wellFormedUtf16(original) || !wellFormedUtf16(expected) || !safeUtf16Boundary(original, start) || !safeUtf16Boundary(original, end) || start > end) throw new Error("Unsafe UTF-16 boundary")
  const nodes = content(parse(xml), false, false).nodes
  if (nodes.length !== 1) throw new Error("Expected one authored paragraph")
  const tokens: Array<{ kind: "same" | "add" | "delete"; value: string }> = []
  const visit = (node: Element) => {
    if (node.namespaceURI === W && ["pPr", "rPr"].includes(node.localName)) return
    if (node.namespaceURI === W && ["ins", "del"].includes(node.localName)) {
      if (node.getElementsByTagNameNS(W, "ins").length || node.getElementsByTagNameNS(W, "del").length || node.getAttributeNS(W, "author") !== "doc.haus") throw new Error("Unsupported nested wording revision")
      const added = node.localName === "ins", value = text(node, added ? "Current" : "Original")
      if (!value || /[\r\n\t\u000b]/.test(value)) throw new Error("Unsupported wording revision")
      tokens.push({ kind: added ? "add" : "delete", value }); return
    }
    if (node.namespaceURI === W && ["t", "delText", "tab", "br", "cr"].includes(node.localName)) {
      const value = text(node, "Original"); if (value) tokens.push({ kind: "same", value }); return
    }
    elements(node).forEach(visit)
  }
  visit(nodes[0]!)
  const hunks: WordChatWordingHunk[] = []; let at = 0, rebuilt = ""
  for (let i = 0; i < tokens.length;) {
    const token = tokens[i]!
    if (token.kind === "same") { if (original.slice(at, at + token.value.length) !== token.value) throw new Error("Authored source differs"); at += token.value.length; rebuilt += token.value; i++; continue }
    const group = [] as typeof tokens
    while (i < tokens.length && tokens[i]!.kind !== "same") group.push(tokens[i++]!)
    const deletions = group.filter(x => x.kind === "delete"), additions = group.filter(x => x.kind === "add")
    if ((deletions.length > 1 && additions.length > 0) || additions.length > 1 || !group.length) throw new Error("Unsupported wording revision group")
    const before = deletions.map(x => x.value).join(""), after = additions[0]?.value ?? "", next = at + before.length
    if (!wellFormedUtf16(before) || !wellFormedUtf16(after) || !safeUtf16Boundary(original, at) || !safeUtf16Boundary(original, next) || at < start || next > end || original.slice(at, next) !== before) throw new Error("Wording revision outside reviewed selection")
    hunks.push({ start: at, end: next, before, after }); at = next; rebuilt += after
  }
  if (!hunks.length || at !== original.length || rebuilt !== expected) throw new Error("Authored wording differs from reviewed result")
  return hunks
}

/** An insertion inherits format only if the adjacent source context is
 * unambiguous; paragraph edges require one actual neighbouring character. */
export function unambiguousWordChatInsertionFormatting(xml: string, offset: number): boolean {
  const runs = singleParagraphRuns(xml, "Original")
  const source = runs.flatMap(run => Array.from({ length: run.text.length }, (_, i) => ({ char: run.text[i]!, format: run.format })))
  const value = source.map(x => x.char).join("")
  if (!source.length || !wellFormedUtf16(value) || !safeUtf16Boundary(value, offset)) return false
  const left = source[offset - 1]?.format, right = source[offset]?.format
  return left !== undefined || right !== undefined ? left === undefined || right === undefined || left === right : false
}

/** Check unchanged characters and each hunk's source-bound format. */
export function sameWordChatHunkFormatting(before: string, after: string, hunks: WordChatWordingHunk[]): boolean {
  const prior = singleParagraphRuns(before, "Original"), actual = singleParagraphRuns(after, "Current")
  const source = prior.flatMap(run => Array.from({ length: run.text.length }, (_, i) => ({ char: run.text[i]!, format: run.format })))
  const afterOriginal = singleParagraphRuns(after, "Original").flatMap(run => Array.from({ length: run.text.length }, (_, i) => ({ char: run.text[i]!, format: run.format })))
  if (JSON.stringify(source) !== JSON.stringify(afterOriginal)) return false
  const original = source.map(x => x.char).join("")
  const expected = [...source]
  let previousEnd = 0
  for (const hunk of hunks) {
    if (hunk.start < previousEnd || !safeUtf16Boundary(original, hunk.start) || !safeUtf16Boundary(original, hunk.end) || hunk.end < hunk.start || !wellFormedUtf16(hunk.after) || original.slice(hunk.start, hunk.end) !== hunk.before || (!hunk.before && !hunk.after)) return false
    previousEnd = hunk.end
  }
  for (const hunk of [...hunks].reverse()) {
    const slice = source.slice(hunk.start, hunk.end)
    if (hunk.after && slice.length && slice.some(x => x.format !== slice[0]!.format)) return false
    if (!slice.length && !unambiguousWordChatInsertionFormatting(before, hunk.start)) return false
    const format = slice[0]?.format ?? source[hunk.start - 1]?.format ?? source[hunk.start]?.format
    if (format === undefined) return false
    expected.splice(hunk.start, hunk.end - hunk.start, ...Array.from({ length: hunk.after.length }, (_, i) => ({ char: hunk.after[i]!, format })))
  }
  const result = actual.flatMap(run => Array.from({ length: run.text.length }, (_, i) => ({ char: run.text[i]!, format: run.format })))
  return JSON.stringify(expected) === JSON.stringify(result)
}

/** A native replacement inherits one run format. Refuse a target span that
 * crosses different character formatting rather than silently restyling it. */
export function uniformWordChatFormatting(xml: string, start: number, end: number): boolean {
  const runs = singleParagraphRuns(xml, "Original")
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start) return false
  let at = 0, format: string | undefined
  for (const run of runs) {
    const from = at, to = at + run.text.length
    at = to
    if (to > start && from < end) {
      if (format !== undefined && format !== run.format) return false
      format = run.format
    }
  }
  return start === end ? unambiguousWordChatInsertionFormatting(xml, start) : at >= end && format !== undefined
}

/** The accepted result must have exactly the reviewed text and preserve the
 * target span's character formatting, with no extra format revision. */
export function sameWordChatReplacementFormatting(before: string, after: string, start: number, end: number, replacement: string): boolean {
  const prior = singleParagraphRuns(before, "Original"), result = singleParagraphRuns(after, "Current")
  let at = 0, format: string | undefined
  for (const run of prior) {
    const from = at, to = at + run.text.length
    at = to
    if (to <= start || from >= end) continue
    if (format !== undefined && format !== run.format) return false
    format = run.format
  }
  if (format === undefined) return false
  const expected: Array<{ text: string; format: string }> = []
  const push = (text: string, value: string) => {
    if (!text) return
    const last = expected.at(-1)
    if (last?.format === value) last.text += text
    else expected.push({ text, format: value })
  }
  at = 0
  let placed = false
  for (const run of prior) {
    const from = at, to = at + run.text.length
    at = to
    if (to <= start) { push(run.text, run.format); continue }
    if (from < start) push(run.text.slice(0, start - from), run.format)
    if (!placed) { push(replacement, format!); placed = true }
    if (from >= end) { push(run.text, run.format); continue }
    if (to > end) push(run.text.slice(end - from), run.format)
  }
  if (!placed) push(replacement, format!)
  return JSON.stringify(expected) === JSON.stringify(result.map(run => ({ text: run.text, format: run.format })))
}

/** Native text insertion must not rewrite the reviewed paragraph properties
 * (style, spacing, numbering reference) as a side effect. */
export function sameWordChatParagraphProperties(before: string, after: string): boolean {
  const properties = (xml: string) => content(parse(xml), true, true).nodes.map(node => {
    const paragraph = node.namespaceURI === W && node.localName === "p" ? node : null
    const pPr = paragraph && elements(paragraph).find(child => child.namespaceURI === W && child.localName === "pPr")
    return canonical(pPr ?? nullNode, true)
  })
  return JSON.stringify(properties(before)) === JSON.stringify(properties(after))
}

/** For layered wording imports compare every inline revision, including older
 * ones, in its own physical paragraph. Do not infer preservation from author. */
export function sameWordChatInlineRevisions(authored: string, observed: string): boolean {
  const revisions = (xml: string) => content(parse(xml), true, true).nodes.map(paragraph => {
    const result: unknown[] = []
    const visit = (node: Element) => {
      if (node.namespaceURI === W && ["pPr", "rPr"].includes(node.localName)) return
      if (node.namespaceURI === W && ["ins", "del"].includes(node.localName)) { result.push(canonical(node)); return }
      elements(node).forEach(visit)
    }
    visit(paragraph)
    return result
  })
  return JSON.stringify(revisions(authored)) === JSON.stringify(revisions(observed))
}

function unchangedFormattingDependencies(before: string, after: string) {
  const captured = parse(before), delivered = parse(after)
  for (const [name, part] of delivered) {
    if (name === MAIN) continue
    if (["/word/settings.xml", "/word/webSettings.xml"].includes(name) || name.startsWith("/docProps/") || /(?:header|footer|webextension)/i.test(name)) throw new Error("Formatting exceeds the paragraph scope")
    const original = captured.get(name)
    if (!original || original.getAttributeNS(PKG, "contentType") !== part.getAttributeNS(PKG, "contentType")) throw new Error("Formatting dependency changed")
    const data = elements(part)[0], prior = elements(original)[0]
    if (!data || !prior) throw new Error("Missing formatting dependency")
    if (name.endsWith(".rels")) {
      const relations = elements(elements(data)[0]!)
      if (name === "/_rels/.rels") {
        if (relations.length !== 1 || relations[0]!.getAttribute("Target") !== "word/document.xml" || relations[0]!.getAttribute("Type") !== "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" || relations[0]!.getAttribute("TargetMode") === "External") throw new Error("Invalid scoped root relationship")
      } else {
        const sourceRelations = elements(elements(prior)[0]!)
        if (relations.some(rel => !sourceRelations.some(old => JSON.stringify(canonical(rel)) === JSON.stringify(canonical(old))))) throw new Error("Formatting relationship changed")
      }
    } else if (data.localName === "binaryData" ? data.textContent?.replace(/\s/g, "") !== prior.textContent?.replace(/\s/g, "")
      : JSON.stringify(elements(data).map(node => canonical(node))) !== JSON.stringify(elements(prior).map(node => canonical(node)))) throw new Error("Formatting dependency changed")
  }
}

/** The native font operation must preserve the definition parts that determine
 * effective formatting. Other Word-maintained package parts are not compared. */
export function assertNativeWordChatInlineDependencies(before: string, after: string) {
  const relevant = (name: string) => ["/word/styles.xml", "/word/numbering.xml", "/word/fontTable.xml"].includes(name) || /^\/word\/theme\/[^/]+\.xml$/.test(name)
  const select = (xml: string) => new Map([...parse(xml,16_000_000)].filter(([name]) => relevant(name)))
  const original = select(before), observed = select(after)
  if (original.size !== observed.size || [...original.keys()].some(name => !observed.has(name))) throw new Error("inline_formatting_dependencies_changed")
  for (const [name, part] of original) {
    const next = observed.get(name)!, data = elements(part), nextData = elements(next)
    if (part.getAttributeNS(PKG,"contentType") !== next.getAttributeNS(PKG,"contentType") || data.length !== 1 || nextData.length !== 1 || JSON.stringify(canonical(data[0]!)) !== JSON.stringify(canonical(nextData[0]!))) throw new Error("inline_formatting_dependencies_changed")
  }
}

/** Native review IDs may be reassigned, but remain unique within the story. */
export function assertNativeInlineWordChatRevisionIds(xml: string) {
  const main = parse(xml,16_000_000).get(MAIN)!
  const ids = ["ins","del","rPrChange","pPrChange"].flatMap(name => Array.from(main.getElementsByTagNameNS(W,name)).map(node => node.getAttributeNS(W,"id")))
  if (ids.some(id => !id || !/^\d+$/.test(id)) || new Set(ids).size !== ids.length) throw new Error("inline_revision_id_collision")
}

/** Match prior formatting carriers by full payload, changing only the old ID.
 * The subsequent paragraph/target comparison still verifies their positions. */
function normalizeInlinePriorFormatIds(original: Element, current: Element): boolean {
  const revisions = (paragraph: Element, names: string[]) => names.flatMap(name => Array.from(paragraph.getElementsByTagNameNS(W,name)))
  const allNames = ["ins", "del", "rPrChange", "pPrChange"]
  for (const paragraph of [original,current]) {
    const ids = revisions(paragraph,allNames).map(node => node.getAttributeNS(W,"id"))
    if (ids.some(id => !id || !/^\d+$/.test(id)) || new Set(ids).size !== ids.length) return false
  }
  const old = revisions(original,["rPrChange","pPrChange"]), now = revisions(current,["rPrChange","pPrChange"])
  const signature = (node: Element) => { const clone = node.cloneNode(true) as Element; clone.removeAttributeNS(W,"id"); return JSON.stringify(canonical(clone)) }
  const used = new Set<Element>()
  for (const prior of old) {
    const matches = now.filter(node => !used.has(node) && signature(node) === signature(prior))
    if (matches.length !== 1) return false
    const match = matches[0]!, attr = prior.getAttributeNodeNS(W,"id")!
    const name = attr.name
    prior.removeAttributeNS(W,"id"); prior.setAttributeNS(W,name,match.getAttributeNS(W,"id")!)
    used.add(match)
  }
  return true
}

/** Inline-only unchanged paragraph comparator; no comment IDs are omitted. */
export function sameNativeInlineWordChatParagraph(before: string, after: string): boolean {
  const fail = () => { throw new Error("invalid paragraph") }
  const parser = () => new DOMParser({errorHandler:{warning:fail,error:fail,fatalError:fail}})
  const original = parser().parseFromString(before,"application/xml").documentElement
  const current = parser().parseFromString(after,"application/xml").documentElement
  if (!original || !current || original.namespaceURI !== W || current.namespaceURI !== W || original.localName !== "p" || current.localName !== "p") return false
  return normalizeInlinePriorFormatIds(original,current) && normalizeNativeRevisionIds([current],[original]) && JSON.stringify(canonical(original,true)) === JSON.stringify(canonical(current,true))
}

/** Check the captured source before the native font operation. */
export function assertNativeWordChatInlineTarget(captured: string, target: WordChatTarget, formatting: WordChatInlineFormat) {
  const paragraphs = content(parse(captured),true,true).nodes
  if (paragraphs.length !== 1) throw new Error("inline_source_paragraph")
  assertWordChatInlineFormattingTarget(paragraphs[0]!,target,formatting)
}

export function verifiedNativeWordChatInlineFormatting(captured: string, observed: string, target: WordChatTarget, formatting: WordChatInlineFormat, effectiveProof: WordChatInlineFormat) {
  const before = content(parse(captured),true,true).nodes, after = content(parse(observed),true,true).nodes
  if (before.length !== 1 || after.length !== 1) throw new Error("inline_native_paragraph")
  // Word updates these session IDs on the paragraph and formatted runs. Keep
  // the shared verifier strict; only the native bridge accepts valid metadata.
  const original = before[0]!.cloneNode(true) as Element, current = after[0]!.cloneNode(true) as Element
  const stripSessionIds = (paragraph: Element) => {
    const remove = (node: Element, name: string) => {
      const attr = node.getAttributeNodeNS(W,name)
      if (attr) { if (!/^[0-9a-fA-F]{8}$/.test(attr.value)) throw new Error("inline_invalid_native_rsid"); node.removeAttributeNS(W,name) }
    }
    remove(paragraph,"rsidR"); remove(paragraph,"rsidRDefault")
    const visible = (node: Element): string => {
      if (node.namespaceURI === W) {
        if (["del","pPr","rPr"].includes(node.localName)) return ""
        if (node.localName === "t") return node.textContent ?? ""
        if (node.localName === "tab") return "\t"
        if (["br","cr"].includes(node.localName)) return "\v"
      }
      return elements(node).map(visible).join("")
    }
    let offset = 0
    for (const node of elements(paragraph)) {
      const length = visible(node).length
      if (node.namespaceURI === W && node.localName === "r" && offset < target.endOffset && offset + length > target.startOffset) remove(node,"rsidRPr")
      offset += length
    }
  }
  stripSessionIds(original); stripSessionIds(current)
  // Map the old IDs to the observed ones, so the verifier checks new carrier
  // collisions against the IDs actually present in Word after the operation.
  if (!normalizeInlinePriorFormatIds(original,current) || !normalizeNativeRevisionIds([current],[original])) throw new Error("inline_prior_revision_changed")
  const actual = verifyWordChatInlineFormatting(original,current,target,formatting,undefined,effectiveProof)
  if (!actual.revisions.length) throw new Error("inline_missing_revision")
  return [...new Set(actual.revisions.map(change => change.author))]
}

/** Strict paragraph comparison, except Word's evidenced inline revision-ID
 * reassignment, using the existing matched-payload normalization. */
export function sameNativeProjectedWordChatParagraph(before: string, after: string): boolean {
  const fail = () => { throw new Error("invalid paragraph") }
  const parser = () => new DOMParser({errorHandler:{warning:fail,error:fail,fatalError:fail}})
  const original = parser().parseFromString(before,"application/xml").documentElement
  const observed = parser().parseFromString(after,"application/xml").documentElement
  if (!original || !observed || original.namespaceURI !== W || observed.namespaceURI !== W || original.localName !== "p" || observed.localName !== "p") return false
  return normalizeNativeRevisionIds([original],[observed]) && JSON.stringify(canonical(original,true)) === JSON.stringify(canonical(observed,true))
}

/** Compare the complete target and every dependency actually imported. Word's
 * selected OOXML includes changing write IDs and unrelated document settings;
 * those do not authorize an edit or form part of the delivered paragraph. */
const diagnosticName = (value: string) => value.replace(/[^a-zA-Z0-9_./:-]/g, "?").slice(0, 120)
function canonicalDifference(before: unknown, after: unknown, path: string): string | undefined {
  if (JSON.stringify(before) === JSON.stringify(after)) return
  if (before === undefined || after === undefined) return `${path}: XML node ${before === undefined ? "added" : "missing"}`
  if (!Array.isArray(before) || !Array.isArray(after)) return `${path}: XML shape changed`
  if (before[0] === "text" || after[0] === "text") return `${path}: text changed (text omitted)`
  // canonical() represents each element as namespace/name/attributes/children.
  if (before.length === 4 && after.length === 4 && typeof before[0] === "string" && typeof after[0] === "string") {
    if (before[0] !== after[0] || before[1] !== after[1]) return `${path}: element name or namespace changed`
    const element = `${path}/${diagnosticName(before[1] as string)}`
    const oldAttributes = before[2] as string[][], newAttributes = after[2] as string[][]
    const key = (attr: string[]) => JSON.stringify(attr.slice(0, 2))
    for (const attr of oldAttributes) {
      const current = newAttributes.find(candidate => key(candidate) === key(attr))
      if (!current) return `${element}/@${diagnosticName(attr[1]!)}: attribute missing`
      if (attr[2] !== current[2]) return `${element}/@${diagnosticName(attr[1]!)}: attribute changed (value omitted)`
    }
    const added = newAttributes.find(attr => !oldAttributes.some(candidate => key(candidate) === key(attr)))
    if (added) return `${element}/@${diagnosticName(added[1]!)}: attribute added`
    return canonicalDifference(before[3], after[3], element)
  }
  for (let index = 0; index < Math.max(before.length, after.length); index++) {
    const difference = canonicalDifference(before[index], after[index], `${path}[${index + 1}]`)
    if (difference) return difference
  }
  return `${path}: XML metadata changed`
}
export function wordChatOoxmlDifference(before: string, current: string, delivered: string): string | undefined {
  const names = [...parse(delivered).keys()].sort()
  if (names.some(name => ["/word/settings.xml", "/word/webSettings.xml"].includes(name) || name.startsWith("/docProps/") || /(?:header|footer|webextension)/i.test(name))) throw new Error("Import exceeds the clause scope")
  const oldParts = parse(before), newParts = parse(current)
  for (const name of names) {
    const oldPart = oldParts.get(name), newPart = newParts.get(name), path = diagnosticName(name)
    if (!oldPart || !newPart) return `${path}: dependency missing from ${!oldPart ? "reviewed" : "current"} Word capture`
    if (name === MAIN) {
      if (oldPart.getAttributeNS(PKG, "contentType") !== newPart.getAttributeNS(PKG, "contentType")) return `${path}: contentType changed`
      const difference = canonicalDifference(content(oldParts, true, true).shape.map(node => canonical(node, true)), content(newParts, true, true).shape.map(node => canonical(node, true)), path)
      if (difference) return difference
    } else {
      const difference = canonicalDifference(canonical(oldPart), canonical(newPart), path)
      if (difference) return difference
    }
  }
}
export function sameWordChatOoxml(before: string, current: string, delivered: string): boolean {
  return wordChatOoxmlDifference(before, current, delivered) === undefined
}

/** Verify both views and the actual authored revision spans. The server cannot
 * expand a selected subspan merely by returning matching final paragraph text. */
export function reviewedWordChatArtifact(artifact: WordChatAuthoredAmendment, proposalId: string, target: WordChatTarget, expected: string[], captured?: string, formatting?: WordChatParagraphFormat[]) {
  if (!artifact || artifact.proposalId !== proposalId || artifact.author !== "doc.haus" || !Array.isArray(artifact.paragraphs) || JSON.stringify(artifact.paragraphs) !== JSON.stringify(expected) || !Array.isArray(artifact.revisions)) throw new Error("The authored amendment differs from the reviewed wording")
  const paragraphs = content(parse(artifact.ooxml), false, false).nodes
  if (formatting !== undefined) {
    if (!captured || !validWordChatFormatting(target, formatting) || JSON.stringify(artifact.formatting) !== JSON.stringify(formatting)) throw new Error("The formatting differs from the reviewed proposal")
    const before = content(parse(captured), true, true).nodes
    if (JSON.stringify(before.map(node => text(node))) !== JSON.stringify(expected)) throw new Error("The formatted passage changed")
    unchangedFormattingDependencies(captured, artifact.ooxml)
    const revisions = reviewedWordChatParagraphFormatting(before, paragraphs, formatting)
    if (JSON.stringify(revisions.flat()) !== JSON.stringify(artifact.revisions)) throw new Error("The formatting revision metadata differs from its Word markup")
    return revisions
  }
  if (artifact.formatting !== undefined) throw new Error("Unreviewed paragraph formatting")
  const current = clauseView(artifact.ooxml, "Current"), original = clauseView(artifact.ooxml, "Original")
  const baseline = captured ? clauseView(captured, "Original") : undefined
  if (current.length !== expected.length || original.length !== target.paragraphs.length ||
    JSON.stringify(current.map(p => p.runs.map(r => r.text).join(""))) !== JSON.stringify(expected) ||
    !baseline && JSON.stringify(original.map(p => p.runs.map(r => r.text).join(""))) !== JSON.stringify(target.paragraphs.map(p => p.text)))
    throw new Error("The authored amendment differs from the reviewed passage")
  if (baseline && JSON.stringify(baseline) !== JSON.stringify(original)) throw new Error("The authored amendment changes original formatting or bookmark anchors")
  const prior: Array<Map<string, number>> = []
  if (captured) for (const paragraph of content(parse(captured), true, true).nodes) {
    const marks = new Map<string, number>()
    const visit = (node: Element) => {
      if (node.namespaceURI === W && ["pPr", "rPr"].includes(node.localName)) return
      if (node.namespaceURI === W && ["ins", "del"].includes(node.localName)) {
        const key = JSON.stringify(canonical(node)), count = marks.get(key) ?? 0
        marks.set(key, count + 1); return
      }
      elements(node).forEach(visit)
    }
    visit(paragraph)
    prior.push(marks)
  }
  const revisions: Array<Array<{ author: string; type: "Added" | "Deleted"; text: string }>> = []
  let offset = 0
  const start = target.startOffset, end = target.paragraphs.slice(0, -1).reduce((at, p) => at + p.text.length + 1, 0) + target.endOffset
  let originalIndex = 0
  for (const [index, paragraph] of paragraphs.entries()) {
    const changes: typeof revisions[number] = []
    const mark = paragraphMark(paragraph)
    const priorInParagraph = mark?.localName === "ins" ? undefined : prior[originalIndex]
    const visit = (node: Element) => {
      if (node.namespaceURI === W && ["pPr", "rPr"].includes(node.localName)) return
      if (node.namespaceURI === W && ["ins", "del"].includes(node.localName)) {
        const added = node.localName === "ins", wording = text(node, added ? "Current" : "Original")
        const key = JSON.stringify(canonical(node)), priorCount = priorInParagraph?.get(key) ?? 0
        if (priorCount) {
          priorInParagraph!.set(key, priorCount - 1)
          if (added) offset += wording.length
          return
        }
        if (!wording || node.getAttributeNS(W, "author") !== "doc.haus" || offset < start || offset + (added ? 0 : wording.length) > end || node.getElementsByTagNameNS(W, "ins").length || node.getElementsByTagNameNS(W, "del").length) throw new Error("The authored revision exceeds the approved selection")
        changes.push({ author: "doc.haus", type: added ? "Added" : "Deleted", text: wording })
        if (!added) offset += wording.length
        return
      }
      if (node.namespaceURI === W && ["t", "delText", "tab", "br", "cr"].includes(node.localName)) { offset += text(node).length; return }
      for (const child of elements(node)) visit(child)
    }
    visit(paragraph)
    if (mark) {
      const added = mark.localName === "ins"
      if (index === paragraphs.length - 1 || mark.getAttributeNS(W, "author") !== "doc.haus" || offset < start || offset + (added ? 0 : 1) > end)
        throw new Error("The paragraph revision exceeds the approved selection")
      changes.push({ author: "doc.haus", type: added ? "Added" : "Deleted", text: "\r" })
    }
    if (index < paragraphs.length - 1 && mark?.localName !== "ins") offset++
    if (mark?.localName !== "ins") originalIndex++
    revisions.push(changes)
  }
  if (prior.some(marks => [...marks.values()].some(Boolean))) throw new Error("The authored amendment changes existing revisions")
  const signature = (changes: typeof artifact.revisions) => JSON.stringify(changes.map(change => [change.author, change.type, change.text]))
  if (!revisions.flat().length || signature(revisions.flat()) !== signature(artifact.revisions)) throw new Error("The authored revision metadata differs from its Word markup")
  return revisions
}

/** Verify a backend-authored exact whole-control replacement. The caller must
 * separately pin the live Word paragraph/control IDs and recapture freshness. */
export function reviewedWordChatInlineControl(artifact: WordChatAuthoredAmendment, proposalId: string, target: WordChatTarget, replacement: string, captured: string) {
  const fail = (): never => { throw new Error("Cannot verify the reviewed inline control") }
  if (!target.id || target.paragraphs.length !== 1 || !target.paragraphs[0]?.id || !replacement || /[\r\n\t\v]/.test(replacement) ||
    target.text[0] === replacement[0] || target.text.at(-1) === replacement.at(-1)) fail()
  const sourceParts = parse(captured), resultParts = parse(artifact.ooxml)
  const source = content(sourceParts, true, true).nodes, result = content(resultParts, false, false).nodes
  if (source.length !== 1 || result.length !== 1) fail()
  const oneControl = (paragraph: Element, parts: Map<string, Element>) => {
    const main = dependencyRoot(parts, MAIN)
    if (!main) return fail()
    const controls = Array.from(main.getElementsByTagNameNS(W, "sdt"))
    if (controls.length !== 1 || controls[0]!.parentNode !== paragraph) return fail()
    const protectedNames = ["commentRangeStart", "commentRangeEnd", "commentReference", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "fldChar", "fldSimple", "instrText", "drawing", "pict", "object", "hyperlink", "footnoteReference", "endnoteReference"]
    if (protectedNames.some(name => main.getElementsByTagNameNS(W, name).length)) return fail()
    for (const name of ["ins", "del"]) for (const node of Array.from(main.getElementsByTagNameNS(W, name))) {
      let ancestor: Node | null = node
      while (ancestor && ancestor !== paragraph) ancestor = ancestor.parentNode
      if (!ancestor) return fail()
    }
    const control = controls[0]!, pieces = elements(control)
    if (pieces.length !== 2 || pieces[0]!.namespaceURI !== W || pieces[0]!.localName !== "sdtPr" || pieces[1]!.namespaceURI !== W || pieces[1]!.localName !== "sdtContent") return fail()
    const properties = elements(pieces[0]!), ids = named(pieces[0]!, "id")
    if (ids.length !== 1 || !/^-?\d+$/.test(ids[0]!.getAttributeNS(W, "val") ?? "") ||
      properties.some(p => p.namespaceURI !== W || !["id", "tag", "alias"].includes(p.localName) || elements(p).length) ||
      new Set(properties.map(p => p.localName)).size !== properties.length) return fail()
    return { control, properties: pieces[0]!, body: pieces[1]!, id: ids[0]!.getAttributeNS(W, "val")! }
  }
  const before = oneControl(source[0]!, sourceParts), after = oneControl(result[0]!, resultParts)
  const plainText = (node: Element) => elements(node).length === 0 && Array.from(node.attributes).every(attr => attr.namespaceURI === XMLNS ||
    attr.namespaceURI === "http://www.w3.org/XML/1998/namespace" && attr.localName === "space" && ["preserve", "default"].includes(attr.value)) &&
    (!/^\s|\s$/.test(node.textContent ?? "") || node.getAttributeNS("http://www.w3.org/XML/1998/namespace", "space") === "preserve")
  const priorRuns = elements(before.body), originalRun = priorRuns[0], originalFields = originalRun ? elements(originalRun) : []
  if (priorRuns.length !== 1 || originalRun?.namespaceURI !== W || originalRun.localName !== "r" ||
    originalFields.some(e => e.namespaceURI !== W || !["rPr", "t"].includes(e.localName)) || named(originalRun, "rPr").length > 1 || named(originalRun, "t").length !== 1 ||
    !plainText(named(originalRun, "t")[0]!) || text(before.body) !== target.text) fail()
  const siblings = elements(source[0]!), start = siblings.slice(0, siblings.indexOf(before.control)).map(node => text(node)).join("").length
  if (start !== target.startOffset || start + target.text.length !== target.endOffset || text(source[0]!) !== target.paragraphs[0]!.text) fail()
  const expected = target.paragraphs[0]!.text.slice(0, start) + replacement + target.paragraphs[0]!.text.slice(target.endOffset)
  reviewedWordChatArtifact(artifact, proposalId, target, [expected], captured)
  unchangedFormattingDependencies(captured, artifact.ooxml)
  const marks = elements(after.body)
  if (marks.length !== 2 || marks[0]!.namespaceURI !== W || marks[0]!.localName !== "del" || marks[1]!.namespaceURI !== W || marks[1]!.localName !== "ins") fail()
  if (marks[0]!.getAttributeNS(W, "id") === marks[1]!.getAttributeNS(W, "id")) fail()
  const runAttributes = JSON.stringify(canonical(originalRun!.cloneNode(false))), runProperties = JSON.stringify(canonical(named(originalRun!, "rPr")[0] ?? nullNode))
  for (const [index, mark] of marks.entries()) {
    const runs = elements(mark)
    if (Array.from(mark.attributes).some(attr => attr.namespaceURI !== XMLNS && (attr.namespaceURI !== W || !["id", "author", "date"].includes(attr.localName))) ||
      mark.getAttributeNS(W, "author") !== "doc.haus" || !/^\d+$/.test(mark.getAttributeNS(W, "id") ?? "") ||
      !Number.isFinite(Date.parse(mark.getAttributeNS(W, "date") ?? "")) || runs.length !== 1 || runs[0]!.namespaceURI !== W || runs[0]!.localName !== "r" ||
      JSON.stringify(canonical(runs[0]!.cloneNode(false))) !== runAttributes || JSON.stringify(canonical(named(runs[0]!, "rPr")[0] ?? nullNode)) !== runProperties) fail()
    const fields = elements(runs[0]!), value = named(runs[0]!, index === 0 ? "delText" : "t")
    if (value.length !== 1 || fields.some(field => field.namespaceURI !== W || !["rPr", index === 0 ? "delText" : "t"].includes(field.localName)) ||
      named(runs[0]!, "rPr").length > 1 || !plainText(value[0]!) || value[0]!.textContent !== (index === 0 ? target.text : replacement)) fail()
  }
  const masked = (paragraph: Element) => {
    const clone = paragraph.cloneNode(true) as Element, sdt = clone.getElementsByTagNameNS(W, "sdt")[0]!, body = named(sdt, "sdtContent")[0]!
    while (body.firstChild) body.removeChild(body.firstChild)
    return JSON.stringify(canonical(clone, true))
  }
  if (before.id !== after.id || masked(source[0]!) !== masked(result[0]!)) fail()
  return { paragraphId: target.paragraphs[0]!.id, controlId: before.id, startOffset: start, endOffset: target.endOffset, before: target.text, after: replacement }
}
