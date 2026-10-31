import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { unzipSync } from "fflate"
import type { WordChatTableRowDeletionProposal } from "../../../dochaus/lib/word-chat-amendment-contract"
import type { WordChatAuthoredTableRowDeletion } from "../../../dochaus/lib/word-chat-contract"
import { assertNativeWordChatInlineDependencies, projectNativeWordChatBody } from "./word-chat-host-ooxml"
import { normalizeNativeWordChatCommentIds } from "./word-chat-native-comments"

type Inventory = Array<{ id: string; text: string; tableNestingLevel?: number }>
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", W14 = "http://schemas.microsoft.com/office/word/2010/wordml", W16 = "http://schemas.microsoft.com/office/word/2023/wordml/word16du", PKG = "http://schemas.microsoft.com/office/2006/xmlPackage", XMLNS = "http://www.w3.org/2000/xmlns/"
const elements = (n: Node) => Array.from(n.childNodes).filter((c): c is Element => c.nodeType === 1)
const named = (n: Node, name: string) => elements(n).filter(e => e.namespaceURI === W && e.localName === name)
const only = (n: Node, names: string[]) => elements(n).every(e => e.namespaceURI === W && names.includes(e.localName))
const fail = (reason = "invalid XML or package"): never => { throw Error(`Native row deletion could not be verified: ${reason}`) }
const eq = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const numericId = (id: string) => id.replace(/^0+(?=\d)/, "")
function parse(xml: string) {
  if (xml.length > 16_000_000 || /<!DOCTYPE/i.test(xml)) return fail()
  const reject = () => { throw Error("Invalid Word XML") }
  const doc = new DOMParser({ errorHandler: { warning: reject, error: reject, fatalError: reject } }).parseFromString(xml, "application/xml")
  if (doc.doctype || !doc.documentElement) return fail()
  return doc
}
function flat(xml: string) {
  const doc = parse(xml)
  if (doc.documentElement.namespaceURI !== PKG || doc.documentElement.localName !== "package") return fail()
  const parts = elements(doc.documentElement).filter(e => e.namespaceURI === PKG && e.localName === "part" && e.getAttributeNS(PKG, "name") === "/word/document.xml")
  if (parts.length !== 1) return fail()
  const bodies = Array.from(parts[0]!.getElementsByTagNameNS(W, "body")); if (bodies.length !== 1) return fail()
  return { doc, body: bodies[0]! }
}
function canonical(n: Node, native = false): unknown {
  if (n.nodeType === 1) {
    const e = n as Element
    return [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.namespaceURI !== XMLNS && !(native && a.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(a.localName))).map(a => [a.namespaceURI, a.localName, a.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))), Array.from(e.childNodes).map(x => canonical(x, native)).filter(x => x !== undefined)]
  }
  if (n.nodeType === 3 || n.nodeType === 4) return n.nodeValue?.trim() || ["t", "delText"].includes((n.parentNode as Element)?.localName) ? ["text", n.nodeValue] : undefined
  if (n.nodeType === 7 || n.nodeType === 8) return [n.nodeType, n.nodeName, n.nodeValue]
  return undefined
}
function tableRow(body: Element, p: WordChatTableRowDeletionProposal) {
  const tables = named(body, "tbl")
  if (!p || p.kind !== "table_row_delete" || !Number.isSafeInteger(p.table?.tableIndex) || p.table.tableIndex < 0 || p.table.tableIndex >= tables.length || !Number.isSafeInteger(p.rowIndex) || p.rowIndex < 0 || !Array.isArray(p.cells) || !p.cells.length || p.cells.length > 32) return fail("table geometry")
  const table = tables[p.table.tableIndex]!, rows = named(table, "tr")
  if (!only(table, ["tblPr", "tblGrid", "tr"]) || rows.length < 2 || rows.length !== p.table.rows.length || p.rowIndex >= rows.length || p.table.rows.some(row => row.length !== p.cells.length)) return fail("table geometry")
  return { table, rows, row: rows[p.rowIndex]! }
}
function target(body: Element, p: WordChatTableRowDeletionProposal) {
  const found = tableRow(body, p)
  const currentText = (node: Node): string => {
    if (node.nodeType !== 1) return ""
    const e = node as Element
    if (e.namespaceURI === W && ["del", "moveFrom"].includes(e.localName)) return ""
    if (e.namespaceURI === W && e.localName === "t") return e.textContent ?? ""
    return elements(e).map(currentText).join("")
  }
  // A range that begins outside the row can still intersect its contents.
  const active = new Set<string>(); let fieldDepth = 0
  const scan = (e: Element) => {
    if (e === found.row && (active.size || fieldDepth)) return fail("target qualification")
    if (e.namespaceURI === W) {
      const starts: Record<string, string> = { commentRangeStart: "comment", bookmarkStart: "bookmark", permStart: "perm", moveFromRangeStart: "moveFrom", moveToRangeStart: "moveTo", customXmlInsRangeStart: "customIns", customXmlDelRangeStart: "customDel" }, ends: Record<string, string> = { commentRangeEnd: "comment", bookmarkEnd: "bookmark", permEnd: "perm", moveFromRangeEnd: "moveFrom", moveToRangeEnd: "moveTo", customXmlInsRangeEnd: "customIns", customXmlDelRangeEnd: "customDel" }
      const family = starts[e.localName] ?? ends[e.localName]
      if (family) {
        const id = e.getAttributeNS(W, "id"), key = `${family}:${id}`
        if (!id || starts[e.localName] && active.has(key) || ends[e.localName] && !active.has(key)) return fail("target qualification")
        if (starts[e.localName]) active.add(key); else active.delete(key)
      }
      if (e.localName === "fldChar") {
        const kind = e.getAttributeNS(W, "fldCharType")
        if (kind === "begin") fieldDepth++; else if (kind === "end") { if (!fieldDepth) return fail("target qualification"); fieldDepth-- }
      }
    }
    for (const child of elements(e)) scan(child)
    if (e === found.row && (active.size || fieldDepth)) return fail("target qualification")
  }
  scan(body)
  if (active.size || fieldDepth) return fail("target qualification")
  for (const [i, row] of found.rows.entries()) {
    const cells = named(row, "tc"), props = named(row, "trPr")
    if (!only(row, ["trPr", "tc"]) || props.length > 1 || cells.length !== p.cells.length || cells.some(c => ["gridSpan", "vMerge", "hMerge"].some(n => c.getElementsByTagNameNS(W, n).length))) return fail("target qualification")
    if (cells.some((cell, j) => currentText(cell) !== p.table.rows[i]![j]!.text)) return fail("target qualification")
    if (i === p.rowIndex) {
      if (props.some(pr => pr.getElementsByTagNameNS(W, "tblHeader").length) || !only(row, ["trPr", "tc"]) || Array.from(row.getElementsByTagNameNS(W, "*")).some(e => ["ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange", "trPrChange", "tcPrChange", "cellIns", "cellDel", "cellMerge", "fldChar", "instrText", "vanish", "webHidden", "commentRangeStart", "commentRangeEnd", "commentReference", "bookmarkStart", "bookmarkEnd", "permStart", "permEnd"].includes(e.localName))) return fail("target qualification")
      for (const [j, cell] of cells.entries()) {
        const ps = named(cell, "p"), cps = named(cell, "tcPr")
        if (!only(cell, ["tcPr", "p"]) || cps.length > 1 || ps.length !== 1 || !only(ps[0]!, ["pPr", "r"]) || named(ps[0]!, "pPr").length > 1) return fail("target qualification")
        const paragraph = ps[0]!, runs = named(paragraph, "r")
        if (!runs.length || runs.some(r => !only(r, ["rPr", "t"]) || named(r, "rPr").length > 1 || named(r, "t").length !== 1 || elements(named(r, "t")[0]!).length)) return fail("target qualification")
        if (Array.from(paragraph.getElementsByTagNameNS(W, "*")).some(e => ["ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange", "fldChar", "instrText", "vanish", "webHidden", "commentRangeStart", "commentRangeEnd", "commentReference", "bookmarkStart", "bookmarkEnd"].includes(e.localName))) return fail("target qualification")
        if (runs.map(r => named(r, "t")[0]!.textContent ?? "").join("") !== p.cells[j] || p.table.rows[i]![j]!.text !== p.cells[j]) return fail("target qualification")
      }
    }
  }
  return found
}
export function assertNativeWordChatRowDeletionTarget(beforeXml: string, inventory: Inventory, proposal: WordChatTableRowDeletionProposal) {
  const projection = projectNativeWordChatBody(beforeXml, inventory), { body } = flat(beforeXml), { table, rows } = target(body, proposal)
  const tablePos = elements(body).indexOf(table), seen = new Set<string>()
  for (const [i, row] of rows.entries()) for (const [j, cell] of named(row, "tc").entries()) {
    const ps = named(cell, "p"); if (ps.length !== 1) return fail("target binding")
    const path = `body/tbl[${tablePos}]/tr[${elements(table).indexOf(row)}]/tc[${elements(row).indexOf(cell)}]/p[${elements(cell).indexOf(ps[0]!)}]`
    const match = projection.paragraphs.filter(x => x.path === path), bound = proposal.table.rows[i]![j]!
    if (match.length !== 1 || match[0]!.id !== bound.id || match[0]!.text !== bound.text || !bound.id || seen.has(bound.id)) return fail("target binding")
    seen.add(bound.id)
  }
}
/** Reconstruct the original row by removing only validated deletion carriers. */
function inverse(original: Element, observed: Element, native: boolean, expectedAuthor?: string) {
  const copy = observed.cloneNode(true) as Element, oldCells = named(original, "tc"), cells = named(copy, "tc"), ids = new Set<string>(), authors = new Set<string>()
  let stamp: string | undefined
  const mark = (e: Element, content: boolean) => {
    const id = e.getAttributeNS(W, "id"), author = e.getAttributeNS(W, "author"), date = e.getAttributeNS(W, "date"), utc = e.getAttributeNS(W16, "dateUtc")
    if (!id || !/^\d+$/.test(id) || ids.has(numericId(id)) || !author?.trim() || expectedAuthor && author !== expectedAuthor || !date || Number.isNaN(Date.parse(date)) || utc && utc !== date || Array.from(e.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === W && ["id", "author", "date"].includes(a.localName)) && !(native && a.namespaceURI === W16 && a.localName === "dateUtc")) || !content && e.childNodes.length) return fail("deletion carriers or row inverse")
    const key = JSON.stringify([author, date]); if (stamp && stamp !== key) return fail("deletion carriers or row inverse"); stamp = key; ids.add(numericId(id)); authors.add(author)
  }
  const rowProps = named(copy, "trPr"), oldRowProps = named(original, "trPr"), rowMarks = rowProps.length === 1 ? named(rowProps[0]!, "del") : []
  if (rowProps.length !== 1 || rowMarks.length !== 1 || cells.length !== oldCells.length) return fail("deletion carriers or row inverse")
  mark(rowMarks[0]!, false); rowProps[0]!.removeChild(rowMarks[0]!)
  if (!oldRowProps.length && !elements(rowProps[0]!).length && !Array.from(rowProps[0]!.attributes).some(a => a.namespaceURI !== XMLNS)) copy.removeChild(rowProps[0]!)
  for (let i = 0; i < cells.length; i++) {
    const oldP = named(oldCells[i]!, "p")[0]!, ps = named(cells[i]!, "p"); if (ps.length !== 1) return fail("deletion carriers or row inverse")
    const p = ps[0]!, pp = named(p, "pPr"), oldPP = named(oldP, "pPr"), pr = pp.length === 1 ? named(pp[0]!, "rPr") : [], oldPR = oldPP.length === 1 ? named(oldPP[0]!, "rPr") : [], pm = pr.length === 1 ? named(pr[0]!, "del") : [], wrappers = named(p, "del")
    if (pp.length !== 1 || pr.length !== 1 || pm.length !== 1 || !wrappers.length || !only(p, ["pPr", "del"])) return fail("deletion carriers or row inverse")
    mark(pm[0]!, false); pr[0]!.removeChild(pm[0]!)
    if (!oldPR.length && !elements(pr[0]!).length && !Array.from(pr[0]!.attributes).some(a => a.namespaceURI !== XMLNS)) pp[0]!.removeChild(pr[0]!)
    if (!oldPP.length && !elements(pp[0]!).length && !Array.from(pp[0]!.attributes).some(a => a.namespaceURI !== XMLNS)) p.removeChild(pp[0]!)
    for (const wrapper of wrappers) {
      mark(wrapper, true); const runs = named(wrapper, "r")
      if (!runs.length || !only(wrapper, ["r"])) return fail("deletion carriers or row inverse")
      for (const run of runs) {
        if (!only(run, ["rPr", "delText"]) || named(run, "rPr").length > 1 || named(run, "delText").length !== 1) return fail("deletion carriers or row inverse")
        if (native && run.hasAttributeNS(W, "rsidDel")) { if (!/^[0-9a-fA-F]{8}$/.test(run.getAttributeNS(W, "rsidDel")!)) return fail("deletion carriers or row inverse"); run.removeAttributeNS(W, "rsidDel") }
        const dt = named(run, "delText")[0]!, t = copy.ownerDocument!.createElementNS(W, "w:t")
        for (const a of Array.from(dt.attributes)) t.setAttributeNS(a.namespaceURI, a.name, a.value)
        while (dt.firstChild) t.appendChild(dt.firstChild)
        run.replaceChild(t, dt); p.insertBefore(run, wrapper)
      }
      p.removeChild(wrapper)
    }
  }
  if (native) {
    const metadata = (a: Element, b: Element) => {
      if (a.getAttributeNS(W14, "paraId") !== b.getAttributeNS(W14, "paraId") || Boolean(a.getAttributeNS(W14, "textId")) !== Boolean(b.getAttributeNS(W14, "textId"))) return fail("deletion carriers or row inverse")
      if (b.hasAttributeNS(W, "rsidDel")) { if (!/^[0-9a-fA-F]{8}$/.test(b.getAttributeNS(W, "rsidDel")!)) return fail("deletion carriers or row inverse"); b.removeAttributeNS(W, "rsidDel") }
      if (a.hasAttributeNS(W14, "textId")) b.setAttributeNS(W14, "w14:textId", a.getAttributeNS(W14, "textId")!)
    }
    metadata(original, copy); for (let i = 0; i < cells.length; i++) metadata(named(oldCells[i]!, "p")[0]!, named(cells[i]!, "p")[0]!)
  }
  if (!eq(canonical(original, native), canonical(copy, native))) throw Error("row inverse mismatch")
  return { ids, authors }
}
function revisionNodes(body: Element) {
  const result: Array<{ path: string; node: Element }> = []
  const visit = (e: Element, path: string) => {
    if (e.namespaceURI === W && ["ins", "del", "pPrChange", "rPrChange"].includes(e.localName)) result.push({ path, node: e })
    elements(e).forEach((c, i) => visit(c, `${path}/${c.localName}[${i}]`))
  }
  visit(body, "body"); return result
}
export function verifyNativeWordChatDeletedRow(beforeXml: string, afterXml: string, beforeInventory: Inventory, afterInventory: Inventory, proposal: WordChatTableRowDeletionProposal): { authors: string[] } {
  assertNativeWordChatRowDeletionTarget(beforeXml, beforeInventory, proposal)
  const before = projectNativeWordChatBody(beforeXml, beforeInventory), after = projectNativeWordChatBody(afterXml, afterInventory)
  const deleted = new Set(proposal.table.rows[proposal.rowIndex]!.map(c => c.id))
  if (before.paragraphs.length !== after.paragraphs.length || before.paragraphs.some((x, i) => { const y = after.paragraphs[i]; return !y || x.id !== y.id || (!deleted.has(x.id) && x.path !== y.path) || x.nesting !== y.nesting || y.text !== (deleted.has(x.id) ? "" : x.text) })) return fail("native preservation")
  assertNativeWordChatInlineDependencies(beforeXml, afterXml)
  const old = flat(beforeXml), now = flat(afterXml), oldRow = target(old.body, proposal).row, currentRow = tableRow(now.body, proposal).row
  const proof = inverse(oldRow, currentRow, true)
  currentRow.parentNode!.replaceChild(oldRow.cloneNode(true), currentRow)
  const normalized = flat(normalizeNativeWordChatCommentIds(beforeXml, new XMLSerializer().serializeToString(now.doc)))
  const dependencyParts = (doc: Document) => {
    const result = new Map<string, Element>()
    for (const part of elements(doc.documentElement)) {
      const name = part.getAttributeNS(PKG, "name")
      if (part.namespaceURI !== PKG || part.localName !== "part" || !name || result.has(name) || result.size >= 64) return fail("native preservation")
      result.set(name, part)
    }
    return result
  }
  const leftParts = dependencyParts(old.doc), rightParts = dependencyParts(normalized.doc)
  const extra = [...rightParts.keys()].filter(name => !leftParts.has(name))
  if ([...leftParts.keys()].some(name => !rightParts.has(name)) || extra.some(name => name !== "/word/people.xml") || extra.length > 1) return fail("native preservation")
  if (extra.length) {
    const p = rightParts.get(extra[0]!)!, ns = "http://schemas.microsoft.com/office/word/2012/wordml"
    const people = Array.from(p.getElementsByTagNameNS(ns, "people")), persons = Array.from(p.getElementsByTagNameNS(ns, "person"))
    const data = elements(p)
    if (data.length !== 1 || data[0]!.namespaceURI !== PKG || data[0]!.localName !== "xmlData" || elements(data[0]!).length !== 1 || people.length !== 1 || elements(data[0]!)[0] !== people[0] || elements(people[0]!).length !== 1 || persons.length !== 1 || persons[0]!.parentNode !== people[0] || !proof.authors.has(persons[0]!.getAttributeNS(ns, "author") ?? "") || Array.from(people[0]!.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === "http://schemas.openxmlformats.org/markup-compatibility/2006" && a.localName === "Ignorable"))) return fail("native preservation")
    const children = elements(persons[0]!)
    if (children.length !== 1 || children[0]!.namespaceURI !== ns || children[0]!.localName !== "presenceInfo" || elements(children[0]!).length || !children[0]!.getAttributeNS(ns, "providerId") || !children[0]!.getAttributeNS(ns, "userId") || Array.from(persons[0]!.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === ns && a.localName === "author")) || Array.from(children[0]!.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === ns && ["providerId", "userId"].includes(a.localName)))) return fail("native preservation")
  }
  for (const [name, left] of leftParts) {
    if (name === "/word/document.xml") continue
    const right = rightParts.get(name)!
    const clean = (part: Element) => {
      const clone = part.cloneNode(true) as Element
      if (extra.length && name === "/word/_rels/document.xml.rels" && part === right) {
        const ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        const links = Array.from(clone.getElementsByTagNameNS(ns, "Relationship")), additions = links.filter(e => e.getAttribute("Type") === "http://schemas.microsoft.com/office/2011/relationships/people" && e.getAttribute("Target") === "people.xml")
        if (additions.length !== 1 || !additions[0]!.getAttribute("Id") || Array.from(additions[0]!.attributes).some(a => a.namespaceURI !== XMLNS && !["Id", "Type", "Target"].includes(a.name)) || links.filter(e => e.getAttribute("Id") === additions[0]!.getAttribute("Id")).length !== 1 || elements(additions[0]!).length) return fail("native preservation")
        additions[0]!.parentNode!.removeChild(additions[0]!)
      }
      if (name === "/word/settings.xml") {
        const rsids = Array.from(clone.getElementsByTagNameNS(W, "rsids")); if (rsids.length > 1) return fail("native preservation")
        for (const list of rsids) {
          if (Array.from(list.attributes).some(a => a.namespaceURI !== XMLNS) || !only(list, ["rsid", "rsidRoot"]) || elements(list).some(e => elements(e).length || !/^[0-9a-fA-F]{8}$/.test(e.getAttributeNS(W, "val") ?? "") || Array.from(e.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === W && a.localName === "val")))) return fail("native preservation")
          list.parentNode!.removeChild(list)
        }
      }
      return canonical(clone)
    }
    if (name === "/word/_rels/document.xml.rels" && extra.length) {
      const relationship = "http://schemas.openxmlformats.org/package/2006/relationships", relAttr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      const entries = (part: Element) => Array.from(part.getElementsByTagNameNS(relationship, "Relationship")).filter(e => e.getAttribute("Type") !== "http://schemas.microsoft.com/office/2011/relationships/people")
      const allLeft = Array.from(left.getElementsByTagNameNS(relationship, "Relationship")), allRight = Array.from(right.getElementsByTagNameNS(relationship, "Relationship"))
      const peopleType = "http://schemas.microsoft.com/office/2011/relationships/people"
      const additions = allRight.filter(e => e.getAttribute("Type") === peopleType)
      const validIds = (items: Element[]) => items.every(e => Boolean(e.getAttribute("Id"))) && new Set(items.map(e => e.getAttribute("Id"))).size === items.length
      if (allLeft.some(e => e.getAttribute("Type") === peopleType) || additions.length !== 1 || additions[0]!.getAttribute("Target") !== "people.xml" || elements(additions[0]!).length || Array.from(additions[0]!.attributes).some(a => a.namespaceURI !== XMLNS && (Boolean(a.namespaceURI) || !["Id", "Type", "Target"].includes(a.name))) || !validIds(allLeft) || !validIds(allRight)) return fail("native preservation")
      const l = entries(left), r = entries(right), used = new Set<Element>()
      const shell = (part: Element) => { const clone = part.cloneNode(true) as Element; const rels = Array.from(clone.getElementsByTagNameNS(relationship, "Relationship")); for (const item of rels) item.parentNode!.removeChild(item); return canonical(clone) }
      if (l.length !== r.length || !eq(shell(left), shell(right))) return fail("native preservation")
      for (const x of l) {
        const sig = (e: Element) => { const c = e.cloneNode(true) as Element; c.removeAttribute("Id"); return canonical(c) }
        const matches = r.filter(y => !used.has(y) && eq(sig(x), sig(y)))
        if (matches.length !== 1) return fail("native preservation")
        const y = matches[0]!, oldId = x.getAttribute("Id"), newId = y.getAttribute("Id")
        if (!oldId || !newId || l.filter(e => e.getAttribute("Id") === oldId).length !== 1 || r.filter(e => e.getAttribute("Id") === newId).length !== 1) return fail("native preservation")
        if (oldId !== newId) {
          if (x.getAttribute("Type") !== "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") return fail("native preservation")
          const referenced = (doc: Document, id: string) => Array.from(doc.getElementsByTagName("*")).some(e => Array.from(e.attributes).some(a => a.namespaceURI === relAttr && a.value === id))
          if (referenced(old.doc, oldId) || referenced(normalized.doc, newId)) return fail("native preservation")
        }
        used.add(y)
      }
      continue
    }
    if (name === "/word/settings.xml" && left.getElementsByTagNameNS(W, "rsids").length !== right.getElementsByTagNameNS(W, "rsids").length || !eq(clean(left), clean(right))) throw Error(`dependency changed: ${name}`)
  }
  const trimTerminal = (body: Element, projection: typeof before) => {
    const children = elements(body), section = children.at(-1), tail = children.at(-2)
    if (section?.namespaceURI !== W || section.localName !== "sectPr" || tail?.namespaceURI !== W || tail.localName !== "p" || projection.paragraphs.some(p => p.path === `body/p[${children.length - 2}]`)) return
    if (Array.from(tail.childNodes).some(n => n.nodeType !== 3 || Boolean(n.nodeValue?.trim())) || Array.from(tail.attributes).some(a => a.namespaceURI !== XMLNS && !(a.namespaceURI === W && ["rsidR", "rsidRDefault"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)) && !(a.namespaceURI === W14 && ["paraId", "textId"].includes(a.localName) && /^[0-9a-fA-F]{8}$/.test(a.value)))) return fail("native preservation")
    body.removeChild(tail)
  }
  trimTerminal(old.body, before); trimTerminal(normalized.body, after)
  const oldRevs = revisionNodes(old.body), newRevs = revisionNodes(normalized.body), oldIds = oldRevs.map(x => x.node.getAttributeNS(W, "id")), newIds = newRevs.map(x => x.node.getAttributeNS(W, "id"))
  if (oldRevs.length !== newRevs.length || [...oldIds, ...newIds].some(id => !id || !/^\d+$/.test(id)) || new Set(oldIds.map(id => numericId(id!))).size !== oldIds.length || new Set(newIds.map(id => numericId(id!))).size !== newIds.length || [...proof.ids].some(id => newIds.some(existing => existing && numericId(existing) === id))) return fail("native preservation")
  for (let i = 0; i < oldRevs.length; i++) {
    const a = oldRevs[i]!, b = newRevs[i]!; if (a.path !== b.path || a.node.localName !== b.node.localName) return fail("native preservation")
    b.node.setAttributeNS(W, b.node.getAttributeNodeNS(W, "id")!.name, oldIds[i]!)
  }
  for (const body of [old.body, normalized.body]) { const section = elements(body).at(-1); if (section?.namespaceURI === W && section.localName === "sectPr") section.removeAttributeNS(W, "rsidSect") }
  if (!eq(canonical(old.body, true), canonical(normalized.body, true))) throw Error("body mismatch")
  return { authors: [...proof.authors] }
}
/** The reviewed SDK artifact is evidence only and is never imported into Word. */
export async function verifyReviewedWordChatRowDeletionOracle(source: Uint8Array, oracle: WordChatAuthoredTableRowDeletion, proposal: WordChatTableRowDeletionProposal): Promise<void> {
  const max = 16 * 1024 * 1024, aggregate = 64 * 1024 * 1024
  if (!(source instanceof Uint8Array) || !source.length || source.length > max || !oracle || !proposal || oracle.kind !== "table_row_delete" || proposal.kind !== "table_row_delete" || oracle.proposalId !== proposal.id || oracle.author !== "doc.haus" || !eq(oracle.table, proposal.table) || oracle.rowIndex !== proposal.rowIndex || !eq(oracle.cells, proposal.cells)) return fail("reviewed oracle")
  const hash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", source))).map(x => x.toString(16).padStart(2, "0")).join("")
  if (oracle.sourceSha256 !== hash || oracle.writerRequest?.source_sha256 !== hash || oracle.writerRequest?.author !== "doc.haus" || !oracle.writerRequest?.target_handle || !Array.isArray(oracle.revisions) || oracle.revisions.length !== 1 || !oracle.revisions[0]!.id || oracle.revisions[0]!.family !== "row_delete") return fail("reviewed oracle")
  const encoded = oracle.authoredBase64
  if (typeof encoded !== "string" || !encoded.length || encoded.length > Math.ceil(max / 3) * 4 || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(encoded)) return fail("reviewed oracle")
  let decoded: string; try { decoded = atob(encoded) } catch { return fail("reviewed oracle") }
  if (!decoded.length || decoded.length > max || btoa(decoded) !== encoded) return fail("reviewed oracle")
  const unpack = (bytes: Uint8Array) => {
    let size = 0, count = 0; const names = new Set<string>()
    try { return unzipSync(bytes, { filter(entry) { if (++count > 4096 || names.has(entry.name) || entry.originalSize > max || (size += entry.originalSize) > aggregate) throw Error("archive limit"); names.add(entry.name); return true } }) } catch { return fail("reviewed oracle") }
  }
  const a = unpack(source), b = unpack(Uint8Array.from(decoded, c => c.charCodeAt(0))), keys = Object.keys(a).sort()
  if (!eq(keys, Object.keys(b).sort()) || !a["word/document.xml"] || !a["word/settings.xml"]) return fail("reviewed oracle")
  for (const key of keys) if (key !== "word/document.xml" && key !== "word/settings.xml" && (a[key]!.length !== b[key]!.length || a[key]!.some((v, i) => v !== b[key]![i]))) return fail("reviewed oracle")
  const document = (bytes: Uint8Array) => parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes))
  const sa = document(a["word/settings.xml"]!), sb = document(b["word/settings.xml"]!)
  for (const [doc, required] of [[sa, false], [sb, true]] as const) {
    const marks = Array.from(doc.getElementsByTagNameNS(W, "trackRevisions"))
    if (marks.length > 1 || required && marks.length !== 1 || marks.some(m => m.parentNode !== doc.documentElement || m.childNodes.length || Array.from(m.attributes).some(x => x.namespaceURI !== XMLNS && !(x.namespaceURI === W && x.localName === "val" && ["1", "true", "on"].includes(x.value))))) return fail("reviewed oracle")
    for (const m of marks) m.parentNode!.removeChild(m)
  }
  if (!eq(canonical(sa.documentElement), canonical(sb.documentElement))) return fail("reviewed oracle")
  const da = document(a["word/document.xml"]!), db = document(b["word/document.xml"]!), ab = Array.from(da.getElementsByTagNameNS(W, "body")), bb = Array.from(db.getElementsByTagNameNS(W, "body"))
  if (ab.length !== 1 || bb.length !== 1) return fail("reviewed oracle")
  const old = target(ab[0]!, proposal).row, actual = tableRow(bb[0]!, proposal).row, proof = inverse(old, actual, false, "doc.haus")
  const oldIds = new Set(Array.from(da.getElementsByTagNameNS(W, "*")).filter(e => ["ins", "del", "pPrChange", "rPrChange"].includes(e.localName)).map(e => numericId(e.getAttributeNS(W, "id") ?? "")))
  if ([...proof.ids].some(id => oldIds.has(id))) return fail("reviewed oracle")
  actual.parentNode!.replaceChild(old.cloneNode(true), actual)
  if (!eq(canonical(da.documentElement), canonical(db.documentElement))) return fail("reviewed oracle")
}
