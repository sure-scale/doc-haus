import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"

// One requested output file for a memo and its instrument: a derived, read-only
// Word file with the memo first, then a section break, then the final body. The
// working body .docx and the notes .docx stay separate; paragraph binding, the
// redline, amendments and the Word round trip stay on them, and this file is
// rebuilt from them after each change and never split back. Paragraph pointers
// are kept, except that a memo pointer the body already uses gets a fresh one.
//
// The memo is the registered notes .docx itself, spliced in, not re-rendered
// from markdown. The notes may carry accepted Word changes and addenda that
// exist only in that file, and the Docxodus projection is lossy as a markdown
// source (block quotes lose their style; a nested bullet under a numbered item
// comes back as a plain "o ..." paragraph). The body package stays as it was:
// its styles, numbering, headers, footers and other parts.
//
// A memo style is reused only when the body defines it identically; otherwise
// it is copied under a new id, so memo headings never pick up a template's
// heading numbering or look. The memo's default styles map to the body's.
// Memo lists get their own numbering definitions, so neither list continues
// the other. The memo becomes its own first section with the body's page setup
// and no header or footer; the body's first section starts on a new page and
// restarts page numbering at 1. Output is deterministic: no random ids, fixed
// zip entry times.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const R = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const BODY = "word/document.xml"
const RELS = "word/_rels/document.xml.rels"
const TYPES = "[Content_Types].xml"
const STYLES = { rel: `${O}/styles`, file: "word/styles.xml", root: "styles",
  type: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" }
const NUMBERING = { rel: `${O}/numbering`, file: "word/numbering.xml", root: "numbering",
  type: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" }

/** Memo first, a section break, then the body, as new .docx bytes. `memo` is
 * the registered notes .docx; `body` is the registered body .docx. Pure: no
 * file or state reads, and neither input changes. */
export function combineMemoAndBody(input: { memo: Uint8Array; body: Uint8Array }) {
  const parts = unzipSync(input.body)
  const source = unzipSync(input.memo)
  if (!parts[BODY] || !parts[TYPES] || !source[BODY]) throw new Error("invalid_docx_package")
  const document = parse(parts[BODY])
  const target = only(document, "body")
  const relationships = parse(parts[RELS] ?? strToU8(`<Relationships xmlns="${R}"/>`))
  const contentTypes = parse(parts[TYPES])
  const memoRelationships = source[RELS] ? parse(source[RELS]) : undefined
  const changed = new Map<string, Document>([[BODY, document]])

  const memoDocument = parse(source[BODY])
  const imported = children(only(memoDocument, "body")).filter(node => node.localName !== "sectPr")
    .map(node => document.importNode(node, true) as Element)
  if (!imported.some(block => (block.textContent ?? "").trim())) throw new Error("empty_memo")
  const scanned = imported.flatMap(block => [block, ...elements(block, W, "*")])
  // Note text lives in parts this file does not carry; dropping it would lose
  // memo content, so such a memo is refused.
  if (scanned.some(node => ["footnoteReference", "endnoteReference"].includes(node.localName))) throw new Error("unsupported_memo_dependency")
  // Memo bookmarks (its pointers included) carry over under ids no body part
  // uses: bookmark ids and names are package-wide, headers, footers and notes
  // included. Each end closes the latest open start of its id; an end or a
  // start left without its partner goes. Word needs unique names (case aside):
  // a memo bookmark whose name any body part already uses is renamed, a
  // paragraph pointer to another well-formed pointer, and the memo's
  // references follow it below (Word's own _GoBack is just dropped); a name
  // the memo repeats keeps its first mark. Review comments on the notes are not
  // memo content and their markers go. Word paragraph ids go so they stay unique.
  const dropped = new Set<Node>()
  const drop = (node: Node) => {
    dropped.add(node)
    node.parentNode?.removeChild(node)
  }
  const marks = [...elements(document, W, "bookmarkStart"), ...elements(document, W, "bookmarkEnd")]
  const lower = (rows: Element[]) => rows.filter(row => row.localName === "bookmarkStart").map(row => (row.getAttributeNS(W, "name") ?? "").toLowerCase())
  const others = Object.entries(parts).filter(([name]) => name.startsWith("word/") && name.endsWith(".xml") && name !== BODY).map(([, bytes]) => strFromU8(bytes))
  const names = new Set([...lower(marks), ...others.flatMap(xml => [...xml.matchAll(/<(?:[\w.-]+:)?bookmarkStart\b[^>]*?\s(?:[\w.-]+:)?name\s*=\s*(["'])(.*?)\1/g)]
    .map(match => decodeReferences(match[2]!).toLowerCase()))])
  // Names any bookmark field, link or expression in either file refers to, even
  // when nothing defines them: a memo bookmark never takes one (a body
  // reference would land in the memo) and a minted name never equals one (a
  // reference would land on the renamed bookmark).
  const bodyReferences = new Set<string>()
  const memoReferences = new Set<string>()
  const taken = new Set([...names, ...lower(scanned)])
  // A name no bookmark in either file uses, within Word's 40 characters. A
  // paragraph pointer (_dh_ and a pid, as minted in drafting-amendment.ts)
  // takes the next pid no bookmark uses.
  const rename = (name: string) => {
    const pointer = /^_dh_p[0-9a-f]{12}$/.test(name)
    let fresh = name
    for (let n = 1; taken.has(fresh.toLowerCase()); n++) fresh = pointer
      ? `_dh_p${((Number.parseInt(name.slice(5), 16) + n) % 2 ** 48).toString(16).padStart(12, "0")}`
      : `${cut(name, 39 - String(n).length)}_${n}`
    taken.add(fresh.toLowerCase())
    return fresh
  }
  readNames(elements(document, W, "*"), bodyReferences)
  for (const xml of others) if (/instr|anchor|fldChar/.test(xml)) {
    try { readNames(elements(parse(strToU8(xml)), W, "*"), bodyReferences) } catch { /* a part that is not well-formed has no readable references */ }
  }
  readNames(scanned, memoReferences)
  for (const name of [...bodyReferences, ...memoReferences]) taken.add(name)
  const renamed = new Map<string, string | undefined>() // name (lower case) -> the name memo references to it take, if another
  const ends = new Map<Element, Element>() // memo start -> the end that closes it
  const open = new Map<string, Element[]>() // memo id -> its open starts
  for (const node of scanned.filter(row => row.namespaceURI === W && ["bookmarkStart", "bookmarkEnd"].includes(row.localName))) {
    const id = node.getAttributeNS(W, "id") ?? ""
    const start = node.localName === "bookmarkEnd" ? open.get(id)?.pop() : undefined
    if (node.localName === "bookmarkStart") open.set(id, [...open.get(id) ?? [], node])
    else if (start) ends.set(start, node)
    else drop(node)
  }
  const used = others.flatMap(xml => [...xml.matchAll(/<(?:[\w.-]+:)?bookmark(?:Start|End)\b[^>]*?\s(?:[\w.-]+:)?id\s*=\s*["'](-?\d+)["']/g)].map(match => Number(match[1])))
  let nextMark = [...marks.map(mark => Number(mark.getAttributeNS(W, "id")) || 0), ...used].reduce((max, id) => Math.max(max, id), -1) + 1
  for (const node of scanned) {
    for (const name of ["paraId", "textId"]) node.removeAttributeNS(W14, name)
    if (node.namespaceURI !== W) continue
    if (node.localName === "bookmarkStart") {
      const name = node.getAttributeNS(W, "name") ?? ""
      const key = name.toLowerCase()
      const end = ends.get(node)
      if (!end || renamed.has(key) || names.has(key) && key === "_goback") {
        for (const row of [node, end]) if (row) drop(row)
        continue
      }
      const fresh = names.has(key) || bodyReferences.has(key) ? rename(name) : name
      renamed.set(key, fresh === name ? undefined : fresh)
      if (fresh !== name) setW(node, "name", fresh)
      const mark = String(nextMark++)
      for (const row of [node, end]) setW(row, "id", mark)
    } else if (["commentRangeStart", "commentRangeEnd"].includes(node.localName)) drop(node)
    else if (node.localName === "commentReference") {
      const run = node.parentNode as Element | null
      drop(run?.localName === "r" && children(run).every(row => ["rPr", "commentReference"].includes(row.localName)) ? run : node)
    }
  }
  // Memo cross-references and internal links to a renamed bookmark follow it:
  // the bookmark arguments and expression names `targets` finds, and
  // HYPERLINK \l or a hyperlink anchor unless the link names a file or
  // address, whose location is in that other file. A reference to a name the
  // memo does not keep (its mark went, or it never had one) that the body uses
  // follows a fresh name no bookmark has, so it never lands in the body.
  const follow = (name: string, strict = false) => {
    const key = name.toLowerCase()
    if (!strict && !renamed.has(key) && names.has(key)) renamed.set(key, rename(name))
    return renamed.get(key) ?? name
  }
  // The edits that make an instruction name the renamed bookmarks, if any.
  const edits = (code: string) => targets(code).flatMap(target => {
    const fresh = follow(target.name, !target.loose)
    return fresh !== target.name ? [{ ...target, fresh }] : []
  })
  const rewrite = (code: string) => edits(code).reduceRight((text, edit) => text.slice(0, edit.start) + edit.fresh + text.slice(edit.end), code)
  const retarget = (node: Element, attribute: string, change: (text: string) => string) => {
    const text = node.getAttributeNS(W, attribute)
    const next = text ? change(text) : text
    if (next && next !== text) setW(node, attribute, next)
  }
  // A complex field's instruction keeps its runs: a new name goes in the run
  // where the old one starts, and other text stays where it was.
  const settle = (fieldParts: { node?: Element; text: string }[]) => {
    const found = edits(fieldParts.map(part => part.text).join(""))
    if (!found.length) return
    let offset = 0
    for (const part of fieldParts) {
      let text = ""
      let at = 0
      for (const edit of found) {
        const [from, to] = [edit.start - offset, edit.end - offset]
        if (from >= part.text.length || to <= 0) continue
        text += part.text.slice(at, Math.max(0, from)) + (from >= 0 ? edit.fresh : "")
        at = Math.min(part.text.length, to)
      }
      if (part.node && at) part.node.textContent = text + part.text.slice(at)
      offset += part.text.length
    }
  }
  walkFields(scanned, {
    code: settle,
    simple: node => retarget(node, "instr", rewrite),
    link: node => { if (!node.hasAttributeNS(O, "id")) retarget(node, "anchor", name => follow(name)) },
  })
  const blocks = imported.filter(block => !dropped.has(block))
  const nodes = blocks.flatMap(block => [block, ...elements(block, W, "*")])

  // Styles: reuse identical definitions, copy the rest under fresh ids.
  const memoStyles = related(source, memoRelationships, STYLES.rel)
  const styles = memoStyles ? ensurePart(parts, relationships, contentTypes, STYLES) : undefined
  const copied: Element[] = []
  if (memoStyles && styles) {
    const theirs = styleIndex(styles.doc)
    const mine = styleIndex(memoStyles)
    const mapped = new Map<string, string | undefined>()
    // Copies take ids and names the body does not use, compared as Word does,
    // without regard to case.
    const inUse = { ids: new Set<string>(), names: new Set<string>() }
    for (const style of elements(styles.doc, W, "style")) {
      inUse.ids.add((style.getAttributeNS(W, "styleId") ?? "").toLowerCase())
      inUse.names.add(value(child(style, "name")).toLowerCase())
    }
    const reusable = (id: string) => {
      const style = mine.byId.get(id)!
      const existing = theirs.byId.get(id)
      if (!existing || canonical(existing) !== canonical(style) || elements(style, W, "numPr").length) return false
      const basedOn = value(child(style, "basedOn"))
      return !basedOn || mapStyle(basedOn) === basedOn
    }
    const mapStyle = (id: string): string | undefined => {
      if (mapped.has(id)) return mapped.get(id)
      const style = mine.byId.get(id)
      const type = style?.getAttributeNS(W, "type") || "paragraph"
      if (!style || mine.defaults.get(type) === id) {
        mapped.set(id, style ? theirs.defaults.get(type) : undefined)
        return mapped.get(id)
      }
      // A linked paragraph and character style pair is reused or copied as a
      // pair, so neither the copies nor the body's pair link across.
      const link = value(child(style, "link"))
      const pair = mine.byId.has(link) && !mapped.has(link) && ![...mine.defaults.values()].includes(link) ? [id, link] : [id]
      for (const member of pair) mapped.set(member, undefined) // a basedOn cycle resolves to no base
      if (pair.every(reusable) && (pair.length > 1 || !mine.byId.has(link) || mapStyle(link) === link)) {
        for (const member of pair) mapped.set(member, member)
        return id
      }
      const copies = pair.map(member => {
        const copy = styles.doc.importNode(mine.byId.get(member)!, true) as Element
        const name = child(copy, "name")
        const tag = (suffix: string): [string, string] => [`Memo${member}${suffix}`, `Memo ${value(name) || member}${suffix && ` ${suffix}`}`]
        let [fresh, label] = tag("")
        for (let suffix = 2; inUse.ids.has(fresh.toLowerCase()) || inUse.names.has(label.toLowerCase()); suffix++) [fresh, label] = tag(String(suffix))
        inUse.ids.add(fresh.toLowerCase())
        inUse.names.add(label.toLowerCase())
        mapped.set(member, fresh)
        theirs.byId.set(fresh, copy)
        setW(copy, "styleId", fresh)
        setW(copy, "customStyle", "1")
        copy.removeAttributeNS(W, "default")
        if (name) setW(name, "val", label)
        return copy
      })
      for (const copy of copies) {
        for (const kind of ["basedOn", "next", "link"]) {
          const reference = child(copy, kind)
          if (reference) rebind(reference, mapStyle(value(reference)))
        }
        styles.doc.documentElement.appendChild(copy)
        copied.push(copy)
      }
      return mapped.get(id)
    }
    for (const reference of nodes.filter(node => ["pStyle", "rStyle", "tblStyle"].includes(node.localName)))
      rebind(reference, mapStyle(value(reference)))
    if (copied.length) adoptNamespaces(styles.doc.documentElement, memoStyles.documentElement, copied)
    if (copied.length || !parts[styles.name]) changed.set(styles.name, styles.doc)
  } else {
    for (const reference of nodes.filter(node => ["pStyle", "rStyle", "tblStyle"].includes(node.localName))) rebind(reference)
  }

  // Numbering: every memo list gets fresh num and abstractNum definitions.
  const listed = [...nodes, ...copied.flatMap(style => elements(style, W, "numId"))].filter(node => node.localName === "numId")
  const memoNumbering = listed.some(node => value(node) !== "0") ? related(source, memoRelationships, NUMBERING.rel) : undefined
  const numMap = new Map<string, string>()
  if (memoNumbering) {
    const numbering = ensurePart(parts, relationships, contentTypes, NUMBERING)
    const root = numbering.doc.documentElement
    const ids = (name: string, attribute: string) => children(root).filter(node => node.localName === name).map(node => Number(node.getAttributeNS(W, attribute)) || 0)
    const next = { num: Math.max(0, ...ids("num", "numId")) + 1, abstract: Math.max(-1, ...ids("abstractNum", "abstractNumId")) + 1 }
    const nsids = new Set(elements(numbering.doc, W, "nsid").map(node => value(node).toUpperCase()))
    const abstractMap = new Map<string, string>()
    const lists: Element[] = []
    const before = (...names: string[]) => children(root).find(node => names.includes(node.localName)) ?? null
    for (const id of new Set(listed.map(value))) {
      const num = elements(memoNumbering, W, "num").find(node => node.getAttributeNS(W, "numId") === id)
      const abstractId = value(num && child(num, "abstractNumId"))
      const abstract = elements(memoNumbering, W, "abstractNum").find(node => node.getAttributeNS(W, "abstractNumId") === abstractId)
      if (!num || !abstract) continue
      if (!abstractMap.has(abstractId)) {
        const copy = numbering.doc.importNode(abstract, true) as Element
        if (elements(copy, W, "numStyleLink").length) throw new Error("unsupported_memo_numbering")
        unlinkNumbering(copy)
        setW(copy, "abstractNumId", String(next.abstract))
        abstractMap.set(abstractId, String(next.abstract++))
        const nsid = child(copy, "nsid")
        if (nsid) {
          let code = Number.parseInt(value(nsid), 16) >>> 0
          while (nsids.has(hex(code))) code = (code + 1) >>> 0
          nsids.add(hex(code))
          setW(nsid, "val", hex(code))
        }
        lists.push(root.insertBefore(copy, before("num", "numIdMacAtCleanup")))
      }
      const copy = numbering.doc.importNode(num, true) as Element
      unlinkNumbering(copy)
      setW(copy, "numId", String(next.num))
      setW(child(copy, "abstractNumId")!, "val", abstractMap.get(abstractId)!)
      lists.push(root.insertBefore(copy, before("numIdMacAtCleanup")))
      numMap.set(id, String(next.num++))
    }
    if (lists.length) adoptNamespaces(root, memoNumbering.documentElement, lists)
    changed.set(numbering.name, numbering.doc)
  }
  for (const node of listed) {
    if (value(node) === "0") continue
    const mapped = numMap.get(value(node))
    const numPr = node.parentNode
    if (mapped) setW(node, "val", mapped)
    else if (numPr?.parentNode) numPr.parentNode.removeChild(numPr)
  }

  // The memo carries no header or footer: each of its own section breaks keeps
  // only its page setup, as its last section does below, and also its type and
  // columns, so a continuous or multi-column part keeps its layout.
  for (const section of blocks.flatMap(block => elements(block, W, "sectPr"))) pageSetupOnly(section, ["type", "cols"])
  adoptNamespaces(document.documentElement, memoDocument.documentElement, blocks)

  // External hyperlinks carry over. Drawings (images, charts, shapes, whose
  // ids the body may already use) and embedded objects are not part of the
  // notes render and are refused, as is any other element, in any namespace,
  // that points at a relationship.
  const added = new Map<string, string>()
  for (const node of blocks.flatMap(block => [block, ...elements(block, "*", "*")])) {
    if (node.namespaceURI === W && ["drawing", "pict", "object"].includes(node.localName)) throw new Error("unsupported_memo_dependency")
    for (const attribute of Array.from(node.attributes).filter(attribute => attribute.namespaceURI === O)) {
      const relationship = memoRelationships && elements(memoRelationships, R, "Relationship").find(row => row.getAttribute("Id") === attribute.value)
      if (node.namespaceURI !== W || node.localName !== "hyperlink" || relationship?.getAttribute("Type") !== `${O}/hyperlink`
        || relationship.getAttribute("TargetMode") !== "External") throw new Error("unsupported_memo_dependency")
      const key = `${relationship.getAttribute("Type")}\n${relationship.getAttribute("Target")}`
      if (!added.has(key)) added.set(key, addRelationship(relationships, relationship.getAttribute("Type")!, relationship.getAttribute("Target")!, "External"))
      node.setAttributeNS(O, attribute.name, added.get(key)!)
    }
  }

  const memoSection = startBodySection(target)
  const first = children(target)[0] ?? null
  for (const block of blocks) target.insertBefore(block, first)
  const last = blocks.at(-1)
  const holder = last?.localName === "p" && !elements(last, W, "sectPr").length
    ? last : target.insertBefore(document.createElementNS(W, "w:p"), first)
  const properties = child(holder, "pPr") ?? holder.insertBefore(document.createElementNS(W, "w:pPr"), holder.firstChild)
  properties.insertBefore(memoSection, child(properties, "pPrChange") ?? null)

  changed.set(RELS, relationships).set(TYPES, contentTypes)
  const serializer = new XMLSerializer()
  for (const [name, doc] of changed) parts[name] = strToU8(serializer.serializeToString(doc))
  return zipSync(parts, { mtime: new Date(1980, 0, 1) })
}

// The body's first section starts on a new page and numbers its pages from 1.
// Returns the memo's section properties: the same page setup (see
// pageSetupOnly), which as the first section it cannot inherit, as an explicit
// next-page break, so the body also starts a new page for a reader that takes a
// break's type from the section it ends. The body's own type is left as it is
// unless it would keep the body on the memo's page.
function startBodySection(target: Element) {
  const document = target.ownerDocument!
  // A section break ends a top-level paragraph, also one inside content
  // controls or custom XML blocks; the body's own sectPr ends the last one.
  const top = (node: Node | null): boolean => node === target
    || ["sdt", "sdtContent", "customXml"].includes((node as Element | null)?.localName ?? "") && top(node!.parentNode)
  const section = elements(target, W, "sectPr").find(node => node.parentNode === target
    || (node.parentNode as Element).localName === "pPr" && top(node.parentNode?.parentNode?.parentNode ?? null))
    ?? target.appendChild(document.createElementNS(W, "w:sectPr"))
  const memo = pageSetupOnly(section.cloneNode(true) as Element)
  const start = document.createElementNS(W, "w:type")
  start.setAttributeNS(W, "w:val", "nextPage")
  memo.insertBefore(start, memo.firstChild)
  const type = child(section, "type")
  if (type && ["continuous", "nextColumn"].includes(value(type))) setW(type, "val", "nextPage")
  const numbering = child(section, "pgNumType") ?? section.insertBefore(document.createElementNS(W, "w:pgNumType"),
    children(section).find(row => ["cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"].includes(row.localName)) ?? null)
  if (!numbering.hasAttributeNS(W, "start")) numbering.setAttributeNS(W, "w:start", "1")
  return memo
}

// A memo section keeps only its page size, margins, paper source, text
// direction and grid (in schema order), plus any `layout` children named: no
// header, footer, printer settings or other layout. Changes the section in place.
function pageSetupOnly(section: Element, layout: string[] = []) {
  const kept = ["pgSz", "pgMar", "paperSrc", "textDirection", "docGrid", ...layout]
  for (const row of children(section).filter(row => row.namespaceURI !== W || !kept.includes(row.localName)))
    section.removeChild(row)
  return section
}

// A copied list level must not bind back to a body paragraph style (that would
// number the body's paragraphs) or to a picture bullet the body does not own.
function unlinkNumbering(node: Element) {
  for (const link of elements(node, W, "*").filter(row => ["styleLink", "pStyle", "lvlPicBulletId"].includes(row.localName)))
    link.parentNode?.removeChild(link)
}

function styleIndex(doc: Document) {
  const byId = new Map<string, Element>()
  const defaults = new Map<string, string>()
  for (const style of elements(doc, W, "style")) {
    const id = style.getAttributeNS(W, "styleId")
    if (!id) continue
    byId.set(id, style)
    if (["1", "true", "on"].includes(style.getAttributeNS(W, "default") ?? "")) defaults.set(style.getAttributeNS(W, "type") || "paragraph", id)
  }
  return { byId, defaults }
}

function canonical(node: Element) {
  return new XMLSerializer().serializeToString(node).replace(/>\s+</g, "><")
}

function rebind(reference: Element, id?: string) {
  if (id) setW(reference, "val", id)
  else reference.parentNode?.removeChild(reference)
}

/** Numeric and the five named character references, as an XML parser reads them. */
function decodeReferences(text: string) {
  const named: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" }
  return text.replace(/&(?:#(\d+)|#[xX]([0-9a-fA-F]+)|(amp|lt|gt|quot|apos));/g, (match, decimal, hex, word) => {
    if (word) return named[word]!
    try { return String.fromCodePoint(Number.parseInt(decimal ?? hex, decimal ? 10 : 16)) } catch { return match }
  })
}

/** At most `max` UTF-16 units of `text`, never ending inside a surrogate pair. */
function cut(text: string, max: number) {
  let out = ""
  for (const character of text) {
    if (out.length + character.length > max) break
    out += character
  }
  return out
}

// The keywords are Word's field types ([MS-DOC] flt and WdFieldType).
const keywords = new Set(`ADDIN ADDRESSBLOCK ADVANCE ASK AUTHOR AUTONUM AUTONUMLGL AUTONUMOUT AUTOTEXT AUTOTEXTLIST
  BARCODE BIBLIOGRAPHY BIDIOUTLINE CITATION COMMENTS COMPARE CONTROL CREATEDATE DATA DATABASE DATE DDE DDEAUTO
  DISPLAYBARCODE DOCPROPERTY DOCVARIABLE EDITTIME EMBED EQ FILENAME FILESIZE FILLIN FORMCHECKBOX FORMDROPDOWN FORMTEXT
  FTNREF GLOSSARY GOTOBUTTON GREETINGLINE HTMLACTIVEX HTMLCONTROL HYPERLINK IF IMPORT INCLUDE INCLUDEPICTURE INCLUDETEXT
  INDEX INFO KEYWORDS LASTSAVEDBY LINK LISTNUM MACROBUTTON MERGEBARCODE MERGEFIELD MERGEREC MERGESEQ NEXT NEXTIF NOTEREF
  NUMCHARS NUMPAGES NUMWORDS OCX PAGE PAGEREF PRINT PRINTDATE PRIVATE QUOTE RD REF REVNUM SAVEDATE SECTION SECTIONPAGES
  SEQ SET SHAPE SKIPIF STYLEREF SUBJECT SUBSCRIBER SYMBOL TA TC TEMPLATE TIME TITLE TOA TOC USERADDRESS USERINITIALS
  USERNAME XE`.split(/\s+/))

type Target = { start: number; end: number; name: string; loose: boolean }

/** Where in a field instruction the bookmark names it reads sit. A "loose"
 * target is a field's own bookmark argument (REF, PAGEREF, NOTEREF, FTNREF,
 * GOTOBUTTON, TOC \b, SET, ASK, a bare name, HYPERLINK \l); the others are
 * identifiers in an `=` or IF expression, which may be anything else, so they
 * count only when a memo bookmark of that name is renamed. A field's
 * instruction can be split over runs and around a nested field; a nested field
 * reads as one opaque word, so a name built with one never matches. A quoted
 * word may hold \". One linear pass. */
function targets(code: string): Target[] {
  const words = [...code.matchAll(/"(?:\\[\s\S]|[^"\\])*"?|\\[^\s"\\]*|[^\s"\\]+/g)]
  const kind = words[0]?.[0].toUpperCase() ?? ""
  const named = (word: RegExpMatchArray | undefined, loose = true): Target[] => {
    if (!word) return []
    const skip = word[0].startsWith('"') ? 1 : 0
    const name = word[0].slice(skip).replace(/"$/, "")
    const start = word.index! + skip
    return name ? [{ start, end: start + name.length, name, loose }] : []
  }
  if (kind.startsWith("=") || kind === "IF") {
    const found: Target[] = []
    const scan = /"(?:\\[\s\S]|[^"\\])*"?|\\[*#@]\s*(?:"(?:\\[\s\S]|[^"\\])*"?|[^\s"\\]*)|\\[^\s"\\]*|[A-Za-z0-9_]+/g
    scan.lastIndex = words[0]!.index! + (kind === "IF" ? 2 : 1)
    for (let match = scan.exec(code); match; match = scan.exec(code)) {
      const text = match[0]
      if (!/^[A-Za-z_]/.test(text) || code[match.index - 1] === ".") continue
      const call = /\s*\(/y
      call.lastIndex = match.index + text.length
      if (!call.test(code)) found.push({ start: match.index, end: match.index + text.length, name: text, loose: false })
    }
    return found
  }
  const hyperlink = kind === "HYPERLINK"
  const args: RegExpMatchArray[] = []
  let anchor: RegExpMatchArray | undefined
  let bookmark: RegExpMatchArray | undefined
  for (let index = 1; index < words.length; index++) {
    const word = words[index]![0]
    if (!word.startsWith("\\")) args.push(words[index]!)
    else if (kind === "TOC" && /^\\b$/i.test(word)) bookmark = words[++index]
    else if ((hyperlink ? /^\\[*#@lot]$/i : /^\\[*#@d]$/i).test(word)) {
      index++ // a switch that takes the next word as its argument
      if (hyperlink && /^\\l$/i.test(word)) anchor = words[index]
    }
  }
  return named(hyperlink ? args.length ? undefined : anchor
    : /^(?:PAGE|NOTE|FTN)?REF$|^(?:GOTOBUTTON|SET|ASK)$/.test(kind) ? args[0]
    : kind === "TOC" ? bookmark
    : keywords.has(kind) || /^["\\]/.test(kind) ? undefined : words[0])
}

type FieldPart = { node?: Element; text: string }

/** Visits, in document order, each complex field's instruction (its text
 * pieces, a nested field one opaque piece), each simple field and each
 * hyperlink. */
function walkFields(nodes: Element[], visit: { code: (parts: FieldPart[]) => void; simple: (node: Element) => void; link: (node: Element) => void }) {
  const fields: (FieldPart[] | undefined)[] = [] // open complex fields, innermost last; undefined after the instruction
  for (const node of nodes) {
    if (node.namespaceURI !== W) continue
    const field = fields.at(-1)
    if (node.localName === "instrText") field?.push({ node, text: node.textContent ?? "" })
    else if (node.localName === "fldChar") {
      const type = node.getAttributeNS(W, "fldCharType")
      if (type === "begin") {
        field?.push({ text: "\u0001" })
        fields.push([])
      } else if (field && ["separate", "end"].includes(type ?? "")) {
        fields[fields.length - 1] = undefined
        visit.code(field)
      }
      if (type === "end") fields.pop()
    } else if (node.localName === "fldSimple") visit.simple(node)
    else if (node.localName === "hyperlink") visit.link(node)
  }
}

/** Adds (lower case) every bookmark name the fields and links among `nodes` read. */
function readNames(nodes: Element[], into: Set<string>) {
  const add = (found: Target[]) => found.forEach(target => into.add(target.name.toLowerCase()))
  walkFields(nodes, {
    code: parts => add(targets(parts.map(part => part.text).join(""))),
    simple: node => add(targets(node.getAttributeNS(W, "instr") ?? "")),
    link: node => { const anchor = node.getAttributeNS(W, "anchor"); if (anchor) into.add(anchor.toLowerCase()) },
  })
}

// Sets a WordprocessingML attribute through the one the element already has.
// xmldom's setAttributeNS matches by qualified name, so a part that binds the
// namespace to a prefix other than w would get a second attribute of the same
// name, which Word and Docxodus reject. A new one takes the element's prefix.
function setW(node: Element, name: string, text: string) {
  const attribute = node.getAttributeNodeNS(W, name)
  if (attribute) attribute.value = attribute.nodeValue = text
  else node.setAttributeNS(W, `${node.namespaceURI === W && node.prefix || "w"}:${name}`, text)
}

function related(parts: Record<string, Uint8Array>, relationships: Document | undefined, type: string) {
  const relationship = relationships && elements(relationships, R, "Relationship").find(row => row.getAttribute("Type") === type)
  const name = relationship && partName(relationship.getAttribute("Target")!)
  return name && parts[name] ? parse(parts[name]) : undefined
}

function ensurePart(parts: Record<string, Uint8Array>, relationships: Document, contentTypes: Document, kind: typeof STYLES) {
  const relationship = elements(relationships, R, "Relationship").find(row => row.getAttribute("Type") === kind.rel)
  const name = relationship ? partName(relationship.getAttribute("Target")!) : kind.file
  if (!relationship) addRelationship(relationships, kind.rel, path.posix.relative("word", name))
  if (!elements(contentTypes, CT, "Override").some(row => row.getAttribute("PartName") === `/${name}`)) {
    const override = contentTypes.createElementNS(CT, "Override")
    override.setAttribute("PartName", `/${name}`)
    override.setAttribute("ContentType", kind.type)
    contentTypes.documentElement.appendChild(override)
  }
  return { name, doc: parse(parts[name] ?? strToU8(`<w:${kind.root} xmlns:w="${W}"/>`)) }
}

function addRelationship(relationships: Document, type: string, target: string, mode?: string) {
  const used = new Set(elements(relationships, R, "Relationship").map(row => row.getAttribute("Id")))
  let index = 1
  while (used.has(`rId${index}`)) index++
  const relationship = relationships.createElementNS(R, "Relationship")
  relationship.setAttribute("Id", `rId${index}`)
  relationship.setAttribute("Type", type)
  relationship.setAttribute("Target", target)
  if (mode) relationship.setAttribute("TargetMode", mode)
  relationships.documentElement.appendChild(relationship)
  return `rId${index}`
}

// The memo part root's namespace declarations and ignorable prefixes move onto
// the body part root that receives `content`, memo elements imported under it.
// Markup-compatibility attributes (Ignorable, Requires, ProcessContent and the
// rest) name prefixes, which the serializer would otherwise declare only on the
// elements that use them. A prefix the body binds to another namespace is never
// rebound: the memo is refused if its content uses the prefix, and otherwise the
// memo's declaration and ignorable entry for it are left out.
function adoptNamespaces(root: Element, memo: Element, content: Element[]) {
  const declared = (node: Element) => new Map(Array.from(node.attributes)
    .filter(row => row.namespaceURI === XMLNS && row.prefix === "xmlns").map(row => [row.localName, row.value]))
  const theirs = declared(root)
  const mine = declared(memo)
  const used = new Set<string>()
  for (const node of content.flatMap(block => [block, ...elements(block, "*", "*")])) {
    used.add(node.prefix ?? "")
    for (const attribute of Array.from(node.attributes)) {
      if (attribute.namespaceURI !== XMLNS) used.add(attribute.prefix ?? "")
      if (attribute.namespaceURI === MC || node.namespaceURI === MC && node.localName === "Choice" && attribute.name === "Requires")
        for (const name of attribute.value.split(/\s+/)) used.add(name.split(":")[0]!)
    }
  }
  for (const [prefix, uri] of mine) {
    if (!theirs.has(prefix)) root.setAttributeNS(XMLNS, `xmlns:${prefix}`, uri)
    else if (theirs.get(prefix) !== uri && used.has(prefix)) throw new Error("unsupported_memo_namespace")
  }
  const bound = declared(root)
  const ignorable = (node: Element) => (node.getAttributeNS(MC, "Ignorable") ?? "").split(/\s+/).filter(Boolean)
  const current = ignorable(root)
  const extra = [...new Set(ignorable(memo))].filter(prefix => mine.has(prefix) && bound.get(prefix) === mine.get(prefix) && !current.includes(prefix))
  if (!extra.length) return
  const prefix = [...bound].find(([, uri]) => uri === MC)?.[0]
  if (!prefix) throw new Error("unsupported_memo_namespace")
  root.setAttributeNS(MC, root.getAttributeNodeNS(MC, "Ignorable")?.name ?? `${prefix}:Ignorable`, [...current, ...extra].join(" "))
}

function partName(target: string) {
  return path.posix.resolve("/word", target).slice(1)
}

function parse(bytes: Uint8Array) {
  const fail = (message: string) => { throw new Error(`invalid_docx_xml:${message}`) }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(strFromU8(bytes), "application/xml")
}

function only(doc: Document, name: string) {
  const node = elements(doc, W, name)[0]
  if (!node) throw new Error("invalid_docx_package")
  return node
}

function elements(node: Document | Element, namespace: string, name: string) {
  return Array.from(node.getElementsByTagNameNS(namespace, name))
}

function children(node: Element) {
  return Array.from(node.childNodes).filter((row): row is Element => row.nodeType === 1)
}

function child(node: Element, name: string) {
  return children(node).find(row => row.namespaceURI === W && row.localName === name)
}

function value(node?: Element) {
  return node?.getAttributeNS(W, "val") ?? ""
}

function hex(code: number) {
  return code.toString(16).toUpperCase().padStart(8, "0")
}
