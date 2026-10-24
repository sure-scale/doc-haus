import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { resolvedReferences, type ReferenceFields } from "../workflows/lean/consistency"
import { openRevision } from "./docx-amendments"
import { owningParagraph, paragraphPointers } from "./drafting-amendment"
import { paragraphFields, refInstruction, runText, unlink, type ParagraphField } from "./docx-fields"
import { clauseNumbering } from "./docx-numbering"
import { shownText } from "./list-label"

// Live cross-references. Once Word list numbering numbers a document's
// clauses, a reference to one of its clauses or attachments is a REF field to
// a bookmark on the paragraph it points at, so Word (on F9 or print) and
// LibreOffice (on opening) show the clause's current number after an edit
// renumbers it. The reference's target is the one the cross-reference check
// resolves (consistency.ts); a reference it does not resolve stays text and is
// reported there. Only the number is the field: "clause " stays text. Its
// cached result is the number the list prints for the target, without a
// trailing period, which is what `REF … \r` shows for the clause list's
// labels; an attachment's number is typed in its heading, so the field shows
// the bookmarked number itself. A document whose clause numbers are typed has
// no clause list, and so no fields. A reference to the paragraph it is in
// ("this clause 5.3" in clause 5.3) stays text: a REF field inside its own
// bookmark is an error in Word.
//
// Each pass also refreshes the fields already in the body: a cached number
// follows its target. A field whose bookmark is gone (its clause deleted, or
// left out of a template) showed a number that may now be another clause's:
// it becomes a visible open-item marker naming that number, "[open item:
// refers to 2.1, which was removed]", which the body's open-item check
// reports like any other. A field to a target that lost its list number
// becomes the plain text it showed, which the cross-reference check then reads
// as any other reference. A reference unlinked so is not linked again in the
// same pass to the clause that now carries its number. A clause split in parts
// keeps its bookmark on its first part, the one that keeps its number.
//
// A pass may link only some paragraphs (`relink`), as after an edit: each is
// read again as the text it shows, refreshed, and its references are linked as
// a drafted paragraph's are. A refreshed field that shows a whole reference and
// points at its clause stays; one showing part of a reference, or pointing at
// another clause, is replaced by it, so a number the edit extended ("clause 6"
// to "clause 6.2") points at the clause it now names. A field no reference
// read from the text covers stays as refreshed. The other paragraphs are only
// refreshed.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const XML = "http://www.w3.org/XML/1998/namespace"
const BODY = "word/document.xml"

export type CrossReferenceOptions = {
  /** The document's title; its first paragraph when not given. */
  title?: string
  /** Its planned form: advice and letters cite other documents' clauses and get no fields. */
  form?: string
  /** The titles of the other documents of the pack: a reference placed after one is to that document. */
  titles?: string[]
  /** Refresh the fields already there without linking new ones. */
  refreshOnly?: true
  /** Link only the paragraphs this selects, read again from their text; refresh the others. */
  relink?: (paragraph: Element) => boolean
  /** Told the open-item marker written for each field whose bookmark is gone. */
  onRemovedTarget?: (marker: string) => void
}

/** The body with its resolved cross-references as REF fields and its existing
 * REF fields refreshed. Returns `bytes` itself when nothing changes, or when
 * the body cannot be read (tracked changes included). */
export async function linkCrossReferences(bytes: Uint8Array, options: CrossReferenceOptions = {}): Promise<Uint8Array> {
  const clauses = clauseNumbering(bytes).numIds
  if (!clauses.size) return bytes
  const opened = await openRevision(bytes).catch(() => undefined)
  if (!opened) return bytes
  try {
    const doc = opened.documents[BODY]!
    const rows = opened.paragraphs.filter(row => row.part === BODY).map(row => {
      const label = row.metadata.list?.generatedLabel?.trim() || undefined
      return { node: row.node, text: row.match.text, label, clause: label !== undefined && clauses.has(String(row.metadata.list!.numId)),
        exact: runText(row.node) === row.match.text }
    })
    const rowOf = new Map(rows.map((row, at) => [row.node, at]))
    const marks = new Map<string, { start: Element; at?: number }>()
    for (const start of elements(doc, "bookmarkStart")) {
      const name = start.getAttributeNS(W, "name")
      if (!name || marks.has(name)) continue
      let paragraph: Node | null = start.parentNode
      while (paragraph && !(paragraph.nodeType === 1 && (paragraph as Element).localName === "p")) paragraph = paragraph.parentNode
      marks.set(name, { start, at: paragraph ? rowOf.get(paragraph as Element) : undefined })
    }
    let changed = false
    const ends = new Map(elements(doc, "bookmarkEnd").map(end => [end.getAttributeNS(W, "id"), end]))
    for (const [name, mark] of marks) {
      const paragraph = mark.start.parentNode as Element
      const end = ends.get(mark.start.getAttributeNS(W, "id"))
      if (!/^_Ref\d+$/.test(name) || mark.at === undefined || paragraph !== rows[mark.at]!.node || !end || end.parentNode === paragraph) continue
      if (elements(paragraph, "*").findIndex(node => node.localName === "t") < elements(paragraph, "*").indexOf(mark.start)) continue
      paragraph.appendChild(end)
      changed = true
    }

    // Refresh. A number follows its target's label; bookmarked text follows the bookmark.
    const unlinked = new Map<number, { start: number; end: number }[]>()
    rows.forEach((row, at) => {
      for (const field of paragraphFields(row.node)) {
        const ref = refInstruction(field)
        if (!ref) continue
        const mark = marks.get(ref.name)
        const showed = field.result.trim()
        if (!mark && field.separate && showed) {
          const marker = `[open item: refers to ${showed}, which was removed]`
          setResult(doc, field, marker)
          unlink(field)
          unlinked.set(at, [...unlinked.get(at) ?? [], { start: field.from, end: field.from + marker.length }])
          options.onRemovedTarget?.(marker)
          changed = true
          continue
        }
        if (!row.exact) continue
        const target = mark?.at === undefined ? undefined : rows[mark.at]
        if (!mark || (ref.number && target && !target.label)) {
          unlink(field)
          unlinked.set(at, [...unlinked.get(at) ?? [], { start: field.from, end: field.to }])
          changed = true
          continue
        }
        if (!target) continue
        const shows = ref.number ? numberOf(target.label) : covered(mark.start)
        if (shows === undefined || shows === field.result) continue
        setResult(doc, field, shows)
        changed = true
      }
      if (changed && row.exact) row.text = runText(row.node)
    })

    const relinked = new Set(options.relink && !options.refreshOnly ? rows.flatMap((row, at) => row.exact && options.relink!(row.node) ? [at] : []) : [])

    if (!options.refreshOnly && (!options.relink || relinked.size)) {
      const paras = rows.map(row => ({ text: row.clause ? shownText(row) : row.text }))
      const title = options.title ?? rows.find(row => row.text.trim())?.text.trim() ?? ""
      const names = new Map<string, string>()
      const taken = new Set(marks.keys())
      let next = nextBookmarkId(opened.documents)
      for (const ref of resolvedReferences(title, paras, options.titles, options.form)) {
        if ((options.relink && !relinked.has(ref.at)) || ref.target === ref.at) continue
        const source = rows[ref.at]!
        const target = rows[ref.target]!
        const start = ref.start - (source.clause ? source.label!.length + 1 : 0)
        const end = start + ref.length
        if (!source.exact || !target.exact || start < 0 || source.text.slice(start, end) !== ref.label) continue
        if ((unlinked.get(ref.at) ?? []).some(span => start <= span.end && span.start <= end)) continue
        // A paragraph to relink keeps a field that shows this whole reference
        // and points at its clause; one showing part of it, or pointing at
        // another clause, gives way to the reference its text now makes.
        const overlapping = paragraphFields(source.node).filter(field => start <= field.to && field.from <= end)
        if (overlapping.length && (!relinked.has(ref.at) || overlapping.some(field => !refInstruction(field)
          || (field.from === start && field.to === end && marks.get(refInstruction(field)!.name)?.at === ref.target)))) continue
        const attachment = ref.number !== undefined
        const key = `${ref.target}\0${attachment}`
        let name = names.get(key)
        if (!name) {
          if (attachment) {
            const from = ref.number! - (target.clause ? target.label!.length + 1 : 0)
            if (target.clause || target.text.slice(from, from + ref.length) !== ref.label) continue
            name = existingNumberMark(target.node, ref.label)
            if (!name) {
              const runs = isolate(target.node, from, from + ref.length)
              if (!runs) continue
              name = uniqueName(`${ref.kind}\0${target.text}`, taken)
              const id = String(next++)
              runs[0]!.parentNode!.insertBefore(bookmark(doc, "bookmarkStart", id, name), runs[0]!)
              runs.at(-1)!.parentNode!.insertBefore(bookmark(doc, "bookmarkEnd", id), runs.at(-1)!.nextSibling)
            }
          } else {
            if (!target.clause || numberOf(target.label) !== ref.label) continue
            name = existingParagraphMark(target.node)
            if (!name) {
              name = uniqueName(`clause\0${shownText(target)}`, taken)
              const id = String(next++)
              const properties = childElements(target.node, "pPr")[0]
              target.node.insertBefore(bookmark(doc, "bookmarkStart", id, name), properties ? properties.nextSibling : target.node.firstChild)
              target.node.appendChild(bookmark(doc, "bookmarkEnd", id))
            }
          }
          names.set(key, name)
        }
        // A field that gave way stays typed if the new one cannot be made.
        for (const field of overlapping) unlink(field)
        if (overlapping.length) changed = true
        const runs = isolate(source.node, start, end)
        if (!runs) continue
        wrapInField(doc, runs, attachment ? ` REF ${name} \\h ` : ` REF ${name} \\r \\h `)
        changed = true
      }
    }
    if (!changed) return bytes
    const parts = { ...opened.parts, [BODY]: strToU8(new XMLSerializer().serializeToString(doc)) }
    return zipSync(parts, { mtime: new Date(1980, 0, 1) })
  } finally {
    opened.session.close()
  }
}

/** The REF fields of a body's paragraphs, by paragraph pointer: where each
 * cached result is in the paragraph's text, and the paragraph its bookmark is
 * on (a number field) or the text the bookmark covers. For the
 * cross-reference check (consistency.ts). */
export function referenceFields(bytes: Uint8Array): ReferenceFields {
  const doc = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[BODY]!), "application/xml")
  const marks = new Map<string, Element>()
  for (const start of elements(doc, "bookmarkStart")) {
    const name = start.getAttributeNS(W, "name")
    if (name && !marks.has(name)) marks.set(name, start)
  }
  const pid = (paragraph: Node | null) => paragraph ? paragraphPointers(paragraph as Element)[0]?.pid : undefined
  const found: ReferenceFields = new Map()
  for (const paragraph of elements(doc, "p")) {
    const own = pid(paragraph)
    if (!own) continue
    const links = paragraphFields(paragraph).flatMap(field => {
      const ref = refInstruction(field)
      if (!ref) return []
      const mark = marks.get(ref.name)
      return [{ start: field.from, length: field.to - field.from, number: ref.number !== undefined,
        ...(mark && ref.number ? { target: pid(owningParagraph(mark)) } : {}), ...(mark && !ref.number ? { covers: covered(mark) } : {}) }]
    })
    if (links.length) found.set(own, links)
  }
  return found
}

// ------------------------------------------------------------------ fields

function setResult(doc: Document, field: ParagraphField, text: string) {
  const [first, ...rest] = field.texts
  for (const node of rest) {
    const run = node.parentNode as Element
    run.removeChild(node)
    if (!Array.from(run.childNodes).some(child => child.nodeType === 1 && (child as Element).localName !== "rPr")) run.parentNode?.removeChild(run)
  }
  if (first) {
    first.textContent = text
    first.setAttributeNS(XML, "xml:space", "preserve")
    return
  }
  const separate = field.separate!.parentNode as Element
  const run = doc.createElementNS(W, "w:r")
  const properties = childElements(separate, "rPr")[0]
  if (properties) run.appendChild(properties.cloneNode(true))
  const node = doc.createElementNS(W, "w:t")
  node.setAttributeNS(XML, "xml:space", "preserve")
  node.textContent = text
  run.appendChild(node)
  separate.parentNode!.insertBefore(run, separate.nextSibling)
}

/** `runs` (consecutive, under one parent) as the cached result of a new field. */
function wrapInField(doc: Document, runs: Element[], instruction: string) {
  const properties = childElements(runs[0]!, "rPr")[0]
  const control = (node: Element) => {
    const run = doc.createElementNS(W, "w:r")
    if (properties) run.appendChild(properties.cloneNode(true))
    run.appendChild(node)
    return run
  }
  const char = (type: string) => {
    const node = doc.createElementNS(W, "w:fldChar")
    node.setAttributeNS(W, "w:fldCharType", type)
    return node
  }
  const code = doc.createElementNS(W, "w:instrText")
  code.setAttributeNS(XML, "xml:space", "preserve")
  code.textContent = instruction
  const parent = runs[0]!.parentNode!
  for (const node of [control(char("begin")), control(code), control(char("separate"))]) parent.insertBefore(node, runs[0]!)
  parent.insertBefore(control(char("end")), runs.at(-1)!.nextSibling)
}

// -------------------------------------------------------------------- runs

/** The text [start, end) of a paragraph as whole runs of their own, split
 * from their neighbours with the same formatting. Undefined unless the text is
 * in plain text runs (properties and one w:t) next to each other under one
 * parent. */
function isolate(paragraph: Element, start: number, end: number): Element[] | undefined {
  const covering = textRuns(paragraph).filter(row => row.start < end && start < row.end)
  if (!covering.length || covering.some(row => !row.plain) || covering.some(row => row.run.parentNode !== covering[0]!.run.parentNode)) return
  for (let at = 1; at < covering.length; at += 1) {
    for (let node = covering[at - 1]!.run.nextSibling; node && node !== covering[at]!.run; node = node.nextSibling)
      if (node.nodeType === 1 && !["bookmarkStart", "bookmarkEnd", "proofErr"].includes((node as Element).localName)) return
  }
  const last = covering.at(-1)!
  if (end < last.end) splitRun(last.run, end - last.start)
  const first = covering[0]!
  if (start > first.start) splitRun(first.run, start - first.start)
  return textRuns(paragraph).filter(row => row.start >= start && row.end <= end && row.end > row.start).map(row => row.run)
}

function textRuns(paragraph: Element) {
  const rows: { run: Element; start: number; end: number; plain: boolean }[] = []
  let offset = 0
  for (const node of elements(paragraph, "t")) {
    const run = node.parentNode as Element
    const length = (node.textContent ?? "").length
    const kids = Array.from(run.childNodes).filter(child => child.nodeType === 1) as Element[]
    const plain = run.localName === "r" && kids.every(child => child.localName === "rPr" || child === node)
    rows.push({ run, start: offset, end: offset + length, plain })
    offset += length
  }
  return rows
}

function splitRun(run: Element, at: number) {
  const copy = run.cloneNode(true) as Element
  const [head, tail] = [childElements(run, "t")[0]!, childElements(copy, "t")[0]!]
  const text = head.textContent ?? ""
  head.textContent = text.slice(0, at)
  tail.textContent = text.slice(at)
  for (const node of [head, tail]) node.setAttributeNS(XML, "xml:space", "preserve")
  run.parentNode!.insertBefore(copy, run.nextSibling)
}

// --------------------------------------------------------------- bookmarks

function bookmark(doc: Document, kind: "bookmarkStart" | "bookmarkEnd", id: string, name?: string) {
  const node = doc.createElementNS(W, `w:${kind}`)
  node.setAttributeNS(W, "w:id", id)
  if (name) node.setAttributeNS(W, "w:name", name)
  return node
}

/** A cross-reference bookmark over a whole paragraph's text. */
function existingParagraphMark(paragraph: Element) {
  for (const start of childElements(paragraph, "bookmarkStart")) {
    const name = start.getAttributeNS(W, "name") ?? ""
    if (/^_Ref\d+$/.test(name) && covered(start) === runText(paragraph)) return name
  }
}

/** A cross-reference bookmark over exactly `label` in a heading. */
function existingNumberMark(paragraph: Element, label: string) {
  for (const start of elements(paragraph, "bookmarkStart")) {
    const name = start.getAttributeNS(W, "name") ?? ""
    if (/^_Ref\d+$/.test(name) && covered(start) === label) return name
  }
}

/** The text a bookmark covers within its paragraph; undefined when it ends
 * outside it. */
function covered(start: Element) {
  const id = start.getAttributeNS(W, "id")
  let paragraph: Node | null = start.parentNode
  while (paragraph && !(paragraph.nodeType === 1 && (paragraph as Element).localName === "p")) paragraph = paragraph.parentNode
  if (!paragraph) return
  let text = ""
  let inside = false
  for (const node of elements(paragraph as Element, "*")) {
    if (node === start) inside = true
    else if (inside && node.localName === "bookmarkEnd" && node.getAttributeNS(W, "id") === id) return text
    else if (inside && node.localName === "t") text += node.textContent ?? ""
  }
}

/** A Word-style hidden bookmark name, `_Ref` and nine digits, the same for the
 * same target each time a document is rendered. */
function uniqueName(seed: string, taken: Set<string>) {
  let hash = 0x811c9dc5
  for (const char of seed) hash = Math.imul(hash ^ char.codePointAt(0)!, 0x01000193) >>> 0
  let number = hash % 1_000_000_000
  let name = `_Ref${String(number).padStart(9, "0")}`
  while (taken.has(name)) {
    number = (number + 1) % 1_000_000_000
    name = `_Ref${String(number).padStart(9, "0")}`
  }
  taken.add(name)
  return name
}

function nextBookmarkId(documents: Record<string, Document>) {
  const ids = Object.values(documents).flatMap(doc => elements(doc, "bookmarkStart").map(node => Number(node.getAttributeNS(W, "id")) || 0))
  return Math.max(0, ...ids) + 1
}

// ------------------------------------------------------------------- shared

/** The number a list label shows in a REF `\r` field: "3." as "3", "3.1" as
 * itself. Undefined for a label that is not a plain dotted number. */
function numberOf(label: string | undefined) {
  const number = label?.replace(/\.+$/, "")
  return number && /^\d+(?:\.\d+)*$/.test(number) ? number : undefined
}


function elements(node: Document | Element, name: string) {
  return Array.from(node.getElementsByTagNameNS(W, name))
}

function childElements(node: Element, name: string) {
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1 && (child as Element).namespaceURI === W && (child as Element).localName === name)
}
