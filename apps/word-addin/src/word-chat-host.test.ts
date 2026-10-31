import { expect, test } from "bun:test"
import { WordChatHost, WordChatHostError, type WordChatContext, type WordChatWord, type WordChatApplyInput } from "./word-chat-host"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { validWordChatFormatting, type WordChatAuthoredAmendment, type WordChatParagraphFormat, type WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { wordChanges } from "../../../dochaus/lib/word-changes"

// This integration uses the real Bun/shared writer without pulling its server
// dependencies and ES2024 library requirements into the browser's type graph.
async function authorWordChatTarget(ooxml: string, target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[]) {
  const writer = await import(new URL("../../../dochaus/lib/docx-flat-opc.ts", import.meta.url).href) as {
    authorWordChatTarget(xml: string, target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[]): Promise<Omit<WordChatAuthoredAmendment, "proposalId">>
  }
  return writer.authorWordChatTarget(ooxml, target, replacement, formatting)
}

const WORD_XML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const PACKAGE_XML = "http://schemas.microsoft.com/office/2006/xmlPackage"
// Minimal reviewed synthetic shape of actual Mac Word Range.getOoxml(): even
// one ordinary word exports Flat OPC with a final body-level section layout.
const escapeXml = (text: string) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
const flatOpc = (content: string, extra = "") => `<pkg:package xmlns:pkg="${PACKAGE_XML}" xmlns:w="${WORD_XML}"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document><w:body>${content}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document></pkg:xmlData></pkg:part>${extra}</pkg:package>`

test("optional native table capture binds geometry to the same paragraph inventory and fails closed", async () => {
  const paragraph = (id: string, text: string) => ({uniqueLocalId:id,text,tableNestingLevel:1})
  const cells = [[paragraph("a","One"),paragraph("b","Two")], [paragraph("c","Three"),paragraph("d","Four")]]
  const collection = <T>(items: T[]) => ({items,load() {}})
  const table = {rowCount:2,isUniform:true,values:cells.map(row => row.map(p => p.text)),load() {},getCell(r: number,c: number) { return {body:{paragraphs:collection([cells[r]![c]!])}} }}
  const cellXml = (p: {text:string}, merged = false) => `<w:tc>${merged ? '<w:tcPr><w:vMerge/></w:tcPr>' : ''}<w:p><w:r><w:t>${p.text}</w:t></w:r></w:p></w:tc>`
  const tableXml = (merged = false) => `<w:tbl>${cells.map((row, i) => `<w:tr>${row.map((p, j) => cellXml(p, merged && i === 0 && j === 0)).join("")}</w:tr>`).join("")}</w:tbl>`
  let xml = flatOpc(tableXml()), tables = [table], broken = false
  const body = { text:"One\rTwo\rThree\rFour",load() {},paragraphs:collection(cells.flat()),get tables() { if (broken) throw Error("unavailable"); return collection(tables) },getRange() { return {getOoxml() {return {value:xml}}} } }
  const word = {run: async (fn: (context: unknown) => Promise<unknown>) => fn({document:{saved:false,load() {},body},sync:async () => {}})} as WordChatWord
  const host = new WordChatHost(word, () => true)
  expect((await host.capture()).table).toEqual({tableIndex:0,rows:cells.map(row => row.map(p => ({id:p.uniqueLocalId,text:p.text})))})
  xml = flatOpc(tableXml(true))
  expect((await host.capture()).table).toBeUndefined()
  xml = flatOpc(tableXml() + tableXml()); tables = [table,table]
  expect((await host.capture()).table).toBeUndefined()
  broken = true
  const ordinary = await host.capture()
  expect(ordinary.text).toBe(body.text)
  expect(ordinary.table).toBeUndefined()
})

/** A text-position Office host model: selections and paragraph ranges are real
 * offsets, searches return every occurrence, and writes execute only on sync. */
function office(paragraphs = ["A clause with a clear obligation.", "Another clause."], nativeText = false) {
  type NativeParagraph = { id: string; text: string; style?: string; styleBuiltIn?: string; outlineLevel?: number;
    list?: { id: number; level: number; label: string }; inTable?: boolean }
  const state = {
    paragraphs: paragraphs.map((text, i): NativeParagraph => ({ id: `paragraph-${i}`, text })),
    selection: [0, paragraphs[0]!.length], separator: "\r", saved: false, tracking: "Off", trackFormatting: true as boolean | undefined, modeChanges: [] as string[],
    writes: [] as Array<{ before: string; text: string; tracking: string; start: number; end: number }>, syncBatchSizes: [] as number[],
    paragraphXml: new Map<string, string>(), comments: [] as Array<[number, number]>, revisions: [] as Array<[number, number]>, xml: undefined as string | undefined,
    selected: 0, highlights: 0, outside: false, annotationCounter: 0,
    annotations: new Map<string, unknown>(), handlers: new Map<string, Set<(args: never) => void>>(),
    throwAfterWrite: false, syncFailure: undefined as Error | undefined, corruptWrite: false, afterSync: undefined as undefined | (() => void),
    driftInsertedRanges: false, renewImportedIds: false, omitNewRevisions: false, authorCalls: 0, duringAuthor: undefined as undefined | (() => void | Promise<void>),
    nativeRevisions: [] as Array<{ start: number; end: number; author: string; type: string; text: string }>,
    extraParts: "", formattedTextEmpty: false, leakAdjacentFormatting: false, wholeOoxmlTrailingParagraph: false,
    inlineBold: false, inlineAfter: "", inlineWrites: 0,
    paragraphInsertions: 0, corruptParagraphInsertion: false, insertedParagraphStyle: "Normal",
    listInsertions: 0, listOriginal: "", listBadFont: false, listBadMarkup: false,
    paragraphDeletions: 0, corruptParagraphDeletion: false, wrongDeletionCurrent: false,
    nativeFormatting: [] as string[], dropNativeFormatting: false, terminalBodyMark: false, cellWholeExpandsRow: false, refuseEmptyContentXml: false,
    omitContentProperties: false, paragraphLoads: [] as Array<{ id: string; properties: string }>, metadataUnavailable: false, metadataFailures: 0,
  }
  const text = () => state.paragraphs.map((p) => p.text).join(state.separator)
  const position = (id: string) => {
    let start = 0
    for (const p of state.paragraphs) {
      if (p.id === id) return { p, start, end: start + p.text.length }
      start += p.text.length + state.separator.length
    }
    throw new Error("ItemNotFound")
  }
  const intersects = (a: number, b: number, c: number, d: number) => a === b ? c <= a && a <= d : c === d ? a <= c && c <= b : a < d && c < b
  const queue: Array<() => void> = []
  const event = (name: string) => ({ add(callback: (args: never) => void) {
    const handlers = state.handlers.get(name) ?? new Set()
    state.handlers.set(name, handlers)
    handlers.add(callback)
    return { context, remove() { handlers.delete(callback) } }
  } })
  const collection = <T>(items: T[]) => ({ items, load() {} })
  const ranges = new Set<Range>()
  class Range {
    constructor(public start: number, public end: number, public bodyWhole = false, public cellWhole = false, public paragraphContent = false, public bodyExport = false) { ranges.add(this) }
    get text() { return (text() + "\r").slice(this.start, this.end) }
    get font() { return { load() {}, get bold() { return state.inlineBold }, set bold(value: boolean | null) { queue.push(() => { state.inlineWrites++; state.inlineBold = Boolean(value); if (state.inlineAfter) state.paragraphXml.set("paragraph-0", state.inlineAfter) }) }, italic: false, underline: "None", name:"Arial", size:11, get color() { return state.listBadFont && state.listInsertions ? "#FF0000" : "#000000" } } }
    get paragraphs() {
      return collection(state.paragraphs.filter((p) => {
        const at = position(p.id)
        return this.start === this.end ? at.start <= this.start && this.start <= at.end : at.start < this.end && this.start <= at.end
      }).map(paragraph))
    }
    load() {}
    getRange(location = "Whole") {
      if (location === "Start") return new Range(this.start, this.start)
      if (location === "End") return new Range(this.bodyWhole ? this.end - 1 : this.end, this.bodyWhole ? this.end - 1 : this.end)
      return new Range(this.start, this.end)
    }
    expandTo(other: Range) { return new Range(Math.min(this.start, other.start), Math.max(this.end, other.end)) }
    compareLocationWith(other: Range) {
      return { value: state.outside ? "Unrelated" : this.start >= other.start && this.end <= other.end ? "Inside" : "Unrelated" }
    }
    search(query: string) {
      // Decode only the literal/control Word Find forms used by the adapter.
      const literal = query.replace(/\^([\^plt])/g, (_, code: string) => ({ "^": "^", p: "\r", l: "\u000b", t: "\t" })[code]!)
      const matches: Range[] = []
      let at = this.text.indexOf(literal)
      while (at >= 0) { matches.push(new Range(this.start + at, this.start + at + literal.length)); at = this.text.indexOf(literal, at + 1) }
      return collection(matches)
    }
    insertText(value: string, location: "Before" | "After" | "Start" | "End" | "Replace" = "After") {
      const replace = location === "Replace"
      const start = replace ? this.start : this.end, end = replace ? this.end : this.end
      const result = new Range(start, start + value.length)
      queue.push(() => {
        const before = text(), removed = before.slice(start, end)
        const delta = value.length - removed.length
        state.nativeRevisions = state.nativeRevisions.map(revision => {
          if (revision.start >= end) { revision.start += delta; revision.end += delta }
          return revision
        })
        if (state.tracking !== "Off" && removed) state.nativeRevisions.push({ start, end: start, author: "Word user", type: "Deleted", text: removed })
        if (state.tracking !== "Off" && value) state.nativeRevisions.push({ start, end: start + value.length, author: "Word user", type: "Added", text: value })
        const next = (before.slice(0, start) + value + before.slice(end)).split(state.separator)
        state.paragraphs = state.paragraphs.map((p, index) => ({ ...p, text: next[index] ?? "" }))
        state.comments = state.comments.map(([a, b]) => [a >= end ? a + delta : a, b >= end ? b + delta : b])
        state.revisions = state.revisions.map(([a, b]) => [a >= end ? a + delta : a, b >= end ? b + delta : b])
        state.writes.push({ before, text: value, tracking: state.tracking, start, end })
        state.saved = false
      })
      return result
    }
    delete() {
      const start = this.start, end = this.end
      queue.push(() => {
        const before = text(), removed = before.slice(start, end)
        if (!removed) return
        if (state.tracking !== "Off") state.nativeRevisions.push({ start, end: start, author: "Word user", type: "Deleted", text: removed })
        state.nativeRevisions = state.nativeRevisions.map(revision => {
          if (revision.start >= end) { revision.start -= removed.length; revision.end -= removed.length }
          return revision
        })
        const next = (before.slice(0, start) + before.slice(end)).split(state.separator)
        state.paragraphs = state.paragraphs.map((p, index) => ({ ...p, text: next[index] ?? "" }))
        state.comments = state.comments.map(([a, b]) => [a >= end ? a - removed.length : a, b >= end ? b - removed.length : b])
        state.revisions = state.revisions.map(([a, b]) => [a >= end ? a - removed.length : a, b >= end ? b - removed.length : b])
        state.writes.push({ before, text: "", tracking: state.tracking, start, end })
        state.saved = false
      })
    }
    insertOoxml(xml: string) {
      const result = new Range(this.start, this.end)
      queue.push(() => {
        const doc = new DOMParser().parseFromString(xml, "application/xml")
        const blocks = Array.from(doc.getElementsByTagNameNS(WORD_XML, "body")[0]!.childNodes).filter((n): n is Element => n.nodeType === 1 && (n as Element).localName === "p")
        const owners = state.paragraphs.filter(p => { const at = position(p.id); return this.start <= at.start && at.end <= this.end })
        const start = this.start, end = this.end, before = this.text
        const current = (node: Node): string => {
          const element = node as Element
          if (element.nodeType !== 1 || ["del", "pPr", "rPr"].includes(element.localName)) return ""
          if (element.localName === "t") return element.textContent || ""
          if (element.localName === "tab") return "\t"
          if (["br", "cr"].includes(element.localName)) return "\v"
          return Array.from(element.childNodes).map(current).join("")
        }
        const texts = blocks.map(current), whole = this.end === position(owners.at(-1)!.id).end + 1, replacement = texts.join("\r") + (whole ? "\r" : "")
        state.nativeRevisions = state.nativeRevisions.filter(revision => !intersects(start, end, revision.start, revision.end))
        const imported = blocks.map((_, index) => owners[index] ?? { id: `new-${state.writes.length}-${index}`, text: "" })
        const at = state.paragraphs.indexOf(owners[0]!)
        state.paragraphs.splice(at, owners.length, ...imported)
        for (const [index, owner] of imported.entries()) {
          owner.text = state.corruptWrite && index === 0 ? "WRONG" : texts[index]!
          if (state.renewImportedIds) owner.id = `imported-${owner.id}`
          state.paragraphXml.set(owner.id, new XMLSerializer().serializeToString(blocks[index]!))
        }
        const delta = replacement.length - (end - start)
        state.writes.push({ before, text: replacement, tracking: state.tracking, start, end })
        for (const kind of ["comments", "revisions"] as const) state[kind] = state[kind].map(([a, b]) => a >= end ? [a + delta, b + delta] : [a, b])
        for (const revision of state.nativeRevisions) if (revision.start >= end) { revision.start += delta; revision.end += delta }
        if (!state.omitNewRevisions) for (const [index, block] of blocks.entries()) {
          let at = position(imported[index]!.id).start
          const visit = (node: Node) => {
            const element = node as Element
            if (element.nodeType !== 1) return
            if (["pPr", "rPr"].includes(element.localName)) return
            if (["ins", "del"].includes(element.localName)) {
              const added = element.localName === "ins"
              const wording = Array.from(element.getElementsByTagNameNS(WORD_XML, added ? "t" : "delText")).map(n => n.textContent).join("")
              state.nativeRevisions.push({ start: at, end: at + (added ? wording.length : 0), author: element.getAttributeNS(WORD_XML, "author") || "", type: added ? "Added" : "Deleted", text: wording })
              if (added) at += wording.length
              return
            }
            if (element.localName === "t") { at += element.textContent?.length || 0; return }
            for (const child of Array.from(element.childNodes)) visit(child)
          }
          visit(block)
          const mark = Array.from(block.getElementsByTagNameNS(WORD_XML, "pPr")[0]?.getElementsByTagNameNS(WORD_XML, "rPr")[0]?.childNodes ?? []).find(n => n.nodeType === 1 && ["ins", "del"].includes((n as Element).localName)) as Element | undefined
          if (mark) state.nativeRevisions.push({ start: at, end: at + 1, author: mark.getAttributeNS(WORD_XML, "author") || "", type: mark.localName === "ins" ? "Added" : "Deleted", text: "\r" })
          const format = Array.from(block.getElementsByTagNameNS(WORD_XML, "pPr")[0]?.childNodes ?? []).find(n => n.nodeType === 1 && (n as Element).localName === "pPrChange") as Element | undefined
          if (format) state.nativeRevisions.push({ start: position(imported[index]!.id).start, end: position(imported[index]!.id).end,
            author: format.getAttributeNS(WORD_XML, "author") || "", type: "Formatted", text: state.formattedTextEmpty ? "" : texts[index]! })
        }
        if (state.leakAdjacentFormatting) {
          const next = state.paragraphs[at + imported.length]
          if (next) state.paragraphXml.set(next.id, `<w:p><w:pPr><w:spacing w:after="480"/></w:pPr><w:r><w:t>${escapeXml(next.text)}</w:t></w:r></w:p>`)
        }
        state.saved = false
        if (state.driftInsertedRanges) { result.start = 0; result.end = text().length }
      })
      return result
    }
    getOoxml() {
      if (state.refuseEmptyContentXml && this.paragraphContent && this.start === this.end) throw new Error("empty Content XML unavailable")
      const paragraphs = this.paragraphs.items
      const whole = paragraphs.length && this.end === position(paragraphs.at(-1)!.uniqueLocalId).end + 1
      const markup = (p: { uniqueLocalId: string; text: string }) => {
        const stored = state.paragraphXml.get(p.uniqueLocalId)
        if (stored) return stored
        const bounds = position(p.uniqueLocalId)
        const changes = state.nativeRevisions.filter(change => ["Added", "Deleted"].includes(change.type) && change.start >= bounds.start && change.end <= bounds.end)
          .sort((a, b) => a.start - b.start)
        const run = (value: string, deleted = false) => value ? `<w:r><w:${deleted ? "delText" : "t"} xml:space="preserve">${escapeXml(value)}</w:${deleted ? "delText" : "t"}></w:r>` : ""
        let cursor = bounds.start, result = ""
        for (const [index, change] of changes.entries()) {
          result += run(text().slice(cursor, change.start))
          const tag = change.type === "Added" ? "ins" : "del"
          result += `<w:${tag} w:id="${index + 1}" w:author="${escapeXml(change.author)}">${run(change.text, tag === "del")}</w:${tag}>`
          cursor = change.end
        }
        return `<w:p>${result}${run(text().slice(cursor, bounds.end))}</w:p>`
      }
      if (this.cellWhole) return { value: flatOpc(`<w:tbl><w:tr>${state.paragraphs.filter(p => p.inTable).map(p => `<w:tc>${markup(paragraph(p))}</w:tc>`).join("")}</w:tr></w:tbl><w:p/>`, state.extraParts) }
      let body = paragraphs.map(markup).join("")
      if (this.bodyExport && state.paragraphs.some(p => p.inTable)) {
        body = ""; let cells: string[] = []
        const flush = () => { if (cells.length) { body += `<w:tbl><w:tr>${cells.join("")}</w:tr></w:tbl>`; cells = [] } }
        for (const p of state.paragraphs) { if (p.inTable) cells.push(`<w:tc>${markup(paragraph(p))}</w:tc>`); else { flush(); body += markup(paragraph(p)) } }
        flush()
      }
      if (this.paragraphContent && state.omitContentProperties) body = body.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/g, "")
      return { value: !state.writes.length && state.xml !== undefined ? state.xml : flatOpc(body + (state.wholeOoxmlTrailingParagraph && whole ? "<w:p/>" : ""), state.extraParts) }
    }
    getReviewedText(version?: string) {
      if (version === "Original" && state.listInsertions && this.start === 0 && this.end === text().length) return {value:state.listOriginal}
      const paragraphs = state.paragraphs.filter(p => { const at = position(p.id); return this.start <= at.start && at.end <= this.end })
      const complete = paragraphs.length && position(paragraphs[0]!.id).start === this.start && position(paragraphs.at(-1)!.id).end === this.end
      const deletedMark = (p: NativeParagraph) => {
        const doc = new DOMParser().parseFromString(flatOpc(state.paragraphXml.get(p.id) ?? `<w:p/>`), "application/xml")
        const properties = doc.getElementsByTagNameNS(WORD_XML, "p")[0]!.getElementsByTagNameNS(WORD_XML, "pPr")[0]
        const mark = properties && Array.from(properties.childNodes).find(n => n.nodeType === 1 && (n as Element).namespaceURI === WORD_XML && (n as Element).localName === "rPr")
        return mark && Array.from(mark.childNodes).some(n => n.nodeType === 1 && (n as Element).namespaceURI === WORD_XML && (n as Element).localName === "del")
      }
      return { value: state.paragraphDeletions && state.wrongDeletionCurrent && this.start === 0 && this.end === text().length ? "Wrong current document" :
        complete ? paragraphs.filter(p => !deletedMark(p)).map(p => p.text).join(state.separator) : this.text }
    }
    getComments() { return collection(state.comments.filter(([start, end]) => intersects(this.start, this.end, start, end)).map((_, index) => ({ author: "Commenter", content: `Comment ${index + 1}`, resolved: false }))) }
    getTrackedChanges() {
      return collection([
        ...state.revisions.filter(([start, end]) => intersects(this.start, this.end, start, end)).map(([start, end]) => ({ author: "Earlier reviewer", type: "Added", text: text().slice(start, end) })),
        ...state.nativeRevisions.filter(revision => intersects(this.start, this.end, revision.start, revision.end)).sort((a, b) => a.start - b.start),
      ])
    }
    select() { state.selected++ }
    highlight() { state.highlights++ }
    removeHighlight() { state.highlights-- }
    track() { return this }
    untrack() { return this }
  }
  const setFormat = (p: NativeParagraph, field: string, value: string) => queue.push(() => {
    if (state.dropNativeFormatting) return
    const xml = state.paragraphXml.get(p.id) ?? `<w:p><w:r><w:t>${escapeXml(p.text)}</w:t></w:r></w:p>`
    const doc = new DOMParser().parseFromString(flatOpc(xml), "application/xml")
    const block = doc.getElementsByTagNameNS(WORD_XML, "p")[0]!
    let properties = Array.from(block.childNodes).find(n => n.nodeType === 1 && (n as Element).localName === "pPr") as Element | undefined
    if (!properties) { properties = doc.createElementNS(WORD_XML, "w:pPr"); block.insertBefore(properties, block.firstChild) }
    let change = Array.from(properties.childNodes).find(n => n.nodeType === 1 && (n as Element).localName === "pPrChange") as Element | undefined
    if (!change && state.trackFormatting && state.tracking === "TrackAll") {
      const previous = properties.cloneNode(true)
      change = doc.createElementNS(WORD_XML, "w:pPrChange")
      change.setAttributeNS(WORD_XML, "w:id", "900")
      change.setAttributeNS(WORD_XML, "w:author", "Word user")
      change.setAttributeNS(WORD_XML, "w:date", "2026-01-01T00:00:00Z")
      change.appendChild(previous); properties.appendChild(change)
      const at = position(p.id)
      state.nativeRevisions.push({ start: at.start, end: at.end, author: "Word user", type: "Formatted", text: p.text })
      state.nativeFormatting.push(p.id)
    }
    const name = field === "style" ? "pStyle" : "spacing"
    let element = Array.from(properties.childNodes).find(n => n.nodeType === 1 && (n as Element).localName === name) as Element | undefined
    if (!element) { element = doc.createElementNS(WORD_XML, `w:${name}`); properties.insertBefore(element, field === "style" ? properties.firstChild : change ?? null) }
    element.setAttributeNS(WORD_XML, `w:${field === "style" ? "val" : field}`, value)
    if (field === "line") element.setAttributeNS(WORD_XML, "w:lineRule", "auto")
    state.paragraphXml.set(p.id, new XMLSerializer().serializeToString(block))
    state.saved = false
  })
  const paragraph = (p: NativeParagraph) => ({
    get text() { return p.text }, get uniqueLocalId() { return p.id },
    get style() { return p.style ?? "Normal" }, get styleBuiltIn() { return p.styleBuiltIn ?? "Normal" },
    set styleBuiltIn(value: string) { setFormat(p, "style", value) },
    get spaceBefore() { return 0 }, get spaceAfter() { return 10 }, get lineSpacing() { return 12 }, get leftIndent() { return 18 }, get firstLineIndent() { return -18 },
    set spaceBefore(value: number) { setFormat(p, "before", String(value * 20)) },
    set spaceAfter(value: number) { setFormat(p, "after", String(value * 20)) },
    set lineSpacing(value: number) { setFormat(p, "line", String(value * 20)) },
    get outlineLevel() { return p.outlineLevel ?? 10 }, get isListItem() { return Boolean(p.list) },
    get tableNestingLevel() { return p.inTable ? 1 : 0 },
    get list() {
      if (!p.list) throw new Error("ItemNotFound")
      const id=p.list.id
      return { id, load() {}, get paragraphs() { return collection(state.paragraphs.filter(q => q.list?.id === id).map(paragraph)) }, insertParagraph(value:string,location:string) {
        if (location !== "End") throw new Error("invalid location")
        queue.push(() => {
          const last=state.paragraphs.filter(q => q.list?.id===id).at(-1)!, index=state.paragraphs.indexOf(last), oldId=last.id
          state.listOriginal=text(); state.listInsertions++
          const source=state.paragraphXml.get(oldId)!, doc=new DOMParser().parseFromString(flatOpc(source),"application/xml"), left=doc.getElementsByTagNameNS(WORD_XML,"p")[0]!
          const oldParaId=left.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml","paraId")
          left.setAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml","w14:paraId","00000009")
          const pp=left.getElementsByTagNameNS(WORD_XML,"pPr")[0]!, marker=doc.createElementNS(WORD_XML,"w:rPr"), ins=doc.createElementNS(WORD_XML,"w:ins")
          ins.setAttributeNS(WORD_XML,"w:id","11"); ins.setAttributeNS(WORD_XML,"w:author","Native user"); ins.setAttributeNS(WORD_XML,"w:date","2026-01-01T00:00:00Z"); marker.appendChild(ins);pp.appendChild(marker)
          last.id="retained-list"; state.paragraphXml.delete(oldId);state.paragraphXml.set(last.id,new XMLSerializer().serializeToString(left))
          state.paragraphs.splice(index+1,0,{id:oldId,text:value,list:{id,level:0,label:`${state.paragraphs.filter(q=>q.list?.id===id).length+1}.`}})
          state.paragraphXml.set(oldId,`<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${oldParaId}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:ins w:id="12" w:author="Native user" w:date="2026-01-01T00:00:00Z"><w:r><w:t>${escapeXml(state.listBadMarkup ? value+"wrong" : value)}</w:t></w:r></w:ins></w:p>`)
          state.saved=false
        }); return paragraph(p)
      } }
    },
    get listItemOrNullObject() { if (!p.list) return undefined; return {level:p.list.level,listString:p.list.label,siblingIndex:state.paragraphs.filter(q=>q.list?.id===p.list!.id).indexOf(p),load(){}} },
    get listItem() { if (!p.list) throw new Error("ItemNotFound"); return { level: p.list.level, listString: p.list.label, load() {} } },
    load(properties: string) {
      state.paragraphLoads.push({ id: p.id, properties })
      if (state.metadataUnavailable && properties.includes("style")) {
        state.metadataFailures++
        queue.push(() => { throw new Error("InvalidArgument: metadata unavailable") })
      }
    },
    getRange(location = "Whole") {
      const at = position(p.id)
      if (location === "Start") return new Range(at.start, at.start)
      if (location === "End") return new Range(at.end, at.end)
      return new Range(at.start, at.end + (location === "Whole" ? 1 : 0), false, state.cellWholeExpandsRow && Boolean(p.inTable) && location === "Whole", location === "Content")
    },
    delete() {
      queue.push(() => {
        state.paragraphDeletions++
        const stamp = 'w:author="Native user" w:date="2026-01-01T00:00:00Z"'
        const oldXml = state.paragraphXml.get(p.id) ?? `<w:p><w:r><w:t xml:space="preserve">${escapeXml(p.text)}</w:t></w:r></w:p>`
        const doc = new DOMParser().parseFromString(flatOpc(oldXml), "application/xml")
        const block = doc.getElementsByTagNameNS(WORD_XML, "p")[0]!
        const originalRun = block.getElementsByTagNameNS(WORD_XML, "r")[0]!.cloneNode(true) as Element
        const originalText = originalRun.getElementsByTagNameNS(WORD_XML, "t")[0]!
        const deletedText = doc.createElementNS(WORD_XML, "w:delText")
        for (const attr of Array.from(originalText.attributes)) deletedText.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
        deletedText.textContent = p.text + (state.corruptParagraphDeletion ? "wrong" : "")
        originalRun.replaceChild(deletedText, originalText)
        while (block.firstChild) block.removeChild(block.firstChild)
        const markup = `<w:p><w:pPr><w:rPr><w:del w:id="1" ${stamp}/></w:rPr></w:pPr><w:del w:id="2" ${stamp}>${new XMLSerializer().serializeToString(originalRun)}</w:del></w:p>`
        const deleted = new DOMParser().parseFromString(flatOpc(markup), "application/xml").getElementsByTagNameNS(WORD_XML, "p")[0]!
        for (const child of Array.from(deleted.childNodes)) block.appendChild(doc.importNode(child, true))
        state.paragraphXml.set(p.id, new XMLSerializer().serializeToString(block))
        p.text = ""
        state.saved = false
      })
    },
    insertParagraph(value: string, location: string) {
      if (location !== "After") throw new Error("Unsupported insertion location")
      queue.push(() => {
        const index = state.paragraphs.indexOf(p), oldId = p.id
        const oldXml = state.paragraphXml.get(oldId) ?? `<w:p><w:r><w:t>${escapeXml(p.text)}</w:t></w:r></w:p>`
        const stamp = 'w:author="Native user" w:date="2026-01-01T00:00:00Z"'
        state.paragraphInsertions++
        p.id = `retained-${state.paragraphInsertions}`
        state.paragraphs.splice(index + 1, 0, { id: oldId, text: value, styleBuiltIn: state.insertedParagraphStyle })
        state.paragraphXml.delete(oldId)
        const left = new DOMParser().parseFromString(flatOpc(oldXml), "application/xml").getElementsByTagNameNS(WORD_XML, "p")[0]!
        left.setAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "w14:paraId", "00000003")
        const properties = left.ownerDocument!.createElementNS(WORD_XML, "w:pPr")
        const marker = new DOMParser().parseFromString(flatOpc(`<w:p><w:pPr><w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr></w:p>`), "application/xml").getElementsByTagNameNS(WORD_XML, "rPr")[0]!
        properties.appendChild(left.ownerDocument!.importNode(marker, true))
        left.insertBefore(properties, left.firstChild)
        state.paragraphXml.set(p.id, new XMLSerializer().serializeToString(left))
        state.paragraphXml.set(oldId, `<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000001"><w:ins w:id="2" ${stamp}><w:r><w:t>${escapeXml(state.corruptParagraphInsertion ? value + "wrong" : value)}</w:t></w:r></w:ins></w:p>`)
        state.saved = false
      })
      return paragraph(p)
    },
    insertAnnotations(set: unknown) {
      const id = `annotation-${++state.annotationCounter}`
      state.annotations.set(id, set)
      return { value: [id] }
    },
  })
  const context = {
    document: {
      get saved() { return state.saved },
      get trackFormatting(): boolean { return state.trackFormatting as boolean },
      set trackFormatting(value: boolean) { queue.push(() => { state.trackFormatting = value }) },
      get changeTrackingMode() { return state.tracking },
      set changeTrackingMode(value: string) { state.modeChanges.push(value); queue.push(() => { state.tracking = value }) },
      load() {},
      getSelection: () => new Range(state.selection[0]!, state.selection[1]!),
      getParagraphByUniqueLocalId(id: string) { return paragraph(position(id).p) },
      getAnnotationById(id: string) { return { delete() { state.annotations.delete(id) } } },
      body: { get text() { return text() }, load() {}, get paragraphs() { return collection(state.paragraphs.map(paragraph)) }, getRange(location?: string) { const whole = state.terminalBodyMark && location !== "Content"; return new Range(0, text().length + (whole ? 1 : 0), whole, false, false, location === "Whole") } },
      onAnnotationClicked: event("annotation"), onParagraphChanged: event("change"), onParagraphAdded: event("add"), onParagraphDeleted: event("delete"),
    },
    async sync() {
      const writes = queue.splice(0)
      state.syncBatchSizes.push(writes.length)
      for (const write of writes) write()
      state.afterSync?.()
      if (state.throwAfterWrite && writes.length) throw state.syncFailure ?? new Error("lost native response")
    },
  } as WordChatContext
  const word = { async run<T>(first: unknown, second?: (context: WordChatContext) => Promise<T>) { return (second ?? first as (context: WordChatContext) => Promise<T>)(context) } } as WordChatWord
  const nativeHost = new WordChatHost(word, () => true, nativeText)
  const author = (input: Omit<WordChatApplyInput, "author">) => async (snapshot: { ooxml: string }): Promise<WordChatAuthoredAmendment> => {
    state.authorCalls++; await state.duringAuthor?.()
    if (input.formatting) return { ...await authorWordChatTarget(snapshot.ooxml, input.target, input.replacement, input.formatting), proposalId: input.proposalId }
    const replacements = input.replacement.split("\r"), revisions: WordChatAuthoredAmendment["revisions"] = []
    let counter = 0
    const paragraphs = input.target.paragraphs.map((p, index) => {
      const start = index === 0 ? input.target.startOffset : 0, end = index === input.target.paragraphs.length - 1 ? input.target.endOffset : p.text.length
      return p.text.slice(0, start) + replacements[index]! + p.text.slice(end)
    })
    const run = (value: string) => value ? `<w:r><w:t xml:space="preserve">${escapeXml(value)}</w:t></w:r>` : ""
    const markup = paragraphs.map((after, index) => {
      const before = input.target.paragraphs[index]!.text
      let at = 0, xml = ""
      for (const change of wordChanges(before, after)) {
        xml += run(before.slice(at, change.offset))
        if (change.replaced) { xml += `<w:del w:id="${++counter}" w:author="doc.haus"><w:r><w:delText xml:space="preserve">${escapeXml(change.replaced)}</w:delText></w:r></w:del>`; revisions.push({ author: "doc.haus", type: "Deleted", text: change.replaced }) }
        if (change.changed) { xml += `<w:ins w:id="${++counter}" w:author="doc.haus">${run(change.changed)}</w:ins>`; revisions.push({ author: "doc.haus", type: "Added", text: change.changed }) }
        at = change.offset + change.replaced.length
      }
      return `<w:p>${xml}${run(before.slice(at))}</w:p>`
    }).join("")
    return { proposalId: input.proposalId, author: "doc.haus", paragraphs, revisions, ooxml: flatOpc(markup).replace(/<w:sectPr>[\s\S]*?<\/w:sectPr>/, "") }
  }
  type TestInput = Omit<WordChatApplyInput, "author"> & { author?: WordChatApplyInput["author"] }
  const host = new Proxy(nativeHost, { get(target, key) {
    if (key === "apply") return (input: TestInput) => target.apply({ ...input, author: input.author ?? author(input) })
    const value = (target as unknown as Record<PropertyKey, unknown>)[key]
    return typeof value === "function" ? value.bind(target) : value
  } }) as Omit<WordChatHost, "apply"> & { apply(input: TestInput): ReturnType<WordChatHost["apply"]> }
  return { state, text, word, host, nativeHost, author,
    emit(name: string, args = {}) { for (const callback of state.handlers.get(name) ?? []) callback(args as never) },
  }
}

test("current context reads unsaved body without requiring a selection and empty selection stays empty", async () => {
  const { state, host } = office(["Unsaved wording.", "A different clause."])
  state.selection = [0,0]
  expect(await host.capture()).toMatchObject({text:"Unsaved wording.\rA different clause.", saved:false,truncated:false})
  expect(await host.selection()).toBeUndefined()
  expect(state.writes).toHaveLength(0)
  expect(state.selected).toBe(0)
})

test("scoped OOXML imports stop before authoring or mutation while document-wide numbering safety is unresolved", async () => {
  for (const kind of ["wording", "structural", "mixed", "nonfinal-formatting"] as const) {
    const { state, host } = office(["An opening paragraph.", "A retained paragraph."])
    state.selection = [0, state.paragraphs.map(p => p.text).join("\r").length]
    const target = (await host.selection())!
    const replacement = kind === "structural" ? `A new paragraph.\r${target.text}` :
      kind === "nonfinal-formatting" ? target.text : target.text.replace("opening", "revised")
    const formatting = kind === "mixed" || kind === "nonfinal-formatting" ? [{ paragraph: 0, spacingAfter: 6 }] : undefined
    await expect(host.apply({ proposalId: kind, target, replacement, formatting })).rejects.toThrow("directly in Word")
    expect(state.authorCalls).toBe(0)
    expect(state.writes).toHaveLength(0)
    expect(state.modeChanges).toEqual([])
  }
})

test("native text transport applies a single-paragraph wording change without importing OOXML", async () => {
  const duplicate = "A clause with a clear obligation."
  const { state, host, text } = office([duplicate, "Another clause."], true)
  state.selection = [duplicate.indexOf("clear"), duplicate.indexOf("clear") + "clear".length]
  state.revisions = [[duplicate.length + 1, duplicate.length + 1 + 7]]
  state.comments = [[duplicate.length + 1, duplicate.length + 7]]
  const target = (await host.selection())!
  const result = await host.apply({ proposalId: "native-wording", target, replacement: "specific" })
  expect(text()).toBe("A clause with a specific obligation.\rAnother clause.")
  expect(result).toMatchObject({ status: "applied", tracked: true, saved: false, authors: ["Word user"], target: { text: "specific", startOffset: duplicate.indexOf("clear"), endOffset: duplicate.indexOf("clear") + "specific".length } })
  expect(state.writes).toHaveLength(1)
  expect(state.writes[0]).toMatchObject({ text: "specific", tracking: "TrackAll", start: duplicate.indexOf("clear") })
  expect(state.modeChanges).toEqual(["TrackAll", "Off"])
  expect(state.comments).toHaveLength(1)
  expect(state.revisions).toHaveLength(1)
  expect(state.nativeRevisions.map(revision => [revision.type, revision.author, revision.text])).toEqual([
    ["Deleted", "Word user", "clear"], ["Added", "Word user", "specific"],
  ])
})

test("native text transport refuses a target that changed after review before authoring", async () => {
  const { state, host } = office(["A clause with a clear obligation."], true)
  state.selection = [15, 20]
  const target = (await host.selection())!
  state.paragraphs[0]!.text = "A clause with a changed obligation."
  await expect(host.apply({ proposalId: "stale", target, replacement: "specific" })).rejects.toThrow("could not match reviewed paragraph")
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toEqual([])
})

test("current document context beyond 96k retains a selected later passage without truncation", async () => {
  const paragraphs = Array.from({length:6},(_,index) => `Section ${index}: ${"Ordinary current wording. ".repeat(800)}`)
  paragraphs.push("A later selected clause.")
  const {state,host,text} = office(paragraphs)
  const at = text().indexOf("later selected")
  state.selection = [at,at + "later selected".length]
  const context = await host.capture(true)
  expect(context.text.length).toBeGreaterThan(96_000)
  expect(context.text).toBe(text())
  expect(context.truncated).toBe(false)
  expect(context.selection?.text).toBe("later selected")
  expect(context.selection?.paragraphs.map(p => p.id)).toEqual(["paragraph-6"])
  expect(state.writes).toHaveLength(0)
})

test("over-limit current documents refuse clearly without sending a silently truncated buffer", async () => {
  const {host,state} = office(Array.from({length:26},() => "Ordinary text. ".repeat(1500)))
  await expect(host.capture()).rejects.toThrow("512,000 characters")
  expect(state.writes).toHaveLength(0)
})

// Import-path model tests remain here for restoration after the numbering safety gate is lifted.
// They are intentionally skipped: they cannot exercise the disabled host route.
test.skip("paragraph identity targets a repeated clause without touching the other copy", async () => {
  const duplicate = "A clear obligation applies."
  const { state, host, text } = office([duplicate, duplicate])
  state.selection = [duplicate.length + 1, duplicate.length * 2 + 1]
  const target = (await host.selection())!
  const result = await host.apply({ proposalId: "reviewed-1", target, replacement: duplicate.replace("clear", "specific") })
  expect(text()).toBe(duplicate + "\rA specific obligation applies.")
  expect(result).toMatchObject({ status: "applied", tracked: true, saved: false, authors: ["doc.haus"] })
  expect(state.writes.map((w) => [w.before, w.text, w.tracking])).toEqual([[duplicate, "A specific obligation applies.", "Off"]])
  expect(state.tracking).toBe("Off")
})

test.skip("multiline authored markup uses one scoped import and restores TrackMineOnly", async () => {
  const { state, host, text } = office(["The first obligation is clear.", "The second obligation is clear."])
  state.selection = [0, text().length]
  state.tracking = "TrackMineOnly"
  const target = (await host.selection())!
  await host.apply({ proposalId: "multiline", target, replacement: target.text.replaceAll("clear", "specific") })
  expect(text()).toBe("The first obligation is specific.\rThe second obligation is specific.")
  expect(state.writes).toHaveLength(1)
  expect(state.writes[0]!.text).toBe(text())
  expect(state.modeChanges).toEqual(["Off", "TrackMineOnly"])
  expect(state.tracking).toBe("TrackMineOnly")
})

test.skip("real shared authoring imports reviewed whole-paragraph insertion/removal without touching outside unsaved text", async () => {
  for (const kind of ["insert", "delete"] as const) {
    const texts = ["Unsaved preceding clause.", "Working note Tuesday.", "Unrelated BLUE FERN.", "Unsaved following clause."]
    const { state, host, word } = office(texts)
    state.selection = [texts[0]!.length + 1, texts.slice(0, 3).join("\r").length]
    state.tracking = "TrackAll"
    const selected = (await host.selection())!, originalXml = '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Working note Tuesday.</w:t></w:r></w:p>'
    state.paragraphXml.set("paragraph-1", originalXml)
    const replacement = kind === "insert" ? "Working notes — non-operative — retain, move or delete:\r" + selected.text : texts[2]!
    const author: WordChatApplyInput["author"] = async snapshot => {
      const authored = await authorWordChatTarget(snapshot.ooxml, selected, replacement)
      // A cursor move and unrelated edits during preparation do not retarget
      // the reviewed clause or get replaced by an older whole document.
      state.selection = [0, 5]
      state.paragraphs[3]!.text += " Local drafting."
      return { ...authored, proposalId: kind }
    }
    const result = await host.apply({ proposalId: kind, target: selected, replacement, author })
    expect(result.status).toBe("applied")
    expect(state.writes).toHaveLength(1)
    expect(state.writes[0]!.before).toBe(selected.text)
    expect(state.tracking).toBe("TrackAll")
    const current = await word.run(async context => context.document.body.getRange().getReviewedText("Current").value)
    expect(current).toBe(texts[0] + "\r" + replacement + "\r" + texts[3] + " Local drafting.")
    expect(state.paragraphs[0]!.id).toBe("paragraph-0")
    expect(state.paragraphs.at(-1)!.id).toBe("paragraph-3")
    expect(state.nativeRevisions.every(change => change.author === "doc.haus")).toBe(true)
    expect(state.nativeRevisions.some(change => change.text === "\r" && change.type === (kind === "insert" ? "Added" : "Deleted"))).toBe(true)
    expect([...state.paragraphXml.values()].some(xml => xml.includes('<w:ilvl w:val="1"') && xml.includes("<w:b"))).toBe(true)
    expect(state.saved).toBe(false)
    await host.apply({ proposalId: kind, target: selected, replacement, author })
    expect(state.writes).toHaveLength(1)
  }
})

test.skip("structural import still refuses a formatting race before native mutation", async () => {
  const { state, host } = office(["Working note Tuesday.", "Unrelated BLUE FERN."])
  state.selection = [0, state.paragraphs.map(p => p.text).join("\r").length]
  const target = (await host.selection())!, replacement = "A non-operative heading:\r" + target.text
  const author: WordChatApplyInput["author"] = async snapshot => {
    const authored = await authorWordChatTarget(snapshot.ooxml, target, replacement)
    state.paragraphXml.set("paragraph-0", '<w:p><w:pPr><w:spacing w:after="240"/></w:pPr><w:r><w:t>Working note Tuesday.</w:t></w:r></w:p>')
    return { ...authored, proposalId: "format-race" }
  }
  await expect(host.apply({ proposalId: "format-race", target, replacement, author })).rejects.toThrow("formatting changed")
  expect(state.writes).toHaveLength(0)
})

function formattedOffice() {
  const texts = ["Unsaved preceding clause.", "The obligation is ten days.", "Retained qualification.", "Unchanged selected boundary paragraph.", "Unrelated BLUE FERN."]
  const fixture = office(texts), { state } = fixture
  const start = texts[0]!.length + 1, revision = start + "The obligation is ".length
  state.selection = [start, texts.slice(0, 4).join("\r").length]
  state.tracking = "TrackAll"
  const paragraph = (text: string, properties: string, run: string) => `<w:p><w:pPr>${properties}</w:pPr><w:r><w:rPr>${run}</w:rPr><w:t>${escapeXml(text)}</w:t></w:r></w:p>`
  state.paragraphXml.set("paragraph-0", paragraph(texts[0]!, '<w:spacing w:before="40"/>', '<w:i/>'))
  state.paragraphXml.set("paragraph-1", '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>The obligation is </w:t></w:r><w:del w:id="42" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>five</w:delText></w:r></w:del><w:ins w:id="43" w:author="Earlier reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t>ten</w:t></w:r></w:ins><w:r><w:t> days.</w:t></w:r></w:p>')
  state.paragraphXml.set("paragraph-2", paragraph(texts[2]!, '<w:keepNext/>', '<w:i/>'))
  state.paragraphXml.set("paragraph-3", paragraph(texts[3]!, '<w:spacing w:after="100"/>', '<w:i/>'))
  state.paragraphXml.set("paragraph-4", paragraph(texts[4]!, '<w:spacing w:after="180"/>', '<w:b/><w:i/>'))
  state.nativeRevisions = [
    { start: revision, end: revision, author: "Earlier reviewer", type: "Deleted", text: "five" },
    { start: revision, end: revision + 3, author: "Earlier reviewer", type: "Added", text: "ten" },
  ]
  const O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  state.extraParts = `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData><w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${O}/styles" Target="styles.xml"/></Relationships></pkg:xmlData></pkg:part>`
  const formatting: WordChatParagraphFormat[] = [{ paragraph: 0, style: "Heading1", spacingBefore: 12, spacingAfter: 8, lineSpacing: 1.25 }, { paragraph: 1, spacingAfter: 6 }]
  return { ...fixture, texts, formatting }
}

test.skip("real shared paragraph formatting preserves an unchanged final paragraph, prior text revisions and adjacent unsaved formatting", async () => {
  for (const emptyText of [false, true]) {
    const { state, host, text, texts, formatting } = formattedOffice()
    state.formattedTextEmpty = emptyText
    const target = (await host.selection())!, outside = [state.paragraphXml.get("paragraph-0"), state.paragraphXml.get("paragraph-4")]
    state.duringAuthor = () => { state.selection = [0, 4] }
    const input = { proposalId: `format-${emptyText}`, target, replacement: target.text, formatting }
    expect((await host.apply(input)).status).toBe("applied")
    expect(state.writes).toHaveLength(1)
    expect(state.writes[0]).toMatchObject({ before: target.text + "\r", text: target.text + "\r", start: texts[0]!.length + 1, end: texts.slice(0, 4).join("\r").length + 1, tracking: "Off" })
    expect(text()).toBe(texts.join("\r"))
    expect(state.paragraphs.map(p => p.id)).toEqual(["paragraph-0", "paragraph-1", "paragraph-2", "paragraph-3", "paragraph-4"])
    expect([state.paragraphXml.get("paragraph-0"), state.paragraphXml.get("paragraph-4")]).toEqual(outside)
    expect(state.nativeRevisions.filter(r => r.author === "Earlier reviewer").map(({ author, type, text }) => ({ author, type, text }))).toEqual([
      { author: "Earlier reviewer", type: "Deleted", text: "five" }, { author: "Earlier reviewer", type: "Added", text: "ten" },
    ])
    expect(state.nativeRevisions.filter(r => r.author === "doc.haus").map(r => r.type)).toEqual(["Formatted", "Formatted"])
    const last = new DOMParser().parseFromString(state.paragraphXml.get("paragraph-2")!, "application/xml")
    expect(last.getElementsByTagNameNS(WORD_XML, "spacing")[0]!.getAttributeNS(WORD_XML, "after")).toBe("120")
    expect(last.getElementsByTagNameNS(WORD_XML, "keepNext")).toHaveLength(2) // current and original properties
    const boundary = new DOMParser().parseFromString(state.paragraphXml.get("paragraph-3")!, "application/xml")
    expect(boundary.getElementsByTagNameNS(WORD_XML, "spacing")[0]!.getAttributeNS(WORD_XML, "after")).toBe("100")
    expect(boundary.getElementsByTagNameNS(WORD_XML, "pPrChange")).toHaveLength(0)
    expect(state.tracking).toBe("TrackAll")
    expect(state.saved).toBe(false)
    await host.apply(input)
    expect(state.writes).toHaveLength(1)
  }
})

test.skip("paragraph formatting refuses a target race before mutation and fences uncertain adjacent formatting changes", async () => {
  for (const race of [true, false]) {
    const { state, host, formatting } = formattedOffice()
    const target = (await host.selection())!
    if (race) state.duringAuthor = () => { state.paragraphXml.set("paragraph-2", state.paragraphXml.get("paragraph-2")!.replace('<w:keepNext/>', '<w:keepNext/><w:spacing w:before="60"/>')) }
    else state.leakAdjacentFormatting = true
    const input = { proposalId: `format-boundary-${race}`, target, replacement: target.text, formatting }
    await expect(host.apply(input)).rejects.toThrow(race ? "formatting changed" : "could not confirm")
    expect(state.writes).toHaveLength(race ? 0 : 1)
    if (race) expect(state.modeChanges).toHaveLength(0)
    else {
      await expect(host.apply(input)).rejects.toThrow("may already contain")
      expect(state.writes).toHaveLength(1)
    }
  }
})

test("native properties track the final selected paragraph without replacing the range", async () => {
  const { state, host, text, texts } = formattedOffice()
  const target = (await host.selection())!, formatting = [{ paragraph: target.paragraphs.length - 1, spacingAfter: 6 }]
  const outside = [state.paragraphXml.get("paragraph-0"), state.paragraphXml.get("paragraph-4")]
  expect(validWordChatFormatting(target, formatting)).toBe(true)
  expect(await host.apply({ proposalId: "final-paragraph-format", target, replacement: target.text, formatting })).toMatchObject({ status: "applied", authors: ["Word user"] })
  expect(state.authorCalls).toBe(1)
  expect(state.writes).toHaveLength(0)
  expect(state.nativeFormatting).toEqual(["paragraph-3"])
  expect(state.modeChanges).toEqual(["TrackAll", "TrackAll"])
  expect([state.paragraphXml.get("paragraph-0"), state.paragraphXml.get("paragraph-4")]).toEqual(outside)
  expect(state.nativeRevisions.filter(r => r.author === "Earlier reviewer")).toHaveLength(2)
  expect(text()).toBe(texts.join("\r"))
})

test("native formatting restores the original tracking mode in the same mutation flush", async () => {
  for (const mode of ["Off", "TrackAll"]) {
    const { state, host } = formattedOffice()
    state.tracking = mode
    const target = (await host.selection())!
    const result = await host.apply({ proposalId: `format-flush-${mode}`, target, replacement: target.text,
      formatting: [{ paragraph: target.paragraphs.length - 1, spacingAfter: 6 }] })
    expect(result.status).toBe("applied")
    expect(state.tracking).toBe(mode)
    expect(state.modeChanges).toEqual(["TrackAll", mode])
    expect(state.syncBatchSizes.filter(size => size > 0)).toEqual([3])
  }
})

test("native formatting projects the body when an empty neighbor has no Content XML", async () => {
  const { state, host } = office(["", "Final paragraph."], true)
  state.selection = [1, 17]
  state.refuseEmptyContentXml = true
  state.paragraphXml.set("paragraph-1", '<w:p><w:pPr><w:spacing w:after="200"/></w:pPr><w:r><w:t>Final paragraph.</w:t></w:r></w:p>')
  const target = (await host.selection())!
  const result = await host.apply({ proposalId:"format-empty-neighbor", target, replacement:target.text, formatting:[{paragraph:0,spacingAfter:18}] })
  expect(result.status).toBe("applied")
  expect(state.nativeFormatting).toEqual(["paragraph-1"])
  expect(state.paragraphXml.get("paragraph-1")).toContain('w:after="360"')
})

test("native formatting permits only matched neighbouring revision ID reassignment", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  const old = '<w:p><w:pPr><w:spacing w:after="180"/></w:pPr><w:del w:id="31" w:author="Earlier" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Old </w:delText></w:r></w:del><w:ins w:id="32" w:author="Earlier" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Next </w:t></w:r></w:ins><w:r><w:t>paragraph.</w:t></w:r></w:p>'
  state.paragraphs[4]!.text = "Next paragraph."
  state.paragraphXml.set("paragraph-4", old)
  state.afterSync = () => { if (state.nativeFormatting.length) { state.paragraphXml.set("paragraph-4", old.replace('w:id="31"','w:id="41"').replace('w:id="32"','w:id="42"')); state.afterSync = undefined } }
  expect((await host.apply({ proposalId:"format-neighbor-renumber", target, replacement:target.text,
    formatting:[{paragraph:target.paragraphs.length - 1,spacingAfter:6}] })).status).toBe("applied")
  expect(state.nativeFormatting).toHaveLength(1)
})

test("projected native formatting rejects a neighboring paragraph property change", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  state.afterSync = () => {
    if (state.nativeFormatting.length) {
      state.paragraphXml.set("paragraph-4", state.paragraphXml.get("paragraph-4")!.replace('w:after="180"', 'w:after="240"'))
      state.afterSync = undefined
    }
  }
  const input = { proposalId:"format-neighbor-properties", target, replacement:target.text,
    formatting:[{paragraph:target.paragraphs.length - 1,spacingAfter:6}] }
  let failure: unknown
  try { await host.apply(input) } catch (error) { failure = error }
  expect(failure).toMatchObject({mutationState:"uncertain", cause: { mutationState:"uncertain", cause: { message:"format-neighbours" } }})
  expect(state.nativeFormatting).toHaveLength(1)
  await expect(host.apply(input)).rejects.toMatchObject({mutationState:"uncertain", message:expect.stringContaining("may already contain")})
  expect(state.nativeFormatting).toHaveLength(1)
})

test("projected native formatting rejects a missing property revision", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  state.afterSync = () => {
    if (state.nativeFormatting.length) {
      state.paragraphXml.set("paragraph-3", state.paragraphXml.get("paragraph-3")!.replace(/<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/, ""))
      state.afterSync = undefined
    }
  }
  let failure: unknown
  try { await host.apply({ proposalId:"format-missing-carrier", target, replacement:target.text,
    formatting:[{paragraph:target.paragraphs.length - 1,spacingAfter:6}] }) } catch (error) { failure = error }
  expect(failure).toMatchObject({mutationState:"uncertain", cause: { mutationState:"uncertain", cause: { message:"format-authored-verification", cause: expect.any(Error) } }})
  expect(state.nativeFormatting).toHaveLength(1)
})

test("unknown formatting tracking refuses native formatting writes", async () => {
  for (const enabled of [undefined]) {
    const { state, host } = formattedOffice()
    state.trackFormatting = enabled
    const target = (await host.selection())!
    await expect(host.apply({ proposalId: `format-tracking-${enabled}`, target, replacement: target.text,
      formatting: [{ paragraph: target.paragraphs.length - 1, spacingAfter: 6 }] })).rejects.toMatchObject({
      mutationState: "unchanged", message: expect.stringContaining("cannot confirm formatting tracking"),
    })
    expect(state.modeChanges).toEqual([])
    expect(state.nativeFormatting).toEqual([])
    expect(state.syncBatchSizes.filter(size => size > 0)).toEqual([])
  }
})

test("disabled formatting tracking is temporarily enabled and restored in the fresh write", async () => {
 for (const disableDuringAuthor of [false, true]) {
  const { state, host } = formattedOffice()
  if (!disableDuringAuthor) state.trackFormatting = false
  const target = (await host.selection())!
  if (disableDuringAuthor) state.duringAuthor = () => { state.trackFormatting = false }
  expect((await host.apply({ proposalId: `format-tracking-fresh-${disableDuringAuthor}`, target, replacement: target.text,
    formatting: [{ paragraph: target.paragraphs.length - 1, spacingAfter: 6 }] })).status).toBe("applied")
  expect(state.authorCalls).toBe(1)
  expect(state.trackFormatting).toBe(false)
  expect(state.modeChanges).toEqual(["TrackAll", "TrackAll"])
  expect(state.nativeFormatting).toEqual(["paragraph-3"])
  expect(state.syncBatchSizes.filter(size => size > 0)).toEqual([5])
 }
})

test("native final-boundary formatting verifies each requested property and prior revisions", async () => {
  const { state, host, formatting } = formattedOffice(), target = (await host.selection())!
  const operations = [...formatting, { paragraph: 2, spacingAfter: 6 }]
  expect((await host.apply({ proposalId: "native-all-properties", target, replacement: target.text, formatting: operations })).status).toBe("applied")
  expect(state.nativeFormatting).toEqual(["paragraph-1", "paragraph-2", "paragraph-3"])
  expect(state.nativeRevisions.filter(r => r.author === "Earlier reviewer")).toHaveLength(2)
  expect(state.writes).toHaveLength(0)
})

test.skip("mixed inserted heading and wording retain prior inline revisions before native formatting", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  const replacement = "Scope of obligations\r" + target.text.replace("Retained qualification.", "Retained important qualification.")
  const formatting: WordChatParagraphFormat[] = [{ paragraph: 0, style: "Heading1" }, { paragraph: 3, spacingAfter: 6 }]
  const result = await host.apply({ proposalId: "mixed-insert-format", target, replacement, formatting })
  expect(result.status).toBe("applied")
  expect(state.writes).toHaveLength(1)
  expect(state.nativeRevisions.filter(r => r.author === "Earlier reviewer")).toHaveLength(2)
  expect(state.nativeRevisions.filter(r => r.type === "Formatted")).toHaveLength(1)
  expect(state.nativeFormatting).toEqual(["new-0-3"])
  const heading = new DOMParser().parseFromString(state.paragraphXml.get("paragraph-1")!, "application/xml")
  expect(heading.getElementsByTagNameNS(WORD_XML, "pStyle")[0]!.getAttributeNS(WORD_XML, "val")).toBe("Heading1")
  expect(heading.getElementsByTagNameNS(WORD_XML, "pPrChange")).toHaveLength(0)
})

test.skip("inserted-only formatting is carried in reviewed wording without native setters", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  const replacement = "Scope of obligations\r" + target.text
  const input = { proposalId: "inserted-only-format", target, replacement, formatting: [{ paragraph: 0, style: "Heading1" as const }] }
  expect((await host.apply(input)).status).toBe("applied")
  expect(state.writes).toHaveLength(1)
  expect(state.nativeFormatting).toHaveLength(0)
  expect(state.nativeRevisions.filter(r => r.type === "Formatted")).toHaveLength(0)
  const heading = new DOMParser().parseFromString(state.paragraphXml.get("paragraph-1")!, "application/xml")
  expect(heading.getElementsByTagNameNS(WORD_XML, "pStyle")[0]!.getAttributeNS(WORD_XML, "val")).toBe("Heading1")
  expect(heading.getElementsByTagNameNS(WORD_XML, "pPrChange")).toHaveLength(0)
})

test.skip("mixed wording beside an existing own revision preserves that exact revision", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  state.paragraphXml.set("paragraph-1", state.paragraphXml.get("paragraph-1")!.replaceAll('w:author="Earlier reviewer"', 'w:author="doc.haus"'))
  for (const revision of state.nativeRevisions) revision.author = "doc.haus"
  const previous = state.paragraphXml.get("paragraph-1")!
  const replacement = "Scope of obligations\r" + target.text.replace("The obligation is", "The duty is")
  const formatting: WordChatParagraphFormat[] = [{ paragraph: 0, style: "Heading1" }, { paragraph: 3, spacingAfter: 6 }]
  const result = await host.apply({ proposalId: "mixed-existing-own", target, replacement, formatting })
  expect(result.status).toBe("applied")
  expect(state.writes).toHaveLength(1)
  expect(state.nativeRevisions.some(r => r.author === "doc.haus" && r.type === "Deleted" && r.text === "five")).toBe(true)
  expect(state.nativeRevisions.some(r => r.author === "doc.haus" && r.type === "Added" && r.text === "ten")).toBe(true)
  expect(previous).toContain('w:author="doc.haus"')
})

test.skip("mixed wording in a paragraph with another author's revision refuses before mutation", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  const replacement = "Scope of obligations\r" + target.text.replace("The obligation is", "The duty is")
  await expect(host.apply({ proposalId: "mixed-existing-other", target, replacement, formatting: [{ paragraph: 0, style: "Heading1" }] })).rejects.toMatchObject({ mutationState: "unchanged" })
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})

test.skip("mixed artifact tampering refuses before mutation and a lost import cannot be replayed", async () => {
  for (const tamper of ["wording", "prior", "none"] as const) {
    const { state, host, author } = formattedOffice(), target = (await host.selection())!
    const replacement = "Scope of obligations\r" + target.text.replace("Retained qualification.", "Retained important qualification.")
    const input = { proposalId: `mixed-failure-${tamper}`, target, replacement, formatting: [{ paragraph: 0, style: "Heading1" as const }, { paragraph: 3, spacingAfter: 6 }] }
    if (tamper !== "none") {
      const write = author(input)
      await expect(host.apply({ ...input, author: async snapshot => {
        const result = await write(snapshot)
        return { ...result, wording: { ...result.wording!, ooxml: result.wording!.ooxml.replace(tamper === "prior" ? "five" : "qualification", "unexpected") } }
      } })).rejects.toMatchObject({ mutationState: "unchanged" })
      expect(state.writes).toHaveLength(0)
    } else {
      state.throwAfterWrite = true
      await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
      expect(state.writes).toHaveLength(1)
      expect(state.tracking).toBe("TrackAll")
      await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
      expect(state.writes).toHaveLength(1)
    }
  }
})

test.skip("unreviewed inserted or retained paragraph properties refuse before native mutation", async () => {
  for (const tamper of ["baseline-indent", "retained-indent", "heading-style"] as const) {
    const { state, host, author } = formattedOffice(), target = (await host.selection())!
    const replacement = "Scope of obligations\r" + target.text.replace("Retained qualification.", "Retained important qualification.")
    const input = { proposalId: `properties-${tamper}`, target, replacement, formatting: [{ paragraph: 0, style: "Heading1" as const }, { paragraph: 3, spacingAfter: 6 }] }
    const write = author(input)
    await expect(host.apply({ ...input, author: async snapshot => {
      const result = await write(snapshot), wording = result.wording!
      const key = tamper === "baseline-indent" ? "unformattedOoxml" : "ooxml"
      const doc = new DOMParser().parseFromString(wording[key]!, "application/xml")
      const body = doc.getElementsByTagNameNS(WORD_XML, "body")[0]!
      const paragraphs = Array.from(body.childNodes).filter((node): node is Element => node.nodeType === 1 && (node as Element).localName === "p")
      const properties = paragraphs[tamper === "retained-indent" ? 1 : 0]!.getElementsByTagNameNS(WORD_XML, "pPr")[0]!
      if (tamper === "heading-style") properties.getElementsByTagNameNS(WORD_XML, "pStyle")[0]!.setAttributeNS(WORD_XML, "w:val", "Heading2")
      else {
        const ind = doc.createElementNS(WORD_XML, "w:ind")
        ind.setAttributeNS(WORD_XML, "w:left", "300")
        properties.insertBefore(ind, properties.firstChild)
      }
      return { ...result, wording: { ...wording, [key]: new XMLSerializer().serializeToString(doc) } }
    } })).rejects.toMatchObject({ mutationState: "unchanged" })
    expect(state.writes).toHaveLength(0)
    expect(state.modeChanges).toHaveLength(0)
  }
})

test("a missing native format revision is uncertain and cannot be replayed", async () => {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  state.dropNativeFormatting = true
  const input = { proposalId: "missing-native-format", target, replacement: target.text, formatting: [{ paragraph: 2, spacingAfter: 6 }] }
  let caught: unknown
  try { await host.apply(input) } catch (error) { caught = error }
  expect(caught).toMatchObject({ mutationState: "uncertain", message: expect.stringContaining("could not confirm") })
  const inner = (caught as WordChatHostError).cause as WordChatHostError
  expect(inner).toBeInstanceOf(WordChatHostError)
  expect(inner.cause).toBeInstanceOf(Error)
  expect((inner.cause as Error).message).toBeTruthy()
  expect(state.tracking).toBe("TrackAll")
  const batches = state.syncBatchSizes.filter(size => size > 0)
  expect(batches).toHaveLength(1)
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
  expect(state.syncBatchSizes.filter(size => size > 0)).toEqual(batches)
  expect(state.writes).toHaveLength(0)
})

test("a lost native formatting response prevents replay without a compensating write", async () => {
 for (const enabled of [true, false]) {
  const { state, host } = formattedOffice(), target = (await host.selection())!
  state.tracking = "Off"
  state.trackFormatting = enabled
  state.throwAfterWrite = true
  const input = { proposalId: "lost-native-format", target, replacement: target.text, formatting: [{ paragraph: 2, spacingAfter: 6 }] }
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
  expect(state.tracking).toBe("Off")
  expect(state.trackFormatting).toBe(enabled)
  expect(state.nativeFormatting).toEqual(["paragraph-3"])
  expect(state.syncBatchSizes.filter(size => size > 0)).toEqual([enabled ? 3 : 5])
  expect(state.modeChanges).toEqual(["TrackAll", "Off"])
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
  expect(state.nativeFormatting).toHaveLength(1)
 }
})

test.skip("formatting captures and verifies Content XML when Whole export adds a synthetic final paragraph", async () => {
  const { state, host, word, formatting, texts, text } = formattedOffice()
  state.wholeOoxmlTrailingParagraph = true
  const target = (await host.selection())!, outside = [state.paragraphXml.get("paragraph-0"), state.paragraphXml.get("paragraph-4")]
  const counts = await word.run(async context => {
    const first = context.document.getParagraphByUniqueLocalId(target.paragraphs[0]!.id), last = context.document.getParagraphByUniqueLocalId(target.paragraphs.at(-1)!.id)
    return ["Content", "Whole"].map(location => {
      const xml = first.getRange(location).expandTo(last.getRange(location)).getOoxml().value
      return new DOMParser().parseFromString(xml, "application/xml").getElementsByTagNameNS(WORD_XML, "body")[0]!.getElementsByTagNameNS(WORD_XML, "p").length
    })
  })
  expect(counts).toEqual([3, 4])
  expect((await host.apply({ proposalId: "whole-export-final-paragraph", target, replacement: target.text, formatting })).status).toBe("applied")
  expect(state.authorCalls).toBe(1)
  expect(state.writes).toHaveLength(1)
  expect(state.writes[0]!.before).toBe(target.text + "\r")
  expect(state.writes[0]!.end).toBe(texts.slice(0, 4).join("\r").length + 1)
  expect(state.paragraphs.map(p => p.id)).toEqual(["paragraph-0", "paragraph-1", "paragraph-2", "paragraph-3", "paragraph-4"])
  expect(text()).toBe(texts.join("\r"))
  expect([state.paragraphXml.get("paragraph-0"), state.paragraphXml.get("paragraph-4")]).toEqual(outside)
  expect(state.nativeRevisions.filter(r => r.author === "Earlier reviewer")).toHaveLength(2)
  expect(state.nativeRevisions.filter(r => r.type === "Formatted")).toHaveLength(2)
  expect(state.tracking).toBe("TrackAll")
  expect(state.saved).toBe(false)
})

test("paragraph formatting refuses an adjacent insertion between neighbour discovery and the final capture", async () => {
  const { state, host, word, formatting: baseFormatting } = formattedOffice()
  const target = (await host.selection())!
  const formatting = [...baseFormatting, { paragraph: target.paragraphs.length - 1, spacingAfter: 6 }]
  let inserted = false
  await word.run(async context => {
    const body = context.document.body, original = Object.getOwnPropertyDescriptor(body, "paragraphs")!.get!
    Object.defineProperty(body, "paragraphs", { get() {
      let items = original.call(body).items, loads = 0
      return { get items() { return items }, load() {
        // Emulate Word's reloaded collection after a coauthor inserts the new
        // following paragraph, after the earlier neighbours were selected.
        if (++loads === 2 && !inserted) {
          state.paragraphs.splice(4, 0, { id: "new-adjacent", text: "A coauthor’s unsaved paragraph." })
          inserted = true
        }
        items = original.call(body).items
      } }
    } })
  })
  await expect(host.apply({ proposalId: "neighbour-capture-race", target, replacement: target.text, formatting })).rejects.toThrow("paragraphs beside this review changed")
  expect(inserted).toBe(true)
  expect(state.authorCalls).toBe(0)
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
  expect(state.paragraphs[4]!.text).toBe("A coauthor’s unsaved paragraph.")
})

test("formatting body capture refuses an extra paragraph or different exported text before authoring", async () => {
  for (const extraParagraph of [true, false]) {
    const { state, host, formatting: baseFormatting } = formattedOffice()
    const target = (await host.selection())!
    const formatting = [...baseFormatting, { paragraph: target.paragraphs.length - 1, spacingAfter: 6 }]
    const body = state.paragraphs.map(p => state.paragraphXml.get(p.id)!).join("")
    state.xml = flatOpc(extraParagraph ? body + '<w:p><w:r><w:t>Extra</w:t></w:r></w:p><w:sectPr/>' : body.replace("The obligation is ", "The commitment is "), state.extraParts)
    await expect(host.apply({ proposalId: `capture-diagnostic-${extraParagraph}`, target, replacement: target.text, formatting })).rejects.toMatchObject({mutationState:"unchanged", message:expect.stringContaining("could not bind the exported body")})
    expect(state.authorCalls).toBe(0)
    expect(state.writes).toHaveLength(0)
    expect(state.modeChanges).toHaveLength(0)
  }
})

test.skip("fresh paragraph proof confirms multiline edits despite returned insertion-range drift", async () => {
  const { state, host, text } = office(["The supplier will promptly identify an issue.", "The recipient will promptly consider the issue."])
  state.selection = [0, text().length]
  state.driftInsertedRanges = true
  const target = (await host.selection())!
  const result = await host.apply({ proposalId: "native-returned-range-drift", target, replacement: target.text.replaceAll("promptly", "within a reasonable time") })
  expect(result.status).toBe("applied")
  expect(text()).toBe(target.text.replaceAll("promptly", "within a reasonable time"))
  expect(state.nativeRevisions.map(change => change.type).sort()).toEqual(["Added", "Added", "Deleted", "Deleted"])
  expect(state.tracking).toBe("Off")
})

test.skip("native proof covers deletion-only and replaced spans without inserted-range assumptions", async () => {
  for (const replacement of ["A clause with an obligation.", "A clause with a specific duty."]) {
    const { state, host } = office()
    const target = (await host.selection())!
    expect((await host.apply({ proposalId: replacement, target, replacement })).status).toBe("applied")
    expect(state.paragraphs[0]!.text).toBe(replacement)
    expect(state.nativeRevisions.some(change => change.type === "Deleted")).toBe(true)
  }
})

test.skip("two separated edits import in document order and preserve revisions in other paragraphs", async () => {
  const { state, host, word } = office(["A clear obligation preserves the agreed commercial position and a firm deadline.", "Earlier revision remains."])
  const priorStart = state.paragraphs[0]!.text.length + 1
  state.nativeRevisions = [
    { start: priorStart, end: priorStart + 7, author: "Earlier reviewer", type: "Added", text: "Earlier" },
    { start: priorStart + 8, end: priorStart + 8, author: "Earlier reviewer", type: "Deleted", text: "old wording " },
  ]
  const target = (await host.selection())!
  const replacement = target.text.replace("clear", "specific").replace("firm", "fixed")
  expect(wordChanges(target.text, replacement).map(edit => [edit.replaced, edit.changed])).toEqual([["clear", "specific"], ["firm", "fixed"]])
  expect((await host.apply({ proposalId: "two-separated-edits", target, replacement })).status).toBe("applied")
  expect(state.writes.map(write => [write.before, write.text])).toEqual([[target.text, replacement]])
  expect(state.nativeRevisions.filter(change => change.author === "Earlier reviewer").map(({ author, type, text }) => ({ author, type, text }))).toEqual([
    { author: "Earlier reviewer", type: "Added", text: "Earlier" }, { author: "Earlier reviewer", type: "Deleted", text: "old wording " },
  ])
  const changes = await word.run(async context => {
    const collection = context.document.body.getRange().getTrackedChanges()
    collection.load("items/author,items/type,items/text")
    await context.sync()
    return collection.items.filter(change => change.author === "doc.haus").map(change => [change.type, change.text])
  })
  expect(changes).toEqual([["Deleted", "clear"], ["Added", "specific"], ["Deleted", "firm"], ["Added", "fixed"]])
  expect(state.paragraphs[0]!.text).toBe(replacement)
})

test.skip("preexisting revisions elsewhere cannot masquerade as new authored markup", async () => {
  const { state, host } = office(["A clear obligation.", "specific clear."])
  const priorStart = state.paragraphs[0]!.text.length + 1
  state.nativeRevisions = [
    { start: priorStart, end: priorStart + 8, author: "doc.haus", type: "Added", text: "specific" },
    { start: priorStart + 9, end: priorStart + 9, author: "doc.haus", type: "Deleted", text: "clear" },
  ]
  state.omitNewRevisions = true
  const target = (await host.selection())!
  const input = { proposalId: "preexisting-is-not-proof", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain", message: expect.stringContaining("new revision signatures") })
  expect(state.writes).toHaveLength(1)
  await expect(host.apply(input)).rejects.toThrow("may already contain")
})

test.skip("edits outside the target are preserved; changed target paragraphs refuse before writing", async () => {
  const { state, host } = office(["A clause with a clear obligation.", "The first qualification.", "The second qualification.", "A distant clause."])
  const target = (await host.selection())!
  state.paragraphs[3]!.text += " Unsaved local drafting."
  await host.apply({ proposalId: "outside", target, replacement: target.text.replace("clear", "specific") })
  expect(state.paragraphs[3]!.text).toEndWith("Unsaved local drafting.")
  const changed = office()
  const stale = (await changed.host.selection())!
  changed.state.paragraphs[0]!.text += " Changed qualification."
  await expect(changed.host.apply({ proposalId: "stale", target: stale, replacement: "Different wording." })).rejects.toThrow("could not match reviewed paragraph 1")
  expect(changed.state.writes).toHaveLength(0)
  expect(changed.state.modeChanges).toHaveLength(0)
})

test.skip("complete paragraph delivery keeps revision authority inside the exact selection", async () => {
  const { state, host, text } = office(["Unchanged bold prefix. A clear obligation. Unchanged italic suffix."])
  const start = text().indexOf("A clear")
  state.selection = [start, start + "A clear obligation.".length]
  const target = (await host.selection())!
  await host.apply({ proposalId: "subspan", target, replacement: "A specific obligation." })
  expect(text()).toBe("Unchanged bold prefix. A specific obligation. Unchanged italic suffix.")
  expect(state.writes.map((w) => w.before)).toEqual([target.paragraphs[0]!.text])
  expect(state.nativeRevisions.map(change => change.text)).toEqual(["clear", "specific"])
})

test("table paragraphs refuse before authoring or native mutation, including a move into a cell during authoring", async () => {
  for (const moveDuringAuthor of [false, true]) {
    const { state, host } = office(["A fee of 100 applies.", "An adjacent numbered paragraph."])
    state.selection = [0, state.paragraphs[0]!.text.length]
    const target = (await host.selection())!
    if (moveDuringAuthor) state.duringAuthor = () => { state.paragraphs[0]!.inTable = true }
    else state.paragraphs[0]!.inTable = true
    await expect(host.apply({ proposalId: `table-${moveDuringAuthor}`, target, replacement: target.text, formatting: [{ paragraph: 0, spacingAfter: 6 }] })).rejects.toThrow("Word tables")
    expect(state.authorCalls).toBe(moveDuringAuthor ? 1 : 0)
    expect(state.writes).toHaveLength(0)
    expect(state.modeChanges).toEqual([])
    expect(state.paragraphs.map(p => p.text)).toEqual(["A fee of 100 applies.", "An adjacent numbered paragraph."])
  }
})

test.skip("verified imported scope returns new IDs and exact selected offsets without following an adjacent cursor", async () => {
  for (const multiline of [false, true]) {
    const paragraphs = multiline ? ["Prefix clear.", "clear suffix.", "An unrelated adjacent clause."] : ["Prefix clear suffix.", "Prefix clear suffix."]
    const { state, host } = office(paragraphs)
    state.selection = [7, multiline ? paragraphs[0]!.length + 1 + 5 : 12]
    state.renewImportedIds = true
    const target = (await host.selection())!, before = structuredClone(target)
    const replacement = multiline ? "specific.\rprecise" : "very specific"
    state.duringAuthor = () => { const at = paragraphs.slice(0, -1).reduce((n, text) => n + text.length + 1, 0); state.selection = [at, at + paragraphs.at(-1)!.length] }
    const input = { proposalId: `verified-import-scope-${multiline}`, target, replacement }
    const applied = await host.apply(input), next = applied.target!
    expect(next.text).toBe(replacement)
    expect(next.paragraphs).toEqual(multiline ? [
      { id: "imported-paragraph-0", text: "Prefix specific." }, { id: "imported-paragraph-1", text: "precise suffix." },
    ] : [{ id: "imported-paragraph-0", text: "Prefix very specific suffix." }])
    expect(next.startOffset).toBe(7)
    expect(next.endOffset).toBe(multiline ? "precise".length : 7 + replacement.length)
    expect(next.saved).toBe(false)
    expect(next.id).not.toBe(target.id)
    expect(Date.parse(next.capturedAt)).toBeGreaterThanOrEqual(Date.parse(target.capturedAt))
    expect(target).toEqual(before)
    expect(state.paragraphs.at(-1)).toMatchObject({ id: `paragraph-${paragraphs.length - 1}`, text: paragraphs.at(-1)! })
    expect((await host.apply(input)).target).toBeUndefined()
    expect(state.writes).toHaveLength(1)
  }
})

test.skip("deleting the complete selected wording does not invent an expanded follow-up target", async () => {
  const { state, host } = office(["Prefix clear suffix."])
  state.selection = [7, 13]
  const target = (await host.selection())!
  const applied = await host.apply({ proposalId: "empty-follow-up-scope", target, replacement: "" })
  expect(applied.status).toBe("applied")
  expect(applied.target).toBeUndefined()
  expect(state.paragraphs[0]!.text).toBe("Prefix suffix.")
})

test.skip("a retry and simultaneous approvals cannot replay the same local proposal", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  const input = { proposalId: "once", target, replacement: target.text.replace("clear", "specific") }
  await Promise.all([host.apply(input), host.apply(input)])
  expect(state.writes).toHaveLength(1)
  expect((await host.apply(input)).status).toBe("already-applied")
  await expect(host.apply({ ...input, replacement: "Unreviewed replacement." })).rejects.toThrow("different wording")
  expect(state.writes).toHaveLength(1)
})

test.skip("a lost native response restores tracking and never blindly retries a possible edit", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.throwAfterWrite = true
  const input = { proposalId: "uncertain", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply(input)).rejects.toMatchObject({ name: "WordChatHostError", mutationState: "uncertain", message: expect.stringContaining("Inspect the clause") })
  expect(state.tracking).toBe("Off")
  expect(state.writes).toHaveLength(1)
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain", message: expect.stringContaining("may already contain") })
  await expect(host.apply({ ...input, replacement: "A different requested change." })).rejects.toMatchObject({ mutationState: "uncertain" })
  expect(state.writes).toHaveLength(1)
})

test("pre-import refusal stays unchanged even when its message describes uncertainty", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  const input = { proposalId: "read-only-refusal", target, replacement: target.text, formatting: [{ paragraph: 0, spacingAfter: 6 }], author: async () => {
    throw new WordChatHostError("Word may already contain this change. Inspect the clause.")
  } }
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "unchanged" })
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})

test.skip("comments and revisions anywhere in imported paragraphs refuse; siblings remain untouched", async () => {
  for (const kind of ["comments", "revisions"] as const) {
    const { state, host } = office()
    state.selection = [16, 21]
    const target = (await host.selection())!
    state[kind] = [[0, 1]]
    await expect(host.apply({ proposalId: kind, target, replacement: "specific" })).rejects.toThrow("existing revisions, comments")
    expect(state.writes).toHaveLength(0)
    expect(state.authorCalls).toBe(0)
    const sibling = state.paragraphs[0]!.text.length + 1
    state[kind] = [[sibling + 1, sibling + 2]]
    await host.apply({ proposalId: kind + "-outside", target, replacement: "specific" })
    expect(state[kind]).toHaveLength(1)
    expect(state.paragraphs[1]!.text).toBe("Another clause.")
  }
})

test.skip("changed paragraph structure and field objects are disclosed before mutation", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  await expect(host.apply({ proposalId: "structure", target, replacement: target.text + "\rA new paragraph." })).rejects.toThrow("paragraph restructuring")
  state.xml = `<w:p xmlns:w="${WORD_XML}"><w:fldSimple/></w:p>`
  await expect(host.apply({ proposalId: "field", target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("linked Word object")
  expect(state.writes).toHaveLength(0)
})

test.skip("native Flat OPC page layout and unrelated package parts do not block ordinary text", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.xml = flatOpc(`<w:p><w:r><w:t>${escapeXml(target.text)}</w:t></w:r></w:p>`, `<pkg:part pkg:name="/word/header1.xml"><pkg:xmlData><w:hdr><w:p><w:fldSimple w:instr="PAGE"/></w:p></w:hdr></pkg:xmlData></pkg:part>`)
  await host.apply({ proposalId: "native-flat-opc", target, replacement: target.text.replace("clear", "specific") })
  expect(state.writes).toHaveLength(1)
  expect(state.writes[0]!.tracking).toBe("Off")
  expect(state.tracking).toBe("Off")
})

test.skip("selected Flat OPC fields, links, revisions, comments and section breaks stay guarded", async () => {
  const unsafe = [
    "<w:p><w:fldSimple w:instr=\"PAGE\"><w:r><w:t>clear</w:t></w:r></w:fldSimple></w:p>",
    "<w:p><w:hyperlink><w:r><w:t>clear</w:t></w:r></w:hyperlink></w:p>",
    "<w:p><w:ins><w:r><w:t>clear</w:t></w:r></w:ins></w:p>",
    "<w:p><w:commentRangeStart w:id=\"1\"/><w:r><w:t>clear</w:t></w:r></w:p>",
    "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>clear</w:t></w:r></w:p>",
    "<w:sectPr/><w:p><w:r><w:t>clear</w:t></w:r></w:p>",
  ]
  for (const [index, fragment] of unsafe.entries()) {
    const { state, host } = office()
    const target = (await host.selection())!
    state.xml = flatOpc(fragment)
    await expect(host.apply({ proposalId: `native-unsafe-${index}`, target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("linked Word object")
    expect(state.writes).toHaveLength(0)
    expect(state.modeChanges).toHaveLength(0)
  }
})

test.skip("malformed or ambiguous native XML cannot authorize an edit", async () => {
  for (const xml of ["<w:p><broken>", "<unexpected/>", flatOpc("<w:p/>").replace("</pkg:package>", flatOpc("<w:p/>").match(/<pkg:part[\s\S]*<\/pkg:part>/)![0] + "</pkg:package>")]) {
    const { state, host } = office()
    const target = (await host.selection())!
    state.xml = xml
    await expect(host.apply({ proposalId: "malformed-native-xml", target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("linked Word object")
    expect(state.writes).toHaveLength(0)
  }
})

test.skip("unexpected inserted wording never reports a confirmed reviewed amendment", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.corruptWrite = true
  const input = { proposalId: "wrong-host-result", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain", message: expect.stringContaining("current wording") })
  expect(state.writes).toHaveLength(1)
  await expect(host.apply(input)).rejects.toThrow("may already contain")
})

test("the native host requires shared authoring and never falls back to a text edit", async () => {
  const { state, nativeHost } = office()
  const target = (await nativeHost.selection())!
  await expect(nativeHost.apply({ proposalId: "missing-author", target, replacement: "Different wording." } as WordChatApplyInput)).rejects.toThrow("Reconnect to doc.haus")
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})

test.skip("a formatting change during shared authoring stops import before changing tracking", async () => {
  const { state, host } = office()
  const target = (await host.selection())!
  state.duringAuthor = () => { state.xml = flatOpc(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(target.text)}</w:t></w:r></w:p>`) }
  await expect(host.apply({ proposalId: "late-format", target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("formatting changed")
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})

test.skip("a server artifact bound to another proposal is rejected before native mutation", async () => {
  const { state, host, author } = office()
  const target = (await host.selection())!
  const input = { proposalId: "artifact-binding", target, replacement: target.text.replace("clear", "specific") }
  await expect(host.apply({ ...input, author: async () => ({ ...await author(input)({ ooxml: "" }), proposalId: "other-proposal" }) })).rejects.toThrow("matches the reviewed passage")
  expect(state.writes).toHaveLength(0)
  expect(state.modeChanges).toHaveLength(0)
})

test("a disposed pane cannot dispatch formatting after read-only preparation", async () => {
  for (const allowed of [true, false]) {
    const { state, host, nativeHost, author } = formattedOffice(), target = (await host.selection())!
    const input = { proposalId: "disposed-before-dispatch", target, replacement: target.text, formatting: [{ paragraph: 2, spacingAfter: 6 }] }
    const apply = nativeHost.apply({ ...input, author: author(input), mayDispatch: () => allowed })
    if (allowed) {
      await expect(apply).resolves.toMatchObject({ status: "applied" })
      expect(state.nativeFormatting).toEqual(["paragraph-3"])
      expect(state.modeChanges.length).toBeGreaterThan(0)
    } else {
      await expect(apply).rejects.toMatchObject({ mutationState: "unchanged", message: expect.stringContaining("pane closed") })
      expect(state.nativeFormatting).toHaveLength(0)
      expect(state.modeChanges).toHaveLength(0)
    }
    expect(state.authorCalls).toBe(1)
  }
})

test("native wording preflights and queues disjoint authored hunks in one tracking batch", async () => {
  const { state, host } = office(["A clear obligation with a firm deadline.", "Other paragraph."], true)
  const target = (await host.selection())!
  const replacement = target.text.replace("clear", "specific").replace("firm", "fixed")
  const result = await host.apply({ proposalId: "granular-native", target, replacement })
  expect(result.status).toBe("applied")
  expect(state.paragraphs[0]!.text).toBe(replacement)
  expect(state.writes.map(write => write.text)).toEqual(["fixed", "specific"])
  expect(state.modeChanges).toEqual(["TrackAll", "Off"])
  expect(state.writes.every(write => write.tracking === "TrackAll")).toBe(true)
  expect(state.syncBatchSizes.filter(size => size > 0)).toEqual([4])
})

test("lost granular native sync is uncertain and cannot replay", async () => {
  const { state, host } = office(["A clear obligation with a firm deadline."], true)
  const target = (await host.selection())!
  const input = { proposalId: "granular-lost", target, replacement: target.text.replace("clear", "specific").replace("firm", "fixed") }
  state.throwAfterWrite = true
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
  expect(state.writes).toHaveLength(2)
  await expect(host.apply(input)).rejects.toThrow("may already contain")
  expect(state.writes).toHaveLength(2)
})

test("native wording permits one ordinary cell but refuses a move into it during authoring", async () => {
  for (const moveDuringAuthor of [false, true]) {
    const { state, host } = office(["A clear obligation.", "Adjacent content."], true)
    const target = (await host.selection())!
    if (moveDuringAuthor) state.duringAuthor = () => { state.paragraphs[0]!.inTable = true }
    else state.paragraphs[0]!.inTable = true
    const input = { proposalId: `native-cell-${moveDuringAuthor}`, target, replacement: target.text.replace("clear", "specific") }
    if (moveDuringAuthor) {
      await expect(host.apply(input)).rejects.toThrow("changed while doc.haus prepared")
      expect(state.writes).toHaveLength(0)
    } else {
      expect((await host.apply(input)).status).toBe("applied")
      expect(state.paragraphs[0]!.text).toBe("A specific obligation.")
    }
  }
})

test("native wording preserves actual cell delimiters and uses the real backend author", async () => {
  const { state, nativeHost, text } = office(["A clear obligation.", "Other cell."], true)
  state.separator = "\t"
  state.paragraphs[0]!.inTable = true
  const target = (await nativeHost.selection())!
  const replacement = target.text.replace("clear", "specific")
  const result = await nativeHost.apply({ proposalId: "real-author-cell", target, replacement,
    author: async ({ ooxml }) => ({ ...await authorWordChatTarget(ooxml, target, replacement), proposalId: "real-author-cell" }) })
  expect(result.status).toBe("applied")
  expect(text()).toBe("A specific obligation.\tOther cell.")
  expect(state.writes).toHaveLength(1)
})

test("the default host enables granular wording while explicit opt-out remains guarded", async () => {
  const enabled = office(["A clear obligation."], false)
  const defaultHost = new WordChatHost(enabled.word, () => true)
  const target = (await defaultHost.selection())!
  const input = { proposalId: "default-granular", target, replacement: target.text.replace("clear", "specific") }
  expect((await defaultHost.apply({ ...input, author: enabled.author(input) })).status).toBe("applied")
  const disabled = office(["A clear obligation."], false)
  const oldTarget = (await disabled.nativeHost.selection())!
  const oldInput = { proposalId: "explicit-opt-out", target: oldTarget, replacement: oldTarget.text.replace("clear", "specific") }
  await expect(disabled.nativeHost.apply({ ...oldInput, author: disabled.author(oldInput) })).rejects.toThrow("cannot safely apply")
  expect(disabled.state.writes).toHaveLength(0)
})

test("real authored one-sided wording uses one native batch and preserves surrounding cell text", async () => {
  const cases = [
    { before: "Alpha beta remains.", after: "Alpha new beta remains.", start: 6, end: 10, replacement: "new beta", types: ["Added"] },
    { before: "Alpha beta remains.", after: "Alpha remains.", start: 6, end: 11, replacement: "", types: ["Deleted"] },
    { before: "Alpha beta remains.", after: "New Alpha beta remains.", start: 0, end: 19, replacement: "New Alpha beta remains.", types: ["Added"] },
    { before: "Alpha beta remains.", after: "Alpha beta remains. Added.", start: 0, end: 19, replacement: "Alpha beta remains. Added.", types: ["Added"] },
    { before: "Alpha beta and gamma remain.", after: "Alpha new beta and remain.", start: 0, end: 28, replacement: "Alpha new beta and remain.", types: ["Added", "Deleted"] },
  ]
  for (const [index, row] of cases.entries()) {
    const { state, nativeHost, text } = office([row.before, "Other cell."], true)
    state.separator = "\t"
    state.paragraphs[0]!.inTable = true
    state.selection = [row.start, row.end]
    const target = (await nativeHost.selection())!
    const proposalId = `one-sided-${index}`
    const input = { proposalId, target, replacement: row.replacement,
      author: async ({ ooxml }: { ooxml: string }) => ({ ...await authorWordChatTarget(ooxml, target, row.replacement), proposalId }) }
    const result = await nativeHost.apply(input)
    expect(result.status).toBe("applied")
    if (row.replacement) expect(result.target?.text).toBe(row.replacement)
    else expect(result.target).toBeUndefined()
    expect(text()).toBe(`${row.after}\tOther cell.`)
    expect(state.nativeRevisions.map(change => change.type).sort()).toEqual([...row.types].sort())
    expect(state.syncBatchSizes.filter(size => size > 0)).toEqual([row.types.length + 2])
    expect(state.modeChanges).toEqual(["TrackAll", "Off"])
    expect(state.writes.every(write => write.tracking === "TrackAll")).toBe(true)
    expect((await nativeHost.apply(input)).status).toBe("already-applied")
    expect(state.writes).toHaveLength(row.types.length)
  }
})

test("a lost pure-deletion response stays uncertain without a second dispatch", async () => {
  const { state, nativeHost, text } = office(["Alpha beta remains."], true)
  state.selection = [6, 11]
  const target = (await nativeHost.selection())!, proposalId = "pure-delete-lost"
  const input = { proposalId, target, replacement: "",
    author: async ({ ooxml }: { ooxml: string }) => ({ ...await authorWordChatTarget(ooxml, target, ""), proposalId }) }
  state.throwAfterWrite = true
  await expect(nativeHost.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
  expect(text()).toBe("Alpha remains.")
  expect(state.writes).toHaveLength(1)
  await expect(nativeHost.apply(input)).rejects.toThrow("may already contain")
  expect(state.writes).toHaveLength(1)
})

test("ordinary wording does not read the unloaded desktop formatting-tracking property", async () => {
  const fixture = office(["Alpha beta remains."], true)
  await fixture.word.run(async context => {
    Object.defineProperty(context.document, "trackFormatting", { get() { throw new Error("PropertyNotLoaded") } })
  })
  const host = new WordChatHost(fixture.word, name => name === "WordApi")
  const target = (await host.selection())!, proposalId = "wording-no-formatting-property"
  const replacement = target.text.replace("beta", "new beta")
  expect((await host.apply({ proposalId, target, replacement,
    author: async ({ ooxml }) => ({ ...await authorWordChatTarget(ooxml, target, replacement), proposalId }) })).status).toBe("applied")
  expect(fixture.text()).toBe("Alpha new beta remains.")
})


test("native wording uses body content when whole body has a terminal paragraph mark", async () => {
  const { state, host } = office(["A clear obligation.", "Unrelated text."], true)
  state.terminalBodyMark = true
  const target = (await host.selection())!
  const result = await host.apply({ proposalId: "body-terminal-mark", target, replacement: target.text.replace("clear", "specific") })
  expect(result.status).toBe("applied")
  expect(state.paragraphs[0]!.text).toBe("A specific obligation.")
})


test("body content comparison still refuses unrelated changes after dispatch", async () => {
  const { state, host } = office(["A clear obligation.", "Unrelated text."], true)
  state.terminalBodyMark = true
  const target = (await host.selection())!
  state.afterSync = () => { if (state.writes.length) { state.paragraphs[1]!.text = "Changed text."; state.afterSync = undefined } }
  await expect(host.apply({ proposalId: "body-outside-change", target, replacement: target.text.replace("clear", "specific") })).rejects.toThrow("could not confirm")
  expect(state.writes).toHaveLength(1)
})


test("native dispatch error remains the cause of an uncertain failure and cannot replay", async () => {
  const { state, host } = office(["A clear obligation."], true)
  const target = (await host.selection())!
  const nativeError = Object.assign(new Error("native sync failed"), { code: "GeneralException" })
  state.throwAfterWrite = true; state.syncFailure = nativeError
  const input = { proposalId: "native-cause", target, replacement: target.text.replace("clear", "specific") }
  let caught: unknown
  try { await host.apply(input) } catch (error) { caught = error }
  expect(caught).toBeInstanceOf(WordChatHostError)
  expect((caught as WordChatHostError).mutationState).toBe("uncertain")
  expect((caught as WordChatHostError).cause).toBe(nativeError)
  expect(state.writes).toHaveLength(1)
  await expect(host.apply(input)).rejects.toMatchObject({ mutationState: "uncertain", message: expect.stringContaining("may already contain") })
  expect(state.writes).toHaveLength(1)
})


test("native cell wording verifies neighbouring paragraph content when Whole expands the row", async () => {
  const { state, host } = office(["Previous cell.", "A clear obligation.", "Next cell."], true)
  state.separator = "\t"; state.cellWholeExpandsRow = true
  for (const p of state.paragraphs) p.inTable = true
  const start = state.paragraphs[0]!.text.length + 1
  state.selection = [start, start + state.paragraphs[1]!.text.length]
  const target = (await host.selection())!
  expect((await host.apply({ proposalId: "cell-whole-row", target, replacement: target.text.replace("clear", "specific") })).status).toBe("applied")
  expect(state.paragraphs.map(p => p.text)).toEqual(["Previous cell.", "A specific obligation.", "Next cell."])
})


test("native cell wording rejects an XML-only formatting change in a neighbour", async () => {
  const { state, host } = office(["Previous cell.", "A clear obligation.", "Next cell."], true)
  state.separator = "\t"; state.cellWholeExpandsRow = true
  for (const p of state.paragraphs) p.inTable = true
  const start = state.paragraphs[0]!.text.length + 1
  state.selection = [start, start + state.paragraphs[1]!.text.length]
  const target = (await host.selection())!
  state.paragraphXml.set("paragraph-2", "<w:p><w:r><w:t>Next cell.</w:t></w:r></w:p>")
  state.afterSync = () => { if (state.writes.length) { state.paragraphXml.set("paragraph-2", "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Next cell.</w:t></w:r></w:p>"); state.afterSync = undefined } }
  await expect(host.apply({ proposalId: "cell-neighbour-change", target, replacement: target.text.replace("clear", "specific") })).rejects.toMatchObject({ mutationState: "uncertain", message: expect.stringContaining("could not confirm") })
  expect(state.paragraphs.map(p => p.text)).toEqual(["Previous cell.", "A specific obligation.", "Next cell."])
  expect(state.writes).toHaveLength(1)
})

test("native wording preserves faithful body, list and styled-cell properties when Content omits them", async () => {
  for (const kind of ["body", "list", "cell"]) {
    const { state, nativeHost, word } = office(["A clear obligation.", "Other paragraph."], true)
    state.omitContentProperties = true
    if (kind === "cell") { state.paragraphs[0]!.inTable = true; state.separator = "\t" }
    const properties = kind === "list" ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="120"/>' : kind === "cell" ? '<w:pStyle w:val="Heading1"/><w:spacing w:after="180"/>' : '<w:spacing w:before="80"/>'
    const xml = `<w:p><w:pPr>${properties}</w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>A clear obligation.</w:t></w:r></w:p>`
    state.paragraphXml.set("paragraph-0", xml)
    await word.run(async context => { const content = context.document.getParagraphByUniqueLocalId("paragraph-0").getRange("Content").getOoxml(); await context.sync(); expect(content.value).not.toContain("<w:pPr>") })
    const target = (await nativeHost.selection())!, replacement = target.text.replace("clear", "specific"), proposalId = `faithful-${kind}`
    let authoredInput = "", authoredParagraph = ""
    state.afterSync = () => { if (state.writes.length) { state.paragraphXml.set("paragraph-0", authoredParagraph); state.afterSync = undefined } }
    const input = { proposalId, target, replacement, author: async ({ ooxml }: { ooxml:string }) => {
      authoredInput = ooxml; const result = await authorWordChatTarget(ooxml, target, replacement)
      const doc = new DOMParser().parseFromString(result.ooxml, "application/xml")
      authoredParagraph = new XMLSerializer().serializeToString(doc.getElementsByTagNameNS(WORD_XML, "p")[0]!)
      return { ...result, proposalId }
    } }
    expect((await nativeHost.apply(input)).status).toBe("applied")
    expect(authoredInput).toContain(properties)
    expect(state.paragraphXml.get("paragraph-0")).toContain("spacing")
    expect(state.paragraphXml.get("paragraph-0")).toContain("w:b")
    expect((await nativeHost.apply(input)).status).toBe("already-applied")
    expect(state.writes).toHaveLength(1)
  }
})

test("native wording refuses faithful property drift before writing and fences postwrite corruption", async () => {
  for (const phase of ["author", "postwrite"]) {
    const { state, nativeHost } = office(["A clear obligation.", "Other paragraph."], true)
    const xml = '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>A clear obligation.</w:t></w:r></w:p>'
    state.paragraphXml.set("paragraph-0", xml)
    const target = (await nativeHost.selection())!, replacement = target.text.replace("clear", "specific"), proposalId = `faithful-drift-${phase}`
    const input = { proposalId, target, replacement, author: async ({ ooxml }: { ooxml:string }) => {
      const result = { ...await authorWordChatTarget(ooxml, target, replacement), proposalId }
      if (phase === "author") state.paragraphXml.set("paragraph-0", xml.replace('w:after="120"', 'w:after="240"'))
      return result
    } }
    if (phase === "postwrite") state.afterSync = () => { if (state.writes.length) { state.paragraphXml.set("paragraph-0", xml.replace("clear", "specific").replace('w:after="120"', 'w:after="240"')); state.afterSync = undefined } }
    await expect(nativeHost.apply(input)).rejects.toMatchObject({ mutationState: phase === "author" ? "unchanged" : "uncertain" })
    expect(state.writes).toHaveLength(phase === "author" ? 0 : 1)
    if (phase === "postwrite") { await expect(nativeHost.apply(input)).rejects.toThrow("may already contain"); expect(state.writes).toHaveLength(1) }
  }
})

test("row author failure is unchanged and simultaneous requests do not dispatch twice", async () => {
  let runs = 0, calls = 0, release!: () => void
  const word = { async run() { runs++; throw Error("must not dispatch") } } as WordChatWord
  const host = new WordChatHost(word, () => true)
  const proposal = { kind: "table_row_insert" as const, id: "row-1", documentId: "doc", explanation: "Append", openItems: [], createdAt: "now", position: "end" as const,
    table: { tableIndex: 0, rows: [[{id:"a",text:"A"},{id:"b",text:"B"}]] }, cells: ["C", "D"] }
  const author = async () => { calls++; await new Promise<void>(resolve => { release = resolve }); throw Error("author failed") }
  const first = host.applyRow(proposal, async () => new Uint8Array([1]), author)
  const second = host.applyRow(proposal, async () => new Uint8Array([1]), author)
  const results = Promise.allSettled([first, second])
  await Promise.resolve(); release()
  for (const result of await results) { expect(result.status).toBe("rejected"); if (result.status === "rejected") expect(result.reason).toMatchObject({ mutationState: "unchanged" }) }
  expect(calls).toBe(1); expect(runs).toBe(0)
})

test("tampered row oracle refuses before opening a Word mutation batch", async () => {
  let runs = 0
  const word = { async run() { runs++; throw Error("must not dispatch") } } as WordChatWord
  const host = new WordChatHost(word, () => true)
  const proposal = { kind: "table_row_insert" as const, id: "row-2", documentId: "doc", explanation: "Append", openItems: [], createdAt: "now", position: "end" as const,
    table: { tableIndex: 0, rows: [[{id:"a",text:"A"}]] }, cells: ["B"] }
  await expect(host.applyRow(proposal, async () => new Uint8Array([1]), async () => ({ kind: "table_row_insert", proposalId: "wrong" } as never))).rejects.toMatchObject({ mutationState:"unchanged" })
  expect(runs).toBe(0)
})

async function nativeRowOrchestrationFixture(lostSync = false, reloadInvalidates = false) {
  const { zipSync, strToU8 } = await import("fflate")
  const writerPath = new URL("../../../dochaus/lib/docx-table-row-insertion.ts", import.meta.url).href
  const { docxTableRowInsertionTargets, authorDocxTableRowInsertion } = await import(writerPath)
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const old = '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Old</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>Old two</w:t></w:r></w:p></w:tc></w:tr>'
  const tableXml = (extra = "") => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${old}${extra}</w:tbl>`
  const source = zipSync({
    "[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    "_rels/.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/></Relationships>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${WORD_XML}"/>`),
    "word/document.xml": strToU8(`<w:document xmlns:w="${WORD_XML}"><w:body>${tableXml()}<w:sectPr/></w:body></w:document>`),
  })
  const packet = await docxTableRowInsertionTargets(source)
  const table = { tableIndex: 0, rows: [[{ id: "a", text: "Old" }, { id: "b", text: "Old two" }]] }
  const proposal = { kind:"table_row_insert" as const, id:"native-row", documentId:"doc", explanation:"Add", openItems:["Disclosed"], createdAt:"now", position:"end" as const, table, cells:["New", "Value"] }
  const writerRequest = { source_sha256:packet.source_sha256, target_handle:packet.targets[0]!.target_handle, position:"after" as const, cells:proposal.cells, author:"doc.haus" }
  const authored = await authorDocxTableRowInsertion(source, writerRequest)
  const oracle = { kind:"table_row_insert" as const, proposalId:proposal.id, sourceSha256:packet.source_sha256, table, cells:proposal.cells, author:"doc.haus" as const, writerRequest, authoredBase64:btoa(Array.from(authored.bytes as Uint8Array, x=>String.fromCharCode(x)).join("")), revisions:authored.revisions }
  const stamp = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"'
  const cell = (text:string, mark:number, content:number) => `<w:tc><w:tcPr/><w:p><w:pPr><w:rPr><w:ins w:id="${mark}" ${stamp}/></w:rPr></w:pPr><w:ins w:id="${content}" ${stamp}><w:r><w:t>${text}</w:t></w:r></w:ins></w:p></w:tc>`
  const addedXml = `<w:tr><w:trPr><w:ins w:id="1" ${stamp}/></w:trPr>${cell("New",2,3)}${cell("Value",4,5)}</w:tr>`
  let applied = false, queued = false, rowWrites = 0, exports = 0, nestingInvalidated = false
  const para = (id:string,text:string) => ({ uniqueLocalId:id, text, get tableNestingLevel() { if (nestingInvalidated) throw Error("PropertyNotLoaded: Paragraph.tableNestingLevel"); return 1 }, styleBuiltIn:"Normal", spaceBefore:0, spaceAfter:0, lineSpacing:12, load() {}, getRange() { return {font:{bold:false,italic:false,underline:"None",load(){}}} } })
  const oldParas = [para("a","Old"),para("b","Old two")], newParas = [para("c","New"),para("d","Value")]
  const collection = <T>(items:T[]) => ({items,load() {}})
  const nativeTable = { get rowCount() { return applied ? 2 : 1 }, isUniform:true, get values() { return applied ? [["Old","Old two"],["New","Value"]] : [["Old","Old two"]] }, load() {},
    getCell(row:number,col:number) { return {body:{paragraphs:collection([row ? newParas[col]! : oldParas[col]!]), insertText() { return {font:{bold:true,italic:true,underline:"Single",load(){}}} }}} },
    addRows() { queued = true; rowWrites++ }, }
  const body = { get paragraphs() { return { items:applied ? [...oldParas,...newParas] : oldParas, load(properties:string) { if (reloadInvalidates && applied) nestingInvalidated = !properties.includes("tableNestingLevel") } } }, tables:collection([nativeTable]),
    getRange() { return {getOoxml() {return {value:flatOpc(tableXml(applied ? addedXml : ""))}}, getReviewedText() {return {value:applied ? "Old\rOld two\rNew\rValue" : "Old\rOld two"}}, getTrackedChanges() {return collection([])}, getComments() {return collection([])} } } }
  const document = {body, changeTrackingMode:"Off", trackFormatting:false, load() {}, getParagraphByUniqueLocalId(id:string) {return [...oldParas,...newParas].find(p=>p.uniqueLocalId===id)!} }
  const word = {async run(fn:(context:unknown)=>Promise<unknown>) {return fn({document, async sync() { if (queued) {queued=false; applied=true; if (lostSync) throw Error("lost final sync") } }})}} as WordChatWord
  return {host:new WordChatHost(word,()=>true), proposal, exportCurrent:async()=>{exports++;return source}, author:async()=>oracle, counts:()=>({rowWrites,exports})}
}

test("native row confirms once and a duplicate does not re-export or append", async () => {
  const f = await nativeRowOrchestrationFixture()
  expect(await f.host.applyRow(f.proposal, f.exportCurrent, f.author)).toEqual({status:"applied",tracked:true,saved:false,authors:["Reviewer"]})
  expect((await f.host.applyRow(f.proposal, f.exportCurrent, f.author)).status).toBe("already-applied")
  expect(f.counts()).toEqual({rowWrites:1,exports:1})
})

test("row confirmation retains loaded paragraph properties when document-state reload invalidates proxies", async () => {
  const f = await nativeRowOrchestrationFixture(false, true)
  expect(await f.host.applyRow(f.proposal, f.exportCurrent, f.author)).toEqual({status:"applied",tracked:true,saved:false,authors:["Reviewer"]})
  expect(f.counts()).toEqual({rowWrites:1,exports:1})
})

test("lost native row mutation sync is uncertain and cannot replay", async () => {
  const f = await nativeRowOrchestrationFixture(true)
  await expect(f.host.applyRow(f.proposal, f.exportCurrent, f.author)).rejects.toMatchObject({mutationState:"uncertain"})
  await expect(f.host.applyRow(f.proposal, f.exportCurrent, f.author)).rejects.toMatchObject({mutationState:"uncertain"})
  expect(f.counts()).toEqual({rowWrites:1,exports:1})
})


test("native inline formatting confirms tracked markup and fences an uncertain write", async () => {
  for (const corrupt of [false, true]) {
    const { state, nativeHost } = office(["abcd"], true)
    const marker = '<w:bookmarkStart w:id="4" w:name="place"/>'
    state.paragraphXml.set("paragraph-0", `<w:p><w:r><w:t>ab</w:t></w:r>${marker}<w:r><w:t>cd</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p>`)
    const target = (await nativeHost.selection())!
    const run = (text: string, id: number, author: string) => `<w:r><w:rPr><w:b/><w:rPrChange w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>${text}</w:t></w:r>`
    const para = (author: string, anchor = marker) => `<w:p>${run("ab", 8, author)}${anchor}${run("cd", 9, author)}<w:bookmarkEnd w:id="4"/></w:p>`
    state.inlineAfter = para("Native user", corrupt ? "" : marker)
    const input = {proposalId:"inline",target,replacement:target.text,inlineFormatting:{bold:true},author:async () => ({kind:"inline_format" as const,proposalId:"inline",target,inlineFormatting:{bold:true}})}
    if (corrupt) {
      await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
      await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
    } else {
      expect((await nativeHost.apply(input)).status).toBe("applied")
      expect((await nativeHost.apply(input)).status).toBe("already-applied")
    }
    expect(state.inlineWrites).toBe(1)
    expect(state.tracking).toBe("Off")
  }
})

test("inline plan binding refuses forgery and satisfied native font is a no-op", async () => {
  const {state,nativeHost} = office(["abcd"], true)
  const target = (await nativeHost.selection())!
  const input = {proposalId:"inline-plan",target,replacement:target.text,inlineFormatting:{bold:true},author:async () => ({kind:"inline_format" as const,proposalId:"wrong",target,inlineFormatting:{bold:true}})}
  await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"unchanged"})
  expect(state.inlineWrites).toBe(0)
  state.inlineBold = true
  const noOp = {...input,author:async () => ({kind:"inline_format" as const,proposalId:"inline-plan",target,inlineFormatting:{bold:true}})}
  expect((await nativeHost.apply(noOp)).status).toBe("applied")
  expect(state.inlineWrites).toBe(0)
  expect((await nativeHost.apply(noOp)).status).toBe("already-applied")
})

test("native paragraph insertion restores tracking and never repeats a dispatched operation", async () => {
  for (const tracking of ["Off", "TrackAll"]) for (const corrupt of [false, true]) for (const formatting of [undefined, [{paragraph:1,style:"Normal" as const}]]) {
    const { state, nativeHost } = office(["An existing obligation."], true)
    state.tracking = tracking
    state.paragraphXml.set("paragraph-0", '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000001"><w:r><w:t>An existing obligation.</w:t></w:r></w:p>')
    const target = (await nativeHost.selection())!
    state.corruptParagraphInsertion = corrupt
    const input = { proposalId: "paragraph-insertion", target, replacement: `${target.text}\rA new obligation.`, formatting,
      author: async () => ({ kind: "paragraph_insert" as const, proposalId: "paragraph-insertion", target, text: "A new obligation.", formatting }) }
    if (corrupt) {
      await expect(nativeHost.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
      await expect(nativeHost.apply(input)).rejects.toMatchObject({ mutationState: "uncertain" })
    } else {
      expect((await nativeHost.apply(input)).status).toBe("applied")
      expect((await nativeHost.apply(input)).status).toBe("already-applied")
      expect(state.paragraphs.map(p => p.text)).toEqual(["An existing obligation.", "A new obligation."])
    }
    expect(state.paragraphInsertions).toBe(1)
    expect(state.tracking).toBe(tracking)
    expect(state.modeChanges).toEqual(["TrackAll", tracking])
  }
})

test("native paragraph insertion refuses an altered prepared command before dispatch", async () => {
  const { state, nativeHost } = office(["An existing obligation."], true)
  state.paragraphXml.set("paragraph-0", '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000001"><w:r><w:t>An existing obligation.</w:t></w:r></w:p>')
  const target = (await nativeHost.selection())!
  let prepared = 0
  await expect(nativeHost.apply({ proposalId: "paragraph-bound", target, replacement: `${target.text}\rReviewed wording.`,
    author: async () => { prepared++; return { kind: "paragraph_insert" as const, proposalId: "paragraph-bound", target, text: "Different wording." } } })).rejects.toMatchObject({ mutationState: "unchanged" })
  expect(prepared).toBe(1)
  await expect(nativeHost.apply({proposalId:"paragraph-style-bound",target,replacement:`${target.text}\rReviewed wording.`,formatting:[{paragraph:1,style:"Normal"}],
    author:async () => ({kind:"paragraph_insert" as const,proposalId:"paragraph-style-bound",target,text:"Reviewed wording."})})).rejects.toMatchObject({mutationState:"unchanged"})
  expect(state.paragraphInsertions).toBe(0)
  expect(state.modeChanges).toEqual([])
})


test("native paragraph insertion reports uncertainty when inserted style differs from Normal", async () => {
  const {state,nativeHost}=office(["An existing obligation."],true)
  state.paragraphXml.set("paragraph-0",'<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000001"><w:r><w:t>An existing obligation.</w:t></w:r></w:p>')
  state.insertedParagraphStyle="Heading1"
  const target=(await nativeHost.selection())!, formatting=[{paragraph:1,style:"Normal" as const}]
  const input={proposalId:"style-drift",target,replacement:target.text+"\rNew paragraph.",formatting,author:async () => ({kind:"paragraph_insert" as const,proposalId:"style-drift",target,text:"New paragraph.",formatting})}
  await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
  await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
  expect(state.paragraphInsertions).toBe(1)
})

test("explicit paragraph deletion tracks text and mark, preserves the following paragraph and never replays", async () => {
  for (const tracking of ["Off", "TrackAll"]) for (const corrupt of ["none", "markup", "current"]) {
    const {state,nativeHost}=office(["An existing obligation.","A following provision."],true)
    state.tracking=tracking
    state.selection=[0,"An existing obligation.".length]
    const target=(await nativeHost.selection())!
    state.corruptParagraphDeletion=corrupt==="markup"
    state.wrongDeletionCurrent=corrupt==="current"
    const input={proposalId:"delete-paragraph",target,replacement:"",paragraphAction:"delete" as const,
      author:async () => ({kind:"paragraph_delete" as const,proposalId:"delete-paragraph",target})}
    if(corrupt==="none") {
      expect((await nativeHost.apply(input)).status).toBe("applied")
      expect((await nativeHost.apply(input)).status).toBe("already-applied")
      expect(state.paragraphs.map(p=>[p.id,p.text])).toEqual([["paragraph-0",""],["paragraph-1","A following provision."]])
    } else {
      await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
      await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
    }
    expect(state.paragraphDeletions).toBe(1)
    expect(state.tracking).toBe(tracking)
    expect(state.modeChanges).toEqual(["TrackAll",tracking])
  }
})

test("paragraph deletion rejects a forged prepared action and final boundary without writing", async () => {
  const {state,nativeHost}=office(["An existing obligation.","A following provision."],true)
  state.selection=[0,"An existing obligation.".length]
  const target=(await nativeHost.selection())!
  await expect(nativeHost.apply({proposalId:"delete-bound",target,replacement:"",paragraphAction:"delete",
    author:async () => ({kind:"paragraph_delete" as const,proposalId:"other",target})})).rejects.toMatchObject({mutationState:"unchanged"})
  expect(state.paragraphDeletions).toBe(0)
  const final=office(["Only paragraph."],true)
  const last=(await final.nativeHost.selection())!
  let authorCalls=0
  await expect(final.nativeHost.apply({proposalId:"final",target:last,replacement:"",paragraphAction:"delete",
    author:async()=>{authorCalls++;return {kind:"paragraph_delete" as const,proposalId:"final",target:last}}})).rejects.toThrow()
  expect(authorCalls).toBe(0)
  expect(final.state.paragraphDeletions).toBe(0)
})

test("paragraph operations classify preflight, author and fresh-capture failures as unchanged", async () => {
  for (const kind of ["insert", "delete"] as const) for (const stage of ["preflight", "author", "fresh"] as const) {
    const {state,nativeHost}=office(["An existing obligation.","A following provision."],true)
    state.selection=[0,"An existing obligation.".length]
    const target=(await nativeHost.selection())!
    if (stage === "preflight") state.paragraphXml.set("paragraph-0",'<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>An existing obligation.</w:t></w:r></w:p>')
    const input={proposalId:`${kind}-${stage}`,target,replacement:kind === "insert" ? `${target.text}\rNew paragraph.` : "",
      ...(kind === "delete" ? {paragraphAction:"delete" as const} : {}),
      author:async () => {
        if (stage === "author") throw new Error("author unavailable")
        if (stage === "fresh") state.afterSync=()=>{throw new Error("fresh capture unavailable")}
        return kind === "insert" ? {kind:"paragraph_insert" as const,proposalId:`${kind}-${stage}`,target,text:"New paragraph."} : {kind:"paragraph_delete" as const,proposalId:`${kind}-${stage}`,target}
      }}
    await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"unchanged"})
    expect(state.paragraphInsertions).toBe(0); expect(state.paragraphDeletions).toBe(0); expect(state.modeChanges).toEqual([])
  }
})

test("native list append binds the list and checks tracked formatting without replay", async () => {
  for (const tracking of ["Off", "TrackAll"]) for (const failure of ["none", "markup", "font", "forged", "lost"] as const) {
    const {state,nativeHost}=office(["Existing list item."],true)
    state.tracking=tracking
    state.paragraphs[0]!.list={id:0,level:0,label:"1."}
    state.paragraphXml.set("paragraph-0",'<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000001"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Existing list item.</w:t></w:r></w:p>')
    state.extraParts='<pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData><w:numbering><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="7"/></w:num></w:numbering></pkg:xmlData></pkg:part>'
    const snapshot=await nativeHost.capture(), list=snapshot.list!
    expect(list?.members.map(m=>m.label)).toEqual(["1."])
    const target=(await nativeHost.selection())!
    state.listBadFont=failure==="font";state.listBadMarkup=failure==="markup";state.throwAfterWrite=failure==="lost"
    const input={proposalId:"list-test",target,replacement:target.text+"\rAdded item.",paragraphAction:"append_list_item" as const,list,
      author:async()=>({kind:"list_insert" as const,proposalId:failure==="forged"?"other":"list-test",target,text:"Added item.",list})}
    if (failure === "none") {
      expect((await nativeHost.apply(input)).status).toBe("applied")
      expect((await nativeHost.apply(input)).status).toBe("already-applied")
    } else {
      expect(await nativeHost.apply(input).catch(e=>e.mutationState)).toBe(failure==="forged"?"unchanged":"uncertain")
      if (failure !== "forged") expect(await nativeHost.apply(input).catch(e=>e.mutationState)).toBe("uncertain")
    }
    expect(state.listInsertions).toBe(failure==="forged"?0:1)
    expect(state.tracking).toBe(tracking)
    expect(state.modeChanges).toEqual(failure==="forged"?[]:["TrackAll",tracking])
  }
})

test("list capture refuses ambiguous lists and dispatch refuses a stale list", async () => {
  const {state,nativeHost}=office(["One.","Other."],true)
  state.paragraphs[0]!.list={id:5,level:0,label:"1."};state.paragraphs[1]!.list={id:6,level:0,label:"1."}
  const p=(id:string,num:number,text:string)=>`<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="${id}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${num}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
  state.paragraphXml.set("paragraph-0",p("00000001",5,"One."));state.paragraphXml.set("paragraph-1",p("00000002",6,"Other."))
  state.extraParts='<pkg:part pkg:name="/word/numbering.xml"><pkg:xmlData><w:numbering><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="7"/></w:num><w:num w:numId="6"><w:abstractNumId w:val="7"/></w:num></w:numbering></pkg:xmlData></pkg:part>'
  expect((await nativeHost.capture()).list).toBeUndefined()
  const selected=await nativeHost.capture(true), list=selected.list!, target=selected.selection!
  expect(list.id).toBe(5)
  const input={proposalId:"stale-list",target,replacement:target.text+"\rAdded.",paragraphAction:"append_list_item" as const,list,
    author:async()=>{state.paragraphs[0]!.list!.label="2.";return {kind:"list_insert" as const,proposalId:"stale-list",target,text:"Added.",list}}}
  expect(await nativeHost.apply(input).catch(e=>e.mutationState)).toBe("unchanged")
  expect(state.listInsertions).toBe(0)
  expect(state.modeChanges).toEqual([])
})

test("row deletion author failure is unchanged and simultaneous requests share preparation", async () => {
  let runs = 0, calls = 0, release!: () => void
  const word = { async run() { runs++; throw Error("must not dispatch") } } as WordChatWord
  const host = new WordChatHost(word, () => true)
  const proposal = { kind:"table_row_delete" as const, id:"delete-row", documentId:"doc", explanation:"Delete", openItems:[], createdAt:"now", rowIndex:1,
    table:{tableIndex:0, rows:[[{id:"a",text:"A"}], [{id:"b",text:"B"}]]}, cells:["B"] }
  const author = async () => { calls++; await new Promise<void>(resolve => { release = resolve }); throw Error("author failed") }
  const first = host.applyRowDeletion(proposal, async () => new Uint8Array([1]), author)
  const second = host.applyRowDeletion(proposal, async () => new Uint8Array([1]), author)
  const results = Promise.allSettled([first, second])
  await Promise.resolve(); release()
  for (const result of await results) { expect(result.status).toBe("rejected"); if (result.status === "rejected") expect(result.reason).toMatchObject({mutationState:"unchanged"}) }
  expect(calls).toBe(1); expect(runs).toBe(0)
})

test("row deletion rejects a tampered oracle before opening Word", async () => {
  let runs = 0
  const host = new WordChatHost({async run() { runs++; throw Error("must not dispatch") }} as WordChatWord, () => true)
  const proposal = {kind:"table_row_delete" as const, id:"delete-row", documentId:"doc", explanation:"Delete", openItems:[], createdAt:"now", rowIndex:1,
    table:{tableIndex:0, rows:[[{id:"a",text:"A"}], [{id:"b",text:"B"}]]}, cells:["B"]}
  await expect(host.applyRowDeletion(proposal, async () => new Uint8Array([1]), async () => ({kind:"table_row_delete",proposalId:"wrong"} as never))).rejects.toMatchObject({mutationState:"unchanged"})
  expect(runs).toBe(0)
})

test("paragraph merge rejects a forged prepared command without writing", async () => {
  const {state,nativeHost}=office(["First clause.","Second clause."],true)
  state.selection=[0,"First clause.\rSecond clause.".length]
  const target=(await nativeHost.selection())!
  let calls=0
  await expect(nativeHost.apply({proposalId:"merge-review",target,replacement:"First clause. Second clause.",paragraphAction:"merge",
    author:async()=>{calls++;return {kind:"paragraph_merge" as const,proposalId:"different",target}}})).rejects.toMatchObject({mutationState:"unchanged"})
  expect(calls).toBe(1);expect(state.writes).toHaveLength(0);expect(state.modeChanges).toEqual([])
})

test("paragraph merge fences an unconfirmed native write against replay", async () => {
  const {state,nativeHost}=office(["First clause.","Second clause."],true)
  state.selection=[0,"First clause.\rSecond clause.".length]
  const target=(await nativeHost.selection())!
  let leaseChecks=0
  const input={proposalId:"merge-boundary",target,replacement:"First clause. Second clause.",paragraphAction:"merge" as const,
    author:async()=>({kind:"paragraph_merge" as const,proposalId:"merge-boundary",target}),mayDispatch:()=>{leaseChecks++;return true}}
  await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
  await expect(nativeHost.apply(input)).rejects.toMatchObject({mutationState:"uncertain"})
  expect(leaseChecks).toBe(1);expect(state.modeChanges).toEqual(["TrackAll","Off"])
})

test("paragraph merge checks the lease before queuing the boundary change", async () => {
  const {state,nativeHost}=office(["First clause.","Second clause."],true)
  state.selection=[0,"First clause.\rSecond clause.".length]
  const target=(await nativeHost.selection())!
  await expect(nativeHost.apply({proposalId:"merge-lease",target,replacement:"First clause. Second clause.",paragraphAction:"merge",
    author:async()=>({kind:"paragraph_merge" as const,proposalId:"merge-lease",target}),mayDispatch:()=>false})).rejects.toMatchObject({mutationState:"unchanged"})
  expect(state.writes).toHaveLength(0);expect(state.modeChanges).toEqual([])
})
