import type { WordChatAuthoredAmendment, WordChatCellRun, WordChatPreparedAmendment, WordChatParagraphFormat, WordChatInlineFormat, WordChatSnapshot, WordChatTableCapture, WordChatTarget, WordChatTableRowInsertionProposal, WordChatTableRowDeletionProposal, WordChatProposal, WordChatListCapture } from "../../../dochaus/lib/word-chat-amendment-contract"
import { validWordChatFormatting, validWordChatInlineFormatting, targetWordChatAmendment, validWordChatTableRowInsertion, validWordChatTableRowDeletion, wordChatApplyPlan, wordChatTableCaptureCurrent } from "../../../dochaus/lib/word-chat-amendment-contract"
import { capturedWordChatParagraphs, importedWordChatDifference, reviewedMixedWordChatArtifact, reviewedWordChatArtifact, wordChatOoxmlDifference, sameImportedWordChat, sameWordChatInlineRevisions, sameWordChatParagraphContent, sameNativeWordChatParagraphContent, sameWordChatParagraphProperties, sameWordChatHunkFormatting, authoredWordChatWordingHunks, sameWordChatTargetSnapshot, sameNativeWordChatBodySnapshot, uniformWordChatFormatting, verifiedNativeWordChatFormatting, projectNativeWordChatBody, verifyNativeWordChatRowRemainder, verifyReviewedWordChatRowOracle, assertWordChatRowFormatNeutral, WordChatRowFormatRefusal, assertNativeWordChatInlineTarget, assertNativeWordChatInlineDependencies, verifiedNativeWordChatInlineFormatting, sameNativeInlineWordChatParagraph, assertNativeInlineWordChatRevisionIds } from "./word-chat-host-ooxml"
import type { WordChatAuthoredTableRow, WordChatAuthoredTableRowDeletion } from "../../../dochaus/lib/word-chat-contract"
import { prepareNativeRow, queueNativeRow, type NativeRowContext } from "./word-chat-native-row"
import { verifyNativeWordChatAddedRow } from "./word-chat-native-row-verification"
import { assertNativeWordChatRowDeletionTarget, verifyReviewedWordChatRowDeletionOracle, verifyNativeWordChatDeletedRow } from "./word-chat-native-row-delete"
import { normalizeNativeWordChatCommentIds as normalizeNativeWordChatAnnotationIds } from "./word-chat-native-comments"
import { assertNativeWordChatParagraphInsertionTarget, verifyNativeWordChatParagraphInsertion } from "./word-chat-native-paragraph-insert"
import { assertNativeWordChatParagraphSplitTarget, verifyNativeWordChatParagraphSplit } from "./word-chat-native-paragraph-split"
import { assertNativeWordChatParagraphMergeTarget, verifyNativeWordChatParagraphMerge } from "./word-chat-native-paragraph-merge"
import { assertNativeWordChatParagraphDeletionTarget, verifyNativeWordChatParagraphDeletion } from "./word-chat-native-paragraph-delete"
import { assertNativeWordChatListInsertionTarget, verifyNativeWordChatListInsertion } from "./word-chat-native-list-insert"
import { nativeWordChatDecimalList } from "./word-chat-native-list-context"
import { DOMParser } from "@xmldom/xmldom"
import { nativeWordChatParagraphSplitEligible, nativeWordChatParagraphMergeEligible, nativeWordChatParagraphInsertionEligible, nativeWordChatParagraphDeletionEligible, nativeWordChatListInsertionEligible, wordChatApplyEligibility } from "./word-chat-apply"

type ClientResult<T> = { value: T }
type Collection<T> = { items: T[]; load(properties: string): void }
type TrackedChange = { author: string; type: string; text: string }
type NativeTable = { rowCount: number; isUniform: boolean; values: string[][]; load(properties: string): void; getCell(row: number, column: number): { body: { paragraphs: Collection<Paragraph> } } }
type NativeList = { id: number; paragraphs?: Collection<Paragraph>; load(properties: string): void; insertParagraph?(text: string, location: "End"): Paragraph }
type NativeListItem = { level: number; listString: string; siblingIndex: number; load(properties: string): void }
type Paragraph = { list?: NativeList; listItemOrNullObject?: NativeListItem; leftIndent?: number; firstLineIndent?: number; text: string; uniqueLocalId: string; tableNestingLevel: number; styleBuiltIn: string; isListItem?: boolean; spaceBefore: number; spaceAfter: number; lineSpacing: number; load(properties: string): void; getRange(location?: string): WordChatRange; insertParagraph?(text: string, location: "Before" | "After"): Paragraph; delete?(): void }
type WordChatRange = {
  text: string
  font?: { bold: boolean | null; italic: boolean | null; underline: string | null; name?: string | null; size?: number | null; color?: string | null; load(properties: string): void }
  paragraphs: Collection<Paragraph>
  load(properties: string): void
  getRange(location?: string): WordChatRange
  expandTo(other: WordChatRange): WordChatRange
  compareLocationWith(other: WordChatRange): ClientResult<string>
  search(query: string, options: Record<string, boolean>): Collection<WordChatRange>
  insertOoxml(xml: string, location: "Replace"): WordChatRange
  insertText(text: string, location: "Before" | "After" | "Start" | "End" | "Replace"): WordChatRange
  delete(): void
  getOoxml(): ClientResult<string>
  getReviewedText(version: "Current" | "Original"): ClientResult<string>
  getComments(): Collection<{ author: string; content: string; resolved: boolean }>
  getTrackedChanges(): Collection<TrackedChange>
}

type WordChatDocumentState = {
  current: string
  paragraphs: Array<{ id: string; text: string; style: string; before: number; after: number; line: number }>
  changes: TrackedChange[]
  comments: Array<{ author: string; content: string; resolved: boolean }>
}
export type WordChatContext = {
  document: {
    saved: boolean
    changeTrackingMode: string
    trackFormatting: boolean
    load(properties: string): void
    getSelection(): WordChatRange
    getParagraphByUniqueLocalId(id: string): Paragraph
    body: { text: string; load(properties: string): void; paragraphs: Collection<Paragraph>; tables?: Collection<NativeTable>; getRange(location?: string): WordChatRange }
  }
  sync(): Promise<void>
}
export type WordChatWord = { run<T>(callback: (context: WordChatContext) => Promise<T>): Promise<T> }
export type WordChatApplyResult = { status: "applied" | "already-applied"; tracked: true; saved: false; authors: string[]; target?: WordChatTarget }
export type WordChatApplyInput = { proposalId: string; target: WordChatTarget; replacement: string; formatting?: WordChatParagraphFormat[]; inlineFormatting?: WordChatInlineFormat; paragraphAction?: WordChatProposal["paragraphAction"]; list?: WordChatListCapture; author: (snapshot: { ooxml: string }) => Promise<WordChatPreparedAmendment>; mayDispatch?: () => boolean }
export class WordChatHostError extends Error {
  constructor(message: string, readonly mutationState: "unchanged" | "uncertain" = "unchanged", options?: ErrorOptions) { super(message, options); this.name = "WordChatHostError" }
}
const LIMIT = 24_000
const PARAGRAPH_LIMIT = 64
const SNAPSHOT_LIMIT = 512_000
const SNAPSHOT_BYTES = 1_750_000
const INSIDE = new Set(["Inside", "InsideStart", "InsideEnd", "Equal"])
const SEARCH_OPTIONS = { matchCase: true, matchWildcards: false, matchWholeWord: false, ignorePunct: false, ignoreSpace: false }
const changedMessage = "This passage has changed. Prepare and review a fresh amendment against its current wording. Your Word changes are safe."
function captureMismatch(reason: string) {
  return new WordChatHostError(`${reason} Your document is unchanged. Review the passage in Word, then prepare it again.`)
}
function validate(target: WordChatTarget): void {
  if (!target || !target.id || !target.text?.trim() || target.text.length > LIMIT || !Array.isArray(target.paragraphs) ||
    !target.paragraphs.length || target.paragraphs.length > PARAGRAPH_LIMIT || target.paragraphs.some(p => !p.id || typeof p.text !== "string" || p.text.length > LIMIT) ||
    !Number.isSafeInteger(target.startOffset) || !Number.isSafeInteger(target.endOffset) || target.startOffset < 0 || target.endOffset < 0 ||
    target.startOffset > target.paragraphs[0]!.text.length || target.endOffset > target.paragraphs.at(-1)!.text.length ||
    target.paragraphs.length === 1 && target.endOffset < target.startOffset)
    throw new WordChatHostError("Select a passage of up to 24,000 characters in the main Word document.")
}
function nativeWord(): WordChatWord {
  const word = (window as unknown as { Word?: WordChatWord }).Word
  if (!word) throw new WordChatHostError("Open the linked document in Microsoft Word to read its current wording.")
  return word
}
function supported(name: string, version: string) { return window.Office?.context.requirements.isSetSupported(name, version) ?? false }
function failure(error: unknown, fallback: string): never { if (error instanceof WordChatHostError) throw error; throw new WordChatHostError(fallback, "unchanged", { cause: error }) }

const wordNamespaces = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
])
const packageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage"
const protectedContent = new Set([
  "commentRangeStart", "commentRangeEnd", "commentReference", "ins", "del", "moveFrom", "moveTo",
  "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
  "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange",
  "fldChar", "fldSimple", "instrText", "drawing", "pict", "object", "sdt", "hyperlink",
  "footnoteReference", "endnoteReference", "sectPr",
])
const children = (node: Node) => Array.from(node.childNodes).filter(child => child.nodeType === 1) as Element[]

/** Revisions added since `before`, and earlier revisions no longer found.
 * Word lists contiguous revisions by one author and type as one item, so a new
 * revision beside an earlier one by the same person absorbs it. An absorbed
 * earlier revision counts as kept only as a leading or trailing piece of such
 * an item; the remainder is the new revision. */
function newWordChatRevisions(before: TrackedChange[], after: TrackedChange[]) {
  const signature = (change: TrackedChange) => JSON.stringify([change.author, change.type, change.text])
  const prior = new Map<string, number>()
  for (const change of before) prior.set(signature(change), (prior.get(signature(change)) ?? 0) + 1)
  const exact = after.filter(change => {
    const key = signature(change), count = prior.get(key) ?? 0
    if (count) { prior.set(key, count - 1); return false }
    return true
  })
  const missing = before.filter(change => {
    const key = signature(change), count = prior.get(key) ?? 0
    if (count) { prior.set(key, count - 1); return true }
    return false
  })
  const added = exact.map(change => {
    let text = change.text
    for (let peeled = true; peeled;) {
      peeled = false
      for (const [index, earlier] of missing.entries()) {
        if (earlier.author !== change.author || earlier.type !== change.type || !earlier.text || earlier.text.length >= text.length) continue
        if (text.startsWith(earlier.text)) text = text.slice(earlier.text.length)
        else if (text.endsWith(earlier.text)) text = text.slice(0, -earlier.text.length)
        else continue
        missing.splice(index, 1); peeled = true; break
      }
    }
    // Office.js properties are prototype getters, so copy them explicitly.
    return { author: change.author, type: change.type, text }
  })
  return { added, missing }
}

function sameNativeProjectionGeometry(a: ReturnType<typeof projectNativeWordChatBody>, b: ReturnType<typeof projectNativeWordChatBody>) {
  return a.geometry === b.geometry && a.paragraphs.length === b.paragraphs.length &&
    a.paragraphs.every((p, i) => p.id === b.paragraphs[i]?.id && p.path === b.paragraphs[i]?.path)
}

/** getOoxml exports a Flat OPC package, including unrelated parts and a final
 * body-level section layout even for one selected word. Inspect only selected
 * main-document content. A section break inside a paragraph remains protected;
 * the final body sectPr is package context and is outside our text-only edits. */
function hasProtectedContent(value: string, authored = false, formatting = false, insertedMarks = false, allowComments = false): boolean {
  try {
    const fail = () => { throw new Error("Invalid native XML") }
    const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(value, "application/xml")
    if (doc.doctype) return true
    let content: Element = doc.documentElement
    if (content.namespaceURI === packageNamespace && content.localName === "package") {
      const parts = children(content).filter(part => part.namespaceURI === packageNamespace && part.localName === "part" && part.getAttributeNS(packageNamespace, "name") === "/word/document.xml")
      if (parts.length !== 1) return true
      const data = children(parts[0]!).filter(part => part.namespaceURI === packageNamespace && part.localName === "xmlData")
      if (data.length !== 1 || children(data[0]!).length !== 1) return true
      content = children(data[0]!)[0]!
      if (content.localName !== "document") return true
    }
    if (!wordNamespaces.has(content.namespaceURI ?? "") || !["document", "body", "p", "r"].includes(content.localName)) return true
    const inspect = (node: Element): boolean => {
      const inlineRevision = ["ins", "del"].includes(node.localName) && !(wordNamespaces.has((node.parentNode as Element | null)?.namespaceURI ?? "") && (node.parentNode as Element).localName === "rPr")
      const insertedMark = insertedMarks && node.localName === "ins" && node.getAttributeNS(node.namespaceURI!, "author") === "doc.haus" &&
        (node.parentNode as Element | null)?.localName === "rPr" && (node.parentNode?.parentNode as Element | null)?.localName === "pPr"
      const comment = allowComments && ["commentRangeStart", "commentRangeEnd", "commentReference"].includes(node.localName)
      if (wordNamespaces.has(node.namespaceURI ?? "") && protectedContent.has(node.localName) && !comment &&
        !(authored && ["ins", "del"].includes(node.localName) || formatting && inlineRevision || insertedMark || authored && formatting && node.localName === "pPrChange")) return true
      const elements = children(node)
      return elements.some((child, index) => {
        if (wordNamespaces.has(node.namespaceURI ?? "") && node.localName === "body" &&
          index === elements.length - 1 && child.namespaceURI === node.namespaceURI && child.localName === "sectPr") return false
        return inspect(child)
      })
    }
    return inspect(content)
  } catch { return true }
}

/** Word Find has a short-query limit. Locate boundaries with a short literal
 * token, then check its exact offset inside the verified paragraph. This never
 * searches for a truncated clause or treats a first match as its identity. */
async function point(context: WordChatContext, range: WordChatRange, text: string, offset: number): Promise<WordChatRange> {
  if (offset === 0) return range.getRange("Start")
  if (offset === text.length) return range.getRange("End")
  if (offset < 0 || offset > text.length) throw new WordChatHostError(changedMessage)
  // Paragraph ranges exclude their paragraph marks. Soft line breaks and tabs
  // are supported through their explicit Word Find codes; carets are escaped.
  const token = text.slice(offset, offset + 64)
  if (/[\u0000-\u0008\u000c\u000e-\u001f]/.test(token))
    throw new WordChatHostError("This passage contains a Word object that cannot be amended safely here. Choose ordinary clause text.")
  const query = token.replace(/\^/g, "^^").replace(/\r/g, "^p").replace(/\n|\u000b/g, "^l").replace(/\t/g, "^t")
  const matches = range.search(query, SEARCH_OPTIONS)
  matches.load("items/text")
  await context.sync()
  if (matches.items.length > 512) throw new WordChatHostError("This passage is too repetitive to target safely. Select a more specific clause.")
  const candidates = matches.items.filter((item) => item.text === token).map((item) => ({
    item,
    prefix: range.getRange("Start").expandTo(item.getRange("Start")),
  }))
  for (const candidate of candidates) candidate.prefix.load("text")
  await context.sync()
  const exact = candidates.filter((candidate) => candidate.prefix.text === text.slice(0, offset))
  if (exact.length !== 1) throw new WordChatHostError(changedMessage)
  return exact[0]!.item.getRange("Start")
}

/** Re-bind each reviewed run of a new cell by its cumulative offset and require
 * its exact text and effective font. A mixed cell's whole-range font is null,
 * so only these per-segment reads can prove the reviewed spans. */
async function verifyNativeCellRuns(context: WordChatContext, range: WordChatRange, text: string, runs: WordChatCellRun[]) {
  const offsets = runs.reduce((all, run) => [...all, all.at(-1)! + run.text.length], [0])
  if (offsets.at(-1) !== text.length) throw new Error("Native row segments do not match the cell")
  const points: WordChatRange[] = []
  for (const offset of offsets) points.push(await point(context, range, text, offset))
  const segments = runs.map((run, i) => {
    const segment = points[i]!.expandTo(points[i + 1]!), font = segment.font
    if (!font) throw new Error("Native row segment font unavailable")
    segment.load("text"); font.load("bold,italic,underline")
    return { run, segment, font }
  })
  await context.sync()
  if (segments.some(({ run, segment, font }) => segment.text !== run.text || font.bold !== run.bold || font.italic !== run.italic ||
    font.underline !== (run.underline ? "Single" : "None"))) throw new Error("Native row segment confirmation failed")
}

/** Host-only adapter. No save, session, source or model authority is held here. */
export class WordChatHost {
  private applied = new Map<string, { identity: string; authors: string[]; uncertain: boolean }>()
  private pending = new Map<string, { identity: string; promise: Promise<WordChatApplyResult> }>()
  constructor(private word: WordChatWord = nativeWord(), private supports = supported, private nativeTextEnabled = true) {}
  private ready() {
    if (!this.supports("WordApi", "1.6")) throw new WordChatHostError("Reviewed amendments require a current Microsoft 365 version of Word. Chat and Save remain available.")
  }
  async capture(includeSelection = false): Promise<WordChatSnapshot> {
    const structured = this.supports("WordApi", "1.6")
    try {
      const snapshot = await this.readSnapshot(structured, includeSelection)
      if (structured) {
        try {
          const table = await this.readTableCapture(snapshot)
          if (table && new TextEncoder().encode(JSON.stringify({ ...snapshot, table })).length <= SNAPSHOT_BYTES) snapshot.table = table
        } catch { /* Optional geometry must never prevent ordinary chat. */ }
        try {
          const list = await this.readListCapture(snapshot)
          if (list && new TextEncoder().encode(JSON.stringify({ ...snapshot, list })).length <= SNAPSHOT_BYTES) snapshot.list = list
        } catch { /* Optional list context must never prevent ordinary chat. */ }
      }
      return snapshot
    }
    catch (error) {
      if (includeSelection || error instanceof WordChatHostError || !structured) failure(error, "Word could not read its current text. Keep your message and try again.")
      // Some licensed hosts advertise the API without providing paragraph IDs.
      // Read current text in a fresh Office batch; never substitute a saved copy.
      return this.readSnapshot(false, false)
    }
  }
  private async observeNativeLists(context: WordChatContext, projection: ReturnType<typeof projectNativeWordChatBody>, selected?: string[]): Promise<WordChatListCapture[]> {
    const all = context.document.body.paragraphs
    all.load("items/text,items/uniqueLocalId,items/tableNestingLevel,items/styleBuiltIn,items/isListItem")
    await context.sync()
    if (all.items.length !== projection.paragraphs.length || all.items.some((p,i) => p.uniqueLocalId !== projection.paragraphs[i]!.id || p.text !== projection.paragraphs[i]!.text || p.tableNestingLevel !== projection.paragraphs[i]!.nesting)) throw captureMismatch("The Word list changed during capture.")
    const items = all.items.filter(p => p.isListItem === true)
    for (const p of items) {
      if (!p.list?.paragraphs || !p.listItemOrNullObject) throw new Error("list_api_unavailable")
      p.list.load("id"); p.listItemOrNullObject.load("level,listString,siblingIndex")
      p.list.paragraphs!.load("items/text,items/uniqueLocalId,items/tableNestingLevel,items/styleBuiltIn,items/isListItem")
    }
    if (!items.length) return []
    await context.sync()
    const groups = new Map<number, Paragraph[]>()
    for (const p of items) {
      const list = p.list!, id = list.id
      if (!Number.isSafeInteger(id) || id < 0) throw new Error("list_id_invalid")
      const existing = groups.get(id) ?? []; existing.push(p); groups.set(id, existing)
    }
    const eligibleGroups = [...groups.entries()].filter(([, members]) => !selected || selected.every(id => members.some(p => p.uniqueLocalId === id)))
    if (eligibleGroups.length !== 1) return []
    const captures: WordChatListCapture[] = []
    for (const [id, members] of eligibleGroups) {
      const nativeMembers = members[0]!.list!.paragraphs!.items
      if (nativeMembers.length !== members.length || nativeMembers.some((p,i) => p.uniqueLocalId !== members[i]!.uniqueLocalId || p.text !== members[i]!.text)) continue
      const decimal = members.every((p,i) => p.listItemOrNullObject!.level === 0 && p.listItemOrNullObject!.siblingIndex === i && p.listItemOrNullObject!.listString === `${i + 1}.`)
      const capture: WordChatListCapture = decimal
        ? { id, format:"decimal", members: members.map(p => ({id:p.uniqueLocalId,text:p.text,level:p.listItemOrNullObject!.level,label:p.listItemOrNullObject!.listString})) }
        : { id, format:"native", members: members.map(p => ({id:p.uniqueLocalId,text:p.text,level:p.listItemOrNullObject!.level,label:p.listItemOrNullObject!.listString,siblingIndex:p.listItemOrNullObject!.siblingIndex})) }
      if (members.some(p => p.tableNestingLevel !== 0 || p.styleBuiltIn !== "Normal") || !nativeWordChatDecimalList(projection,capture)) continue
      captures.push(capture)
    }
    return captures
  }

  private async readListCapture(snapshot: WordChatSnapshot): Promise<WordChatListCapture | undefined> {
    return this.word.run(async context => {
      const all = context.document.body.paragraphs, xml = context.document.body.getRange("Whole").getOoxml()
      all.load("items/text,items/uniqueLocalId,items/tableNestingLevel")
      await context.sync()
      if (all.items.length !== snapshot.paragraphs.length || all.items.some((p,i) => p.uniqueLocalId !== snapshot.paragraphs[i]!.id || p.text !== snapshot.paragraphs[i]!.text)) return
      const projection = projectNativeWordChatBody(xml.value,all.items.map(p => ({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})))
      const selected = snapshot.selection?.paragraphs.map(p => p.id)
      const lists = await this.observeNativeLists(context,projection,selected)
      const matches = selected ? lists.filter(l => selected.every(id => l.members.some(p => p.id === id))) : lists
      return matches.length === 1 ? matches[0] : undefined
    })
  }
  private async readTableCapture(snapshot: WordChatSnapshot): Promise<WordChatTableCapture | undefined> {
    return this.word.run(async context => {
      const body = context.document.body, tables = body.tables, paragraphs = body.paragraphs
      if (!tables) return
      tables.load("items")
      await context.sync()
      if (!tables.items.length || !snapshot.selection && tables.items.length !== 1) return
      const xml = body.getRange("Whole").getOoxml()
      paragraphs.load("items/text,items/uniqueLocalId,items/tableNestingLevel")
      for (const table of tables.items) table.load("rowCount,isUniform,values")
      await context.sync()
      const inventory = paragraphs.items.map(p => ({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel}))
      if (inventory.length !== snapshot.paragraphs.length || inventory.some((p, i) => p.id !== snapshot.paragraphs[i]!.id || p.text !== snapshot.paragraphs[i]!.text)) return
      const projection = projectNativeWordChatBody(xml.value, inventory)
      const fail = () => { throw new Error("Invalid table capture XML") }
      const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml.value, "application/xml")
      const parts = Array.from(doc.getElementsByTagNameNS(packageNamespace, "part")).filter(p => p.getAttributeNS(packageNamespace, "name") === "/word/document.xml")
      if (parts.length !== 1) return
      const bodies = Array.from(parts[0]!.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "body"))
      if (bodies.length !== 1) return
      const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      const bodyChildren = children(bodies[0]!), xmlTables = bodyChildren.filter(n => n.namespaceURI === W && n.localName === "tbl")
      if (xmlTables.length !== tables.items.length) return
      const candidates: Array<{ tableIndex: number; rows: Array<Array<{ collection: Collection<Paragraph>; path: string }>> }> = []
      for (const [tableIndex, table] of tables.items.entries()) {
        if (!table.isUniform || !Number.isSafeInteger(table.rowCount) || table.rowCount < 1 || table.rowCount > 1000 ||
          !Array.isArray(table.values) || table.values.length !== table.rowCount || !Array.isArray(table.values[0])) continue
        const width = table.values[0]!.length
        if (!width || width > 32 || table.rowCount * width > 20_000 || table.values.some(row => !Array.isArray(row) || row.length !== width)) continue
        const tableXml = xmlTables[tableIndex]!, tableChildren = children(tableXml), rows = tableChildren.filter(n => n.namespaceURI === W && n.localName === "tr")
        if (rows.length !== table.rowCount || tableChildren.some(n => n.namespaceURI !== W || !["tblPr", "tblGrid", "tr"].includes(n.localName))) continue
        const xmlRows: Element[][] = []
        let valid = true
        for (const row of rows) {
          const rowChildren = children(row), cells = rowChildren.filter(n => n.namespaceURI === W && n.localName === "tc")
          if (cells.length !== width || rowChildren.some(n => n.namespaceURI !== W || !["trPr", "tc"].includes(n.localName)) ||
            Array.from(row.getElementsByTagNameNS(W, "gridBefore")).length || Array.from(row.getElementsByTagNameNS(W, "gridAfter")).length) { valid = false; break }
          for (const cell of cells) {
            const cellChildren = children(cell)
            if (cellChildren.some(n => n.namespaceURI !== W || !["tcPr", "p"].includes(n.localName)) || cellChildren.filter(n => n.localName === "p").length !== 1 ||
              ["gridSpan", "vMerge", "hMerge"].some(name => cell.getElementsByTagNameNS(W, name).length)) { valid = false; break }
          }
          if (!valid) break
          xmlRows.push(cells)
        }
        if (!valid) continue
        const path = `body/tbl[${bodyChildren.indexOf(tableXml)}]`
        const capturedRows = xmlRows.map((cells, ri) => cells.map((cell, ci) => {
          const collection = table.getCell(ri, ci).body.paragraphs
          collection.load("items/text,items/uniqueLocalId,items/tableNestingLevel")
          return { collection, path: `${path}/tr[${tableChildren.indexOf(rows[ri]!)}]/tc[${children(rows[ri]!).indexOf(cell)}]/p[${children(cell).findIndex(n => n.namespaceURI === W && n.localName === "p")}]` }
        }))
        candidates.push({ tableIndex, rows: capturedRows })
      }
      if (!candidates.length) return
      await context.sync()
      const captures = candidates.flatMap(candidate => {
        const rows = candidate.rows.map(row => row.map(({ collection, path }) => {
          if (collection.items.length !== 1) return undefined
          const p = collection.items[0]!, projected = projection.paragraphs.find(item => item.id === p.uniqueLocalId)
          return p.tableNestingLevel === 1 && projected?.path === path && projected.text === p.text ? {id:p.uniqueLocalId,text:p.text} : undefined
        }))
        if (rows.some(row => row.some(cell => !cell))) return []
        const capture = {tableIndex:candidate.tableIndex,rows:rows as Array<Array<{id:string;text:string}>>}
        return wordChatTableCaptureCurrent(snapshot, capture) ? [capture] : []
      })
      const selected = snapshot.selection?.paragraphs.map(p => p.id)
      const matches = selected ? captures.filter(capture => selected.every(id => capture.rows.some(row => row.some(cell => cell.id === id)))) : captures
      return matches.length === 1 ? matches[0] : undefined
    })
  }
  private async readSnapshot(structured: boolean, includeSelection: boolean): Promise<WordChatSnapshot> {
    return this.word.run(async context => {
      context.document.load("saved")
      const body = context.document.body
      body.load("text")
      if (structured) body.paragraphs.load("items/text,items/uniqueLocalId")
      await context.sync()
      const paragraphs = structured ? body.paragraphs.items.map(p => ({id:p.uniqueLocalId,text:p.text})) : []
      const full = structured ? body.paragraphs.items.map(p => p.text).join("\r") : body.text
      if (full.length > SNAPSHOT_LIMIT || paragraphs.length > 20_000)
        throw new WordChatHostError("This Word document exceeds the chat limit of 512,000 characters or 20,000 paragraphs. Use a shorter linked document, or open the matter in doc.haus to work with it. Your document is unchanged.")
      const snapshot: WordChatSnapshot = { text:full, paragraphs, saved: context.document.saved,
        capturedAt: new Date().toISOString(), truncated:false }
      if (includeSelection) snapshot.selection = await this.readSelection(context)
      // The receiver also carries instructions, attachment identifiers and a
      // pinned selection in its bounded request. Large Unicode buffers consume
      // more bytes than their character count; refuse clearly before transport.
      if (new TextEncoder().encode(JSON.stringify(snapshot)).length > SNAPSHOT_BYTES)
        throw new WordChatHostError("This Word document is too large to send as chat context. Use a shorter linked document, or open the matter in doc.haus. Your document is unchanged.")
      return snapshot
    })
  }
  async selection(): Promise<WordChatTarget | undefined> {
    this.ready()
    try { return await this.word.run(context => this.readSelection(context)) }
    catch (error) { failure(error, "Word could not read that selection. Select ordinary text in the main document and try again.") }
  }
  private async readSelection(context: WordChatContext): Promise<WordChatTarget | undefined> {
    this.ready()
    const selection = context.document.getSelection()
    selection.load("text")
    selection.paragraphs.load("items/text,items/uniqueLocalId")
    context.document.load("saved")
    const location = selection.compareLocationWith(context.document.body.getRange())
    await context.sync()
    if (!selection.text.trim()) return
    if (!INSIDE.has(location.value)) throw new WordChatHostError("Select text in the main document. Headers, footnotes and text boxes are not supported for amendments yet.")
    const items = selection.paragraphs.items
    if (!items.length || items.length > PARAGRAPH_LIMIT) throw new WordChatHostError("Select one clause or a short group of paragraphs.")
    const start = items[0]!.getRange("Start").expandTo(selection.getRange("Start"))
    const end = items.at(-1)!.getRange("Start").expandTo(selection.getRange("End"))
    start.load("text"); end.load("text")
    await context.sync()
    const target: WordChatTarget = { id: crypto.randomUUID(), text: selection.text,
      paragraphs: items.map(p => ({id: p.uniqueLocalId,text: p.text})), startOffset: start.text.length, endOffset: end.text.length,
      saved: context.document.saved, capturedAt: new Date().toISOString() }
    validate(target)
    return target
  }
  private async resolve(context: WordChatContext, target: WordChatTarget, allowNativeCell = false): Promise<{ range: WordChatRange; paragraphs: Paragraph[] }> {
    validate(target)
    const paragraphs = target.paragraphs.map((p) => context.document.getParagraphByUniqueLocalId(p.id))
    for (const paragraph of paragraphs) paragraph.load("text,uniqueLocalId,tableNestingLevel")
    await context.sync()
    // Word can alter list numbering outside a table when importing a scoped
    // paragraph fragment into a cell. Until that transport is proven safe,
    // refuse before authoring or mutation, including after the freshness read.
    if (paragraphs.some(paragraph => paragraph.tableNestingLevel !== 0 && !(allowNativeCell && paragraphs.length === 1 && paragraph.tableNestingLevel === 1)))
      throw new WordChatHostError("Amendments in Word tables are not supported safely yet. Make this change directly in Word; your document is unchanged.")
    const changed = paragraphs.findIndex((p, i) => p.uniqueLocalId !== target.paragraphs[i]!.id || p.text !== target.paragraphs[i]!.text)
    if (changed >= 0)
      throw captureMismatch(`Word could not match reviewed paragraph ${changed + 1} (${paragraphs[changed]!.text.length} Word characters; ${target.paragraphs[changed]!.text.length} reviewed).`)
    const union = paragraphs[0]!.getRange("Content").expandTo(paragraphs.at(-1)!.getRange("Content"))
    union.paragraphs.load("items/uniqueLocalId")
    const location = union.compareLocationWith(context.document.body.getRange())
    await context.sync()
    if (!INSIDE.has(location.value) || union.paragraphs.items.length !== paragraphs.length || union.paragraphs.items.some((p, i) => p.uniqueLocalId !== paragraphs[i]!.uniqueLocalId))
      throw captureMismatch(`Word returned a different paragraph range (${union.paragraphs.items.length} paragraphs; ${paragraphs.length} reviewed).`)
    const start = await point(context, paragraphs[0]!.getRange("Content"), paragraphs[0]!.text, target.startOffset)
    const end = await point(context, paragraphs.at(-1)!.getRange("Content"), paragraphs.at(-1)!.text, target.endOffset)
    const range = start.expandTo(end)
    range.load("text")
    await context.sync()
    if (range.text !== target.text) throw captureMismatch(`Word returned different text for the reviewed range (${range.text.length} Word characters; ${target.text.length} reviewed).`)
    return { range, paragraphs }
  }

  /** Whole-document evidence for the native text transport: accepted text,
   * every paragraph's visible properties, prior revision signatures and the
   * existing comment bodies. This never includes unrelated Word metadata. */
  private async documentState(context: WordChatContext): Promise<WordChatDocumentState> {
    const body = context.document.body, all = body.paragraphs
    all.load("items/text,items/uniqueLocalId,items/styleBuiltIn,items/spaceBefore,items/spaceAfter,items/lineSpacing")
    const current = body.getRange("Content").getReviewedText("Current")
    const changes = body.getRange().getTrackedChanges()
    changes.load("items/author,items/type,items/text")
    const comments = body.getRange().getComments()
    comments.load("items/author,items/content,items/resolved")
    await context.sync()
    return {
      current: current.value,
      paragraphs: all.items.map(p => ({ id:p.uniqueLocalId, text:p.text, style:p.styleBuiltIn, before:p.spaceBefore, after:p.spaceAfter, line:p.lineSpacing })),
      changes: changes.items.map(change => ({ author:change.author, type:change.type, text:change.text })),
      comments: comments.items.map(comment => ({ author:comment.author, content:comment.content, resolved:comment.resolved })),
    }
  }

  async apply(input: WordChatApplyInput): Promise<WordChatApplyResult> {
    this.ready()
    validate(input.target)
    if (typeof input.author !== "function") throw new WordChatHostError("Reconnect to doc.haus to author this amendment before applying it. Your document is unchanged.")
    if (!input.proposalId || typeof input.replacement !== "string" || input.replacement.length > LIMIT || input.replacement === input.target.text && input.formatting === undefined && input.inlineFormatting === undefined)
      throw new WordChatHostError("Prepare and review a change before applying it.")
    const identity = JSON.stringify([input.target, input.replacement, input.formatting, input.inlineFormatting, input.paragraphAction, input.list])
    const pending = this.pending.get(input.proposalId)
    if (pending) {
      if (pending.identity !== identity) throw new WordChatHostError("This proposal has different wording. Prepare a fresh proposal.")
      return pending.promise
    }
    const observed = this.applied.get(input.proposalId)
    if (observed) {
      if (observed.uncertain) throw new WordChatHostError("Word may already contain this change. Inspect the clause and use Undo if needed before preparing a fresh proposal.", "uncertain")
      if (observed.identity !== identity) throw new WordChatHostError("This proposal has different wording. Prepare a fresh proposal.")
      return { status: "already-applied", tracked: true, saved: false, authors: observed.authors }
    }
    const insertion = this.nativeTextEnabled && nativeWordChatParagraphInsertionEligible(input.target, input.replacement, input.formatting, input.inlineFormatting)
    const listInsertion = input.paragraphAction === "append_list_item" || input.paragraphAction === "insert_list_item_before"
    if (listInsertion && (!this.nativeTextEnabled || !nativeWordChatListInsertionEligible(input.target, input.replacement, input.paragraphAction, input.list, input.formatting, input.inlineFormatting)))
      throw new WordChatHostError("Review a supported existing-list insertion before applying it. Your document is unchanged.")
    if (input.paragraphAction === "split" && (!this.nativeTextEnabled || !nativeWordChatParagraphSplitEligible(input.target,input.replacement,input.paragraphAction,input.formatting,input.inlineFormatting)))
      throw new WordChatHostError("Review a complete paragraph split before applying it. Your document is unchanged.")
    if (input.paragraphAction === "merge" && (!this.nativeTextEnabled || !nativeWordChatParagraphMergeEligible(input.target, input.replacement, input.paragraphAction, input.formatting, input.inlineFormatting)))
      throw new WordChatHostError("Review two complete adjacent paragraphs before merging them. Your document is unchanged.")
    if (input.paragraphAction !== undefined && input.paragraphAction !== "split" && input.paragraphAction !== "merge" && !listInsertion && (!this.nativeTextEnabled || !nativeWordChatParagraphDeletionEligible(input.target, input.replacement, input.paragraphAction, input.formatting, input.inlineFormatting)))
      throw new WordChatHostError("Review deletion of one complete paragraph before applying it. Your document is unchanged.")
    const promise = input.paragraphAction === "split" ? this.applyParagraphSplitOnce(input, identity) : input.paragraphAction === "merge" ? this.applyParagraphMergeOnce(input, identity) : listInsertion ? this.applyListInsertionOnce(input, identity) : input.paragraphAction !== undefined ? this.applyParagraphDeletionOnce(input, identity) : input.inlineFormatting !== undefined ? this.applyInlineOnce(input, identity) : insertion ? this.applyParagraphInsertionOnce(input, identity) : this.applyOnce(input, identity)
    this.pending.set(input.proposalId, { identity, promise })
    try { return await promise } finally { this.pending.delete(input.proposalId) }
  }

  async applyRow(proposal: WordChatTableRowInsertionProposal, exportCurrent: () => Promise<Uint8Array>,
    author: (bytes: Uint8Array) => Promise<WordChatAuthoredTableRow>, mayDispatch?: () => boolean): Promise<WordChatApplyResult> {
    this.ready()
    if (!this.supports("WordApiDesktop", "1.4")) throw new WordChatHostError("Appending a reviewed row requires a supported desktop version of Word.")
    const rows = proposal?.table?.rows
    const snapshot: WordChatSnapshot = { text: "", paragraphs: Array.isArray(rows) ? rows.flat() : [], saved: false, capturedAt: "", truncated: false, table: proposal?.table }
    if (!validWordChatTableRowInsertion(snapshot, proposal) || typeof exportCurrent !== "function" || typeof author !== "function") throw new WordChatHostError("Review a complete table row before applying it. Your document is unchanged.")
    const identity = JSON.stringify(proposal), pending = this.pending.get(proposal.id), observed = this.applied.get(proposal.id)
    if (pending) {
      if (pending.identity !== identity) throw new WordChatHostError("This proposal has changed. Prepare a fresh proposal.")
      return pending.promise
    }
    if (observed) {
      if (observed.uncertain) throw new WordChatHostError("Word may already contain this row. Inspect it and use Undo if needed before preparing a fresh proposal.", "uncertain")
      if (observed.identity !== identity) throw new WordChatHostError("This proposal has changed. Prepare a fresh proposal.")
      return { status: "already-applied", tracked: true, saved: false, authors: observed.authors }
    }
    const promise = this.applyRowOnce(proposal, identity, exportCurrent, author, mayDispatch)
    this.pending.set(proposal.id, { identity, promise })
    try { return await promise } finally { this.pending.delete(proposal.id) }
  }

  private async applyRowOnce(proposal: WordChatTableRowInsertionProposal, identity: string, exportCurrent: () => Promise<Uint8Array>,
    author: (bytes: Uint8Array) => Promise<WordChatAuthoredTableRow>, mayDispatch?: () => boolean): Promise<WordChatApplyResult> {
    try {
      const source = await exportCurrent()
      if (!(source instanceof Uint8Array) || !source.length || source.length > 16 * 1024 * 1024) throw new WordChatHostError("The current Word document is too large or unavailable to prepare this row. Your document is unchanged.")
      assertWordChatRowFormatNeutral(source, proposal)
      const oracle = await author(source)
      await verifyReviewedWordChatRowOracle(source, oracle, proposal)
      return await this.word.run(async context => {
        if (!context.document.body.tables) throw new WordChatHostError("Word cannot read the reviewed table. Your document is unchanged.")
        const native = context as unknown as NativeRowContext, prepared = await prepareNativeRow(native, proposal.table)
        const before = await this.documentState(context), all = context.document.body.paragraphs
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        const body = context.document.body.getRange("Whole").getOoxml()
        await context.sync()
        const inventory = all.items.map(p => ({ id:p.uniqueLocalId, text:p.text, tableNestingLevel:p.tableNestingLevel }))
        const beforeXml = body.value
        const projection = projectNativeWordChatBody(beforeXml, inventory)
        const fail = () => { throw new Error("Invalid table XML") }
        const xml = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(beforeXml, "application/xml")
        const parts = Array.from(xml.getElementsByTagNameNS(packageNamespace, "part")).filter(p => p.getAttributeNS(packageNamespace, "name") === "/word/document.xml")
        const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", bodies = parts.length === 1 ? Array.from(parts[0]!.getElementsByTagNameNS(W, "body")) : []
        if (bodies.length !== 1) throw captureMismatch("Word returned a different table structure.")
        const bodyChildren = children(bodies[0]!), tables = bodyChildren.filter(e => e.namespaceURI === W && e.localName === "tbl"), table = tables[proposal.table.tableIndex]
        if (!table) throw captureMismatch("Word returned a different table structure.")
        const tableChildren = children(table), rows = tableChildren.filter(e => e.namespaceURI === W && e.localName === "tr")
        if (rows.length !== proposal.table.rows.length || tableChildren.some(e => e.namespaceURI !== W || !["tblPr", "tblGrid", "tr"].includes(e.localName))) throw captureMismatch("Word returned a different table structure.")
        for (const [ri, row] of rows.entries()) {
          const rowChildren = children(row), cells = rowChildren.filter(e => e.namespaceURI === W && e.localName === "tc")
          if (cells.length !== proposal.cells.length || rowChildren.some(e => e.namespaceURI !== W || !["trPr", "tc"].includes(e.localName)) || ["gridBefore", "gridAfter"].some(n => row.getElementsByTagNameNS(W,n).length)) throw captureMismatch("Word returned a different table structure.")
          for (const [ci, cell] of cells.entries()) {
            const cellChildren = children(cell), paragraphs = cellChildren.filter(e => e.namespaceURI === W && e.localName === "p")
            if (paragraphs.length !== 1 || cellChildren.some(e => e.namespaceURI !== W || !["tcPr", "p"].includes(e.localName)) || ["gridSpan", "vMerge", "hMerge"].some(n => cell.getElementsByTagNameNS(W,n).length)) throw captureMismatch("Word returned an unsupported table structure.")
            const path = `body/tbl[${bodyChildren.indexOf(table)}]/tr[${tableChildren.indexOf(row)}]/tc[${rowChildren.indexOf(cell)}]/p[${cellChildren.indexOf(paragraphs[0]!)}]`, expected = proposal.table.rows[ri]![ci]!, found = projection.paragraphs.filter(p => p.path === path)
            if (found.length !== 1 || found[0]!.id !== expected.id || found[0]!.text !== expected.text) throw captureMismatch("Word returned a different reviewed table.")
          }
        }
        if (mayDispatch && !mayDispatch()) throw new WordChatHostError("The Word pane closed before the row was applied. Your document is unchanged.")
        this.applied.set(proposal.id, { identity, authors: [], uncertain: true })
        queueNativeRow(native, prepared, proposal.cells, proposal.runs)
        await context.sync()
        context.document.load("changeTrackingMode,trackFormatting")
        prepared.table.load("rowCount,isUniform,values")
        const added = proposal.cells.map((_, col) => { const paragraphs = prepared.table.getCell(prepared.rowCount, col).body.paragraphs; paragraphs.load("items/text,items/uniqueLocalId,items/tableNestingLevel"); return paragraphs })
        const afterAll = context.document.body.paragraphs
        afterAll.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        const afterXml = context.document.body.getRange("Whole").getOoxml()
        await context.sync()
        // A later body.paragraphs load can refresh aliased Office proxies and
        // invalidate properties already loaded on these collections.
        const observedTracking = context.document.changeTrackingMode, observedTrackFormatting = context.document.trackFormatting
        const observedTable = { rowCount:prepared.table.rowCount, isUniform:prepared.table.isUniform, values:prepared.table.values.map(row => [...row]) }
        const addedParagraphs = added.map(collection => {
          if (collection.items.length !== 1) throw new Error("Unexpected new cell paragraphs")
          const p = collection.items[0]!
          return { id:p.uniqueLocalId, text:p.text, tableNestingLevel:p.tableNestingLevel }
        })
        const afterInventory = afterAll.items.map(p => ({ id:p.uniqueLocalId, text:p.text, tableNestingLevel:p.tableNestingLevel }))
        const observedXml = afterXml.value
        // A cell with any reviewed formatting is proved per segment below;
        // every other new cell must read uniformly plain as a whole.
        const formatted = proposal.cells.map((_, i) => proposal.runs?.[i]?.some(run => run.bold || run.italic || run.underline) ? proposal.runs[i] : undefined)
        const contents = addedParagraphs.map(paragraph => context.document.getParagraphByUniqueLocalId(paragraph.id).getRange("Content") as WordChatRange & { font: { bold: boolean | null; italic: boolean | null; underline: string | null; load(properties: string): void } })
        const fonts = contents.map((range, i) => { if (formatted[i]) return undefined; range.font.load("bold,italic,underline"); return range.font })
        const after = await this.documentState(context)
        if (observedTracking !== prepared.tracking || observedTrackFormatting !== prepared.trackFormatting || !observedTable.isUniform || observedTable.rowCount !== prepared.rowCount + 1 || observedTable.values.length !== prepared.rowCount + 1 || observedTable.values.some(row => row.length !== prepared.width) || addedParagraphs.some((paragraph, i) => paragraph.text !== proposal.cells[i] || paragraph.tableNestingLevel !== 1) || fonts.some(font => font && (font.bold !== false || font.italic !== false || font.underline !== "None")) || JSON.stringify(before.comments) !== JSON.stringify(after.comments)) throw new Error("Native row confirmation failed")
        for (const [i, runs] of formatted.entries()) if (runs) await verifyNativeCellRuns(context, contents[i]!, proposal.cells[i]!, runs)
        const remainder = verifyNativeWordChatRowRemainder(beforeXml, observedXml, inventory, afterInventory, proposal.table, proposal.cells)
        if (remainder.addedParagraphIds.some((id, i) => id !== addedParagraphs[i]!.id)) throw new Error("Native row identity mismatch")
        const authors = verifyNativeWordChatAddedRow(beforeXml, observedXml, proposal.table, proposal.cells, proposal.runs)
        this.applied.set(proposal.id, { identity, authors, uncertain: false })
        return { status: "applied", tracked: true, saved: false, authors }
      })
    } catch (error) {
      if (this.applied.get(proposal.id)?.uncertain) throw new WordChatHostError("Word may have applied the row, but its tracked result could not be confirmed. Inspect it and use Undo if needed; do not apply it again.", "uncertain", { cause: error })
      if (error instanceof WordChatRowFormatRefusal) throw new WordChatHostError(error.message, "unchanged", { cause: error })
      failure(error, "Word could not prepare the reviewed row. Your document is unchanged.")
    }
  }

  async applyRowDeletion(proposal: WordChatTableRowDeletionProposal, exportCurrent: () => Promise<Uint8Array>,
    author: (bytes: Uint8Array) => Promise<WordChatAuthoredTableRowDeletion>, mayDispatch?: () => boolean): Promise<WordChatApplyResult> {
    this.ready()
    if (!this.supports("WordApiDesktop", "1.4")) throw new WordChatHostError("Deleting a reviewed row requires a supported desktop version of Word.")
    const rows = proposal?.table?.rows
    const snapshot: WordChatSnapshot = {text:"", paragraphs:Array.isArray(rows) ? rows.flat() : [], saved:false, capturedAt:"", truncated:false, table:proposal?.table}
    if (!validWordChatTableRowDeletion(snapshot, proposal) || typeof exportCurrent !== "function" || typeof author !== "function") throw new WordChatHostError("Review a complete table row before deleting it. Your document is unchanged.")
    const identity=JSON.stringify(proposal), pending=this.pending.get(proposal.id), observed=this.applied.get(proposal.id)
    if (pending) { if (pending.identity!==identity) throw new WordChatHostError("This proposal has changed. Prepare a fresh proposal."); return pending.promise }
    if (observed) {
      if (observed.uncertain) throw new WordChatHostError("Word may already contain this deletion. Inspect it and use Undo if needed before preparing a fresh proposal.","uncertain")
      if (observed.identity!==identity) throw new WordChatHostError("This proposal has changed. Prepare a fresh proposal.")
      return {status:"already-applied",tracked:true,saved:false,authors:observed.authors}
    }
    const promise=this.applyRowDeletionOnce(proposal,identity,exportCurrent,author,mayDispatch)
    this.pending.set(proposal.id,{identity,promise})
    try { return await promise } finally { this.pending.delete(proposal.id) }
  }

  private async applyRowDeletionOnce(proposal: WordChatTableRowDeletionProposal, identity: string, exportCurrent: () => Promise<Uint8Array>,
    author: (bytes: Uint8Array) => Promise<WordChatAuthoredTableRowDeletion>, mayDispatch?: () => boolean): Promise<WordChatApplyResult> {
    try {
      const source=await exportCurrent()
      if (!(source instanceof Uint8Array) || !source.length || source.length>16*1024*1024) throw new WordChatHostError("The current Word document is too large or unavailable to prepare this row. Your document is unchanged.")
      const oracle=await author(source)
      await verifyReviewedWordChatRowDeletionOracle(source,oracle,proposal)
      return await this.word.run(async context => {
        if (!context.document.body.tables) throw captureMismatch("Word cannot read the reviewed table.")
        const native=context as unknown as NativeRowContext, prepared=await prepareNativeRow(native,proposal.table)
        type DeleteRow={rowIndex:number;cellCount:number;isHeader:boolean;values:string[][];delete():void}
        const table=prepared.table as typeof prepared.table & { rows:Collection<DeleteRow>; getCell(row:number,column:number): {body:{paragraphs:Collection<Paragraph>;getRange(location:string):WordChatRange}} }
        table.load("rowCount,isUniform,values");table.rows.load("items/rowIndex,items/cellCount,items/isHeader,items/values")
        const all=context.document.body.paragraphs, body=context.document.body.getRange("Whole").getOoxml(), original=context.document.body.getRange("Content").getReviewedText("Original")
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        const cells=proposal.cells.map((_,i)=>{const r=(table.getCell(proposal.rowIndex,i).body as unknown as {getRange(location:string):WordChatRange}).getRange("Content");return {current:r.getReviewedText("Current"),original:r.getReviewedText("Original")}})
        await context.sync()
        const inventory=all.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})), beforeXml=body.value, originalText=original.value
        const expectedValues=proposal.table.rows.map(row=>row.map(cell=>cell.text))
        if (!table.isUniform || table.rowCount!==proposal.table.rows.length || JSON.stringify(table.values)!==JSON.stringify(expectedValues) || table.rows.items.length!==table.rowCount || table.rows.items.some((r,i)=>r.rowIndex!==i || r.cellCount!==prepared.width) || table.rows.items[proposal.rowIndex]!.isHeader!==false || JSON.stringify(table.rows.items[proposal.rowIndex]!.values)!==JSON.stringify([proposal.cells]) || cells.some((c,i)=>c.current.value!==proposal.cells[i] || c.original.value!==proposal.cells[i])) throw captureMismatch("The reviewed row or table changed.")
        assertNativeWordChatRowDeletionTarget(beforeXml,inventory,proposal)
        const beforeComments=await this.documentState(context)
        const freshBody=context.document.body.getRange("Whole").getOoxml(), freshAll=context.document.body.paragraphs
        freshAll.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        table.load("rowCount,isUniform,values");table.rows.load("items/rowIndex,items/cellCount,items/isHeader,items/values")
        await context.sync()
        const freshInventory=freshAll.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel}))
        if (!sameNativeWordChatBodySnapshot(beforeXml,freshBody.value) || JSON.stringify(inventory)!==JSON.stringify(freshInventory) || !table.isUniform || table.rowCount!==prepared.rowCount || JSON.stringify(table.values)!==JSON.stringify(expectedValues) || table.rows.items.length!==prepared.rowCount || table.rows.items.some((r,i)=>r.rowIndex!==i || r.cellCount!==prepared.width) || table.rows.items[proposal.rowIndex]!.isHeader!==false || JSON.stringify(table.rows.items[proposal.rowIndex]!.values)!==JSON.stringify([proposal.cells])) throw captureMismatch("The reviewed table changed before deletion.")
        assertNativeWordChatRowDeletionTarget(freshBody.value,freshInventory,proposal)
        if (mayDispatch && !mayDispatch()) throw new WordChatHostError("The Word pane closed before the row was deleted. Your document is unchanged.")
        this.applied.set(proposal.id,{identity,authors:[],uncertain:true})
        context.document.changeTrackingMode="TrackAll"
        table.rows.items[proposal.rowIndex]!.delete()
        context.document.changeTrackingMode=prepared.tracking
        context.document.trackFormatting=prepared.trackFormatting
        await context.sync()
        context.document.load("changeTrackingMode,trackFormatting")
        table.load("rowCount,isUniform,values")
        const afterAll=context.document.body.paragraphs, afterXml=context.document.body.getRange("Whole").getOoxml(), afterOriginal=context.document.body.getRange("Content").getReviewedText("Original")
        afterAll.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        const afterCells=proposal.cells.map((_,i)=>{const r=(table.getCell(proposal.rowIndex,i).body as unknown as {getRange(location:string):WordChatRange}).getRange("Content");return {current:r.getReviewedText("Current"),original:r.getReviewedText("Original")}})
        await context.sync()
        const tracking=context.document.changeTrackingMode, trackFormatting=context.document.trackFormatting, values=table.values.map(row=>[...row]), count=table.rowCount, uniform=table.isUniform
        const afterInventory=afterAll.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})), afterXmlValue=afterXml.value, bodyOriginal=afterOriginal.value
        const reviewed=afterCells.map(c=>({current:c.current.value,original:c.original.value}))
        const expectedAfter=expectedValues.map((row,i)=>i===proposal.rowIndex ? row.map(()=>"") : row)
        if (tracking!==prepared.tracking || trackFormatting!==prepared.trackFormatting || !uniform || count!==prepared.rowCount || JSON.stringify(values)!==JSON.stringify(expectedAfter) || bodyOriginal!==originalText || reviewed.some((c,i)=>c.current!=="" || c.original!==proposal.cells[i])) throw new Error("Native row deletion confirmation failed")
        const afterComments=await this.documentState(context)
        if (JSON.stringify(beforeComments.comments)!==JSON.stringify(afterComments.comments)) throw new Error("Native row comments changed")
        const verified=verifyNativeWordChatDeletedRow(beforeXml,afterXmlValue,inventory,afterInventory,proposal)
        this.applied.set(proposal.id,{identity,authors:verified.authors,uncertain:false})
        return {status:"applied",tracked:true,saved:false,authors:verified.authors}
      })
    } catch(error) {
      if (this.applied.get(proposal.id)?.uncertain) throw new WordChatHostError("Word may have deleted the row, but its tracked result could not be confirmed. Inspect it and use Undo if needed; do not apply it again.","uncertain",{cause:error})
      failure(error,"Word could not prepare the reviewed row deletion. Your document is unchanged.")
    }
  }

  /** Complete paragraphs form the import transport. The approved edit still
   * names only target.text/startOffset/endOffset, checked in both authored views. */
  private async importSnapshot(context: WordChatContext, target: WordChatTarget, formatting = false, wordingImport = false, insertedMarks = false, allowNativeCell = false, faithfulProjection = false, inline = false) {
    const { range: selected, paragraphs } = await this.resolve(context, target, allowNativeCell)
    const all = context.document.body.paragraphs
    all.load("items/uniqueLocalId")
    await context.sync()
    const at = all.items.findIndex(p => p.uniqueLocalId === target.paragraphs[0]!.id)
    if (at < 0) throw captureMismatch("Word could not locate the reviewed paragraph among the document’s current paragraphs.")
    // A cell Whole export may include the row, the target and a trailing paragraph.
    const neighbours: Array<{ id: string; range: WordChatRange }> = [all.items[at - 1], all.items[at + target.paragraphs.length]]
      .filter((p): p is Paragraph => Boolean(p)).map(p => ({ id: p.uniqueLocalId, range: p.getRange(allowNativeCell ? "Content" : "Whole") }))
    // Native formatting and wording use the bound whole-story projection below. Legacy
    // import paths retain their existing scoped Content export.
    const content = paragraphs[0]!.getRange("Content").expandTo(paragraphs.at(-1)!.getRange("Content"))
    const range = formatting && !wordingImport ? paragraphs[0]!.getRange("Whole").expandTo(paragraphs.at(-1)!.getRange("Whole")) : content
    // Every preservation baseline comes from this final read batch. Context
    // discovery yields to Word and must never leave an earlier XML snapshot as
    // the authority for a later complete-paragraph replacement.
    const xml = faithfulProjection ? undefined : content.getOoxml(), bodyXml = faithfulProjection ? context.document.body.getRange("Whole").getOoxml() : undefined, comments = range.getComments(), revisions = range.getTrackedChanges()
    comments.load("items"); revisions.load("items/author,items/type")
    const current = paragraphs.map(paragraph => paragraph.getRange("Content").getReviewedText("Current"))
    context.document.load(formatting ? "changeTrackingMode,trackFormatting" : "changeTrackingMode")
    all.load(faithfulProjection ? "items/uniqueLocalId,items/text,items/tableNestingLevel" : "items/uniqueLocalId")
    const adjacent = faithfulProjection ? [] : neighbours.map(p => ({ id: p.id, xml: p.range.getOoxml() }))
    const bodyChanges = context.document.body.getRange().getTrackedChanges()
    bodyChanges.load("items/author,items/type,items/text")
    selected.load("text")
    for (const paragraph of paragraphs) paragraph.load("text")
    await context.sync()
    const finalIds = all.items.map(p => p.uniqueLocalId)
    const finalAt = finalIds.indexOf(target.paragraphs[0]!.id)
    if (faithfulProjection && (finalAt !== at || JSON.stringify([finalIds[finalAt - 1], finalIds[finalAt + target.paragraphs.length]].filter((id): id is string => id !== undefined)) !== JSON.stringify(neighbours.map(p => p.id))))
      throw captureMismatch("The paragraphs beside this review changed during capture.")
    let projection: ReturnType<typeof projectNativeWordChatBody> | undefined
    try { projection = bodyXml ? projectNativeWordChatBody(bodyXml.value, all.items.map(p => ({ id:p.uniqueLocalId, text:p.text, tableNestingLevel:p.tableNestingLevel }))) : undefined }
    catch { throw captureMismatch("Word could not bind the exported body to its native paragraphs.") }
    const scoped = projection ? projection.scope(target.paragraphs.map((_, i) => at + i)) : xml!.value
    if (formatting && typeof context.document.trackFormatting !== "boolean")
      throw new WordChatHostError("This version of Word cannot confirm formatting tracking. Your document is unchanged.")
    if (!inline && (comments.items.length || (!formatting && revisions.items.length) || formatting && revisions.items.some(revision => !["Added", "Deleted"].includes((revision as TrackedChange).type)) || hasProtectedContent(scoped, false, formatting, insertedMarks)))
      throw new WordChatHostError("This paragraph contains existing revisions, comments or a linked Word object. Review that passage in Word first, then prepare a fresh suggestion for ordinary text.")
    let captured: string[]
    try { captured = capturedWordChatParagraphs(scoped) }
    catch { throw new WordChatHostError("Word could not capture these complete paragraphs safely. Choose ordinary text within one clause or table cell; your document is unchanged.") }
    if (captured.length !== target.paragraphs.length)
      throw captureMismatch(`Word exported ${captured.length} paragraphs for this ${target.paragraphs.length}-paragraph review.`)
    const exportedMismatch = captured.findIndex((text, index) => text !== target.paragraphs[index]!.text)
    if (exportedMismatch >= 0)
      throw captureMismatch(`Word’s exported text differs at reviewed paragraph ${exportedMismatch + 1} (${captured[exportedMismatch]!.length} exported characters; ${target.paragraphs[exportedMismatch]!.text.length} reviewed).`)
    const currentMismatch = current.findIndex((text, index) => text.value !== target.paragraphs[index]!.text)
    if (currentMismatch >= 0)
      throw captureMismatch(`Word’s current tracked-change text differs at reviewed paragraph ${currentMismatch + 1} (${current[currentMismatch]!.value.length} current characters; ${target.paragraphs[currentMismatch]!.text.length} reviewed).`)
    if (selected.text !== target.text || paragraphs.some((p, index) => p.text !== target.paragraphs[index]!.text))
      throw captureMismatch("Word’s paragraph or range text changed during capture.")
    const ids = all.items.map(p => p.uniqueLocalId)
    const index = ids.indexOf(target.paragraphs[0]!.id)
    if (index < 0 || target.paragraphs.some((p, offset) => ids[index + offset] !== p.id))
      throw captureMismatch("The reviewed paragraphs changed identity or order during capture.")
    if (formatting && !faithfulProjection && JSON.stringify([ids[index - 1], ids[index + target.paragraphs.length]].filter((id): id is string => id !== undefined)) !== JSON.stringify(adjacent.map(p => p.id)))
      throw captureMismatch("The paragraphs beside this review changed during capture.")
    const nearby = projection ? [index - 1, index + target.paragraphs.length].filter(i => i >= 0 && i < ids.length).map(i => ({ id:ids[i]!, ooxml:projection.scope([i]) })) : adjacent.map(p => ({ id:p.id, ooxml:p.xml.value }))
    return { range, selected, ooxml: scoped, ids, index, nesting: paragraphs.map(p => p.tableNestingLevel), adjacent: nearby, projection, tracking: context.document.changeTrackingMode, trackFormatting: formatting ? context.document.trackFormatting : undefined,
      changes: bodyChanges.items.map(change => ({ author: change.author, type: change.type, text: change.text })) }
  }

  private async applyListInsertionOnce(input: WordChatApplyInput, identity: string): Promise<WordChatApplyResult> {
    try {
      if (!input.list) throw captureMismatch("Word could not bind the reviewed list.")
      const position = input.paragraphAction === "insert_list_item_before" ? "before_final" : "end"
      const metrics = (p: Paragraph, font: NonNullable<WordChatRange["font"]>) => {
        const values = [font.bold,font.italic,font.underline,font.name,font.size,font.color,p.spaceBefore,p.spaceAfter,p.lineSpacing,p.leftIndent,p.firstLineIndent]
        if (typeof font.bold !== "boolean" || typeof font.italic !== "boolean" || typeof font.underline !== "string" || typeof font.name !== "string" || !font.name || typeof font.color !== "string" || !font.color || typeof font.size !== "number" || !Number.isFinite(font.size) || values.slice(6).some(x => typeof x !== "number" || !Number.isFinite(x))) throw new Error("list_effective_format_unavailable")
        return values
      }
      const capture = async (context: WordChatContext) => {
        const snap = await this.importSnapshot(context,input.target,false,false,false,false,true,true)
        if (!snap.projection) throw captureMismatch("Word could not capture the reviewed list.")
        assertNativeWordChatListInsertionTarget(snap.projection,input.target,position)
        const lists = await this.observeNativeLists(context,snap.projection,[input.target.paragraphs[0]!.id])
        if (lists.length !== 1 || JSON.stringify(lists[0]) !== JSON.stringify(input.list)) throw captureMismatch("The reviewed list changed.")
        const anchor = context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id)
        anchor.load("styleBuiltIn,isListItem,spaceBefore,spaceAfter,lineSpacing,leftIndent,firstLineIndent")
        const font = anchor.getRange("Content").font
        if (!font) throw new Error("list_font_unavailable")
        font.load("bold,italic,underline,name,size,color")
        const current = context.document.body.getRange("Content").getReviewedText("Current"), original = context.document.body.getRange("Content").getReviewedText("Original")
        const all = context.document.body.paragraphs, xml = context.document.body.getRange("Whole").getOoxml()
        context.document.load("trackFormatting")
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        await context.sync()
        if (typeof context.document.trackFormatting !== "boolean") throw captureMismatch("Word could not confirm formatting tracking. Your document is unchanged.")
        const final = projectNativeWordChatBody(xml.value,all.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})))
        if (!sameNativeProjectionGeometry(snap.projection,final) || !sameNativeWordChatBodySnapshot(snap.projection.ooxml,final.ooxml) || !nativeWordChatDecimalList(final,lists[0]!)) throw captureMismatch("The list or document changed during capture.")
        assertNativeWordChatInlineDependencies(snap.projection.ooxml,final.ooxml)
        if (anchor.styleBuiltIn !== "Normal" || anchor.isListItem !== true) throw captureMismatch("The reviewed list formatting changed.")
        const span = input.target.text, at = current.value.indexOf(span)
        if (at < 0 || at > 0 && current.value[at-1] !== "\r" || at+span.length < current.value.length && current.value[at+span.length] !== "\r" || current.value.indexOf(span,at+1) !== -1) throw captureMismatch("Word could not uniquely confirm the list boundary.")
        return {...snap,projection:final,current:current.value,original:original.value,at,format:metrics(anchor,font),trackFormatting:context.document.trackFormatting}
      }
      const initial = await this.word.run(capture)
      const plan = await input.author({ooxml:initial.ooxml})
      if (!("kind" in plan) || plan.kind !== "list_insert" || plan.proposalId !== input.proposalId || JSON.stringify(plan.target) !== JSON.stringify(input.target) || JSON.stringify(plan.list) !== JSON.stringify(input.list) || (plan.position ?? "end") !== position || input.replacement !== (position === "end" ? `${input.target.text}\r${plan.text}` : `${plan.text}\r${input.target.text}`)) throw captureMismatch("The prepared list insertion differs from the reviewed change.")
      return await this.word.run(async context => {
        const fresh = await capture(context)
        if (!sameNativeWordChatBodySnapshot(initial.projection.ooxml,fresh.projection.ooxml) || !sameNativeProjectionGeometry(initial.projection,fresh.projection) || initial.current !== fresh.current || initial.original !== fresh.original || JSON.stringify(initial.format) !== JSON.stringify(fresh.format) || initial.trackFormatting !== fresh.trackFormatting) throw captureMismatch("The list changed during review.")
        assertNativeWordChatInlineDependencies(initial.projection.ooxml,fresh.projection.ooxml)
        const anchor = context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id), list = anchor.list
        if (!list || (position === "end" ? typeof list.insertParagraph !== "function" : typeof anchor.insertParagraph !== "function")) throw captureMismatch("Word cannot insert into this list.")
        if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
        this.applied.set(input.proposalId,{identity,authors:[],uncertain:true})
        try {
          context.document.changeTrackingMode = "TrackAll"
          if (position === "end") list.insertParagraph!(plan.text,"End")
          else anchor.insertParagraph!(plan.text,"Before")
          context.document.changeTrackingMode = fresh.tracking
          context.document.trackFormatting = fresh.trackFormatting
          await context.sync()
          context.document.load("changeTrackingMode,trackFormatting")
          const all = context.document.body.paragraphs, body = context.document.body.getRange("Whole").getOoxml(), current = context.document.body.getRange("Content").getReviewedText("Current"), original = context.document.body.getRange("Content").getReviewedText("Original")
          all.load("items/uniqueLocalId,items/text,items/tableNestingLevel,items/styleBuiltIn,items/isListItem,items/spaceBefore,items/spaceAfter,items/lineSpacing,items/leftIndent,items/firstLineIndent")
          await context.sync()
          const insertionAt = fresh.at + (position === "end" ? input.target.text.length : 0)
          const addition = position === "end" ? "\r" + plan.text : plan.text + "\r"
          const expected = fresh.current.slice(0,insertionAt) + addition + fresh.current.slice(insertionAt)
          if (context.document.changeTrackingMode !== fresh.tracking || context.document.trackFormatting !== fresh.trackFormatting || current.value !== expected || original.value !== fresh.original) throw new Error("list_reviewed_text")
          const observed = projectNativeWordChatBody(body.value,all.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})))
          const authors = verifyNativeWordChatListInsertion(fresh.projection,observed,input.target,plan.text,position)
          const inserted = all.items[fresh.index + (position === "end" ? 1 : 0)]
          if (!inserted || inserted.styleBuiltIn !== "Normal" || inserted.isListItem !== true) throw new Error("list_style")
          const font = inserted.getRange("Content").font
          if (!font) throw new Error("list_font_unavailable")
          font.load("bold,italic,underline,name,size,color")
          const lists = await this.observeNativeLists(context,observed,[inserted.uniqueLocalId])
          await context.sync()
          const expectedMembers = input.list!.members
          const source = expectedMembers.at(-1)!
          const result = lists[0], previous = result?.members[expectedMembers.length-1], last = result?.members[expectedMembers.length]
          const sibling = (member: typeof source | undefined, ordinal: number) => !!member && (input.list!.format === "decimal" ? member.level === 0 && member.label === `${ordinal+1}.` : "siblingIndex" in member && member.siblingIndex === ordinal)
          const originalSibling = input.list!.format === "native" ? input.list!.members.at(-1)!.siblingIndex : expectedMembers.length-1
          const preserved = result?.members.slice(0,expectedMembers.length-1).every((member,i)=>JSON.stringify(member)===JSON.stringify(expectedMembers[i]))
          const identityMatches = position === "end"
            ? previous && last && previous.id !== source.id && previous.text === source.text && previous.level === source.level && previous.label === source.label && sibling(previous,originalSibling) && last.id === source.id && last.text === plan.text && last.level === source.level && sibling(last,originalSibling+1)
            : previous && last && previous.id === inserted.uniqueLocalId && !expectedMembers.some(member=>member.id===previous.id) && previous.text === plan.text && previous.level === source.level && previous.label === source.label && sibling(previous,originalSibling) && last.id === source.id && last.text === source.text && last.level === source.level && sibling(last,originalSibling+1)
          if (!result || result.id !== input.list!.id || result.format !== input.list!.format || result.members.length !== expectedMembers.length+1 || !preserved || !identityMatches || JSON.stringify(metrics(inserted,font)) !== JSON.stringify(fresh.format)) throw new Error("list_members_or_format")
          this.applied.set(input.proposalId,{identity,authors,uncertain:false})
          return {status:"applied",tracked:true,saved:false,authors}
        } catch(cause) { throw new WordChatHostError("Word may contain the new list item, but could not confirm its tracked changes. Inspect it and use Undo if needed. Do not apply this proposal again.","uncertain",{cause}) }
      })
    } catch(error) {
      if (this.applied.get(input.proposalId)?.uncertain) { if (error instanceof WordChatHostError && error.mutationState === "uncertain") throw error; throw new WordChatHostError("Word may already contain this change. Inspect it and use Undo if needed.","uncertain",{cause:error}) }
      failure(error,"Word could not prepare the reviewed list insertion. Your document is unchanged.")
    }
  }

  private async applyParagraphSplitOnce(input: WordChatApplyInput, identity: string): Promise<WordChatApplyResult> {
    try {
      const [leftText,rightText]=input.replacement.split("\r")
      if (!leftText || !rightText || input.target.text!==leftText+" "+rightText) throw captureMismatch("The reviewed split does not identify one existing space.")
      const format = (p:Paragraph,f:NonNullable<WordChatRange["font"]>) => {
        const metrics=[p.spaceBefore,p.spaceAfter,p.lineSpacing,p.leftIndent,p.firstLineIndent]
        if (p.styleBuiltIn!=="Normal" || p.isListItem!==false || ![0,1].includes(p.tableNestingLevel) || metrics.some(v=>typeof v!=="number" || !Number.isFinite(v))) throw captureMismatch("Word cannot confirm the paragraph formatting.")
        const knownString=(v:unknown) => typeof v==="string" && v.trim() && v.toLowerCase()!=="mixed" ? v : null
        return {nesting:p.tableNestingLevel,metrics,font:[knownString(f.name),typeof f.size==="number" && Number.isFinite(f.size) && f.size>0 ? f.size : null,knownString(f.color)]}
      }
      const loadFormat = (p:Paragraph) => {
        p.load("styleBuiltIn,isListItem,tableNestingLevel,spaceBefore,spaceAfter,lineSpacing,leftIndent,firstLineIndent")
        const font=p.getRange("Content").font; if (!font) throw Error("split_font_unavailable")
        font.load("bold,italic,underline,name,size,color");return font
      }
      const capture = async (context: WordChatContext) => {
        const snapshot=await this.importSnapshot(context,input.target,false,false,false,true,true,true)
        if (!snapshot.projection) throw captureMismatch("Word could not capture the reviewed paragraph.")
        const beforeState=await this.documentState(context), paragraph=context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id)
        const body=context.document.body.getRange("Content"), current=body.getReviewedText("Current"), original=body.getReviewedText("Original"), range=paragraph.getRange("Content")
        const targetCurrent=range.getReviewedText("Current"), targetOriginal=range.getReviewedText("Original"), xml=context.document.body.getRange("Whole").getOoxml(), all=context.document.body.paragraphs
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel");context.document.load("changeTrackingMode,trackFormatting")
        await context.sync()
        const inventory=all.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})), originalXml=xml.value, currentText=current.value, originalText=original.value, targetCurrentText=targetCurrent.value, targetOriginalText=targetOriginal.value
        const font=loadFormat(paragraph), finalXml=context.document.body.getRange("Whole").getOoxml();await context.sync()
        const values=format(paragraph,font)
        if (typeof context.document.trackFormatting!=="boolean" || !sameNativeWordChatBodySnapshot(originalXml,finalXml.value)) throw captureMismatch("The paragraph changed during capture.")
        const projection=projectNativeWordChatBody(finalXml.value,inventory)
        if (!sameNativeProjectionGeometry(snapshot.projection,projection) || !sameNativeWordChatBodySnapshot(snapshot.projection.ooxml,projection.ooxml)) throw captureMismatch("The paragraph changed during capture.")
        const index=assertNativeWordChatParagraphSplitTarget(projection,input.target), at=currentText.indexOf(input.target.text)
        if (targetCurrentText!==input.target.text || targetOriginalText!==input.target.text || at<0 || values.nesting===0 && (at>0 && currentText[at-1]!=="\r" || at+input.target.text.length<currentText.length && currentText[at+input.target.text.length]!=="\r") || currentText.indexOf(input.target.text,at+1)!==-1) throw captureMismatch("Word could not uniquely confirm the paragraph to split.")
        return {...snapshot,projection,index,current:currentText,original:originalText,at,format:values,tracking:context.document.changeTrackingMode,trackFormatting:context.document.trackFormatting,comments:beforeState.comments}
      }
      const initial=await this.word.run(capture), plan=await input.author({ooxml:initial.ooxml})
      if (!("kind" in plan) || plan.kind!=="paragraph_split" || plan.proposalId!==input.proposalId || JSON.stringify(plan.target)!==JSON.stringify(input.target)) throw captureMismatch("The prepared split differs from the reviewed change.")
      return await this.word.run(async context=>{
        const fresh=await capture(context)
        if (!sameNativeWordChatBodySnapshot(initial.projection.ooxml,fresh.projection.ooxml) || !sameNativeProjectionGeometry(initial.projection,fresh.projection) || initial.current!==fresh.current || initial.original!==fresh.original || JSON.stringify(initial.format)!==JSON.stringify(fresh.format) || JSON.stringify(initial.comments)!==JSON.stringify(fresh.comments)) throw captureMismatch("The paragraph changed during review.")
        const paragraph=context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id), range=paragraph.getRange("Content")
        const prefix=range.search(leftText,{matchWildcards:false,matchCase:true,matchWholeWord:false}), suffix=range.search(rightText,{matchWildcards:false,matchCase:true,matchWholeWord:false})
        prefix.load("items/text");suffix.load("items/text");await context.sync()
        if (prefix.items.length!==1 || prefix.items[0]!.text!==leftText || suffix.items.length!==1 || suffix.items[0]!.text!==rightText) throw captureMismatch("Word could not uniquely identify both sides of the split.")
        const separator=prefix.items[0]!.getRange("End").expandTo(suffix.items[0]!.getRange("Start"));separator.load("text");await context.sync()
        if (separator.text!==" ") throw captureMismatch("Word could not confirm the original space to split.")
        if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
        this.applied.set(input.proposalId,{identity,authors:[],uncertain:true})
        context.document.changeTrackingMode="TrackAll";separator.insertText("\r","Replace");context.document.changeTrackingMode=fresh.tracking;context.document.trackFormatting=fresh.trackFormatting
        await context.sync()
        const afterState=await this.documentState(context), all=context.document.body.paragraphs, xml=context.document.body.getRange("Whole").getOoxml(), body=context.document.body.getRange("Content")
        const current=body.getReviewedText("Current"), original=body.getReviewedText("Original")
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel");context.document.load("changeTrackingMode,trackFormatting");await context.sync()
        const inventory=all.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})), observedXml=xml.value, currentText=current.value, originalText=original.value, tracking=context.document.changeTrackingMode, trackFormatting=context.document.trackFormatting
        const old=fresh.projection.paragraphs, prefixParagraph=inventory[fresh.index], suffixParagraph=inventory[fresh.index+1]
        if (inventory.length!==old.length+1 || !prefixParagraph || !suffixParagraph || prefixParagraph.text!==leftText || suffixParagraph.text!==rightText || suffixParagraph.id!==input.target.paragraphs[0]!.id || !prefixParagraph.id || old.some(p=>p.id===prefixParagraph.id) || inventory.some((p,i)=>i!==fresh.index && (i===fresh.index+1 ? false : p.id!==old[i<fresh.index?i:i-1]?.id))) throw Error("paragraph_split_identity")
        const pieces=[prefixParagraph,suffixParagraph].map(p=>context.document.getParagraphByUniqueLocalId(p.id)), fonts=pieces.map(loadFormat)
        const span=pieces[0]!.getRange("Content").expandTo(pieces[1]!.getRange("Content")), targetCurrent=span.getReviewedText("Current"), targetOriginal=span.getReviewedText("Original"), finalXml=context.document.body.getRange("Whole").getOoxml()
        await context.sync()
        const formats=pieces.map((p,i)=>format(p,fonts[i]!)), observed=projectNativeWordChatBody(finalXml.value,inventory)
        const expected=fresh.current.slice(0,fresh.at)+input.replacement+fresh.current.slice(fresh.at+input.target.text.length)
        if (!sameNativeWordChatBodySnapshot(observedXml,finalXml.value) || tracking!==fresh.tracking || trackFormatting!==fresh.trackFormatting || currentText!==expected || originalText!==fresh.original || targetCurrent.value!==input.replacement || targetOriginal.value!==input.target.text || formats.some(v=>v.nesting!==fresh.format.nesting || JSON.stringify(v.metrics)!==JSON.stringify(fresh.format.metrics) || fresh.format.font.some((known,i)=>known!==null && known!==v.font[i])) || JSON.stringify(afterState.comments)!==JSON.stringify(fresh.comments)) throw Error("paragraph_split_confirmation")
        const authors=verifyNativeWordChatParagraphSplit(fresh.projection,observed,input.target,input.replacement)
        this.applied.set(input.proposalId,{identity,authors,uncertain:false})
        return {status:"applied",tracked:true,saved:false,authors}
      })
    } catch(error) {
      if (this.applied.get(input.proposalId)?.uncertain) throw new WordChatHostError("Word may contain the split, but its tracked result could not be confirmed. Inspect it and use Undo if needed; do not apply it again.","uncertain",{cause:error})
      failure(error,"Word could not prepare the reviewed paragraph split. Your document is unchanged.")
    }
  }

  private async applyParagraphMergeOnce(input: WordChatApplyInput, identity: string): Promise<WordChatApplyResult> {
    try {
      const capture = async (context: WordChatContext) => {
        const snapshot = await this.importSnapshot(context, input.target, false, false, false, false, true, true)
        if (!snapshot.projection) throw captureMismatch("Word could not capture the reviewed paragraphs.")
        const beforeState = await this.documentState(context)
        const paragraphs = input.target.paragraphs.map(p => context.document.getParagraphByUniqueLocalId(p.id))
        const fonts = paragraphs.map(p => { const f = p.getRange("Content").font; if (!f) throw Error("merge_font_unavailable"); return f })
        context.document.load("changeTrackingMode,trackFormatting")
        const body = context.document.body.getRange("Content"), whole = context.document.body.getRange("Whole").getOoxml()
        const current = body.getReviewedText("Current"), original = body.getReviewedText("Original")
        const span = paragraphs[0]!.getRange("Content").expandTo(paragraphs[1]!.getRange("Content"))
        const targetCurrent = span.getReviewedText("Current"), targetOriginal = span.getReviewedText("Original")
        const all = context.document.body.paragraphs
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        await context.sync()
        const inventory = all.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel}))
        const originalXml = whole.value, currentText = current.value, originalText = original.value, targetCurrentText = targetCurrent.value, targetOriginalText = targetOriginal.value
        for (const p of paragraphs) p.load("styleBuiltIn,isListItem,tableNestingLevel,spaceBefore,spaceAfter,lineSpacing,leftIndent,firstLineIndent")
        for (const f of fonts) f.load("bold,italic,underline,name,size,color")
        const finalXml = context.document.body.getRange("Whole").getOoxml()
        await context.sync()
        if (!sameNativeWordChatBodySnapshot(originalXml,finalXml.value)) throw captureMismatch("The paragraphs changed during capture.")
        const format = paragraphs.map((p,i) => {
          const f=fonts[i]!, values=[f.bold,f.italic,f.underline,f.name,f.size,f.color,p.spaceBefore,p.spaceAfter,p.lineSpacing,p.leftIndent,p.firstLineIndent]
          if (p.styleBuiltIn!=="Normal" || p.isListItem!==false || p.tableNestingLevel!==0 || typeof f.bold!=="boolean" || typeof f.italic!=="boolean" || typeof f.underline!=="string" || typeof f.name!=="string" || !f.name || typeof f.color!=="string" || !f.color || typeof f.size!=="number" || !Number.isFinite(f.size) || values.slice(6).some(v=>typeof v!=="number" || !Number.isFinite(v))) throw captureMismatch("Word cannot confirm compatible ordinary paragraph formatting.")
          return values
        })
        if (JSON.stringify(format[0])!==JSON.stringify(format[1]) || typeof context.document.trackFormatting!=="boolean") throw captureMismatch("The paragraph formats are not compatible.")
        const projection = projectNativeWordChatBody(finalXml.value,inventory)
        if (!sameNativeProjectionGeometry(snapshot.projection,projection) || !sameNativeWordChatBodySnapshot(snapshot.projection.ooxml,projection.ooxml)) throw captureMismatch("The paragraphs changed during capture.")
        assertNativeWordChatParagraphMergeTarget(projection,input.target)
        const at=currentText.indexOf(input.target.text)
        if (targetCurrentText!==input.target.text || targetOriginalText!==input.target.text || at<0 || at>0 && currentText[at-1]!=="\r" || at+input.target.text.length<currentText.length && currentText[at+input.target.text.length]!=="\r" || currentText.indexOf(input.target.text,at+1)!==-1) throw captureMismatch("Word could not uniquely confirm the reviewed paragraph boundary.")
        return {...snapshot,projection,current:currentText,original:originalText,at,format,tracking:context.document.changeTrackingMode,trackFormatting:context.document.trackFormatting,comments:beforeState.comments}
      }
      const initial = await this.word.run(capture)
      const plan = await input.author({ooxml:initial.ooxml})
      if (!("kind" in plan) || plan.kind!=="paragraph_merge" || plan.proposalId!==input.proposalId || JSON.stringify(plan.target)!==JSON.stringify(input.target)) throw captureMismatch("The prepared merge differs from the reviewed change.")
      return await this.word.run(async context => {
        const fresh=await capture(context)
        if (!sameNativeWordChatBodySnapshot(initial.projection.ooxml,fresh.projection.ooxml) || !sameNativeProjectionGeometry(initial.projection,fresh.projection) || initial.current!==fresh.current || initial.original!==fresh.original || JSON.stringify(initial.format)!==JSON.stringify(fresh.format) || JSON.stringify(initial.comments)!==JSON.stringify(fresh.comments)) throw captureMismatch("The paragraphs changed during review.")
        const left=context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id)
        const matches=left.getRange("Whole").search("^p",{matchWildcards:false,matchCase:true,matchWholeWord:false})
        matches.load("items/text"); await context.sync()
        if (matches.items.length!==1 || matches.items[0]!.text!=="\r") throw captureMismatch("Word could not uniquely identify the paragraph mark.")
        if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
        this.applied.set(input.proposalId,{identity,authors:[],uncertain:true})
        context.document.changeTrackingMode="TrackAll"
        matches.items[0]!.insertText(" ","After")
        matches.items[0]!.delete()
        context.document.changeTrackingMode=fresh.tracking
        context.document.trackFormatting=fresh.trackFormatting
        await context.sync()
        const afterState=await this.documentState(context)
        const paragraphs=input.target.paragraphs.map(p=>context.document.getParagraphByUniqueLocalId(p.id))
        const span=paragraphs[0]!.getRange("Content").expandTo(paragraphs[1]!.getRange("Content"))
        const targetCurrent=span.getReviewedText("Current"), targetOriginal=span.getReviewedText("Original")
        const body=context.document.body.getRange("Content"), current=body.getReviewedText("Current"), original=body.getReviewedText("Original"), xml=context.document.body.getRange("Whole").getOoxml(), all=context.document.body.paragraphs
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel");context.document.load("changeTrackingMode,trackFormatting")
        await context.sync()
        const observed=projectNativeWordChatBody(xml.value,all.items.map(p=>({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})))
        const expected=fresh.current.slice(0,fresh.at)+input.replacement+fresh.current.slice(fresh.at+input.target.text.length)
        if (context.document.changeTrackingMode!==fresh.tracking || context.document.trackFormatting!==fresh.trackFormatting || current.value!==expected || original.value!==fresh.original || targetCurrent.value!==input.replacement || targetOriginal.value!==input.target.text || JSON.stringify(afterState.comments)!==JSON.stringify(fresh.comments)) throw Error("paragraph_merge_confirmation")
        const authors=verifyNativeWordChatParagraphMerge(fresh.projection,observed,input.target)
        this.applied.set(input.proposalId,{identity,authors,uncertain:false})
        return {status:"applied",tracked:true,saved:false,authors}
      })
    } catch(error) {
      if (this.applied.get(input.proposalId)?.uncertain) throw new WordChatHostError("Word may contain the merge, but its tracked result could not be confirmed. Inspect it and use Undo if needed; do not apply it again.","uncertain",{cause:error})
      failure(error,"Word could not prepare the reviewed paragraph merge. Your document is unchanged.")
    }
  }

  private async applyParagraphDeletionOnce(input: WordChatApplyInput, identity: string): Promise<WordChatApplyResult> {
    try {
    const capture = async (context: WordChatContext) => {
      const snapshot = await this.importSnapshot(context, input.target, false, false, false, false, true, true)
      if (!snapshot.projection) throw captureMismatch("Word could not capture the paragraph to delete.")
      // Any style or list format: the verifier proves the following paragraph's
      // properties and mark are unchanged, so accepting leaves its formatting.
      assertNativeWordChatParagraphDeletionTarget(snapshot.projection, input.target)
      const current = context.document.body.getRange("Content").getReviewedText("Current")
      await context.sync()
      // The text may repeat; the XML inverse binds the paragraph, so any
      // boundary-aligned removal is an acceptable reviewed text.
      const span = input.target.text + "\r", expected: string[] = []
      for (let at = current.value.indexOf(span); at >= 0; at = current.value.indexOf(span, at + 1))
        if (at === 0 || current.value[at - 1] === "\r") expected.push(current.value.slice(0, at) + current.value.slice(at + span.length))
      if (!expected.length) throw captureMismatch("Word could not confirm the paragraph and its following boundary.")
      return { ...snapshot, current: current.value, expected }
    }
    const initial = await this.word.run(capture)
    const plan = await input.author({ ooxml: initial.ooxml })
    if (!("kind" in plan) || plan.kind !== "paragraph_delete" || plan.proposalId !== input.proposalId ||
      JSON.stringify(plan.target) !== JSON.stringify(input.target))
      throw new WordChatHostError("The prepared deletion does not match the reviewed proposal. Your document is unchanged.")
    return await this.word.run(async context => {
      const fresh = await capture(context)
      if (!sameWordChatTargetSnapshot(initial.ooxml, fresh.ooxml) || initial.current !== fresh.current ||
        !sameNativeProjectionGeometry(initial.projection!, fresh.projection!))
        throw captureMismatch("The paragraph or its boundary changed during review.")
      const paragraph = context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id)
      if (typeof paragraph.delete !== "function") throw new WordChatHostError("Word cannot delete this paragraph. Your document is unchanged.")
      if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
      this.applied.set(input.proposalId, { identity, authors: [], uncertain: true })
      try {
        context.document.changeTrackingMode = "TrackAll"
        paragraph.delete()
        context.document.changeTrackingMode = fresh.tracking
        await context.sync()
        context.document.load("changeTrackingMode")
        const all = context.document.body.paragraphs
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        const body = context.document.body.getRange("Whole").getOoxml()
        const current = context.document.body.getRange("Content").getReviewedText("Current")
        await context.sync()
        if (context.document.changeTrackingMode !== fresh.tracking || !fresh.expected.includes(current.value))
          throw new Error("paragraph_delete_current")
        const observed = projectNativeWordChatBody(body.value, all.items.map(p => ({ id: p.uniqueLocalId, text: p.text, tableNestingLevel: p.tableNestingLevel })))
        const authors = verifyNativeWordChatParagraphDeletion(fresh.projection!, observed, input.target)
        this.applied.set(input.proposalId, { identity, authors, uncertain: false })
        return { status: "applied", tracked: true, saved: false, authors }
      } catch (cause) {
        throw new WordChatHostError("Word may contain the deletion, but could not confirm its tracked changes. Inspect it and use Undo if needed. Do not apply this proposal again.", "uncertain", { cause })
      }
    })
    } catch (error) {
      if (this.applied.get(input.proposalId)?.uncertain) {
        if (error instanceof WordChatHostError && error.mutationState === "uncertain") throw error
        throw new WordChatHostError("Word may already contain this change. Inspect it and use Undo if needed.", "uncertain", { cause: error })
      }
      failure(error, "Word could not prepare the reviewed paragraph deletion. Your document is unchanged.")
    }
  }

  private async applyParagraphInsertionOnce(input: WordChatApplyInput, identity: string): Promise<WordChatApplyResult> {
    try {
    const capture = async (context: WordChatContext) => {
      const anchor = context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id)
      anchor.load("styleBuiltIn,isListItem")
      // Headings and style-linked lists are accepted; direct numbering is refused by the XML target rule.
      const snapshot = await this.importSnapshot(context, input.target, false, false, false, false, true, true)
      context.document.load("trackFormatting")
      await context.sync()
      return { ...snapshot, style: anchor.styleBuiltIn, listItem: anchor.isListItem, trackFormatting: context.document.trackFormatting }
    }
    const requested = input.formatting?.[0]?.style
    const initial = await this.word.run(capture)
    if (!initial.projection) throw captureMismatch("Word could not capture the insertion point.")
    assertNativeWordChatParagraphInsertionTarget(initial.projection, input.target)
    const plan = await input.author({ ooxml: initial.ooxml })
    if (!("kind" in plan) || plan.kind !== "paragraph_insert" || plan.proposalId !== input.proposalId ||
      JSON.stringify(plan.target) !== JSON.stringify(input.target) || input.replacement !== `${input.target.text}\r${plan.text}` ||
      JSON.stringify(plan.formatting) !== JSON.stringify(input.formatting))
      throw new WordChatHostError("The prepared paragraph does not match the reviewed proposal. Your document is unchanged.")
    return await this.word.run(async context => {
      const fresh = await capture(context)
      if (!fresh.projection || !sameWordChatTargetSnapshot(initial.ooxml, fresh.ooxml) ||
        !sameNativeProjectionGeometry(initial.projection!, fresh.projection) || fresh.style !== initial.style || fresh.listItem !== initial.listItem ||
        fresh.trackFormatting !== initial.trackFormatting) throw captureMismatch("The insertion point changed during review.")
      assertNativeWordChatParagraphInsertionTarget(fresh.projection, input.target)
      const restyle = requested !== undefined && requested !== fresh.style
      if (restyle && typeof fresh.trackFormatting !== "boolean") throw new WordChatHostError("This version of Word cannot confirm formatting tracking. Your document is unchanged.")
      const anchor = context.document.getParagraphByUniqueLocalId(input.target.paragraphs[0]!.id)
      if (typeof anchor.insertParagraph !== "function") throw new WordChatHostError("Word cannot insert this paragraph. Your document is unchanged.")
      if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
      this.applied.set(input.proposalId, { identity, authors: [], uncertain: true })
      try {
        context.document.changeTrackingMode = "TrackAll"
        const added = anchor.insertParagraph(plan.text, "After")
        // The requested style is a second tracked action on the new paragraph only.
        if (restyle) { context.document.trackFormatting = true; added.styleBuiltIn = requested; context.document.trackFormatting = fresh.trackFormatting }
        context.document.changeTrackingMode = fresh.tracking
        await context.sync()
        context.document.load("changeTrackingMode,trackFormatting")
        const all = context.document.body.paragraphs
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel,items/styleBuiltIn,items/isListItem")
        const bodyXml = context.document.body.getRange("Whole").getOoxml()
        await context.sync()
        if (context.document.changeTrackingMode !== fresh.tracking || context.document.trackFormatting !== fresh.trackFormatting) throw new Error("paragraph_insert_tracking")
        const observed = projectNativeWordChatBody(bodyXml.value, all.items.map(p => ({ id: p.uniqueLocalId, text: p.text, tableNestingLevel: p.tableNestingLevel })))
        const authors = verifyNativeWordChatParagraphInsertion(fresh.projection, observed, input.target, plan.text, restyle)
        const inserted = all.items[fresh.index + 1]
        if (!inserted || inserted.styleBuiltIn !== (requested ?? fresh.style) || !restyle && inserted.isListItem !== fresh.listItem) throw new Error("paragraph_insert_style")
        this.applied.set(input.proposalId, { identity, authors, uncertain: false })
        return { status: "applied", tracked: true, saved: false, authors }
      } catch (cause) {
        throw new WordChatHostError("Word may contain the new paragraph, but could not confirm its tracked changes. Inspect it and use Undo if needed. Do not apply this proposal again.", "uncertain", { cause })
      }
    })
    } catch (error) {
      if (this.applied.get(input.proposalId)?.uncertain) {
        if (error instanceof WordChatHostError && error.mutationState === "uncertain") throw error
        throw new WordChatHostError("Word may already contain this change. Inspect it and use Undo if needed.", "uncertain", { cause: error })
      }
      failure(error, "Word could not prepare the reviewed paragraph insertion. Your document is unchanged.")
    }
  }

  private async applyInlineOnce(input: WordChatApplyInput, identity: string): Promise<WordChatApplyResult> {
    const format = input.inlineFormatting!
    const snapshot: WordChatSnapshot = { text: input.target.paragraphs.map(p => p.text).join("\r"), paragraphs: input.target.paragraphs, selection: input.target, saved: input.target.saved, capturedAt: input.target.capturedAt, truncated: false }
    const target = targetWordChatAmendment(snapshot, input.target.text)
    if (!target || !validWordChatInlineFormatting(snapshot, target, format, input.replacement, input.formatting) || target.startOffset !== input.target.startOffset || target.endOffset !== input.target.endOffset)
      throw new WordChatHostError("Review a valid inline formatting change before applying it.")
    if (!this.supports("WordApiDesktop", "1.4")) throw new WordChatHostError("This version of Word cannot confirm formatting tracking. Your document is unchanged.")
    const eligibility = wordChatApplyEligibility(input.target, input.replacement, input.formatting, this.nativeTextEnabled, format)
    if (!eligibility.supported) throw new WordChatHostError(eligibility.reason)
    const capture = (context: WordChatContext) => this.importSnapshot(context, input.target, true, false, false, true, true, true)
    const initial = await this.word.run(capture)
    if (!initial.projection) throw captureMismatch("Word could not capture the document for formatting.")
    assertNativeInlineWordChatRevisionIds(initial.projection.ooxml)
    normalizeNativeWordChatAnnotationIds(initial.projection.ooxml, initial.projection.ooxml)
    const plan = await input.author({ooxml: initial.ooxml})
    if (!("kind" in plan) || plan.kind !== "inline_format" || plan.proposalId !== input.proposalId || JSON.stringify(plan.target) !== JSON.stringify(input.target) || JSON.stringify(plan.inlineFormatting) !== JSON.stringify(format)) throw new WordChatHostError("The prepared inline change does not match the reviewed proposal. Your document is unchanged.")
    try { assertNativeWordChatInlineTarget(initial.ooxml, input.target, format) }
    catch (cause) { throw new WordChatHostError("This passage cannot be formatted safely. Your document is unchanged.", "unchanged", {cause}) }
    return this.word.run(async context => {
      const fresh = await capture(context)
      if (!sameWordChatTargetSnapshot(initial.ooxml, fresh.ooxml) || !initial.projection || !fresh.projection || !sameNativeProjectionGeometry(initial.projection, fresh.projection)) throw captureMismatch("This passage changed during review.")
      const font = fresh.selected.font
      if (!font) throw new WordChatHostError("Word cannot inspect this inline formatting. Your document is unchanged.")
      font.load("bold,italic,underline")
      await context.sync()
      const matches = () => (format.bold === undefined || font.bold === format.bold) && (format.italic === undefined || font.italic === format.italic) && (format.underline === undefined || font.underline === (format.underline ? "Single" : "None"))
      const already = matches()
      const final = await capture(context)
      if (!sameWordChatTargetSnapshot(fresh.ooxml, final.ooxml) || !final.projection || !sameNativeProjectionGeometry(fresh.projection, final.projection)) throw captureMismatch("This passage changed before formatting.")
      if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
      if (already) {
        this.applied.set(input.proposalId, {identity, authors: [], uncertain: false})
        return {status: "applied", tracked: true, saved: false, authors: [], target: {...input.target, id: crypto.randomUUID(), saved: false, capturedAt: new Date().toISOString()}}
      }
      const writeFont = final.selected.font
      if (!writeFont) throw new WordChatHostError("Word cannot apply this inline formatting. Your document is unchanged.")
      this.applied.set(input.proposalId, {identity, authors: [], uncertain: true})
      try {
        if (!final.trackFormatting) context.document.trackFormatting = true
        context.document.changeTrackingMode = "TrackAll"
        if (format.bold !== undefined) writeFont.bold = format.bold
        if (format.italic !== undefined) writeFont.italic = format.italic
        if (format.underline !== undefined) writeFont.underline = format.underline ? "Single" : "None"
        context.document.changeTrackingMode = final.tracking
        if (!final.trackFormatting) context.document.trackFormatting = false
        await context.sync()
        context.document.load("changeTrackingMode,trackFormatting")
        const all = context.document.body.paragraphs
        all.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
        const bodyXml = context.document.body.getRange("Whole").getOoxml()
        writeFont.load("bold,italic,underline")
        await context.sync()
        assertNativeInlineWordChatRevisionIds(bodyXml.value)
        const verifiedAnnotations = normalizeNativeWordChatAnnotationIds(final.projection.ooxml, bodyXml.value)
        const observed = projectNativeWordChatBody(verifiedAnnotations, all.items.map(p => ({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel})))
        if (context.document.changeTrackingMode !== final.tracking || context.document.trackFormatting !== final.trackFormatting || !sameNativeProjectionGeometry(final.projection, observed)) throw new Error("inline_geometry_or_tracking")
        if (!observed.paragraphs.every((p, i) => i === final.index || sameNativeInlineWordChatParagraph(final.projection!.paragraphs[i]!.xml, p.xml))) throw new Error("inline_unchanged_paragraph")
        if ((format.bold !== undefined && writeFont.bold !== format.bold) || (format.italic !== undefined && writeFont.italic !== format.italic) || (format.underline !== undefined && writeFont.underline !== (format.underline ? "Single" : "None"))) throw new Error("inline_effective_font")
        assertNativeWordChatInlineDependencies(final.ooxml, observed.scope([final.index]))
        const authors = verifiedNativeWordChatInlineFormatting(final.ooxml, observed.scope([final.index]), input.target, format, { bold: writeFont.bold ?? undefined, italic: writeFont.italic ?? undefined, underline: writeFont.underline === "Single" ? true : writeFont.underline === "None" ? false : undefined })
        this.applied.set(input.proposalId, {identity, authors, uncertain: false})
        return {status: "applied", tracked: true, saved: false, authors, target: {...input.target, id: crypto.randomUUID(), saved: false, capturedAt: new Date().toISOString()}}
      } catch (cause) { throw new WordChatHostError("Word changed the formatting, but could not confirm the reviewed tracked markup. Inspect it and use Undo if needed. Do not apply this proposal again.", "uncertain", {cause}) }
    })
  }

  private async applyOnce(input: WordChatApplyInput, identity: string, wordingImport = false): Promise<WordChatApplyResult> {
    try {
      if (input.formatting !== undefined && !this.supports("WordApiDesktop", "1.4"))
        throw new WordChatHostError("This version of Word cannot confirm formatting tracking. Your document is unchanged.")
      const eligibility = wordChatApplyEligibility(input.target, input.replacement, input.formatting, this.nativeTextEnabled)
      if (!eligibility.supported) throw new WordChatHostError(eligibility.reason)
      // Scoped OOXML imports may change numbering elsewhere in Word, even
      // outside the reviewed range. They stay guarded; wording uses the
      // separate native text executor rather than importing the artifact.
      const nativeFinalFormatting = input.formatting !== undefined
      // Granular native wording remains fenced by authored-hunk and fresh-scope checks.
      const nativeText = this.nativeTextEnabled && !nativeFinalFormatting && input.replacement !== input.target.text
      const mixed = input.formatting !== undefined && input.replacement !== input.target.text
      const afterParts = input.replacement.split("\r"), plan = wordChatApplyPlan(input.target, input.replacement, nativeFinalFormatting || mixed ? undefined : input.formatting)
      if (!plan.supported) throw new WordChatHostError(plan.reason)
      if (nativeText && (plan.structure || plan.paragraphs.length !== 1)) throw new WordChatHostError("Word can apply wording changes within one paragraph. Refine this change to one paragraph, or make the structural change directly in Word.")
      if (mixed && (!validWordChatFormatting(input.target, input.formatting, input.replacement) || plan.structure?.kind === "delete")) throw new WordChatHostError("This change needs a smaller reviewed amendment. Your Word document is unchanged.")
      const expected = plan.paragraphs
      const formatting = input.formatting !== undefined
      const snapshot = await this.word.run(context => this.importSnapshot(context, input.target, formatting || wordingImport, wordingImport, false, nativeText, (nativeText || formatting && !mixed)))
      // Shared Docxodus authoring is read-only. No Office proxy or canonical
      // write is kept pending while that service authors the reviewed result.
      const authored = await input.author({ ooxml: snapshot.ooxml })
      if ("kind" in authored) throw new WordChatHostError("The prepared amendment does not match the reviewed proposal. Your document is unchanged.")
      let authoredChanges: ReturnType<typeof reviewedWordChatArtifact>
      let mixedArtifact: ReturnType<typeof reviewedMixedWordChatArtifact> | undefined
      try {
        mixedArtifact = mixed ? reviewedMixedWordChatArtifact(authored, input.proposalId, input.target, expected, snapshot.ooxml, input.formatting!) : undefined
        authoredChanges = mixedArtifact ? [] : reviewedWordChatArtifact(authored, input.proposalId, input.target, expected, snapshot.ooxml, input.formatting)
        if (hasProtectedContent(authored.ooxml, true, formatting)) throw new Error("Protected authored content")
      } catch { throw new WordChatHostError("doc.haus could not confirm this authored amendment matches the reviewed passage. Prepare it again; your Word document is unchanged.") }
      if (mixedArtifact) {
        const verified = mixedArtifact
        // The wording import has its own fresh snapshot and independently
        // verified artifact. Fence the parent before any possible mutation.
        this.applied.set(input.proposalId, { identity, authors: [], uncertain: true })
        const wordingId = `${input.proposalId}:wording`
        const wordingInput: WordChatApplyInput = { proposalId: wordingId, target: input.target, replacement: input.replacement,
          author: async () => ({ ...verified.wording, proposalId: wordingId }) }
        const applied = await this.applyOnce(wordingInput, JSON.stringify([input.target, input.replacement]), true)
        if (!applied.target) throw new WordChatHostError("Word could not locate the amended paragraphs for the reviewed formatting. Inspect and use Undo if needed.")
        const next = applied.target
        return await this.word.run(async context => {
          const fresh = await this.importSnapshot(context, next, true, false, true)
          if (!sameImportedWordChat(verified.wording.ooxml, fresh.ooxml) || !sameWordChatInlineRevisions(verified.wording.ooxml, fresh.ooxml))
            throw new WordChatHostError("Word changed the amended passage before formatting. Inspect it and use Undo if needed.")
          return this.applyNativeFormatting(context, { ...input, target: next, replacement: next.text }, identity, fresh, authored, verified.inserted, verified.nativeFormatCount)
        })
      }
      return await this.word.run(async context => {
        const fresh = await this.importSnapshot(context, input.target, formatting || wordingImport, wordingImport, false, nativeText, nativeText || formatting && !mixed)
        if (nativeText) {
          if (JSON.stringify(snapshot.nesting) !== JSON.stringify(fresh.nesting) || !sameWordChatTargetSnapshot(snapshot.ooxml, fresh.ooxml) || !snapshot.projection || !fresh.projection || !sameNativeProjectionGeometry(snapshot.projection, fresh.projection))
            throw new WordChatHostError("This paragraph changed while doc.haus prepared the amendment. Prepare and review it again; your Word changes are safe.")
          return this.applyNativeText(context, input, identity, fresh, authored, expected)
        }
        let difference: string | undefined
        try {
          difference = wordChatOoxmlDifference(snapshot.ooxml, fresh.ooxml, authored.ooxml)
        } catch { throw new WordChatHostError("This paragraph or its formatting changed while doc.haus prepared the amendment. Prepare and review it again; your Word changes are safe.") }
        if (difference) throw new WordChatHostError(`This paragraph or its formatting changed while doc.haus prepared the amendment. Word’s recheck found ${difference}. Prepare and review it again; your Word changes are safe.`)
        if (nativeFinalFormatting) return this.applyNativeFormatting(context, input, identity, fresh, authored)
        this.applied.set(input.proposalId, { identity, authors: [], uncertain: true })
        let imported = false
        try {
          // Revisions already exist in the Docxodus payload. Importing with
          // tracking disabled avoids tracking the whole replacement again.
          context.document.changeTrackingMode = "Off"
          fresh.range.insertOoxml(authored.ooxml, "Replace")
          context.document.changeTrackingMode = fresh.tracking
          await context.sync()
          imported = true
        } finally {
          if (!imported) {
            context.document.changeTrackingMode = fresh.tracking
            await context.sync()
          }
        }
        const all = context.document.body.paragraphs
        all.load("items/uniqueLocalId")
        const changes = context.document.body.getRange().getTrackedChanges()
        changes.load("items/author,items/type,items/text")
        await context.sync()
        const ids = all.items.map(p => p.uniqueLocalId)
        const physicalCount = authoredChanges.length, delta = physicalCount - input.target.paragraphs.length
        const identityConfirmed = ids.length === fresh.ids.length + delta && ids.every((id, index) =>
          index >= fresh.index && index < fresh.index + physicalCount || id === fresh.ids[index < fresh.index ? index : index - delta])
        const actual = all.items.slice(fresh.index, fresh.index + physicalCount).map(paragraph => {
          const content = paragraph.getRange("Content"), changes = content.getTrackedChanges()
          changes.load("items/author,items/type,items/text")
          return { text: content.getReviewedText("Current"), changes }
        })
        const importedParagraphs = all.items.slice(fresh.index, fresh.index + physicalCount)
        const importedRange = importedParagraphs.length === physicalCount
          ? importedParagraphs[0]!.getRange("Content").expandTo(importedParagraphs.at(-1)!.getRange("Content")) : undefined
        const importedXml = importedRange?.getOoxml(), importedText = importedRange?.getReviewedText("Current")
        const adjacent = fresh.adjacent.map(p => ({ before: p.ooxml, after: context.document.getParagraphByUniqueLocalId(p.id).getRange("Whole").getOoxml() }))
        await context.sync()
        const { added, missing } = newWordChatRevisions(fresh.changes, changes.items)
        const matches = (observed: TrackedChange[], wanted: TrackedChange[]) => {
          if (formatting) return observed.length === wanted.length && observed.every(change => change.author === "doc.haus" && change.type === "Formatted") && wanted.every(change => change.type === "Formatted")
          if (observed.some(change => change.author !== "doc.haus" || !["Added", "Deleted"].includes(change.type))) return false
          // Word may group a paragraph mark with its text or expose it as an
          // empty text revision. Exact structural marks/authors are separately
          // verified in the freshly exported OOXML, in both document views.
          const words = (changes: TrackedChange[], type: string) => changes.filter(change => change.type === type).map(change => change.text).join("")
          const native = (value: string) => plan.structure ? value.replace(/\r/g, "") : value
          return ["Added", "Deleted"].every(type => native(words(observed, type)) === native(words(wanted, type)))
        }
        const priorConfirmed = !missing.length, additionsConfirmed = matches(added, authored.revisions)
        const textConfirmed = importedText?.value === expected.join("\r") &&
          (Boolean(plan.structure) || actual.every((item, index) => item.text.value === expected[index] && (formatting || wordingImport || matches(item.changes.items, authoredChanges[index]!))))
        let viewDifference: string | undefined = "unavailable", inlineConfirmed = !wordingImport, neighboursConfirmed = false
        try { viewDifference = importedXml ? importedWordChatDifference(authored.ooxml, importedXml.value) : "unavailable" } catch { viewDifference = "unreadable" }
        try { inlineConfirmed = !wordingImport || Boolean(importedXml && sameWordChatInlineRevisions(authored.ooxml, importedXml.value)) } catch { /* Refuse unreadable inline revisions. */ }
        try { neighboursConfirmed = adjacent.every(p => sameWordChatParagraphContent(p.before, p.after.value)) } catch { /* Refuse unreadable neighbouring content. */ }
        const failed = [!identityConfirmed && "paragraph identity", !priorConfirmed && "prior revision signatures", !additionsConfirmed && "new revision signatures",
          !textConfirmed && "current wording", viewDifference && `authored views: ${viewDifference}`, !inlineConfirmed && "inline revision markup", !neighboursConfirmed && "neighbouring content"].filter(Boolean)
        if (failed.length) throw new WordChatHostError(`Word has imported the amendment, but could not confirm its wording and tracked markup (${failed.join(", ")}). Inspect it and use Undo if needed. Do not apply this proposal again.`)
        this.applied.set(input.proposalId, { identity, authors: ["doc.haus"], uncertain: false })
        // This scope comes only from the verified import, never the cursor.
        // The old model context and remembered authority describe the original
        // wording; they must not be inherited by its changed replacement.
        const target: WordChatTarget | undefined = input.replacement.trim() && actual.length === expected.length && actual.every((item, index) => item.text.value === expected[index]) ? {
          id: crypto.randomUUID(), text: input.replacement,
          paragraphs: actual.map((item, index) => ({ id: ids[fresh.index + index]!, text: item.text.value })),
          startOffset: input.target.startOffset,
          endOffset: (expected.length === 1 ? input.target.startOffset : 0) + afterParts.at(-1)!.length,
          saved: false, capturedAt: new Date().toISOString(),
        } : undefined
        return { status: "applied", tracked: true, saved: false, authors: ["doc.haus"], ...(target ? { target } : {}) }
      })
    } catch (error) {
      const fallback = "Word could not complete the authored amendment. Inspect the clause before preparing a fresh proposal; your existing work has not been replaced by a saved copy."
      if (this.applied.get(input.proposalId)?.uncertain)
        throw new WordChatHostError(error instanceof WordChatHostError ? error.message : fallback, "uncertain", { cause: error })
      failure(error, fallback)
    }
  }

  /** Word replaces only the verified authored hunks with TrackAll active and
   * attributes those revisions to the current Word user. All range discovery
   * finishes before queuing writes; this is not an atomic compare-and-edit. */
  private async applyNativeText(context: WordChatContext, input: WordChatApplyInput, identity: string,
    fresh: Awaited<ReturnType<WordChatHost["importSnapshot"]>>, authored: WordChatAuthoredAmendment, expected: string[]): Promise<WordChatApplyResult> {
    const target = input.target, paragraph = target.paragraphs[0]!
    if (expected.length !== 1 || input.replacement === target.text) throw new WordChatHostError("Refine this change to one paragraph before applying it. Your Word document is unchanged.")
    if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
    const before = await this.documentState(context)
    const at = before.paragraphs.findIndex(item => item.id === paragraph.id)
    if (at < 0 || at !== fresh.index || before.paragraphs[at]!.text !== paragraph.text)
      throw captureMismatch("Word returned a different paragraph before the change.")
    let hunks: ReturnType<typeof authoredWordChatWordingHunks>
    try { hunks = authoredWordChatWordingHunks(authored.ooxml, paragraph.text, expected[0]!, target.startOffset, target.endOffset) }
    catch { throw new WordChatHostError("Word cannot safely map this authored wording to the reviewed passage. Your document is unchanged.") }
    if (hunks.some(hunk => hunk.after.length > 0 && !uniformWordChatFormatting(fresh.ooxml, hunk.start, hunk.end)))
      throw new WordChatHostError("A changed span crosses character formatting, so Word cannot replace it safely here. Your document is unchanged.")
    if (!sameWordChatHunkFormatting(fresh.ooxml, authored.ooxml, hunks))
      throw new WordChatHostError("The authored wording changes character formatting outside the reviewed change. Your document is unchanged.")
    const paragraphRange = context.document.getParagraphByUniqueLocalId(paragraph.id).getRange("Content")
    const ranges: Array<{ hunk: typeof hunks[number]; range: WordChatRange }> = []
    for (const hunk of hunks) {
      const start = await point(context, paragraphRange, paragraph.text, hunk.start)
      const end = await point(context, paragraphRange, paragraph.text, hunk.end)
      const range = start.expandTo(end); range.load("text"); ranges.push({ hunk, range })
    }
    await context.sync()
    if (ranges.some(item => item.range.text !== item.hunk.before)) throw captureMismatch("Word could not confirm each reviewed wording span.")
    const expectedParagraph = paragraph.text.slice(0, target.startOffset) + input.replacement + paragraph.text.slice(target.endOffset)
    // Preserve Word's real story delimiters (including table cell/row marks),
    // rather than reconstructing body text from paragraph names joined by CR.
    const bodyRange = context.document.body.getRange("Content")
    const prefix = bodyRange.getRange("Start").expandTo(paragraphRange.getRange("Start")).getReviewedText("Current")
    const suffix = paragraphRange.getRange("End").expandTo(bodyRange.getRange("End")).getReviewedText("Current")
    const all = context.document.body.paragraphs
    all.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
    const bodyXml = context.document.body.getRange("Whole").getOoxml()
    await context.sync()
    const pre = projectNativeWordChatBody(bodyXml.value, all.items.map(p => ({ id:p.uniqueLocalId, text:p.text, tableNestingLevel:p.tableNestingLevel })))
    if (!fresh.projection || !sameNativeProjectionGeometry(fresh.projection, pre) || !sameWordChatTargetSnapshot(pre.scope([at]), fresh.ooxml) ||
        !fresh.adjacent.every(item => { const i = pre.paragraphs.findIndex(p => p.id === item.id); return i >= 0 && sameWordChatParagraphContent(item.ooxml, pre.scope([i])) }))
      throw captureMismatch("Word returned different XML for the reviewed paragraph or its neighbours before the change.")
    if (before.current !== prefix.value + paragraph.text + suffix.value) throw captureMismatch("Word returned inconsistent surrounding document text.")
    const expectedText = prefix.value + expectedParagraph + suffix.value
    if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
    // All reads are complete. Fence before queuing tracking or any write; a
    // lost final sync is uncertain and must never trigger a second dispatch.
    this.applied.set(input.proposalId, { identity, authors: [], uncertain: true })
    context.document.changeTrackingMode = "TrackAll"
    for (const item of [...ranges].reverse()) {
      // A verified zero-length range inserts without deleting source text.
      // Pure deletions use the native deletion operation, excluding the mark.
      if (item.hunk.after) item.range.insertText(item.hunk.after, "Replace")
      else item.range.delete()
    }
    context.document.changeTrackingMode = fresh.tracking
    await context.sync()
    context.document.load("changeTrackingMode")
    const after = await this.documentState(context)
    const afterAll = context.document.body.paragraphs
    afterAll.load("items/uniqueLocalId,items/text,items/tableNestingLevel")
    const afterBodyXml = context.document.body.getRange("Whole").getOoxml()
    await context.sync()
    const post = projectNativeWordChatBody(afterBodyXml.value, afterAll.items.map(p => ({ id:p.uniqueLocalId, text:p.text, tableNestingLevel:p.tableNestingLevel })))
    const targetAfter = post.scope([at])
    const { added, missing } = newWordChatRevisions(before.changes, after.changes)
    const wanted = authored.revisions.filter(change => change.type !== "Formatted")
    const texts = (changes: Array<{ type: string; text: string }>, type: string) => changes.filter(change => change.type === type).map(change => change.text).sort()
    const authors = [...new Set(added.map(change => change.author).filter(Boolean))]
    const sameParagraphs = after.paragraphs.length === before.paragraphs.length &&
      after.paragraphs.every((item, index) => item.id === before.paragraphs[index]!.id && (index !== at ||
        item.text === expectedParagraph) && (index === at || JSON.stringify(item) === JSON.stringify(before.paragraphs[index])))
    // Checked in order; the first failing stage is kept as the cause.
    const checks: Array<[string, () => boolean]> = [
      ["text-tracking-mode", () => context.document.changeTrackingMode === fresh.tracking],
      ["text-paragraphs", () => sameParagraphs],
      ["text-current", () => after.current === expectedText],
      ["text-prior-revisions", () => !missing.length],
      ["text-new-revisions", () => added.length > 0 && added.every(change => ["Added", "Deleted"].includes(change.type) && Boolean(change.author)) &&
        JSON.stringify(texts(added, "Added")) === JSON.stringify(texts(wanted, "Added")) && JSON.stringify(texts(added, "Deleted")) === JSON.stringify(texts(wanted, "Deleted"))],
      ["text-geometry", () => sameNativeProjectionGeometry(pre, post)],
      ["text-paragraph-properties", () => sameWordChatParagraphProperties(pre.scope([at]), targetAfter)],
      ["text-hunk-formatting", () => sameWordChatHunkFormatting(pre.scope([at]), targetAfter, hunks)],
      ["text-adjacent", () => fresh.adjacent.every(item => { const i = post.paragraphs.findIndex(p => p.id === item.id); return i >= 0 && sameNativeWordChatParagraphContent(pre.scope([i]), post.scope([i])) })],
      ["text-comments", () => JSON.stringify(after.comments) === JSON.stringify(before.comments)],
    ]
    const failedStage = checks.find(([, check]) => !check())?.[0]
    if (failedStage) throw new WordChatHostError("Word applied the wording, but could not confirm that only the reviewed passage and tracked markup changed. Inspect it and use Undo if needed. Do not apply this proposal again.", "unchanged", { cause: new Error(failedStage) })
    this.applied.set(input.proposalId, { identity, authors, uncertain: false })
    return { status: "applied", tracked: true, saved: false, authors, target: input.replacement.trim() ? {
      ...target, id: crypto.randomUUID(), text: input.replacement,
      paragraphs: [{ id: paragraph.id, text: expectedParagraph }],
      startOffset: target.startOffset, endOffset: target.startOffset + input.replacement.length,
      saved: false, capturedAt: new Date().toISOString(),
    } : undefined }
  }

  private async applyNativeFormatting(context: WordChatContext, input: WordChatApplyInput, identity: string,
    fresh: Awaited<ReturnType<WordChatHost["importSnapshot"]>>, authored: WordChatAuthoredAmendment,
    inserted = new Set<number>(), expectedFormatCount = authored.revisions.length): Promise<WordChatApplyResult> {
    const formatting = input.formatting!
    if (input.mayDispatch && !input.mayDispatch()) throw new WordChatHostError("The Word pane closed before the change was applied. Your document is unchanged.")
    const paragraphs = input.target.paragraphs.map(p => context.document.getParagraphByUniqueLocalId(p.id))
    this.applied.set(input.proposalId, { identity, authors: [], uncertain: true })
    const nativeFormatting = formatting.filter(op => !inserted.has(op.paragraph))
    if (nativeFormatting.length && !fresh.trackFormatting) context.document.trackFormatting = true
    if (nativeFormatting.length) context.document.changeTrackingMode = "TrackAll"
    for (const op of nativeFormatting) {
      const paragraph = paragraphs[op.paragraph]!
      if (op.style !== undefined) paragraph.styleBuiltIn = op.style
      if (op.spacingBefore !== undefined) paragraph.spaceBefore = op.spacingBefore
      if (op.spacingAfter !== undefined) paragraph.spaceAfter = op.spacingAfter
      if (op.lineSpacing !== undefined) paragraph.lineSpacing = op.lineSpacing * 12
    }
    if (nativeFormatting.length) context.document.changeTrackingMode = fresh.tracking
    if (nativeFormatting.length && !fresh.trackFormatting) context.document.trackFormatting = false
    await context.sync()
    context.document.load("changeTrackingMode,trackFormatting")
    const all = context.document.body.paragraphs
    all.load(fresh.projection ? "items/uniqueLocalId,items/text,items/tableNestingLevel" : "items/uniqueLocalId")
    const range = paragraphs[0]!.getRange("Content").expandTo(paragraphs.at(-1)!.getRange("Content"))
    const xml = fresh.projection ? undefined : range.getOoxml(), bodyXml = fresh.projection ? context.document.body.getRange("Whole").getOoxml() : undefined, current = range.getReviewedText("Current")
    const changes = context.document.body.getRange().getTrackedChanges()
    changes.load("items/author,items/type,items/text")
    const adjacent = fresh.projection ? [] : fresh.adjacent.map(p => ({ before: p.ooxml, after: context.document.getParagraphByUniqueLocalId(p.id).getRange("Whole").getOoxml() }))
    await context.sync()
    let authors: string[] = []
    let failedStage: string | undefined
    let verificationError: unknown
    let checking = "format-projection"
    // Preserve the original short-circuit order: later checks do not run once
    // an earlier predicate has failed. Codes contain no document information.
    let confirmed = context.document.changeTrackingMode === fresh.tracking
    if (!confirmed) failedStage = "format-tracking-mode"
    if (confirmed && context.document.trackFormatting !== fresh.trackFormatting) { confirmed = false; failedStage = "format-tracking-property" }
    if (confirmed && JSON.stringify(all.items.map(p => p.uniqueLocalId)) !== JSON.stringify(fresh.ids)) { confirmed = false; failedStage = "format-paragraph-identity" }
    if (confirmed && current.value !== input.target.text) { confirmed = false; failedStage = "format-current-text" }
    try {
      const projection = bodyXml ? projectNativeWordChatBody(bodyXml.value, all.items.map(p => ({id:p.uniqueLocalId,text:p.text,tableNestingLevel:p.tableNestingLevel}))) : undefined
      const observed = projection ? projection.scope(input.target.paragraphs.map((_, i) => fresh.index + i)) : xml!.value
      checking = "format-authored-verification"
      authors = verifiedNativeWordChatFormatting(fresh.ooxml, authored.ooxml, observed, formatting, inserted)
      checking = "format-geometry"
      if (confirmed && projection && fresh.projection && projection.geometry !== fresh.projection.geometry) { confirmed = false; failedStage = checking }
      checking = "format-paragraph-paths"
      if (confirmed && projection && fresh.projection && !projection.paragraphs.every((p, i) => p.id === fresh.projection!.paragraphs[i]?.id && p.path === fresh.projection!.paragraphs[i]?.path)) { confirmed = false; failedStage = checking }
      checking = "format-neighbours"
      if (confirmed && !(projection && fresh.projection ? fresh.adjacent.every(p => { const i = projection.paragraphs.findIndex(q => q.id === p.id); return i >= 0 && sameNativeWordChatParagraphContent(p.ooxml, projection.scope([i])) }) : adjacent.every(p => sameNativeWordChatParagraphContent(p.before, p.after.value)))) { confirmed = false; failedStage = checking }
      checking = "format-revisions"
      const { added, missing } = newWordChatRevisions(fresh.changes, changes.items)
      if (confirmed && missing.length) { confirmed = false; failedStage = "format-prior-revisions" }
      if (confirmed && expectedFormatCount !== 0 && added.length === 0) { confirmed = false; failedStage = "format-missing-revision" }
      if (confirmed && !added.every(change => change.type === "Formatted" && authors.includes(change.author))) { confirmed = false; failedStage = "format-new-revisions" }
    } catch (error) { confirmed = false; failedStage ??= checking; verificationError = error }
    if (!confirmed) {
      const detail = new Error(failedStage ?? "format-verification", verificationError === undefined ? undefined : { cause: verificationError })
      throw new WordChatHostError("Word changed the formatting, but could not confirm the reviewed tracked markup. Inspect it and use Undo if needed. Do not apply this proposal again.", "uncertain", { cause: detail })
    }
    if (authored.wording) authors = [...new Set(["doc.haus", ...authors])]
    this.applied.set(input.proposalId, { identity, authors, uncertain: false })
    return { status: "applied", tracked: true, saved: false, authors, target: { ...input.target, id: crypto.randomUUID(), saved: false, capturedAt: new Date().toISOString() } }
  }

}
let defaultHost: WordChatHost | undefined
const host = () => defaultHost ??= new WordChatHost()
export const captureCurrentWordContext = (includeSelection = false) => host().capture(includeSelection)
export const captureWordChatSelection = () => host().selection()
export const applyReviewedProposal = (proposal: import("../../../dochaus/lib/word-chat-amendment-contract").WordChatProposal,
  author: WordChatApplyInput["author"], mayDispatch?: () => boolean) => {
  if (proposal.before !== proposal.target.text) throw new WordChatHostError("Review a fresh amendment before applying it.")
  return host().apply({proposalId: proposal.id,target: proposal.target,replacement: proposal.after,formatting: proposal.formatting,inlineFormatting: proposal.inlineFormatting,paragraphAction:proposal.paragraphAction,list:proposal.list,author,mayDispatch})
}

export const applyReviewedTableRow = (proposal: WordChatTableRowInsertionProposal, exportCurrent: () => Promise<Uint8Array>,
  author: (bytes: Uint8Array) => Promise<WordChatAuthoredTableRow>, mayDispatch?: () => boolean) => host().applyRow(proposal, exportCurrent, author, mayDispatch)

export const applyWordChatRowDeletion = (proposal: WordChatTableRowDeletionProposal, exportCurrent: () => Promise<Uint8Array>,
  author: (bytes: Uint8Array) => Promise<WordChatAuthoredTableRowDeletion>, mayDispatch?: () => boolean) => host().applyRowDeletion(proposal, exportCurrent, author, mayDispatch)
