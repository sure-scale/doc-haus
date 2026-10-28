import type { WordChatAuthoredAmendment, WordChatParagraphFormat, WordChatSnapshot, WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { validWordChatFormatting, wordChatApplyPlan } from "../../../dochaus/lib/word-chat-amendment-contract"
import { capturedWordChatParagraphs, importedWordChatDifference, reviewedMixedWordChatArtifact, reviewedWordChatArtifact, wordChatOoxmlDifference, sameImportedWordChat, sameWordChatInlineRevisions, sameWordChatParagraphContent, verifiedNativeWordChatFormatting } from "./word-chat-host-ooxml"
import { DOMParser } from "@xmldom/xmldom"

type ClientResult<T> = { value: T }
type Collection<T> = { items: T[]; load(properties: string): void }
type TrackedChange = { author: string; type: string; text: string }
type Paragraph = { text: string; uniqueLocalId: string; tableNestingLevel: number; styleBuiltIn: string; spaceBefore: number; spaceAfter: number; lineSpacing: number; load(properties: string): void; getRange(location?: string): WordChatRange }
type WordChatRange = {
  text: string
  paragraphs: Collection<Paragraph>
  load(properties: string): void
  getRange(location?: string): WordChatRange
  expandTo(other: WordChatRange): WordChatRange
  compareLocationWith(other: WordChatRange): ClientResult<string>
  search(query: string, options: Record<string, boolean>): Collection<WordChatRange>
  insertOoxml(xml: string, location: "Replace"): WordChatRange
  getOoxml(): ClientResult<string>
  getReviewedText(version: "Current"): ClientResult<string>
  getComments(): Collection<unknown>
  getTrackedChanges(): Collection<TrackedChange>
}
export type WordChatContext = {
  document: {
    saved: boolean
    changeTrackingMode: string
    load(properties: string): void
    getSelection(): WordChatRange
    getParagraphByUniqueLocalId(id: string): Paragraph
    body: { text: string; load(properties: string): void; paragraphs: Collection<Paragraph>; getRange(location?: string): WordChatRange }
  }
  sync(): Promise<void>
}
export type WordChatWord = { run<T>(callback: (context: WordChatContext) => Promise<T>): Promise<T> }
export type WordChatApplyResult = { status: "applied" | "already-applied"; tracked: true; saved: false; authors: string[]; target?: WordChatTarget }
export type WordChatApplyInput = { proposalId: string; target: WordChatTarget; replacement: string; formatting?: WordChatParagraphFormat[]; author: (snapshot: { ooxml: string }) => Promise<WordChatAuthoredAmendment> }
export class WordChatHostError extends Error {
  constructor(message: string, readonly mutationState: "unchanged" | "uncertain" = "unchanged") { super(message); this.name = "WordChatHostError" }
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
function failure(error: unknown, fallback: string): never { if (error instanceof WordChatHostError) throw error; throw new WordChatHostError(fallback) }

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

/** getOoxml exports a Flat OPC package, including unrelated parts and a final
 * body-level section layout even for one selected word. Inspect only selected
 * main-document content. A section break inside a paragraph remains protected;
 * the final body sectPr is package context and is outside our text-only edits. */
function hasProtectedContent(value: string, authored = false, formatting = false, insertedMarks = false): boolean {
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
      if (wordNamespaces.has(node.namespaceURI ?? "") && protectedContent.has(node.localName) &&
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

/** Host-only adapter. No save, session, source or model authority is held here. */
export class WordChatHost {
  private applied = new Map<string, { identity: string; authors: string[]; uncertain: boolean }>()
  private pending = new Map<string, { identity: string; promise: Promise<WordChatApplyResult> }>()
  constructor(private word: WordChatWord = nativeWord(), private supports = supported) {}
  private ready() {
    if (!this.supports("WordApi", "1.6")) throw new WordChatHostError("Reviewed amendments require a current Microsoft 365 version of Word. Chat and Save remain available.")
  }
  async capture(includeSelection = false): Promise<WordChatSnapshot> {
    const structured = this.supports("WordApi", "1.6")
    try { return await this.readSnapshot(structured, includeSelection) }
    catch (error) {
      if (includeSelection || error instanceof WordChatHostError || !structured) failure(error, "Word could not read its current text. Keep your message and try again.")
      // Some licensed hosts advertise the API without providing paragraph IDs.
      // Read current text in a fresh Office batch; never substitute a saved copy.
      return this.readSnapshot(false, false)
    }
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
  private async resolve(context: WordChatContext, target: WordChatTarget): Promise<{ range: WordChatRange; paragraphs: Paragraph[] }> {
    validate(target)
    const paragraphs = target.paragraphs.map((p) => context.document.getParagraphByUniqueLocalId(p.id))
    for (const paragraph of paragraphs) paragraph.load("text,uniqueLocalId,tableNestingLevel")
    await context.sync()
    // Word can alter list numbering outside a table when importing a scoped
    // paragraph fragment into a cell. Until that transport is proven safe,
    // refuse before authoring or mutation, including after the freshness read.
    if (paragraphs.some(paragraph => paragraph.tableNestingLevel !== 0))
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

  async apply(input: WordChatApplyInput): Promise<WordChatApplyResult> {
    this.ready()
    validate(input.target)
    if (typeof input.author !== "function") throw new WordChatHostError("Reconnect to doc.haus to author this amendment before applying it. Your document is unchanged.")
    if (!input.proposalId || typeof input.replacement !== "string" || input.replacement.length > LIMIT || input.replacement === input.target.text && input.formatting === undefined)
      throw new WordChatHostError("Prepare and review a change before applying it.")
    const identity = JSON.stringify([input.target, input.replacement, input.formatting])
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
    const promise = this.applyOnce(input, identity)
    this.pending.set(input.proposalId, { identity, promise })
    try { return await promise } finally { this.pending.delete(input.proposalId) }
  }

  /** Complete paragraphs form the import transport. The approved edit still
   * names only target.text/startOffset/endOffset, checked in both authored views. */
  private async importSnapshot(context: WordChatContext, target: WordChatTarget, formatting = false, wordingImport = false, insertedMarks = false) {
    const { range: selected, paragraphs } = await this.resolve(context, target)
    const all = context.document.body.paragraphs
    let neighbours: Array<{ id: string; range: WordChatRange }> = []
    if (formatting) {
      all.load("items/uniqueLocalId")
      await context.sync()
      const at = all.items.findIndex(p => p.uniqueLocalId === target.paragraphs[0]!.id)
      if (at < 0) throw captureMismatch("Word could not locate the reviewed paragraph among the document’s current paragraphs.")
      neighbours = [all.items[at - 1], all.items[at + target.paragraphs.length]].filter((p): p is Paragraph => Boolean(p))
        .map(p => ({ id: p.uniqueLocalId, range: p.getRange("Whole") }))
    }
    // Content exports the exact paragraphs and all their pPr. Word's Whole
    // export can add a synthetic trailing paragraph after the final mark.
    // Only formatting replacement owns that mark; never discard exported
    // paragraphs to make a Whole snapshot fit the reviewed target.
    const content = paragraphs[0]!.getRange("Content").expandTo(paragraphs.at(-1)!.getRange("Content"))
    const range = formatting && !wordingImport ? paragraphs[0]!.getRange("Whole").expandTo(paragraphs.at(-1)!.getRange("Whole")) : content
    // Every preservation baseline comes from this final read batch. Context
    // discovery yields to Word and must never leave an earlier XML snapshot as
    // the authority for a later complete-paragraph replacement.
    const xml = content.getOoxml(), comments = range.getComments(), revisions = range.getTrackedChanges()
    comments.load("items"); revisions.load("items/author,items/type")
    const current = paragraphs.map(paragraph => paragraph.getRange("Content").getReviewedText("Current"))
    context.document.load("changeTrackingMode")
    all.load("items/uniqueLocalId")
    const adjacent = neighbours.map(p => ({ id: p.id, xml: p.range.getOoxml() }))
    const bodyChanges = context.document.body.getRange().getTrackedChanges()
    bodyChanges.load("items/author,items/type,items/text")
    selected.load("text")
    for (const paragraph of paragraphs) paragraph.load("text")
    await context.sync()
    if (comments.items.length || (!formatting && revisions.items.length) || formatting && revisions.items.some(revision => !["Added", "Deleted"].includes((revision as TrackedChange).type)) || hasProtectedContent(xml.value, false, formatting, insertedMarks))
      throw new WordChatHostError("This paragraph contains existing revisions, comments or a linked Word object. Review that passage in Word first, then prepare a fresh suggestion for ordinary text.")
    let captured: string[]
    try { captured = capturedWordChatParagraphs(xml.value) }
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
    if (formatting && JSON.stringify([ids[index - 1], ids[index + target.paragraphs.length]].filter((id): id is string => id !== undefined)) !== JSON.stringify(adjacent.map(p => p.id)))
      throw captureMismatch("The paragraphs beside this review changed during capture.")
    return { range, ooxml: xml.value, ids, index, adjacent: adjacent.map(p => ({ id: p.id, ooxml: p.xml.value })), tracking: context.document.changeTrackingMode,
      changes: bodyChanges.items.map(change => ({ author: change.author, type: change.type, text: change.text })) }
  }

  private async applyOnce(input: WordChatApplyInput, identity: string, wordingImport = false): Promise<WordChatApplyResult> {
    try {
      const nativeFinalFormatting = input.formatting !== undefined && input.replacement === input.target.text &&
        validWordChatFormatting(input.target, input.formatting) && input.formatting.some(op => op.paragraph === input.target.paragraphs.length - 1)
      // Scoped OOXML imports may change numbering elsewhere in Word, even
      // outside the reviewed range. Stop before any authoring or mutation
      // until the import's document-wide effects have been established.
      if (!nativeFinalFormatting)
        throw new WordChatHostError("Word cannot safely apply this amendment yet because it may change numbering elsewhere in the document. Make the change directly in Word; your document is unchanged.")
      const mixed = input.formatting !== undefined && input.replacement !== input.target.text
      const afterParts = input.replacement.split("\r"), plan = wordChatApplyPlan(input.target, input.replacement, nativeFinalFormatting || mixed ? undefined : input.formatting)
      if (!plan.supported) throw new WordChatHostError(plan.reason)
      if (mixed && (!validWordChatFormatting(input.target, input.formatting, input.replacement) || plan.structure?.kind === "delete")) throw new WordChatHostError("This change needs a smaller reviewed amendment. Your Word document is unchanged.")
      const expected = plan.paragraphs
      const formatting = input.formatting !== undefined
      const snapshot = await this.word.run(context => this.importSnapshot(context, input.target, formatting || wordingImport, wordingImport))
      // Shared Docxodus authoring is read-only. No Office proxy or canonical
      // write is kept pending while that service authors the reviewed result.
      const authored = await input.author({ ooxml: snapshot.ooxml })
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
        const fresh = await this.importSnapshot(context, input.target, formatting || wordingImport, wordingImport)
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
        const signature = (change: TrackedChange) => JSON.stringify([change.author, change.type, change.text])
        const prior = new Map<string, number>()
        for (const change of fresh.changes) prior.set(signature(change), (prior.get(signature(change)) ?? 0) + 1)
        const added = changes.items.filter(change => {
          const key = signature(change), count = prior.get(key) ?? 0
          if (count) { prior.set(key, count - 1); return false }
          return true
        })
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
        const priorConfirmed = ![...prior.values()].some(Boolean), additionsConfirmed = matches(added, authored.revisions)
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
        throw new WordChatHostError(error instanceof WordChatHostError ? error.message : fallback, "uncertain")
      failure(error, fallback)
    }
  }

  private async applyNativeFormatting(context: WordChatContext, input: WordChatApplyInput, identity: string,
    fresh: Awaited<ReturnType<WordChatHost["importSnapshot"]>>, authored: WordChatAuthoredAmendment,
    inserted = new Set<number>(), expectedFormatCount = authored.revisions.length): Promise<WordChatApplyResult> {
    const formatting = input.formatting!
    const paragraphs = input.target.paragraphs.map(p => context.document.getParagraphByUniqueLocalId(p.id))
    this.applied.set(input.proposalId, { identity, authors: [], uncertain: true })
    const nativeFormatting = formatting.filter(op => !inserted.has(op.paragraph))
    try {
      if (nativeFormatting.length) context.document.changeTrackingMode = "TrackAll"
      for (const op of nativeFormatting) {
        const paragraph = paragraphs[op.paragraph]!
        if (op.style !== undefined) paragraph.styleBuiltIn = op.style
        if (op.spacingBefore !== undefined) paragraph.spaceBefore = op.spacingBefore
        if (op.spacingAfter !== undefined) paragraph.spaceAfter = op.spacingAfter
        if (op.lineSpacing !== undefined) paragraph.lineSpacing = op.lineSpacing * 12
      }
      await context.sync()
    } finally {
      if (nativeFormatting.length) {
        context.document.changeTrackingMode = fresh.tracking
        await context.sync()
      }
    }
    const all = context.document.body.paragraphs
    all.load("items/uniqueLocalId")
    const range = paragraphs[0]!.getRange("Content").expandTo(paragraphs.at(-1)!.getRange("Content"))
    const xml = range.getOoxml(), current = range.getReviewedText("Current")
    const changes = context.document.body.getRange().getTrackedChanges()
    changes.load("items/author,items/type,items/text")
    const adjacent = fresh.adjacent.map(p => ({ before: p.ooxml, after: context.document.getParagraphByUniqueLocalId(p.id).getRange("Whole").getOoxml() }))
    await context.sync()
    let authors: string[] = []
    let confirmed = JSON.stringify(all.items.map(p => p.uniqueLocalId)) === JSON.stringify(fresh.ids) && current.value === input.target.text
    try {
      authors = verifiedNativeWordChatFormatting(fresh.ooxml, authored.ooxml, xml.value, formatting, inserted)
      confirmed &&= adjacent.every(p => sameWordChatParagraphContent(p.before, p.after.value))
      const signature = (change: TrackedChange) => JSON.stringify([change.author, change.type, change.text])
      const prior = new Map<string, number>()
      for (const change of fresh.changes) prior.set(signature(change), (prior.get(signature(change)) ?? 0) + 1)
      const added = changes.items.filter(change => {
        const key = signature(change), count = prior.get(key) ?? 0
        if (count) { prior.set(key, count - 1); return false }
        return true
      })
      confirmed &&= ![...prior.values()].some(Boolean) && (expectedFormatCount === 0 || added.length > 0) &&
        added.every(change => change.type === "Formatted" && authors.includes(change.author))
    } catch { confirmed = false }
    if (!confirmed) throw new WordChatHostError("Word changed the formatting, but could not confirm the reviewed tracked markup. Inspect it and use Undo if needed. Do not apply this proposal again.")
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
  author: WordChatApplyInput["author"]) => {
  if (proposal.before !== proposal.target.text) throw new WordChatHostError("Review a fresh amendment before applying it.")
  return host().apply({proposalId: proposal.id,target: proposal.target,replacement: proposal.after,formatting: proposal.formatting,author})
}
