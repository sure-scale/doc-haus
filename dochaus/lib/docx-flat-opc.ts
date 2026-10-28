import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { applyDocxParagraphEdits, createDocxRevision, type DocxParagraphEdit } from "./docx-amendments"
import { docxodus } from "./docxodus"
import { validWordChatFormatting, wordChatApplyPlan, type WordChatParagraphFormat } from "./word-chat-amendment-contract"
import type { WordChatAuthoredAmendment, WordChatTarget } from "./word-chat-amendment-contract"
import { reviewedWordChatParagraphFormatting, wordChatParagraphFormatChanged } from "./word-chat-formatting-ooxml"
import { assertLayeredRoundTrip } from "./redline-round-trip"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const REL = "http://schemas.openxmlformats.org/package/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const MAIN = "word/document.xml"
const AUTHOR = "doc.haus" as const
export const MAX_WORD_AMENDMENT_OOXML_BYTES = 1024 * 1024
const serializer = new XMLSerializer()
const children = (node: Node) => Array.from(node.childNodes).filter(child => child.nodeType === 1) as Element[]
const all = (node: Element | Document, namespace: string, local: string) => Array.from(node.getElementsByTagNameNS(namespace, local))

export class WordAmendmentError extends Error {
  constructor(readonly code: "word_snapshot_changed" | "word_snapshot_unsupported" | "word_snapshot_large") {
    super(code)
    this.name = "WordAmendmentError"
  }
}
export class WordFormattingError extends Error {
  constructor(message: string) { super(message); this.name = "WordFormattingError" }
}
function require(condition: unknown, code: WordAmendmentError["code"] = "word_snapshot_unsupported"): asserts condition {
  if (!condition) throw new WordAmendmentError(code)
}
function parse(xml: string) {
  require(xml.length <= MAX_WORD_AMENDMENT_OOXML_BYTES * 2 && !/<!DOCTYPE/i.test(xml))
  const fail = () => { throw new WordAmendmentError("word_snapshot_unsupported") }
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(xml, "application/xml")
  require(doc.documentElement && !doc.doctype)
  return doc
}

// A Flat OPC XML part may inherit namespaces from pkg:package/xmlData. Its ZIP
// entry must keep that context when it becomes an independent XML document.
function standalone(element: Element): string {
  const clone = element.cloneNode(true) as Element
  const present = new Set(Array.from(clone.attributes).filter(a => a.namespaceURI === XMLNS).map(a => a.name))
  for (let ancestor = element.parentNode; ancestor && ancestor.nodeType === 1; ancestor = ancestor.parentNode) {
    for (const attr of Array.from((ancestor as Element).attributes)) {
      if (attr.namespaceURI !== XMLNS || present.has(attr.name)) continue
      clone.setAttributeNS(XMLNS, attr.name, attr.value)
      present.add(attr.name)
    }
  }
  return serializer.serializeToString(clone)
}

function packageFromFlatOpc(xml: string) {
  require(new TextEncoder().encode(xml).byteLength <= MAX_WORD_AMENDMENT_OOXML_BYTES, "word_snapshot_large")
  const root = parse(xml).documentElement
  require(root.namespaceURI === PKG && root.localName === "package")
  const parts: Record<string, Uint8Array> = Object.create(null), types = new Map<string, string>()
  const rows = children(root)
  require(rows.length > 0 && rows.length <= 64)
  for (const part of rows) {
    require(part.namespaceURI === PKG && part.localName === "part")
    const name = part.getAttributeNS(PKG, "name") ?? "", type = part.getAttributeNS(PKG, "contentType") ?? ""
    require(name.startsWith("/") && name === path.posix.normalize(name) && !/[\\\x00?#]/.test(name) && !name.endsWith("/") && name !== "/[Content_Types].xml")
    const key = name.slice(1)
    require(!parts[key] && type.length > 0 && type.length <= 256)
    const data = children(part)
    require(data.length === 1 && data[0]!.namespaceURI === PKG)
    if (data[0]!.localName === "xmlData") {
      const content = children(data[0]!)
      require(content.length === 1)
      parts[key] = strToU8(standalone(content[0]!))
    } else {
      require(data[0]!.localName === "binaryData")
      const encoded = (data[0]!.textContent ?? "").replace(/\s/g, "")
      require(/^(?:[a-zA-Z0-9+/]{4})*(?:[a-zA-Z0-9+/]{2}==|[a-zA-Z0-9+/]{3}=)?$/.test(encoded))
      parts[key] = new Uint8Array(Buffer.from(encoded, "base64"))
    }
    types.set(key, type)
  }
  require(parts[MAIN])
  require(parts["_rels/.rels"])
  const roots = parse(strFromU8(parts["_rels/.rels"]!))
  require(roots.documentElement.namespaceURI === REL && children(roots.documentElement).some(rel => rel.namespaceURI === REL &&
    rel.localName === "Relationship" && rel.getAttribute("Type") === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" &&
    rel.getAttribute("TargetMode") !== "External" && path.posix.resolve("/", rel.getAttribute("Target") ?? "").slice(1) === MAIN))
  // Complete a minimal selection package before shared tracked authoring adds
  // its settings relationship. Preservation is checked against this full base.
  if (!parts["word/_rels/document.xml.rels"]) {
    parts["word/_rels/document.xml.rels"] = strToU8(`<Relationships xmlns="${REL}"/>`)
    types.set("word/_rels/document.xml.rels", "application/vnd.openxmlformats-package.relationships+xml")
  }
  const doc = parse(`<Types xmlns="${CT}"/>`)
  for (const [extension, type] of [["rels", "application/vnd.openxmlformats-package.relationships+xml"], ["xml", "application/xml"]]) {
    const fallback = doc.createElementNS(CT, "Default")
    fallback.setAttribute("Extension", extension!); fallback.setAttribute("ContentType", type!)
    doc.documentElement.appendChild(fallback)
  }
  for (const [name, type] of types) {
    const override = doc.createElementNS(CT, "Override")
    override.setAttribute("PartName", `/${name}`); override.setAttribute("ContentType", type)
    doc.documentElement.appendChild(override)
  }
  parts["[Content_Types].xml"] = strToU8(serializer.serializeToString(doc))
  return { parts, types }
}

const protectedContent = ["commentRangeStart", "commentRangeEnd", "commentReference", "ins", "del", "moveFrom", "moveTo",
  "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange",
  "fldChar", "fldSimple", "instrText", "drawing", "pict", "object", "sdt", "hyperlink", "footnoteReference", "endnoteReference", "sectPr"]
function bodyParagraphs(doc: Document, unrevised: boolean, formatting = false) {
  const root = doc.documentElement
  require(root.namespaceURI === W && root.localName === "document")
  const bodies = children(root).filter(e => e.namespaceURI === W && e.localName === "body")
  require(bodies.length === 1)
  const body = bodies[0]!, bodyChildren = children(body), content = bodyChildren.filter(e => !(e.namespaceURI === W && e.localName === "sectPr" && e === bodyChildren.at(-1)))
  require(content.length > 0)
  let paragraphs: Element[]
  if (content.length === 1 && content[0]!.namespaceURI === W && content[0]!.localName === "tbl") {
    // Word may wrap a selected cell's paragraph in its containing table. Only a
    // complete single-cell snapshot can become paragraph-only scoped transport;
    // the native adapter separately proves the actual target stays in that cell.
    const table = content[0]!, cells = all(table, W, "tc")
    require(cells.length === 1 && all(table, W, "tbl").length === 0)
    paragraphs = children(cells[0]!).filter(e => e.namespaceURI === W && e.localName === "p")
    require(paragraphs.length === all(table, W, "p").length)
  } else {
    require(content.every(e => e.namespaceURI === W && e.localName === "p"))
    paragraphs = content
  }
  require(paragraphs.length > 0 && paragraphs.length <= 64)
  if (unrevised) require(!paragraphs.some(p => protectedContent.some(name => !(formatting && ["ins", "del"].includes(name)) && all(p, W, name).length)))
  if (formatting) require(!paragraphs.some(p => all(p, W, "pPr").some(properties => all(properties, W, "ins").length || all(properties, W, "del").length)))
  require(!paragraphs.some(p => all(p, W, "br").some(br => br.getAttributeNS(W, "type") && br.getAttributeNS(W, "type") !== "textWrapping")))
  return { body, paragraphs }
}
function paragraphText(paragraph: Element): string {
  const text = (node: Element): string => node.namespaceURI === W && ["pPr", "rPr", "del"].includes(node.localName) ? ""
    : node.namespaceURI === W && node.localName === "t" ? node.textContent ?? ""
    : node.namespaceURI === W && node.localName === "tab" ? "\t"
    : node.namespaceURI === W && ["br", "cr"].includes(node.localName) ? "\v" : children(node).map(text).join("")
  return text(paragraph)
}

async function authorParagraphFormatting(source: ReturnType<typeof packageFromFlatOpc>, bytes: Uint8Array, before: Element[], formatting: WordChatParagraphFormat[]) {
  const dx = await docxodus(), session = dx.openDocxSession(bytes, { smartQuotes: false, persistAnchorIds: false })
  let baked: Uint8Array
  try {
    const anchors = Object.entries(session.project().anchorIndex).filter(([, value]) => value.partUri === "/word/document.xml" && ["p", "h", "li"].includes(value.kind)).map(([id]) => id)
    require(anchors.length === before.length)
    // Anchor order is ephemeral. The complete result below proves every
    // paragraph against its captured ordinal, including duplicate wording.
    for (const op of formatting) {
      let anchor = anchors[op.paragraph]!
      const membership = (id: string) => { const value = session.getListMembership(id); return value ? { ...value, anchorId: undefined } : null }
      const numbered = JSON.stringify(membership(anchor))
      if (op.style !== undefined) {
        if (!session.listStyles().some(style => style.id === op.style && style.type === "paragraph"))
          throw new WordFormattingError("That paragraph style is not available in this document. Choose an existing style in Word, then prepare the change again.")
        const changed = session.setParagraphStyle(anchor, op.style)
        require(changed.success)
        anchor = changed.modified[0]?.id ?? anchor
      }
      if (op.spacingBefore !== undefined || op.spacingAfter !== undefined || op.lineSpacing !== undefined)
        require(session.setParagraphFormat(anchor, {
          ...(op.spacingBefore !== undefined ? { spacingBefore: op.spacingBefore * 20 } : {}),
          ...(op.spacingAfter !== undefined ? { spacingAfter: op.spacingAfter * 20 } : {}),
          ...(op.lineSpacing !== undefined ? { lineSpacing: op.lineSpacing * 240, lineSpacingRule: "auto" as const } : {}),
        }).success)
      if (JSON.stringify(membership(anchor)) !== numbered)
        throw new WordFormattingError("That style would change this paragraph’s numbering. Keep its current style and adjust the spacing, or change the numbering directly in Word.")
    }
    // Deliver only the SDK-authored main part; styles/numbering and all other
    // dependencies remain the captured source, never rewritten by this route.
    const main = unzipSync(session.save())[MAIN]!
    baked = zipSync({ ...source.parts, [MAIN]: main }, { mtime: new Date(1980, 0, 1) })
  } finally { session.close() }
  const parts = unzipSync(baked), doc = parse(strFromU8(parts[MAIN]!)), after = bodyParagraphs(doc, false).paragraphs
  require(after.length === before.length)
  const operations = new Map(formatting.map(op => [op.paragraph, op]))
  let id = Math.max(0, ...all(doc, W, "*").map(node => Number(node.getAttributeNS(W, "id"))).filter(Number.isSafeInteger))
  let changed = 0
  for (const [index, paragraph] of after.entries()) {
    if (!wordChatParagraphFormatChanged(before[index]!, paragraph, operations.get(index))) continue
    changed++
    const properties = children(paragraph).find(node => node.namespaceURI === W && node.localName === "pPr")!
    const previous = children(before[index]!).find(node => node.namespaceURI === W && node.localName === "pPr")
    const change = doc.createElementNS(W, "w:pPrChange")
    change.setAttributeNS(W, "w:id", String(++id)); change.setAttributeNS(W, "w:author", AUTHOR); change.setAttributeNS(W, "w:date", new Date().toISOString())
    change.appendChild(previous ? doc.importNode(previous, true) : doc.createElementNS(W, "w:pPr"))
    properties.appendChild(change)
  }
  if (!changed) throw new WordFormattingError("These paragraphs already use the proposed formatting. Choose a different change.")
  const revisions = reviewedWordChatParagraphFormatting(before, after, formatting).flat()
  parts[MAIN] = strToU8(serializer.serializeToString(doc))
  const authored = zipSync(parts, { mtime: new Date(1980, 0, 1) })
  await assertLayeredRoundTrip(authored, bytes, baked)
  // The shared round-trip checks content and prior revisions. Also prove exact
  // paragraph/inline XML and properties after resolving only our format marks.
  for (const accept of [false, true]) {
    const resolved = dx.openDocxSession(authored, { persistAnchorIds: false })
    try {
      for (const revision of resolved.listRevisions().filter(revision => revision.type === "format"))
        require((accept ? resolved.acceptRevision(revision.id) : resolved.rejectRevision(revision.id)).success)
      const result = bodyParagraphs(parse(strFromU8(unzipSync(resolved.save())[MAIN]!)), false).paragraphs
      const expected = accept ? bodyParagraphs(parse(strFromU8(unzipSync(baked)[MAIN]!)), false).paragraphs : before
      require(result.length === expected.length)
      result.forEach((paragraph, index) => wordChatParagraphFormatChanged(expected[index]!, paragraph))
    } finally { resolved.close() }
  }
  return { ooxml: importFlatOpc(parts, source.types), paragraphs: before.map(paragraphText), revisions, author: AUTHOR, formatting }
}

function paragraphMark(paragraph: Element) {
  const properties = children(paragraph).find(p => p.namespaceURI === W && p.localName === "pPr")
  const run = properties && children(properties).find(p => p.namespaceURI === W && p.localName === "rPr")
  return run && children(run).find(p => p.namespaceURI === W && ["ins", "del"].includes(p.localName))
}

async function authorParagraphStructure(bytes: Uint8Array, texts: string[], expected: string[], structure: NonNullable<Extract<ReturnType<typeof wordChatApplyPlan>, { supported: true }>["structure"]>) {
  const dx = await docxodus(), session = dx.openDocxSession(bytes, { smartQuotes: false, persistAnchorIds: false, trackedChanges: "render_inline", revisionAuthor: AUTHOR })
  try {
    const matches = session.grep("^[\\s\\S]*$", { scope: dx.ProjectionScopes.Body, whitespace: 0 })
    require(JSON.stringify(matches.map(match => match.text)) === JSON.stringify(texts), "word_snapshot_changed")
    if (structure.kind === "insert") {
      const kept = new Set(structure.retained.map(p => p.after))
      for (let index = 0; index < expected.length; index++) if (!kept.has(index)) {
        const next = structure.retained.find(p => p.after > index)
        require(next)
        // The mutation API consumes Markdown. Escape syntax so the reviewed
        // wording, including punctuation, is inserted as ordinary literal text.
        const literal = expected[index]!.replace(/([\\`*_[\]{}()#+.!>|~\-])/g, "\\$1")
        require(session.insertParagraph(matches[next.before]!.enclosingAnchor.id, "before", literal).success)
      }
    } else {
      const kept = new Set(structure.retained.map(p => p.before))
      for (let index = 0; index < texts.length; index++) if (!kept.has(index))
        require(session.deleteBlock(matches[index]!.enclosingAnchor.id).success)
    }
    return session.save()
  } finally { session.close() }
}

function importFlatOpc(parts: Record<string, Uint8Array>, originalTypes: Map<string, string>) {
  const contentTypes = parse(strFromU8(parts["[Content_Types].xml"]!))
  const typeFor = (name: string) => all(contentTypes, CT, "Override").find(e => e.getAttribute("PartName") === `/${name}`)?.getAttribute("ContentType")
    ?? all(contentTypes, CT, "Default").find(e => e.getAttribute("Extension") === name.split(".").at(-1))?.getAttribute("ContentType") ?? originalTypes.get(name)
  const selected: Record<string, Uint8Array> = Object.create(null), queue = [MAIN]
  const permitted = new Set(["styles", "theme", "fontTable", "numbering", "font"])
  while (queue.length) {
    const name = queue.shift()!
    if (selected[name]) continue
    require(parts[name]); selected[name] = parts[name]!
    const relName = path.posix.join(path.posix.dirname(name), "_rels", `${path.posix.basename(name)}.rels`)
    if (!parts[relName]) continue
    const rels = parse(strFromU8(parts[relName]!))
    require(rels.documentElement.namespaceURI === REL && rels.documentElement.localName === "Relationships")
    for (const rel of children(rels.documentElement)) {
      require(rel.namespaceURI === REL && rel.localName === "Relationship")
      const kind = (rel.getAttribute("Type") ?? "").split("/").at(-1)!
      // Retained numbering/theme parts can reference embedded images (for
      // example a picture bullet). Their relationship closure must travel with
      // the formatting; otherwise Word receives a dangling image reference.
      if (!permitted.has(kind) && !(kind === "image" && name !== MAIN)) { rel.parentNode!.removeChild(rel); continue }
      require(rel.getAttribute("TargetMode") !== "External")
      const target = rel.getAttribute("Target") ?? ""
      require(target && !/[\\\x00?#]/.test(target))
      const resolved = path.posix.resolve("/", path.posix.dirname(name), target).slice(1)
      require(parts[resolved]); queue.push(resolved)
    }
    if (children(rels.documentElement).length) selected[relName] = strToU8(serializer.serializeToString(rels))
  }
  const doc = parse(strFromU8(selected[MAIN]!)), { body } = bodyParagraphs(doc, false)
  for (const child of children(body)) if (child.namespaceURI === W && child.localName === "sectPr") body.removeChild(child)
  selected[MAIN] = strToU8(serializer.serializeToString(doc))
  selected["_rels/.rels"] = strToU8(`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  const flat = parse(`<pkg:package xmlns:pkg="${PKG}"/>`)
  for (const [name, bytes] of Object.entries(selected)) {
    const type = name.endsWith(".rels") ? "application/vnd.openxmlformats-package.relationships+xml" : typeFor(name)
    require(type)
    const part = flat.createElementNS(PKG, "pkg:part")
    part.setAttributeNS(PKG, "pkg:name", `/${name}`); part.setAttributeNS(PKG, "pkg:contentType", type)
    const xml = type.endsWith("+xml") || type === "application/xml"
    const data = flat.createElementNS(PKG, xml ? "pkg:xmlData" : "pkg:binaryData")
    if (xml) data.appendChild(flat.importNode(parse(strFromU8(bytes)).documentElement, true))
    else data.textContent = Buffer.from(bytes).toString("base64")
    part.appendChild(data); flat.documentElement.appendChild(part)
  }
  return serializer.serializeToString(flat)
}

/** Transform a live, complete paragraph snapshot privately through the shared
 * amendment engine. Native session IDs never become Docxodus paragraph IDs;
 * exact ordered text and selected UTF-16 offsets bind this ephemeral package. */
export async function authorWordChatTarget(ooxml: string, target: WordChatTarget, replacement: string, formatting?: WordChatParagraphFormat[]): Promise<Omit<WordChatAuthoredAmendment, "proposalId">> {
  const source = packageFromFlatOpc(ooxml), doc = parse(strFromU8(source.parts[MAIN]!)), { body, paragraphs } = bodyParagraphs(doc, true, Boolean(formatting))
  const texts = paragraphs.map(paragraphText), proposed = replacement.split("\r")
  require(JSON.stringify(texts) === JSON.stringify(target.paragraphs.map(p => p.text)), "word_snapshot_changed")
  // Authoring supports every valid paragraph operation. The native host/UI
  // separately gate transport limitations, including the final paragraph mark.
  if (formatting !== undefined) require(replacement === target.text && validWordChatFormatting(target, formatting))
  const plan = wordChatApplyPlan(target, replacement)
  require(plan.supported)
  require(Number.isSafeInteger(target.startOffset) && Number.isSafeInteger(target.endOffset) && target.startOffset >= 0 && target.endOffset >= 0 &&
    target.startOffset <= texts[0]!.length && target.endOffset <= texts.at(-1)!.length && (texts.length > 1 || target.endOffset >= target.startOffset), "word_snapshot_changed")
  const spans = texts.map((text, i) => ({ start: i === 0 ? target.startOffset : 0, end: i === texts.length - 1 ? target.endOffset : text.length }))
  require(texts.map((text, i) => text.slice(spans[i]!.start, spans[i]!.end)).join("\r") === target.text, "word_snapshot_changed")
  // Strip only an optional table/cell transport wrapper, retaining each complete
  // paragraph XML. No table geometry or unrelated content is imported into Word.
  for (const child of children(body)) if (!(child.namespaceURI === W && child.localName === "sectPr")) body.removeChild(child)
  const section = children(body)[0] ?? null
  for (const p of paragraphs) body.insertBefore(p, section)
  source.parts[MAIN] = strToU8(serializer.serializeToString(doc))
  const bytes = zipSync(source.parts, { mtime: new Date(1980, 0, 1) })
  if (formatting) return authorParagraphFormatting(source, bytes, paragraphs, formatting)
  const options = { pendingRevisions: "layer" as const, revisionAuthor: AUTHOR, ownAuthors: [AUTHOR] }
  const revision = plan.structure ? undefined : await createDocxRevision(bytes, options), projected = revision?.paragraphs.filter(p => p.scope === "body") ?? []
  const nonempty = texts.filter(text => text.replace(/[\v\t]/g, "").length)
  if (!plan.structure) require(JSON.stringify(projected.map(p => p.text)) === JSON.stringify(nonempty.map(text => text.replace(/[\v\t]/g, ""))), "word_snapshot_changed")
  const edits: DocxParagraphEdit[] = [], expected = plan.paragraphs
  let ordinal = 0
  if (!plan.structure) texts.forEach((text, i) => {
    const paragraph = text.replace(/[\v\t]/g, "") ? projected[ordinal++] : undefined, { start, end } = spans[i]!
    const before = text.slice(start, end), after = proposed[i]!
    if (before === after) return
    require(paragraph)
    const pieces = before.split(/([\v\t])/), replacements = after.split(/([\v\t])/)
    require(pieces.length === replacements.length && pieces.every((piece, at) => at % 2 === 0 || piece === replacements[at]))
    let at = text.slice(0, start).replace(/[\v\t]/g, "").length
    pieces.forEach((piece, part) => {
      if (part % 2) return
      if (piece !== replacements[part]) {
        require(piece.length > 0)
        const projectedText = paragraph.text
        let occurrence = 0, found = projectedText.indexOf(piece)
        while (found >= 0 && found < at) { occurrence++; found = projectedText.indexOf(piece, found + piece.length) }
        require(found === at, "word_snapshot_changed")
        edits.push({ id: `paragraph-${i}-span-${part}`, paragraph_handle: paragraph.paragraph_handle, occurrence, before_quote: piece, replacement: replacements[part]! })
      }
      at += piece.length
    })
  })
  require(plan.structure || edits.length > 0)
  const amended = plan.structure ? await authorParagraphStructure(bytes, texts, expected, plan.structure)
    : (await applyDocxParagraphEdits(bytes, edits, undefined, options)).bytes
  const outputParts = unzipSync(amended)
  const outputDoc = parse(strFromU8(outputParts[MAIN]!)), output = bodyParagraphs(outputDoc, false).paragraphs
  const current = output.filter(p => paragraphMark(p)?.localName !== "del")
  // Current text excludes deletions. Keep native soft breaks/tabs in the result.
  for (const deletion of all(outputDoc, W, "del")) deletion.parentNode!.removeChild(deletion)
  require(JSON.stringify(current.map(paragraphText)) === JSON.stringify(expected), "word_snapshot_changed")
  const authored = parse(strFromU8(outputParts[MAIN]!)), revisions: WordChatAuthoredAmendment["revisions"] = []
  for (const p of bodyParagraphs(authored, false).paragraphs) {
    const visit = (node: Element) => {
      if (node.namespaceURI === W && ["ins", "del"].includes(node.localName)) {
        const added = node.localName === "ins", author = node.getAttributeNS(W, "author") ?? ""
        require(author === AUTHOR)
        revisions.push({ type: added ? "Added" : "Deleted", author, text: all(node, W, added ? "t" : "delText").map(t => t.textContent ?? "").join("") })
      } else if (!(node.namespaceURI === W && ["pPr", "rPr"].includes(node.localName))) children(node).forEach(visit)
    }
    visit(p)
    const mark = paragraphMark(p)
    if (mark) { require(mark.getAttributeNS(W, "author") === AUTHOR); revisions.push({ type: mark.localName === "ins" ? "Added" : "Deleted", author: AUTHOR, text: "\r" }) }
  }
  require(revisions.length > 0)
  return { ooxml: importFlatOpc(outputParts, source.types), paragraphs: expected, revisions, author: AUTHOR }
}
