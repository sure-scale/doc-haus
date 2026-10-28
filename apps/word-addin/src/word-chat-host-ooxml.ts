import { DOMParser } from "@xmldom/xmldom"
import { validWordChatFormatting, type WordChatAuthoredAmendment, type WordChatParagraphFormat, type WordChatTarget } from "../../../dochaus/lib/word-chat-amendment-contract"
import { reviewedWordChatParagraphFormatting, wordChatParagraphFormatChanged } from "../../../dochaus/lib/word-chat-formatting-ooxml"

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const XMLNS = "http://www.w3.org/2000/xmlns/"
const MAIN = "/word/document.xml"
const elements = (node: Node) => Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1)
function parse(xml: string) {
  if (typeof xml !== "string" || xml.length > 2_000_000 || /<!DOCTYPE/i.test(xml)) throw new Error("Invalid clause OOXML")
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
    originals.length !== observed.length || originals.some((marks, i) => marks.length !== observed[i]!.length)) return
  const pairs = originals.flatMap((marks, i) => marks.map((mark, j) => [mark, observed[i]![j]!] as const))
  const withoutId = (node: Element) => { const copy = node.cloneNode(true) as Element; copy.removeAttributeNS(W, "id"); return canonical(copy) }
  if (pairs.some(([old, current]) => JSON.stringify(withoutId(old)) !== JSON.stringify(withoutId(current)))) return
  for (const [old, current] of pairs) current.setAttributeNS(W, "w:id", old.getAttributeNS(W, "id")!)
}

/** Verify native paragraph formatting against both the captured source and the
 * independently checked authored artifact. Word owns the revision author for
 * property mutations; no other author or revision content is normalized. */
export function verifiedNativeWordChatFormatting(captured: string, authored: string, observed: string, formatting: WordChatParagraphFormat[], inserted = new Set<number>()): string[] {
  const before = content(parse(captured), true, true).nodes
  const after = content(parse(observed), true, true).nodes.map(node => node.cloneNode(true) as Element)
  const authors: string[] = []
  for (const [index, paragraph] of after.entries()) {
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
