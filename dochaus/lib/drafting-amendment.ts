import { createHash, randomBytes } from "node:crypto"
import { readFile } from "node:fs/promises"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { z } from "zod"
import { openRevision } from "./docx-amendments"
import { writeDurableBytes } from "./durable-json"
import { docxodus } from "./docxodus"

// One amendment explicitly replaces, deletes or splits one paragraph, addressed by the
// paragraph pointer (pid) minted by ensureParagraphPointers. existing_sha256 is
// paragraphSha256 of the paragraph text the author saw; it is the only
// integrity check the applier performs.
export const amendmentSchema = z.object({
  id: z.string().min(1),
  pid: z.string().min(1),
  existing_sha256: z.string().regex(/^[0-9a-f]{64}$/, "expected paragraphSha256 hex of the paragraph text"),
  operation: z.enum(["replace", "delete", "split"]).optional(),
  updated: z.union([z.string(), z.array(z.string())]),
  reason: z.string().refine(text => text.trim().length > 0, "reason is empty"),
  sources: z.array(z.object({ source_id: z.string().min(1), quote: z.string().min(1) }).strict()),
}).strict().superRefine((row, ctx) => {
  const error = (message: string) => ctx.addIssue({ code: "custom", path: ["updated"], message })
  if (row.operation === "delete") {
    if (row.updated !== "") error("explicit deletion requires updated to be the empty string")
    return
  }
  const paragraphs = typeof row.updated === "string" ? [row.updated] : row.updated
  if (row.operation === "split" ? !Array.isArray(row.updated) || row.updated.length < 2 : typeof row.updated !== "string")
    error("use operation split with an array of at least two paragraphs, or a string for replacement")
  if (paragraphs.some(text => !text.trim())) error("updated paragraph text is empty; deletion requires operation delete")
  if (paragraphs.some(text => /[\r\n]/.test(text))) error("updated paragraph text must be one paragraph without line breaks")
})
export type Amendment = z.infer<typeof amendmentSchema>

export function amendmentParagraphs(amendment: Pick<Amendment, "operation" | "updated">): string[] {
  return amendment.operation === "delete" ? [] : typeof amendment.updated === "string" ? [amendment.updated] : amendment.updated
}

export function amendmentText(amendment: Pick<Amendment, "operation" | "updated">) {
  return amendmentParagraphs(amendment).join("\n\n")
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const BODY = "word/document.xml"
// Leading underscore keeps the bookmark hidden in Word's bookmark list.
const PREFIX = "_dh_"
const PID = /^p[0-9a-f]{12}$/

export function normalizeParagraphText(text: string) {
  return text.normalize("NFC").replace(/\s+/g, " ").trim()
}

export function paragraphSha256(text: string) {
  return createHash("sha256").update(normalizeParagraphText(text)).digest("hex")
}

export function paragraphBookmarkName(pid: string) {
  return `${PREFIX}${pid}`
}

// Pids of one opened body paragraph: hidden bookmarks whose start sits in this
// paragraph (not in a nested text-box paragraph or deleted revision history).
export function paragraphPointers(node: Element) {
  return Array.from(node.getElementsByTagNameNS(W, "bookmarkStart"))
    .filter(mark => owningParagraph(mark) === node && !deletedRevisionAncestor(mark))
    .flatMap(mark => {
      const name = mark.getAttributeNS(W, "name") ?? ""
      return name.startsWith(PREFIX) ? [{ pid: name.slice(PREFIX.length), mark }] : []
    })
}

// Gives every non-empty body paragraph exactly one invisible bookmark pointer.
// Existing pointers are kept; a pointer seen on an earlier paragraph (copied or
// split paragraphs) or a second pointer on one paragraph is replaced. The file
// is rewritten atomically only when a pointer was added or replaced.
export async function ensureParagraphPointers(docxPath: string) {
  const source = new Uint8Array(await readFile(docxPath))
  const opened = await openRevision(source)
  try {
    const result = addParagraphPointers(source, opened.parts, opened.documents[BODY]!, opened.paragraphs
      .filter(paragraph => paragraph.part === BODY).map(paragraph => ({ node: paragraph.node, text: paragraph.match.text })))
    if (result.bytes !== source) writeDurableBytes(docxPath, result.bytes)
    return result.pointers
  } finally {
    opened.session.close()
  }
}

// Generation/export metadata only: unlike the amendment reader above, this can
// point at accepted text in a native comparison without accepting its revisions.
// A comparison may reuse the clean body's ids only when its complete accepted
// paragraph sequence matches. Fresh generation otherwise receives fresh ids.
export function withParagraphPointers(source: Uint8Array, reference?: Uint8Array) {
  const parts = unzipSync(source)
  const doc = pointerDocument(parts)
  const paragraphs = pointerParagraphs(doc)
  const referenceParagraphs = reference ? pointerParagraphs(pointerDocument(unzipSync(reference))) : []
  const preferred = paragraphs.length === referenceParagraphs.length && paragraphs.every((row, index) => row.text === referenceParagraphs[index]!.text)
    ? referenceParagraphs.map(row => paragraphPointers(row.node).find(pointer => PID.test(pointer.pid))?.pid)
    : []
  return addParagraphPointers(source, parts, doc, paragraphs, preferred).bytes
}

export async function compareDocxWithPointers(original: Uint8Array, modified: Uint8Array, options: Parameters<Awaited<ReturnType<typeof docxodus>>["compareDocuments"]>[2]) {
  const current = withParagraphPointers(modified)
  return withParagraphPointers(await (await docxodus()).compareDocuments(original, current, options), current)
}

function pointerDocument(parts: Record<string, Uint8Array>) {
  if (!parts[BODY]) throw new Error("invalid_docx_package")
  const fail = (message: string) => { throw new Error(`invalid_docx_xml:${message}`) }
  return new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } })
    .parseFromString(strFromU8(parts[BODY]!), "application/xml")
}

function pointerParagraphs(doc: Document) {
  return Array.from(doc.getElementsByTagNameNS(W, "p")).filter(node => !deletedRevisionAncestor(node))
    .map(node => ({ node, text: pointerText(node, node) }))
    .filter(row => row.text.trim())
}

function deletedRevisionAncestor(node: Node) {
  for (let parent = node.parentNode; parent; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && ["del", "moveFrom"].includes((parent as Element).localName)) return true
  return false
}

function pointerText(node: Node, paragraph: Element): string {
  if (node.nodeType !== 1) return ""
  const element = node as Element
  if (element.namespaceURI === W) {
    if (["del", "moveFrom"].includes(element.localName) || element.localName === "p" && element !== paragraph) return ""
    if (element.localName === "t") return element.textContent ?? ""
    if (element.localName === "tab") return "\t"
    if (["br", "cr"].includes(element.localName)) return "\n"
  }
  return Array.from(node.childNodes).map(child => pointerText(child, paragraph)).join("")
}

function addParagraphPointers(source: Uint8Array, parts: Record<string, Uint8Array>, doc: Document, paragraphs: { node: Element; text: string }[], preferred: (string | undefined)[] = []) {
  // Bookmark numeric ids are package-wide; paragraph pids are artifact-scoped.
  const ids = [
    ...Array.from(doc.getElementsByTagNameNS(W, "bookmarkStart")).map(mark => Number(mark.getAttributeNS(W, "id")) || 0),
    ...Object.entries(parts).filter(([name]) => name !== BODY && name.endsWith(".xml"))
      .flatMap(([, bytes]) => [...strFromU8(bytes).matchAll(/<w:bookmarkStart\b[^>]*?\bw:id="(\d+)"/g)].map(match => Number(match[1]))),
  ]
  const state = { next: Math.max(0, ...ids) + 1, changed: false }
  const seen = new Set<string>()
  const pointers = Object.fromEntries(paragraphs.map((paragraph, index) => {
    const wanted = preferred[index] && !seen.has(preferred[index]!) ? preferred[index] : undefined
    const existing = paragraphPointers(paragraph.node)
    const kept = existing.find(row => PID.test(row.pid) && !seen.has(row.pid) && (!wanted || row.pid === wanted))
    for (const row of existing.filter(row => row !== kept)) {
      removeBookmark(doc, row.mark)
      state.changed = true
    }
    const pid = kept?.pid ?? insertBookmark(doc, paragraph.node, String(state.next++), wanted)
    if (!kept) state.changed = true
    seen.add(pid)
    return [pid, paragraph.text]
  }))
  return { pointers, bytes: state.changed ? zipSync({ ...parts, [BODY]: strToU8(new XMLSerializer().serializeToString(doc)) }, { mtime: new Date(1980, 0, 1) }) : source }
}

function insertBookmark(doc: Document, paragraph: Element, id: string, preferred?: string) {
  const pid = preferred ?? `p${randomBytes(6).toString("hex")}`
  const start = doc.createElementNS(W, "w:bookmarkStart")
  start.setAttributeNS(W, "w:id", id)
  start.setAttributeNS(W, "w:name", paragraphBookmarkName(pid))
  const end = doc.createElementNS(W, "w:bookmarkEnd")
  end.setAttributeNS(W, "w:id", id)
  const properties = Array.from(paragraph.childNodes).find(node => node.nodeType === 1
    && (node as Element).namespaceURI === W && (node as Element).localName === "pPr")
  paragraph.insertBefore(start, properties ? properties.nextSibling : paragraph.firstChild)
  paragraph.appendChild(end)
  return pid
}

// A copied paragraph can repeat the bookmark id, so prefer the end inside the
// same paragraph and never remove an end another start still uses.
function removeBookmark(doc: Document, start: Element) {
  const id = start.getAttributeNS(W, "id")
  const paragraph = owningParagraph(start)
  start.parentNode?.removeChild(start)
  const ends = Array.from(doc.getElementsByTagNameNS(W, "bookmarkEnd")).filter(end => end.getAttributeNS(W, "id") === id)
  const local = ends.filter(end => owningParagraph(end) === paragraph)
  const shared = Array.from(doc.getElementsByTagNameNS(W, "bookmarkStart")).some(mark => mark.getAttributeNS(W, "id") === id)
  const removed = local.length ? local.slice(0, 1) : shared ? [] : ends
  removed.forEach(end => end.parentNode?.removeChild(end))
}

export function owningParagraph(node: Node) {
  for (let parent = node.parentNode; parent; parent = parent.parentNode)
    if (parent.nodeType === 1 && (parent as Element).namespaceURI === W && (parent as Element).localName === "p") return parent
  return null
}
