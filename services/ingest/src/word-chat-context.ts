import { createHash } from "node:crypto"
import { Database } from "bun:sqlite"
import path from "node:path"
import { readFileSync } from "node:fs"
import type { WordChatAttachment, WordChatContext, WordChatMaterial, WordChatScope, WordChatSource } from "../../../dochaus/lib/word-chat-contract"
import { readDeskSource } from "./word-source-desk"
import { readVersion, readVersionBytes } from "./document-versions"
import { extractDocumentText } from "./ingest"
import { listPlaybooks } from "./matter"
import { wordChatParagraphBreaks, wordChatListCaptureCurrent, wordChatTableCaptureCurrent, wordChatTargetCurrent } from "../../../dochaus/lib/word-chat-amendment-contract"
import type { WordGrant } from "./word-grants"

export class WordChatContextError extends Error {}
export type WordChatAuth = { matter: { id: string; title: string; dir: string }; grant: WordGrant; target: { name: string; docPath: string } }
export type WordChatUpload = WordChatMaterial & { text: string }
export const wordChatHash = (value: unknown) => createHash("sha256").update(typeof value === "string" ? value : JSON.stringify(value)).digest("hex")
const words = (text: string) => [...new Set(text.toLowerCase().match(/[\p{L}\p{N}]{3,}/gu) ?? [])].slice(0, 24)

/** Select literal canonical spans, coalescing only actual adjacency. Display
 * separators are never source text or quotation authority. */
export function wordChatExcerpts(text: string, query: string, limit = 14_000): { quote: string; passages: string[]; complete: boolean } {
  if (text.length <= limit) return { quote: text, passages: [text], complete: true }
  const terms = words(query)
  const blocks: { text: string; start: number; end: number; score: number }[] = []
  let offset = 0
  for (const line of text.split(/(?<=\n)/)) {
    for (let at = 0; at < line.length; at += 1600) {
      const block = line.slice(at, at + 1600)
      blocks.push({ text: block, start: offset + at, end: offset + at + block.length,
        score: terms.reduce((score, word) => score + (block.toLowerCase().includes(word) ? 1 : 0), 0) })
    }
    offset += line.length
  }
  const chosen: typeof blocks = []
  const spans = (rows: typeof blocks) => {
    const merged: { start: number; end: number }[] = []
    for (const block of rows.toSorted((a, b) => a.start - b.start)) {
      const previous = merged.at(-1)
      if (previous?.end === block.start) previous.end = block.end
      else merged.push({ start: block.start, end: block.end })
    }
    return merged.map(span => text.slice(span.start, span.end))
  }
  const separator = "\n[…excerpt gap…]\n"
  let passages: string[] = []
  for (const block of blocks.toSorted((a, b) => b.score - a.score || a.start - b.start)) {
    const next = spans([...chosen, block])
    if (next.join(separator).length > limit) continue
    chosen.push(block); passages = next
    if (passages.reduce((size, passage) => size + passage.length, 0) >= limit - 1600) break
  }
  return { quote: passages.join(separator), passages, complete: false }
}

/** Verification passages stay in the existing source snapshot, not in model
 * context a second time. Complete refers only to that admitted extraction. */
export const WORD_CHAT_SOURCE_COVERAGE = "Each source's complete flag means all text of its admitted capture or extraction is supplied, never complete document semantics or a complete review. False means selected excerpts; absent means completeness is unknown. Selected or unknown evidence cannot support whole-document absence or completeness claims. Display excerpts can have presentation separators between omitted passages; never stitch them into a source quotation. If a needed source or qualification is absent, disclose the gap."
export const wordChatPromptSources = (sources: WordChatSource[]) => sources.map(({ passages: _passages, ...source }) => source)

/** A display quote may contain presentation joins. Require one actual literal
 * passage. Older Word/upload snapshots can use their admitted canonical capture;
 * older saved/playbook excerpts without a snapshot are explicitly unavailable. */
export function wordChatQuoteVerified(source: WordChatSource, quote: string, context: WordChatContext, uploads: WordChatUpload[] = []) {
  if (!source.quote.includes(quote)) return false
  if (source.passages !== undefined) return typeof source.complete === "boolean" && Array.isArray(source.passages) &&
    source.passages.every(passage => typeof passage === "string") && source.passages.some(passage => passage.includes(quote))
  if (source.provenance === "word-buffer") {
    if (source.id === "word") return context.text.includes(quote)
    if (source.id === "word-selection") return Boolean(context.selection?.text.includes(quote))
  }
  if (source.provenance === "request-upload") return uploads.some(upload => source.id === `upload:${upload.id}` &&
    upload.hash === source.contentHash && typeof upload.text === "string" && wordChatHash(upload.text) === upload.hash && upload.text.includes(quote))
  return false
}

/** Suggestions historically normalize paragraph separators for native targets.
 * Preserve that spelling compatibility inside each canonical passage only;
 * exact answer quotations continue to use wordChatQuoteVerified unchanged. */
export function wordChatSuggestionVerified(source: WordChatSource, passage: string, context: WordChatContext) {
  if (source.provenance !== "word-buffer") return false
  const normalized = wordChatParagraphBreaks(passage)
  if (!wordChatParagraphBreaks(source.quote).includes(normalized)) return false
  if (source.passages !== undefined) return typeof source.complete === "boolean" && Array.isArray(source.passages) &&
    source.passages.every(value => typeof value === "string") && source.passages.some(value => wordChatParagraphBreaks(value).includes(normalized))
  const captured = source.id === "word" ? context.text : source.id === "word-selection" ? context.selection?.text : undefined
  return captured !== undefined && wordChatParagraphBreaks(captured).includes(normalized)
}

export function validateWordChatContext(value: unknown): value is WordChatContext {
  if (!value || typeof value !== "object") return false
  const input = value as WordChatContext
  return typeof input.text === "string" && input.text.length > 0 && input.text.length <= 512_000 &&
    typeof input.saved === "boolean" && typeof input.capturedAt === "string" && typeof input.truncated === "boolean" && !input.truncated &&
    Array.isArray(input.paragraphs) && input.paragraphs.length <= 20_000 && input.paragraphs.every(paragraph =>
      paragraph && typeof paragraph.id === "string" && typeof paragraph.text === "string" && paragraph.text.length <= 512_000) &&
    (!input.selection || wordChatTargetCurrent(input, input.selection) && input.text.includes(input.selection.text)) &&
    (input.table === undefined || wordChatTableCaptureCurrent(input, input.table)) &&
    (input.list === undefined || wordChatListCaptureCurrent(input, input.list))
}

export function wordChatMaterials(auth: WordChatAuth, uploads: WordChatUpload[] = []): WordChatMaterial[] {
  const db = new Database(path.join(auth.matter.dir, ".dochaus", "legal.db"), { readonly: true })
  let ids: { id: number }[]
  try { ids = db.query("SELECT id FROM documents ORDER BY name LIMIT 200").all() as { id: number }[] } finally { db.close() }
  const sources = ids.flatMap(row => {
    const value = readDeskSource(auth, String(row.id))
    return value ? [{ kind: "source" as const, id: value.id, hash: value.hash, name: value.name, provenance: "saved-version" as const }] : []
  })
  return [...sources, ...listPlaybooks().slice(0, 100).map(book => ({ kind: "playbook" as const, id: book.name, hash: wordChatHash(book.content), name: book.name, description: book.description, provenance: "instructions" as const })),
    ...uploads.map(({ text: _text, ...material }) => material)]
}

export async function wordChatEvidence(auth: WordChatAuth, context: WordChatContext, scope: WordChatScope, attachments: WordChatAttachment[], uploads: WordChatUpload[], query: string) {
  const available = wordChatMaterials(auth, uploads)
  const selected = attachments.map(item => {
    const material = available.find(material => material.id === item.id && material.kind === item.kind && material.hash === item.hash)
    if (!material) throw new WordChatContextError("An attached source changed or is unavailable. Remove it or attach its current version.")
    return material
  })
  const sources: WordChatSource[] = [{ id: "word", name: auth.target.name, ...wordChatExcerpts(context.text, `${query}\n${context.selection?.text ?? ""}`), location: "Current Word buffer", provenance: "word-buffer" }]
  const savedText = new Map<string, string>()
  async function saved(id: string) {
    if (savedText.has(id)) return savedText.get(id)!
    const source = readDeskSource(auth, id)
    const kept = source?.version ? readVersion(auth.matter.dir, source.version.id) : undefined
    const bytes = kept ? readVersionBytes(auth.matter.dir, kept.version) : undefined
    if (!source || !bytes || bytes.length > 16 * 1024 * 1024) throw new WordChatContextError("An attached source is unavailable or too large. Open it in doc.haus.")
    const text = await extractDocumentText(source.name, bytes)
    savedText.set(id, text)
    return text
  }
  for (const material of selected) {
    const text = material.kind === "playbook" ? listPlaybooks().find(book => book.name === material.id)?.content : material.kind === "upload" ? uploads.find(upload => upload.id === material.id)?.text : await saved(material.id)
    if (typeof text !== "string") throw new WordChatContextError("An attachment is unavailable. Attach it again.")
    sources.push({ id: `${material.kind}:${material.id}`, contentHash: material.hash, name: material.name, ...wordChatExcerpts(text, query, 8000), location: material.kind === "playbook" ? "Attached guidance; not a client decision" : "Deliberately attached material", provenance: material.provenance })
  }
  let versions = ""
  if (scope === "matter") {
    const terms = words(query).slice(0, 12)
    const db = new Database(path.join(auth.matter.dir, ".dochaus", "legal.db"), { readonly: true })
    let matches: { document_id: number; text: string; section: string }[] = []
    try {
      if (terms.length) matches = db.query(`SELECT c.document_id,c.text,c.section FROM chunks c JOIN chunks_fts f ON f.rowid=c.id WHERE chunks_fts MATCH ? ORDER BY bm25(chunks_fts) LIMIT 12`).all(terms.map(word => `"${word}"`).join(" OR ")) as typeof matches
    } finally { db.close() }
    for (const match of matches) {
      const source = readDeskSource(auth, String(match.document_id))
      if (!source || source.name === auth.target.name || sources.some(prior => prior.id === `source:${source.id}`)) continue
      const text = await saved(source.id)
      if (!text.includes(match.text)) continue // stale index is never evidence
      sources.push({ id: `source:${source.id}`, contentHash: source.hash, name: source.name, ...wordChatExcerpts(text, query, 4000), location: match.section || "Saved matter source", provenance: "saved-version" })
      if (sources.length >= 10) break
    }
    const materials = available.filter(material => material.kind === "source")
    versions = wordChatHash(materials.map(material => [material.id, material.hash]))
    let defaultPlaybook: string | undefined
    try { defaultPlaybook = (JSON.parse(readFileSync(path.join(auth.matter.dir, "matter.json"), "utf8")) as { playbook?: string }).playbook } catch {}
    const book = defaultPlaybook ? listPlaybooks().find(book => book.name === defaultPlaybook) : undefined
    if (book && !selected.some(material => material.kind === "playbook" && material.id === book.name)) {
      sources.push({ id: `playbook:${book.name}`, contentHash: wordChatHash(book.content), name: book.name, ...wordChatExcerpts(book.content, query, 6000), location: "Matter guidance; not a client decision", provenance: "instructions" })
      versions += wordChatHash(book.content)
    }
  }
  const sourceKey = wordChatHash({ scope, attachments, versions })
  // Visible history survives edits, but an earlier buffer/selection is not
  // current evidence. Exclude capture timestamps and random selection IDs so
  // unchanged context can still carry ordinary follow-up discussion.
  const selection = context.selection ? { text: context.selection.text, paragraphs: context.selection.paragraphs.map(paragraph => paragraph.id), start: context.selection.startOffset, end: context.selection.endOffset } : null
  const table = context.table ? { tableIndex: context.table.tableIndex, rows: context.table.rows.map(row => row.map(cell => [cell.id, cell.text])) } : undefined
  return { sources, attachments: selected, sourceKey, evidenceKey: wordChatHash({ sourceKey, text: context.text, selection, table, list:context.list }) }
}

/** Three small independent passage jobs; a large document reports sampling. */
export function wordChatSuggestionBatches(context: WordChatContext) {
  const paragraphs = context.paragraphs
  const groups: { id: string; text: string }[][] = []
  let group: typeof paragraphs = [], size = 0
  for (const paragraph of paragraphs) {
    const pieces = paragraph.text.length > 8000 ? Array.from({ length: Math.ceil(paragraph.text.length / 8000) }, (_, index) => ({ ...paragraph, text: paragraph.text.slice(index * 8000, (index + 1) * 8000) })) : [paragraph]
    for (const piece of pieces) {
      if (group.length && size + piece.text.length > 8000) { groups.push(group); group = []; size = 0 }
      group.push(piece); size += piece.text.length
    }
  }
  if (group.length) groups.push(group)
  // One flow considers the document and current selection together. In a large
  // document, prioritize its selected passage within the same bounded budget.
  const selectedIDs = new Set(context.selection?.paragraphs.map(paragraph => paragraph.id))
  const selected = groups.findIndex(group => group.some(paragraph => selectedIDs.has(paragraph.id)))
  const indexes = groups.length <= 3 ? groups.map((_, index) => index)
    : [...new Set([...(selected >= 0 ? [selected] : []), 0, groups.length - 1, Math.floor((groups.length - 1) / 2)])].slice(0, 3).sort((a, b) => a - b)
  return { batches: indexes.map(index => groups[index]!), reviewed: indexes.length, total: groups.length }
}

/** A late response cannot revive evidence removed/replaced since preparation. */
export function wordChatSourcesCurrent(auth: WordChatAuth, sources: WordChatSource[]) {
  return sources.every(source => {
    if (source.id.startsWith("source:")) return readDeskSource(auth, source.id.slice(7))?.hash === source.contentHash
    if (source.id.startsWith("playbook:")) { const book = listPlaybooks().find(book => book.name === source.id.slice(9)); return Boolean(book && wordChatHash(book.content) === source.contentHash) }
    return true
  })
}
