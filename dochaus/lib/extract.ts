import { statSync } from "node:fs"
import { writeListLabels } from "./docx-list-labels"
import { markTrackedChanges } from "./docx-tracked-text"
import { normalizeWithMap, sourceReadings } from "./quote-verify"
import { normalizeExtractedText } from "./untrusted"

// Live-document extraction shared by the citation plugin (which re-checks every
// citation's span against the current file) and the `cite` tool (which anchors a
// quotation the model intends to use). Both need the exact same text the ingest
// service indexed, so the logic lives here once.

const CACHE_LIMIT = 64
// Extraction is per-document, not per-citation: cache by path, invalidate on
// mtime so an accepted redline or re-upload is picked up immediately.
const textCache = new Map<string, { mtimeMs: number; text: string }>()

export async function liveText(docPath: string) {
  const mtimeMs = statSync(docPath).mtimeMs
  const hit = textCache.get(docPath)
  if (hit && hit.mtimeMs === mtimeMs) return hit.text
  const buffer = Buffer.from(await Bun.file(docPath).arrayBuffer())
  // Ingest computes offsets by accumulating `line + "\n"` over the normalized
  // extracted text, which equals the normalized raw extraction plus exactly one
  // trailing newline — mirror both steps here or every chunk fails verification.
  let unlabelled = false
  const text = normalizeExtractedText(await extract(docPath, buffer, () => { unlabelled = true })) + "\n"
  // Text read without the list numbers Docxodus failed to give is not kept:
  // ingest may have indexed it with them, and the next read tries again.
  if (unlabelled) return text
  if (!textCache.has(docPath) && textCache.size >= CACHE_LIMIT) {
    const oldest = textCache.keys().next().value
    if (oldest !== undefined) textCache.delete(oldest)
  }
  textCache.set(docPath, { mtimeMs, text })
  return text
}

// Locate a quotation in the live text of `document` (the source file's name or
// path, which says how its text was written). Exact match first; if that fails,
// fall back to a whitespace-tolerant search so a quote that only differs in
// run-length of spaces/newlines (mammoth and pdf.js normalize whitespace
// differently than a human copying text) still anchors; then to each of the
// document's readings in verifyQuote's order (sourceReadings: curly quotes
// folded, then without backslash escapes, then without the format's own markup:
// Markdown in a .md file, an email or a PDF, reply marks in an email, page-break
// rules in a PDF, CSV quoting in a spreadsheet), so `cite` anchors the quotes
// verifyQuote confirms in the same document. Without a document name only the
// readings every format shares are tried. The returned `excerpt` is always the
// RAW matched span from the live text, never the caller's input — downstream
// span checks compare against the live document, so the stored excerpt must be
// what the document actually contains. When `near` is given (the plugin
// re-anchoring a span an edit shifted), the occurrence nearest that offset wins
// among the matches of the first search that finds one, since it is almost
// certainly the same passage; otherwise the first match wins.
export function findQuote(text: string, quote: string, document?: string, near?: number) {
  const matches = findQuotes(text, quote, document)
  const first = matches[0]
  if (!first) return undefined
  if (near === undefined) return first
  return matches.slice(1).reduce((a, b) => (Math.abs(b.start - near) < Math.abs(a.start - near) ? b : a), first)
}

// Every match of the first search (findQuote's order) that finds one, as raw
// spans. Matches from a later reading never join an earlier one's, so a quote
// found once as written stays unique when an escaped copy of it is elsewhere;
// two copies read the same way are two matches. A normalized match maps back
// through normalizeWithMap's offsets: its first character's raw start to its
// last's raw end (past a whole decoded entity).
export function findQuotes(text: string, quote: string, document?: string) {
  const span = (start: number, end: number) => ({ start, end, excerpt: text.slice(start, end) })
  const exact = [...text.matchAll(quoteRegExp(quote))].map((m) => span(m.index, m.index + m[0].length))
  if (exact.length) return exact
  const needle = normalizeWithMap(quote).text
  if (!needle) return []
  for (const reading of sourceReadings(document)) {
    const { text: normalized, map, end } = normalizeWithMap(text, reading)
    const found: ReturnType<typeof span>[] = []
    for (let at = normalized.indexOf(needle); at !== -1; at = normalized.indexOf(needle, at + needle.length))
      found.push(span(map[at]!, end[at + needle.length - 1]!))
    if (found.length) return found
  }
  return []
}

// Match the quote exactly where possible, but tolerate differing whitespace runs.
// Escape every regex metacharacter, then let any whitespace run stand in for any
// other so a copied quote anchors despite extractor-specific line breaks.
function quoteRegExp(quote: string) {
  return new RegExp(quote.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+"), "g")
}

// Must stay in lockstep with extractDocumentText in services/ingest/src/ingest.ts:
// same extractors, same options, or offsets stop lining up. Every branch ingest
// accepts must appear here too — the citation verifier (verifyCitation via
// liveText) re-extracts whatever document a search-document citation points at,
// so a missing branch (e.g. .eml routed to mammoth) throws "Can't find end of
// central directory" inside the tool.execute.after hook and the whole
// search-document call surfaces as failed to the model.
async function extract(docPath: string, buffer: Buffer, unlabelled: () => void) {
  const name = docPath.toLowerCase()
  if (name.endsWith(".md") || name.endsWith(".markdown"))
    return new TextDecoder("utf-8", { fatal: true }).decode(buffer)
  if (name.endsWith(".pdf")) {
    // Citation verification must use the exact extractor chain used at ingest.
    // Ingest now prefers LiteParse/MarkItDown before unpdf; using unpdf alone
    // here made every indexed LiteParse span look stale in the unchanged PDF.
    const { pdfToText } = await import("../../services/ingest/src/pdf")
    return pdfToText(buffer)
  }
  if (name.endsWith(".eml")) {
    // The one email extractor ingest also uses (email.ts).
    const { emlToText } = await import("./email")
    return emlToText(buffer)
  }
  if (name.endsWith(".msg")) {
    // The one Outlook email extractor ingest also uses (msg.ts).
    const { msgToText } = await import("./msg")
    return msgToText(buffer)
  }
  if (name.endsWith(".xlsx")) return xlsxToText(buffer)
  if (name.endsWith(".zip")) {
    // The one zip reader ingest also uses (zip.ts): the listing of its files.
    const { zipToText } = await import("./zip")
    return zipToText(buffer)
  }
  const { default: mammoth } = await import("mammoth")
  return (await mammoth.extractRawText({ buffer: markTrackedChanges(await writeListLabels(buffer, unlabelled)) })).value
}

// Spreadsheet sources (.xlsx) — same tolerant reader as ingest, one labeled CSV
// block per sheet so tabular relationships survive as text the same way the
// indexer saw them.
async function xlsxToText(buffer: Buffer) {
  const { read, utils } = await import("xlsx")
  const workbook = read(buffer, { type: "buffer" })
  return workbook.SheetNames.map((name: string) => {
    const sheet = workbook.Sheets[name]
    if (!sheet) return ""
    const csv = utils.sheet_to_csv(sheet).trim()
    return csv ? `## ${name}\n${csv}` : ""
  })
    .filter(Boolean)
    .join("\n\n")
}
