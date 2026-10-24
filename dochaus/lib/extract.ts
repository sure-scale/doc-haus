import { statSync } from "node:fs"
import { writeListLabels } from "./docx-list-labels"
import { markTrackedChanges } from "./docx-tracked-text"
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
  const text = normalizeExtractedText(await extract(docPath, buffer)) + "\n"
  if (!textCache.has(docPath) && textCache.size >= CACHE_LIMIT) {
    const oldest = textCache.keys().next().value
    if (oldest !== undefined) textCache.delete(oldest)
  }
  textCache.set(docPath, { mtimeMs, text })
  return text
}

// Locate a quotation in the live text. Exact match first; if that fails, fall
// back to a whitespace-tolerant search so a quote that only differs in run-length
// of spaces/newlines (mammoth and pdf.js normalize whitespace differently than a
// human copying text) still anchors. The returned `excerpt` is always the RAW
// matched span from the live text, never the caller's input — downstream span
// checks compare against the live document, so the stored excerpt must be what
// the document actually contains. When `near` is given (the plugin re-anchoring a
// span an edit shifted), the occurrence nearest that offset wins, since it is
// almost certainly the same passage; otherwise the first match wins.
export function findQuote(text: string, quote: string, near?: number) {
  const matches = [...text.matchAll(quoteRegExp(quote))]
  const first = matches[0]
  if (!first) return undefined
  const best =
    near === undefined
      ? first
      : matches.slice(1).reduce(
          (a, b) => (Math.abs(b.index - near) < Math.abs(a.index - near) ? b : a),
          first,
        )
  return { start: best.index, end: best.index + best[0].length, excerpt: best[0] }
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
async function extract(docPath: string, buffer: Buffer) {
  const name = docPath.toLowerCase()
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
  if (name.endsWith(".xlsx")) return xlsxToText(buffer)
  if (name.endsWith(".zip")) {
    // The one zip reader ingest also uses (zip.ts): the listing of its files.
    const { zipToText } = await import("./zip")
    return zipToText(buffer)
  }
  const { default: mammoth } = await import("mammoth")
  return (await mammoth.extractRawText({ buffer: await writeListLabels(markTrackedChanges(buffer)) })).value
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
