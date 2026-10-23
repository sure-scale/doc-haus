import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { getExtractionQuality } from "../../services/ingest/src/db"
import { formatCitations, type DocumentCitation } from "../lib/citations"
import { findQuote, liveText } from "../lib/extract"

// doc.haus citation-anchoring tool. Before an agent puts a quotation from a
// matter document into its answer, it calls this tool with the verbatim text;
// the tool re-extracts the live document and confirms the quote actually appears
// in it, returning a verified citation the model can present. The legal plugin
// re-checks the same span on the way out, so a quote that does not anchor here is
// a quote the lawyer must never see.

export default tool({
  description:
    "Anchor a quotation against a matter document before you quote it in an answer. Call this for EVERY passage you intend to quote from a matter document. The quote must be copied verbatim from the document: no paraphrase, no added or removed ellipses, no edits. If this tool cannot verify the quote, do not present that quotation to the user — re-read the document and quote only text the tool confirmed.",
  args: {
    docPath: tool.schema
      .string()
      .describe(
        "Path of the matter document, exactly as returned by search-document or read-document. A bare document name is resolved against the matter root.",
      ),
    documentName: tool.schema.string().describe("Human-readable name of the document."),
    quote: tool.schema
      .string()
      .min(10)
      .max(600)
      .describe("The verbatim text copied from the document, with no paraphrase or added ellipses."),
    reason: tool.schema.string().describe("Why this passage supports the claim you are making."),
    confidence: tool.schema
      .number()
      .int()
      .min(1)
      .max(5)
      .describe(
        "How strongly the passage supports the claim: 5 = verbatim quote that directly and unambiguously supports it; 3 = supports it but requires interpretation or context; 1 = tangential or weak support.",
      ),
    context: tool.schema
      .string()
      .optional()
      .describe("Optional: the sentence(s) immediately around the quote, for the reader."),
  },
  async execute(args, ctx) {
    // Models routinely pass the document's display name instead of the absolute
    // path search-document returned; resolve it against the matter root rather
    // than failing the citation and forcing a browse-and-retry loop.
    const docPath = path.isAbsolute(args.docPath) ? args.docPath : path.resolve(ctx.directory, args.docPath)
    if (!existsSync(docPath)) {
      return `${args.documentName} is not in this matter (no file at ${docPath}). Do not use this quotation. Confirm the document with search-document or read-document before quoting it.`
    }

    // Ingest's extraction quality gate: a document with no readable text was
    // never indexed and nothing in it can be verified; a degraded one is
    // verified with its warning.
    const quality = storedQuality(ctx.directory, docPath)
    if (quality?.status === "no_readable_text") {
      return `${args.documentName} has no readable text: ingest could not extract any, so it is not indexed and no quotation from it can be verified. Do not use this quotation; tell the user the document could not be read.`
    }

    const match = findQuote(await liveText(docPath), args.quote)
    if (!match) {
      return `The quoted text does not appear in ${args.documentName}. Do not use this quotation — re-read the document with read-document or search-document and quote only text it returns verbatim.`
    }

    const citation: DocumentCitation = {
      documentName: args.documentName,
      docPath,
      // The indexed section holding the passage, so the citation reads and
      // groups like a search hit; "cited passage" when the matter has no index
      // entry for it.
      section: indexedSection(ctx.directory, docPath, match.start) ?? "cited passage",
      // Store the raw matched span, not the model's input, so the plugin's exact
      // slice check against the live document passes on the way out.
      excerpt: match.excerpt,
      charStart: match.start,
      charEnd: match.end,
      reason: args.reason,
      confidence: args.confidence,
      context: args.context,
      verified: true,
      ...(quality?.status === "degraded" ? { degradedExtraction: quality.signals.map((signal) => signal.detail) } : {}),
    }

    return {
      title: `Verified quote in ${args.documentName}`,
      output: `${formatCitations([citation])}\n\n` + (citation.degradedExtraction
        ? "Verified verbatim against the live document, but ingest found this document's extracted text may be incomplete or garbled. Check the original before relying on the quotation, and say so when you present it."
        : "Verified verbatim against the live document — safe to quote."),
      metadata: { citations: [citation] },
    }
  },
})

// The section label of the indexed chunk that holds `offset` in the document's
// text. Ingest's chunk offsets and the live text the quote was found in share
// one projection (lib/extract liveText). Undefined without an index or chunk.
function indexedSection(directory: string, docPath: string, offset: number) {
  const dbPath = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(dbPath)) return
  const db = new Database(dbPath, { readonly: true })
  try {
    const row = db
      .query("SELECT section FROM chunks WHERE doc_path = ? AND char_start <= ? AND char_end > ? ORDER BY char_start DESC LIMIT 1")
      .get(docPath, offset, offset) as { section: string } | null
    return row?.section || undefined
  } finally {
    db.close()
  }
}

// The document's stored extraction quality, read from the matter's index
// without writing to it. Undefined when the matter has no index, the document
// predates the gate, or ingest has not yet added the column.
function storedQuality(directory: string, docPath: string) {
  const dbPath = path.join(directory, ".dochaus", "legal.db")
  if (!existsSync(dbPath)) return
  const db = new Database(dbPath, { readonly: true })
  try {
    const columns = db.query("SELECT name FROM pragma_table_info('documents')").all() as { name: string }[]
    if (!columns.some((column) => column.name === "extraction_quality")) return
    return getExtractionQuality(db, docPath)
  } finally {
    db.close()
  }
}
