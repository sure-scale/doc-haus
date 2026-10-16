import { tool } from "./legacy-definition"
import {
  MAX_RENDERED_HITS,
  SHORT_QUOTE,
  normalize,
  verifyQuote,
} from "../lib/quote-verify"

// doc.haus quote-verification tool. Before the agent presents any text as a
// quotation from a matter document, this tool checks the exact wording against
// the per-matter legal.db that `services/ingest` populates
// (`<matter>/.dochaus/legal.db`), so verification never crosses into another
// matter's privileged material. It answers only from what is stored: a hit
// returns the stored verbatim chunk text and its location so the agent can
// re-quote exactly from storage, and a miss is reported as NOT FOUND — never a
// near match or best guess. This tool is read-only; all writes happen in the
// ingest service. The matching logic (single-chunk and cross-boundary) lives
// in lib/quote-verify.verifyQuote, shared with lib/verification-loop's
// citation-anchors gate.

export default tool({
  description:
    "Verify that a quote appears verbatim in the matter's indexed documents before presenting it. Returns VERIFIED with the stored text and its location (document, section, char offsets) when the exact wording exists, or NOT FOUND when it does not — never a near match. Use this before quoting any document text in an answer or a draft.",
  args: {
    quote: tool.schema.string().describe("The exact text you intend to quote, verbatim"),
    document: tool.schema
      .string()
      .optional()
      .describe("Restrict verification to a single document by its exact name. Omit to check the whole matter."),
  },
  async execute(args, ctx) {
    const result = verifyQuote({ directory: ctx.directory, quote: args.quote, document: args.document })

    if (result.reason === "no-db") return "No documents have been indexed for this matter yet."
    if (result.reason === "empty-quote")
      return "The quote is empty after normalization. Pass the exact text you intend to quote."

    const scope = args.document ? `"${args.document}"` : "the matter"
    const needle = normalize(args.quote)
    const shortNote =
      needle.length < SHORT_QUOTE
        ? `\n\nNote: this quote is under ${SHORT_QUOTE} characters — very short quotes match trivially, so a section-level citation is more useful than quote verification.`
        : ""

    if (result.reason === "no-docs") {
      if (args.document)
        return `No document named "${args.document}" is indexed in this matter. Check the exact name, or omit the document argument to verify against the whole matter.`
      return "No documents have been indexed for this matter yet."
    }

    if (!result.found)
      return (
        `NOT FOUND — this exact text does not appear in ${scope}. Do not present it as a quote. Use search-document to locate the correct wording.` +
        shortNote
      )

    const overflow =
      result.hits.length > MAX_RENDERED_HITS
        ? `\n\n---\n\n...and ${result.hits.length - MAX_RENDERED_HITS} more document(s) contain this text — narrow with the document argument.`
        : ""
    return (
      `VERIFIED — the quote appears verbatim in ${result.hits.length} document(s). Re-quote exactly from the stored text below, never from memory.\n\n` +
      result.hits
        .slice(0, MAX_RENDERED_HITS)
        .map(
          (h) =>
            `${h.doc} — ${h.section} (chars ${h.charStart}-${h.charEnd})\nStored text of the containing region:\n${h.verbatim}`,
        )
        .join("\n\n---\n\n") +
      overflow +
      shortNote
    )
  },
})
