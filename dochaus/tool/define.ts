import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"

// doc.haus defined-term lookup. Reads the deterministic `defined_terms` table
// that `services/ingest` extracts by regex (see services/ingest/src/structure.ts)
// from the per-matter legal.db. Every row is verbatim text plus offsets into the
// document's extracted text, so this tool can only quote what the documents
// actually say — a parse miss is an absent row and a "not found" answer, never a
// wrong fact. This tool is read-only; all writes happen in the ingest service.

const USAGE_LIMIT = 5
// Cap on the contains-fallback only; exact matches stay uncapped because a
// lawyer needs every definition of one term.
const NEAR_MATCH_LIMIT = 10
// Half-window of verbatim context shown around a usage occurrence, bounded to
// the line containing it so every snippet stays single-line and verbatim.
const SNIPPET_RADIUS = 100

type TermRow = { doc_name: string; term: string; definition: string; char_start: number; char_end: number }
type UsageRow = { doc_name: string; section: string; text: string; char_start: number; char_end: number }

export default tool({
  description:
    "Look up a defined term in the current matter's documents. Returns the term's verbatim definition(s) with document name and character offsets, plus up to five usage sites. Use this instead of searching when a question hinges on what a capitalized contract term means. Only verbatim indexed text is returned; if the term is not found, fall back to the search-document tool.",
  args: {
    term: tool.schema.string().describe('The defined term to look up, without quotes (e.g. Change of Control)'),
    document: tool.schema
      .string()
      .optional()
      .describe("Restrict the lookup to a single document by its exact name. Omit to look across the whole matter."),
  },
  async execute(args, ctx) {
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) {
      return "No documents have been indexed for this matter yet."
    }

    const db = new Database(dbPath, { readonly: true })
    const select = "SELECT doc_name, term, definition, char_start, char_end FROM defined_terms WHERE "
    const docScope = args.document ? " AND doc_name = ?" : ""
    const order = " ORDER BY doc_name, char_start"
    const scopeParams = args.document ? [args.document] : []
    // Escape LIKE wildcards so the bound term only ever matches literally —
    // otherwise a `%` term dumps the whole table and breaks the indexOf
    // invariant in the snippet logic below.
    const like = args.term.replaceAll(/[\\%_]/g, "\\$&")
    const exact = db
      .query(select + "term = ? COLLATE NOCASE" + docScope + order)
      .all(args.term, ...scopeParams) as TermRow[]
    // SQLite LIKE is case-insensitive over ASCII by default, so this is the
    // NOCASE contains fallback for when the exact form misses.
    const hits = exact.length
      ? exact
      : (db
          .query(select + "term LIKE '%' || ? || '%' ESCAPE '\\'" + docScope + order + " LIMIT ?")
          .all(like, ...scopeParams, NEAR_MATCH_LIMIT) as TermRow[])
    if (hits.length === 0) {
      db.close()
      const where = args.document ? ` in "${args.document}"` : " in this matter"
      return `No defined term matching "${args.term}"${where}. Try the search-document tool to find related language instead.`
    }

    // Usage sites: chunks whose body contains the bare term. Extracted text
    // quotes terms with straight or curly quotes interchangeably, so matching
    // the unquoted term is the reliable form.
    const usage = db
      .query(
        "SELECT doc_name, section, text, char_start, char_end FROM chunks WHERE text LIKE '%' || ? || '%' ESCAPE '\\'" +
          (args.document ? " AND doc_name = ?" : "") +
          " ORDER BY doc_name, char_start LIMIT ?",
      )
      .all(like, ...scopeParams, USAGE_LIMIT) as UsageRow[]
    db.close()

    const byDoc = Map.groupBy(hits, (hit) => hit.doc_name)
    const heading = exact.length
      ? `Definition(s) of "${hits[0].term}":`
      : `No exact match for "${args.term}" — these are near-matches (defined terms containing it, up to ${NEAR_MATCH_LIMIT}):`
    const definitions = [...byDoc.entries()]
      .map(
        ([doc, rows]) =>
          `## ${doc}\n` +
          rows.map((row) => `"${row.term}" [chars ${row.char_start}-${row.char_end}]:\n${row.definition}`).join("\n\n"),
      )
      .join("\n\n")
    // Only meaningful for exact matches — near-match rows are different terms,
    // so differing definitions are expected, not a conflict to warn about.
    const differNote =
      exact.length > 0 && byDoc.size > 1 && new Set(hits.map((row) => row.definition)).size > 1
        ? "\n\nNote: multiple documents define this term and the definitions are not identical — check which document governs before relying on any one of them."
        : ""
    const usageLines = usage.map((row) => {
      // The row matched LIKE on this same ASCII-case-folded term, so indexOf on
      // the lowercased text always finds an occurrence.
      const idx = row.text.toLowerCase().indexOf(args.term.toLowerCase())
      const lineStart = row.text.lastIndexOf("\n", idx) + 1
      const lineEndIndex = row.text.indexOf("\n", idx)
      const lineEnd = lineEndIndex === -1 ? row.text.length : lineEndIndex
      const start = Math.max(lineStart, idx - SNIPPET_RADIUS)
      const end = Math.min(lineEnd, idx + args.term.length + SNIPPET_RADIUS)
      const snippet =
        (start > lineStart ? "..." : "") + row.text.slice(start, end) + (end < lineEnd ? "..." : "")
      // Chunk-granularity offsets, unlike the term-precise offsets on
      // definition rows — label them so the two are not conflated.
      return `- ${row.doc_name} § ${row.section} [within chars ${row.char_start}-${row.char_end}]: ${snippet}`
    })
    const usageBlock = usageLines.length
      ? `\n\nUsage sites (up to ${USAGE_LIMIT}):\n` + usageLines.join("\n")
      : ""

    return `${heading}\n\n${definitions}${differNote}${usageBlock}`
  },
})
