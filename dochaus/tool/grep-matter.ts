import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"

// doc.haus exact-text search. Greps the per-matter legal.db that
// `services/ingest` populates: every occurrence of a literal string or JS
// regular expression across the matter's chunk text, reported as the verbatim
// line containing the match plus its location (document, section, absolute
// character offset into the document's extracted text).
//
// Everything returned is text already stored in the database — nothing is
// paraphrased or synthesized, so a hit is a fact and a miss is a miss. The
// database lives inside the matter directory (`<matter>/.dochaus/legal.db`),
// scoping results to the active matter. This tool is read-only; all writes
// happen in the ingest service.

const MATCH_CAP = 50
const LINE_CAP = 200

type ChunkRow = { doc_name: string; section: string; text: string; char_start: number }

// The verbatim line (text between the surrounding newlines) containing a
// match, windowed to <= LINE_CAP chars around the match when the line is long.
function lineAt(text: string, index: number) {
  const lineStart = text.lastIndexOf("\n", index) + 1
  const newline = text.indexOf("\n", index)
  const lineEnd = newline === -1 ? text.length : newline
  if (lineEnd - lineStart <= LINE_CAP) return text.slice(lineStart, lineEnd).trim()
  const start = Math.max(lineStart, index - 80)
  return text.slice(start, Math.min(lineEnd, start + LINE_CAP)).trim()
}

function literalMatches(text: string, needle: string) {
  const matches: number[] = []
  let index = text.indexOf(needle)
  while (index !== -1) {
    matches.push(index)
    index = text.indexOf(needle, index + needle.length)
  }
  return matches
}

export default tool({
  description:
    "Exact-text search (grep) across the current matter's documents. Finds every occurrence of a literal string or regular expression and returns the verbatim line containing each match with its document, section, and character offset. Use this for exact strings — section numbers, defined terms, party names, specific wording. For meaning-based questions use search-document instead.",
  args: {
    pattern: tool.schema
      .string()
      .min(1)
      .describe("The text to find — a literal string, or a JS regular expression when regex is true"),
    regex: tool.schema.boolean().optional().describe("Treat pattern as a JavaScript regular expression (default false)"),
    document: tool.schema
      .string()
      .optional()
      .describe("Restrict the search to a single document by its exact name. Omit to search the whole matter."),
    ignoreCase: tool.schema.boolean().optional().describe("Case-insensitive matching (default true)"),
  },
  async execute(args, ctx) {
    // An empty needle would never advance literalMatches' indexOf loop.
    if (!args.pattern) return "Pattern must not be empty."

    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) {
      return "No documents have been indexed for this matter yet."
    }

    const ignoreCase = args.ignoreCase ?? true
    // The pattern is untrusted model output and RegExp() is the one genuine
    // parse boundary here — surface the engine's message verbatim.
    let re: RegExp | null = null
    if (args.regex) {
      try {
        re = new RegExp(args.pattern, ignoreCase ? "gi" : "g")
      } catch (error) {
        // Bun's constructor message already carries this prefix; don't double it.
        const message = error instanceof Error ? error.message : String(error)
        return message.startsWith("Invalid regular expression") ? message : `Invalid regular expression: ${message}`
      }
    }

    const db = new Database(dbPath, { readonly: true })
    const sql =
      "SELECT doc_name, section, text, char_start FROM chunks" +
      (args.document ? " WHERE doc_name = ?" : "") +
      " ORDER BY doc_name, char_start"
    const chunks = (args.document ? db.query(sql).all(args.document) : db.query(sql).all()) as ChunkRow[]
    db.close()

    if (!chunks.length && args.document)
      return `No document named "${args.document}" is indexed in this matter. Omit the document argument to search the whole matter.`
    if (!chunks.length) return "No documents have been indexed for this matter yet."

    const needle = ignoreCase ? args.pattern.toLowerCase() : args.pattern
    const lines: string[] = []
    let scanned = 0
    let capped = false
    for (const chunk of chunks) {
      if (lines.length >= MATCH_CAP) {
        capped = true
        break
      }
      scanned++
      // matchAll copies the regex, so no lastIndex state leaks across chunks,
      // and it advances past zero-length matches on its own.
      const found = re
        ? [...chunk.text.matchAll(re)].map((m) => m.index)
        : literalMatches(ignoreCase ? chunk.text.toLowerCase() : chunk.text, needle)
      for (const index of found) {
        if (lines.length >= MATCH_CAP) {
          capped = true
          break
        }
        lines.push(`${chunk.doc_name} [${chunk.section}] @ ${chunk.char_start + index}: ${lineAt(chunk.text, index)}`)
      }
    }

    const shown = re ? `regex /${args.pattern}/${ignoreCase ? "i" : ""}` : `"${args.pattern}"`
    const where = args.document ? ` in "${args.document}"` : ""
    if (!lines.length)
      return `No matches for ${shown}${where}. This tool only finds exact text; for meaning-based or paraphrased queries use the search-document tool.`

    // Offsets are character positions into the document's extracted text, the
    // same coordinate space as the structure tables and chunk offsets.
    const capNote = capped
      ? `\n\nStopped at the ${MATCH_CAP}-match cap with ${chunks.length - scanned} of ${chunks.length} chunk(s) still unscanned — more matches may exist. Narrow the search with the document argument or a more specific pattern.`
      : ""
    return `${lines.length} match(es) for ${shown}${where}:\n${lines.join("\n")}${capNote}`
  },
})
