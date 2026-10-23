import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { sectionPassages } from "../lib/chunk-text"
import { nameWithProvenance, sourceProvenance } from "../lib/source-provenance"

// doc.haus deterministic section lookup. Reads the per-matter legal.db that
// `services/ingest` populates and returns a section's text verbatim by
// rebuilding it from its chunks (sectionPassages) at the document's own
// offsets. Called without a section it returns the matter outline (every
// document's section labels in document order) so the agent can see structure
// before fetching. Nothing here is generated or paraphrased: a
// label that does not exist is reported as not found, never approximated, and
// the agent is pointed at search-document as the fallback.
//
// The database lives inside the matter directory (`<matter>/.dochaus/legal.db`)
// so lookups are naturally scoped to the active matter. This tool is read-only;
// all writes happen in the ingest service.

// Cap on returned section text per call, so one oversized section (or a broad
// prefix match) cannot flood the context window. Truncation always lands on a
// chunk boundary so the returned text stays verbatim.
const TEXT_CAP = 20000

type Chunk = {
  doc_name: string
  section: string
  text: string
  char_start: number
}

export default tool({
  description:
    'Fetch the full verbatim text of a contract section by its label (e.g. "7.2" or "EXHIBIT A"), or call with no arguments to get an outline of every document\'s sections in the matter. Use this when you know which section you need; use search-document when you only know what you are looking for.',
  args: {
    section: tool.schema
      .string()
      .optional()
      .describe('Section label or heading, e.g. "7.2" or "RECITALS". Omit to get the outline.'),
    document: tool.schema
      .string()
      .optional()
      .describe("Restrict to a single document by its exact name. Omit to cover the whole matter."),
  },
  async execute(args, ctx) {
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) {
      return "No documents have been indexed for this matter yet."
    }
    const db = new Database(dbPath, { readonly: true })

    if (!args.section) {
      const outline = renderOutline(db, sourceProvenance(ctx.directory), args.document)
      db.close()
      return outline
    }

    const scopedSql = (where: string) =>
      `SELECT doc_name, section, text, char_start FROM chunks WHERE ${where}` +
      (args.document ? " AND doc_name = ?" : "") +
      " ORDER BY doc_name, char_start, id"
    const params = args.document ? [args.section, args.document] : [args.section]
    const exact = db.query(scopedSql("section = ? COLLATE NOCASE")).all(...params) as Chunk[]
    // Exact label missed — retry as prefix so "7" still surfaces 7.1–7.4 (the
    // result says which labels actually matched; nothing is silently widened).
    // LIKE metacharacters in the label are escaped so they match literally.
    const prefixParams = args.document
      ? [args.section.replaceAll(/[\\%_]/g, "\\$&"), args.document]
      : [args.section.replaceAll(/[\\%_]/g, "\\$&")]
    const chunks = exact.length
      ? exact
      : (db.query(scopedSql("section LIKE ? || '%' ESCAPE '\\'")).all(...prefixParams) as Chunk[])

    if (!chunks.length) {
      const miss = renderMiss(db, sourceProvenance(ctx.directory), args.section, args.document)
      db.close()
      return miss
    }
    db.close()

    const { passages, rest } = sectionPassages(chunks, TEXT_CAP)
    const parts = passages.map(
      (passage) =>
        `=== ${passage.doc_name} — § ${passage.section} (chars ${passage.char_start}–${passage.char_start + passage.text.length}) ===\n${passage.text}`,
    )

    const matchedLabels = [...new Set(chunks.map((chunk) => `${chunk.doc_name} § ${chunk.section}`))]
    const shownLabels =
      matchedLabels.length > 20
        ? `${matchedLabels.slice(0, 20).join(", ")} … and ${matchedLabels.length - 20} more`
        : matchedLabels.join(", ")
    const prefixNote = exact.length
      ? ""
      : `No section labeled exactly "${args.section}". Matched by prefix: ${shownLabels}.\n\n`
    const truncationNote = rest
      ? `\n\n[Truncated at the ${TEXT_CAP}-character cap, on a chunk boundary before ${rest.doc_name} § ${rest.section} (char ${rest.char_start}). Call again with document: "${rest.doc_name}" and section: "${rest.section}" to fetch the rest, or use search-document for a specific passage.]`
      : ""
    return prefixNote + parts.join("\n\n") + truncationNote
  },
})

function renderOutline(db: Database, provenance: ReadonlyMap<string, string>, document?: string) {
  const sql =
    "SELECT doc_name, section, MIN(char_start) AS at FROM chunks" +
    (document ? " WHERE doc_name = ?" : "") +
    " GROUP BY doc_name, section ORDER BY doc_name, at"
  const rows = (document ? db.query(sql).all(document) : db.query(sql).all()) as Array<{
    doc_name: string
    section: string
    at: number
  }>
  if (!rows.length && document) {
    const names = db.query("SELECT name FROM documents ORDER BY name").all() as Array<{ name: string }>
    return `No document named "${document}" is indexed in this matter. Indexed documents:\n${names.map((row) => `- ${nameWithProvenance(row.name, provenance)}`).join("\n")}`
  }
  if (!rows.length) return "No documents have been indexed for this matter yet."
  const byDoc = new Map<string, Array<{ section: string; at: number }>>()
  for (const row of rows) byDoc.set(row.doc_name, [...(byDoc.get(row.doc_name) ?? []), row])
  return [...byDoc.entries()]
    .map(([doc, sections]) => `${nameWithProvenance(doc, provenance)}\n${sections.map((s) => `  § ${s.section} (char ${s.at})`).join("\n")}`)
    .join("\n\n")
}

function renderMiss(db: Database, provenance: ReadonlyMap<string, string>, section: string, document?: string) {
  // Nearby = labels sharing the stem before the first dot ("7.9" → labels
  // starting with "7"), in document order; fall back to the first labels of the
  // scope when the stem matches nothing either.
  const stem = (section.match(/^[^.\s]+/)?.[0] ?? section).replaceAll(/[\\%_]/g, "\\$&")
  const nearSql =
    "SELECT section, MIN(char_start) AS at FROM chunks WHERE section LIKE ? || '%' ESCAPE '\\'" +
    (document ? " AND doc_name = ?" : "") +
    " GROUP BY section ORDER BY at LIMIT 8"
  const near = (document ? db.query(nearSql).all(stem, document) : db.query(nearSql).all(stem)) as Array<{
    section: string
  }>
  const fallbackSql =
    "SELECT section, MIN(char_start) AS at FROM chunks" +
    (document ? " WHERE doc_name = ?" : "") +
    " GROUP BY section ORDER BY at LIMIT 8"
  const labels = near.length
    ? near
    : ((document ? db.query(fallbackSql).all(document) : db.query(fallbackSql).all()) as Array<{ section: string }>)
  // A doc-scoped miss with zero labels means the document itself is unknown —
  // report that (with the indexed names) instead of a bare section miss.
  if (document && !labels.length) {
    const names = db.query("SELECT name FROM documents ORDER BY name").all() as Array<{ name: string }>
    return `No document named "${document}" is indexed in this matter. Indexed documents:\n${names.map((row) => `- ${nameWithProvenance(row.name, provenance)}`).join("\n")}`
  }
  const scope = document ? `in "${document}"` : "in this matter"
  const nearby = labels.length ? ` Nearby labels: ${labels.map((row) => row.section).join(", ")}.` : ""
  return `No section "${section}" ${scope}.${nearby} Call this tool with no arguments to see the full outline, or use search-document to find the passage by content.`
}
