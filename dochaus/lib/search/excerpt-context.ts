import type { Database } from "bun:sqlite"
import { chunkScope, type ChunkRow, type SearchScope } from "./channels"
import { findTermUse } from "./term-use"

// Supplementary context the search-document tool appends after its citations.

// Auto-attach the verbatim definitions of defined terms that appear inside the
// returned excerpts. A passage like "during the Cure Period" silently depends
// on a definition that lives pages away; attaching it saves the model a lookup
// and keeps it from guessing the meaning. Matching is case-sensitive whole-word
// over every occurrence (see term-use.ts), definition text is verbatim from the defined_terms table, and the list is
// deduped and capped so a definition-dense passage cannot flood the response.
const ATTACH_CAP = 5

export function attachDefinitions(db: Database, ranked: Array<{ row: ChunkRow }>, scope: SearchScope) {
  const docPaths = [...new Set(ranked.map(({ row }) => row.doc_path))]
  if (!docPaths.length) return ""
  const filter = chunkScope(scope, "c.")
  // Definitions have source offsets but no structural path. For a strict
  // scope, require their entire stored span to fit in a scoped chunk. If that
  // coverage is not established, omit the supplementary preview.
  const sql = `SELECT t.doc_path, t.doc_name, t.term, t.definition FROM defined_terms t WHERE t.doc_path IN (${docPaths.map(() => "?").join(",")})` +
    (scope.structural_path === undefined
      ? ""
      : ` AND t.char_start < t.char_end AND EXISTS (SELECT 1 FROM chunks c WHERE c.doc_path = t.doc_path AND c.char_start <= t.char_start AND c.char_end >= t.char_end AND ${filter.sql})`)
  const terms = db.query(sql).all(
    ...docPaths,
    ...(scope.structural_path === undefined ? [] : filter.params),
  ) as Array<{ doc_path: string; doc_name: string; term: string; definition: string }>
  const attached = terms
    .filter((t) =>
      ranked.some(({ row }) => {
        if (row.doc_path !== t.doc_path) return false
        // Skip excerpts that are themselves the definition site — attaching the
        // definition under the passage that states it is noise.
        if (row.text.includes(`"${t.term}"`) || row.text.includes(`“${t.term}”`)) return false
        return findTermUse(row.text, t.term) !== -1
      }),
    )
    .filter((t, i, all) => all.findIndex((o) => o.term === t.term && o.doc_path === t.doc_path) === i)
    // Longer terms first: "Original Credit Agreement" beats "Credit Agreement"
    // for the cap when both match the same excerpt.
    .sort((a, b) => b.term.length - a.term.length)
    .slice(0, ATTACH_CAP)
  if (!attached.length) return ""
  return (
    `\n\nDefined terms used in these passages (definition previews from the documents):\n` +
    attached.map((t) => `- "${t.term}" (${t.doc_name}): ${definitionPreview(t.definition)}`).join("\n")
  )
}

function definitionPreview(definition: string) {
  const max = 600
  if (definition.length <= max) return definition
  return (
    definition.slice(0, max).trimEnd() +
    ` [definition preview clipped at ${max} characters; use cite or read-document for the full definition before quoting beyond this preview.]`
  )
}

// Auto-resolve numbered cross-references appearing in the excerpts ("as set
// forth in Section 8.3") against the citing document's own section labels, so
// the model knows the referenced section is one get-section call away instead
// of searching for it — or that it is not in the indexed text at all.
const EXCERPT_REF_RE =
  /\b(Section|Article|Clause|Exhibit|Schedule|Annex|Appendix)\s+(\d+(?:\.\d+)*(?:\([a-z]+\))*|[A-Z](?![A-Za-z])|[IVXLC]+(?![A-Za-z]))/g

export function attachCrossRefs(db: Database, ranked: Array<{ row: ChunkRow }>, scope: SearchScope) {
  const filter = scope.structural_path === undefined ? { sql: "", params: [] } : chunkScope(scope)
  const resolved = ranked
    .flatMap(({ row }) =>
      [...row.text.matchAll(EXCERPT_REF_RE)].map((m) => ({
        docPath: row.doc_path,
        docName: row.doc_name,
        ownSection: row.section,
        kind: m[1],
        label: m[2],
      })),
    )
    .filter((ref) => ref.label !== ref.ownSection)
    .filter(
      (ref, i, all) => all.findIndex((o) => o.docPath === ref.docPath && o.kind === ref.kind && o.label === ref.label) === i,
    )
    .map((ref) => ({
      ...ref,
      // EXCERPT_REF_RE labels cannot contain %/_ today; escape anyway so a
      // widened ref pattern can never turn the label into a LIKE wildcard.
      hit: db
        .query(
          "SELECT section FROM chunks WHERE doc_path = ?1 AND (section = ?2 OR section LIKE ?3 || '.%' ESCAPE '\\')" +
            (filter.sql ? ` AND ${filter.sql}` : "") + " LIMIT 1",
        )
        .get(ref.docPath, ref.label, ref.label.replaceAll(/[\\%_]/g, "\\$&"), ...filter.params) as { section: string } | null,
    }))
    .filter((ref) => ref.hit)
    .slice(0, ATTACH_CAP)
  if (!resolved.length) return ""
  return (
    `\n\nCross-references in these passages that resolve to indexed sections (fetch with the get-section tool):\n` +
    resolved.map((ref) => `- ${ref.kind} ${ref.label} → section "${ref.hit!.section}" of ${ref.docName}`).join("\n")
  )
}
