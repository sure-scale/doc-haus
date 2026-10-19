import type { Database } from "bun:sqlite"
import { chunkScope, type ChunkRow, type SearchScope } from "./channels"

// Supplementary context the search-document tool appends after its citations.

// Auto-attach the verbatim definitions of defined terms that appear inside the
// returned excerpts. A passage like "during the Cure Period" silently depends
// on a definition that lives pages away; attaching it saves the model a lookup
// and keeps it from guessing the meaning. Matching is case-sensitive whole-word
// (defined terms are capitalized — case folding would match ordinary prose),
// definition text is verbatim from the defined_terms table, and the list is
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
        const at = row.text.indexOf(t.term)
        if (at === -1) return false
        const before = row.text[at - 1]
        const after = row.text[at + t.term.length]
        return (before === undefined || !/[A-Za-z0-9]/.test(before)) && (after === undefined || !/[A-Za-z0-9]/.test(after))
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

// For numbered references in the excerpts ("Section 8.3"), report the citing
// document's own section with a matching label, so the model knows such a
// section is one get-section call away. A matching label is a fact about this
// document, not a resolution: "Section 8.3 of the Credit Agreement" inside a
// guarantee still matches the guarantee's 8.3. The note says so and leaves the
// referent to the reading agent.
const EXCERPT_REF_RE =
  /\b(Section|Article|Clause|Exhibit|Schedule|Annex|Appendix)\s+(\d+(?:\.\d+)*(?:\([a-z]+\))*|[A-Z](?![A-Za-z])|[IVXLC]+(?![A-Za-z]))/g

// Ingest labels a numbered line ("8.3 Notices") by its number alone and an
// unnumbered heading ("EXHIBIT B - FORM OF NOTICE") by its own text. A bare
// number therefore records no kind beyond the section numbering itself, so only
// Section/Clause references may match it; Article, Exhibit, Schedule, Annex and
// Appendix match only a heading label that names that kind and label.
const CLAUSE_NUMBER_LABEL = /^\d+(?:\.\d+)*$/
const CLAUSE_NUMBER_KINDS = new Set(["Section", "Clause"])

type LabelledSection = { section: string; structural_path: string }

const LABEL_CHAR = /[\p{L}\p{N}]/u

// A kind-named heading starts with the kind and the label, both compared
// without letter case. The label may end the heading or be followed by
// whitespace or punctuation ("ARTICLE V—Definitions", "Schedule 2: Pricing"),
// but not by a character that extends it into another label ("EXHIBIT A-1",
// "8.3.1", "ARTICLE VI").
function namesKindAndLabel(section: string, kind: string, label: string) {
  const heading = section.trim()
  if (heading.slice(0, kind.length).toUpperCase() !== kind.toUpperCase()) return false
  const rest = heading.slice(kind.length)
  const labelAt = rest.length - rest.trimStart().length
  if (!labelAt || rest.slice(labelAt, labelAt + label.length).toUpperCase() !== label.toUpperCase()) return false
  const [next, after] = [...rest.slice(labelAt + label.length)]
  if (next === undefined) return true
  if (LABEL_CHAR.test(next)) return false
  // ".", "-" or "(" directly before a letter or digit continues the label.
  return !(".-(".includes(next) && after !== undefined && LABEL_CHAR.test(after))
}

function matchingSection(sections: LabelledSection[], kind: string, label: string) {
  const named = sections.find(({ section }) => namesKindAndLabel(section, kind, label))
  if (named || !CLAUSE_NUMBER_KINDS.has(kind)) return named
  const numbered = sections.filter(({ section }) => CLAUSE_NUMBER_LABEL.test(section))
  return numbered.find(({ section }) => section === label) ?? numbered.find(({ section }) => section.startsWith(`${label}.`))
}

// Headings come from the document; keep the note line short.
const HEADING_MAX = 80

function clipHeading(heading: string) {
  const points = [...heading]
  return points.length > HEADING_MAX ? `${points.slice(0, HEADING_MAX).join("").trimEnd()}…` : heading
}

export function attachCrossRefs(db: Database, ranked: Array<{ row: ChunkRow }>, scope: SearchScope) {
  const filter = scope.structural_path === undefined ? { sql: "", params: [] } : chunkScope(scope)
  const sectionsByDoc = new Map<string, LabelledSection[]>()
  const sectionsOf = (docPath: string) => {
    const cached = sectionsByDoc.get(docPath)
    if (cached) return cached
    // One row per labelled section, in document order, not one per chunk.
    const sections = db
      .query(
        "SELECT section, structural_path FROM chunks WHERE doc_path = ?" +
          (filter.sql ? ` AND ${filter.sql}` : "") +
          " GROUP BY section, structural_path ORDER BY MIN(char_start)",
      )
      .all(docPath, ...filter.params) as LabelledSection[]
    sectionsByDoc.set(docPath, sections)
    return sections
  }
  const notes = ranked
    .flatMap(({ row }) =>
      [...row.text.matchAll(EXCERPT_REF_RE)].map((m) => ({
        row,
        hit: matchingSection(sectionsOf(row.doc_path), m[1]!, m[2]!),
      })),
    )
    .filter(({ row, hit }) => hit && hit.section !== row.section)
    .filter(
      ({ row, hit }, i, all) =>
        all.findIndex((o) => o.row.doc_path === row.doc_path && o.hit!.section === hit!.section) === i,
    )
    .slice(0, ATTACH_CAP)
    .map(({ row, hit }) => {
      const heading = hit!.structural_path.split(" › ").at(-1)
      return `- ${row.doc_name} has a section labelled "${hit!.section}"` + (heading && heading !== hit!.section ? `: ${clipHeading(heading)}` : "")
    })
  if (!notes.length) return ""
  return (
    `\n\nSections whose labels match references in these passages (a matching label does not show which document a reference means; to read one, pass both its document and section label to get-section):\n` +
    notes.join("\n")
  )
}
