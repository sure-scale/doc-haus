import { Database } from "bun:sqlite"
import { createHash } from "node:crypto"
import { getMeta, setMeta } from "./db"

// Deterministic legal-document structure extraction. Each extractor is
// signal-driven and optional: it records marked definitions, numbered
// cross-references, opening entity labels, or formulaic amendment operations
// only when the source text contains that structure. No document family is
// presumed, and a miss is an absent row rather than an inferred fact. Every row
// carries verbatim text plus offsets. Bump VERSION whenever a pattern changes so
// existing matters re-extract at next boot (server.ts) or ingest (ingest.ts).

export const VERSION = "3"

// Quoted defined term: starts uppercase, single line, sane length. Straight or
// curly quotes — extracted text carries both.
const TERM = `[“"]([A-Z][^”"\\n]{0,60}?)[”"]`
// (the "Term") / (each a "Term") / (collectively, the "Terms") parentheticals.
const PAREN_TERM_RE = new RegExp(
  `\\(\\s*(?:the\\s+|this\\s+|each\\s+an?\\s+|an?\\s+|(?:together|collectively),?\\s+(?:the\\s+)?)?${TERM}\\s*\\)`,
  "g",
)
// "Term" means / shall mean / has the meaning ... definitional sentences.
const MEANS_TERM_RE = new RegExp(`${TERM}\\s+(?:shall\\s+)?(?:means?\\b|shall\\s+have\\s+the\\s+meaning|has\\s+the\\s+meaning)`, "g")

// Numbered cross-references: "Section 8.3", "Sections 2.1(a)", "Article IV",
// "Exhibit A", "Schedule 3". Label shapes: dotted numerics with optional
// letter parentheticals, a bare capital letter, or roman numerals.
const REF_RE =
  /\b(Section|Sections|Article|Articles|Clause|Clauses|Exhibit|Exhibits|Schedule|Schedules|Annex|Appendix)\s+(\d+(?:\.\d+)*(?:\([a-z]+\))*|[A-Z](?![A-Za-z])|[IVXLC]+(?![A-Za-z]))/g

// Corporate-suffix entity names in the opening text. Name words carry no internal
// periods — a period would let a match run across a sentence boundary
// and dotted lead-ins drop off rather than corrupting the match. A standalone
// holding-company suffix terminates the name only when no entity suffix follows;
// a leading article never glues onto a name; and an ampersand is a valid connector.
const ENTITY_RE =
  /\b((?!The\b)[A-Z][A-Za-z&'\-]+(?: (?:[A-Z][A-Za-z&'\-]+|&)){0,4},? (?:(?:Holdings\s+)?(?:LLC|LLP|L\.P\.|Inc\.|Ltd\.|Limited|Corporation|Corp\.|Company|GmbH|N\.A\.|plc|S\.A\.)|Holdings))(?=[\s,.;)])/g
// A short label parenthetical following the entity.
const ENTITY_LABEL_RE = new RegExp(`^[^\\n]{0,120}?\\(\\s*(?:the\\s+|an?\\s+)?${TERM}\\s*\\)`)

// Amendment recitals are formulaic and live at the top of the document.
const AMENDS_RES = [
  /\bAmendment\s+(?:No\.\s*\d+\s+|Number\s+\d+\s+)?to\s+(?:the\s+|that\s+certain\s+)?([A-Z][^\n,(;]{4,100}?)(?=\s*[,(;\n]|\s+dated\b|\s+entered\b|\s+by\b)/g,
  /\bamends?\s+and\s+restates?\s+(?:in\s+its\s+entirety\s+)?(?:the\s+|that\s+certain\s+)?([A-Z][^\n,(;]{4,100}?)(?=\s*[,(;\n]|\s+dated\b)/g,
]

// Section-label shape, shared with REF_RE above. One identifier: dotted
// numeric with optional letter parentheticals ("7.2(a)"), a single capital
// letter ("Exhibit A"), or roman numerals ("Article IV").
const OP_LABEL = `(?:\\d+(?:\\.\\d+)*(?:\\([a-z]+\\))*|[A-Z](?![A-Za-z])|[IVXLC]+(?![A-Za-z]))`
// One or more labels separated by comma / "and" / "&" — captured as group 1 so
// "Sections 4, 5, and 7" expands to one row per label below. The separator also
// accepts the Oxford-comma form (", and" / ", &") so "4, 5, and 7" matches the
// full list rather than stopping at the conjunction.
const OP_LABEL_LIST = `(${OP_LABEL}(?:\\s*(?:,|and|&|,\\s*(?:and|&))\\s*${OP_LABEL}){0,9})`
// "of the Agreement"-style qualifier, optional, discarded so the captured label
// is bare. Capped to avoid running past a sentence boundary.
const OP_OF_BASE = `(?:\\s+of\\s+(?:the\\s+|that\\s+certain\\s+)?[A-Z][^\\n,.]{0,80}?)?`

// Operative-verb patterns. AMENDED's negative lookahead stops "amended" from
// also matching the "amended and restated" verb phrase that belongs to RESTATED.
// Each pattern's trailing `[^\\n.]{0,300}` (or the explicit optional groups for
// RESTATED) captures the full amendment statement up to the next period or
// newline so the stored span covers the text needed to verify the operation.
const OP_RESTATED_RE = new RegExp(
  `\\b(?:Section|Sections|Article|Articles)\\s+${OP_LABEL_LIST}${OP_OF_BASE}\\s+(?:is|are)\\s+hereby\\s+amended\\s+and\\s+restated[^\\n.]{0,300}`,
  "g",
)
const OP_AMENDED_RE = new RegExp(
  `\\b(?:Section|Sections|Article|Articles)\\s+${OP_LABEL_LIST}${OP_OF_BASE}\\s+(?:is|are)\\s+hereby\\s+amended(?!\\s+and\\s+restated)[^\\n.]{0,300}`,
  "g",
)
const OP_DELETED_RE = new RegExp(
  `\\b(?:Section|Sections|Article|Articles)\\s+${OP_LABEL_LIST}${OP_OF_BASE}\\s+(?:is|are)\\s+hereby\\s+(?:deleted|stricken)[^\\n.]{0,300}`,
  "g",
)
const OP_ADDED_RE = new RegExp(
  `\\b(?:(?:a|A)\\s+new\\s+)?(?:Section|Sections|Article|Articles)\\s+${OP_LABEL_LIST}(?:\\s+of\\s+(?:the\\s+|that\\s+certain\\s+)?[A-Z][^\\n,.]{0,80}?)?\\s+(?:is|are)\\s+hereby\\s+(?:added|inserted)[^\\n.]{0,300}`,
  "g",
)

// Conservative comma-list split: only well-formed separators (comma / "and" / "&")
// expand. A label fragment that does not match the bare shape drops out rather
// than guess; the resolver then has no row for it and falls back honestly.
const OP_LABEL_PIECE_RE = new RegExp(`^${OP_LABEL}$`)
function splitLabelList(list: string): string[] {
  return list
    .split(/,|\s+and\s+|\s+&\s+/)
    .map((piece) => piece.trim())
    .filter((piece) => OP_LABEL_PIECE_RE.test(piece))
}

export type SectionOp = {
  sectionLabel: string
  operation: "amend" | "restate" | "delete" | "add"
  charStart: number
  charEnd: number
}

export function extractSectionOps(text: string): SectionOp[] {
  const ops: SectionOp[] = []
  const push = (re: RegExp, operation: SectionOp["operation"]) => {
    re.lastIndex = 0
    for (const m of text.matchAll(re)) {
      const list = m[1]!
      const start = m.index!
      for (const label of splitLabelList(list)) {
        // Dedupe within the same operative sentence: a comma-list that
        // double-captures a label does not record the same op twice.
        if (ops.some((o) => o.sectionLabel === label && o.operation === operation && o.charStart === start)) continue
        ops.push({ sectionLabel: label, operation, charStart: start, charEnd: start + m[0].length })
      }
    }
  }
  push(OP_RESTATED_RE, "restate")
  push(OP_AMENDED_RE, "amend")
  push(OP_DELETED_RE, "delete")
  push(OP_ADDED_RE, "add")
  return ops
}

const OPENING_CHARS = 4000
const RECITAL_CHARS = 2000
const DEFINITION_CAP = 1500

// The paragraph (line, in extracted text) containing an offset — the verbatim
// context stored as a term's definition.
function paragraphAt(text: string, offset: number) {
  const start = text.lastIndexOf("\n", offset) + 1
  const endIndex = text.indexOf("\n", offset)
  const end = endIndex === -1 ? text.length : endIndex
  return { start, end: Math.min(end, start + DEFINITION_CAP) }
}

export function extractDefinedTerms(text: string) {
  const seen = new Map<string, { term: string; definition: string; charStart: number; charEnd: number }>()
  for (const re of [MEANS_TERM_RE, PAREN_TERM_RE]) {
    re.lastIndex = 0
    for (const m of text.matchAll(re)) {
      const span = paragraphAt(text, m.index)
      // Capture group 1 always participates in these patterns; TS cannot see that.
      const term = m[1]!
      const key = `${term}|${span.start}`
      if (seen.has(key)) continue
      seen.set(key, {
        term,
        definition: text.slice(span.start, span.end).trim(),
        charStart: span.start,
        charEnd: span.end,
      })
    }
  }
  return [...seen.values()]
}

export function extractSectionRefs(text: string) {
  return [...text.matchAll(REF_RE)].map((m) => ({
    refKind: m[1]!.replace(/s$/, ""),
    refLabel: m[2]!,
    charStart: m.index,
    charEnd: m.index + m[0].length,
  }))
}

export function extractParties(text: string) {
  // Keep the public name and storage schema for compatibility. This is an
  // optional opening-entity signal, not a complete participant roster.
  const opening = text.slice(0, OPENING_CHARS)
  const byName = new Map<string, { name: string; role: string | null; charStart: number; charEnd: number }>()
  for (const m of opening.matchAll(ENTITY_RE)) {
    const name = m[1]!.replace(/,$/, "")
    const role = opening.slice(m.index + m[0].length).match(ENTITY_LABEL_RE)?.[1] ?? null
    const existing = byName.get(name)
    // First mention wins for position; a later mention may still supply the role.
    if (existing) {
      if (!existing.role && role) existing.role = role
      continue
    }
    byName.set(name, { name, role, charStart: m.index, charEnd: m.index + m[0].length })
  }
  return [...byName.values()]
}

export function extractAmendments(text: string) {
  const recitals = text.slice(0, RECITAL_CHARS)
  const seen = new Map<string, { targetName: string; charStart: number; charEnd: number }>()
  for (const re of AMENDS_RES) {
    re.lastIndex = 0
    for (const m of recitals.matchAll(re)) {
      const targetName = m[1]!.trim()
      if (seen.has(targetName)) continue
      seen.set(targetName, { targetName, charStart: m.index, charEnd: m.index + m[0].length })
    }
  }
  return [...seen.values()]
}

export function clearStructure(db: Database, docPath: string) {
  for (const table of ["defined_terms", "section_refs", "parties", "doc_relations", "section_ops"])
    db.run(`DELETE FROM ${table} WHERE doc_path = ?`, [docPath])
}

export function extractStructure(db: Database, docPath: string, docName: string, text: string) {
  clearStructure(db, docPath)
  for (const t of extractDefinedTerms(text))
    db.run(
      "INSERT INTO defined_terms (doc_path, doc_name, term, definition, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)",
      [docPath, docName, t.term, t.definition, t.charStart, t.charEnd],
    )
  for (const r of extractSectionRefs(text))
    db.run(
      "INSERT INTO section_refs (doc_path, doc_name, ref_kind, ref_label, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)",
      [docPath, docName, r.refKind, r.refLabel, r.charStart, r.charEnd],
    )
  for (const p of extractParties(text))
    db.run("INSERT INTO parties (doc_path, doc_name, name, role, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [
      docPath,
      docName,
      p.name,
      p.role,
      p.charStart,
      p.charEnd,
    ])
  for (const a of extractAmendments(text))
    db.run(
      "INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES (?, ?, 'amends', ?, ?, ?)",
      [docPath, docName, a.targetName, a.charStart, a.charEnd],
    )
  for (const op of extractSectionOps(text))
    db.run(
      "INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)",
      [docPath, docName, op.sectionLabel, op.operation, op.charStart, op.charEnd],
    )
}

// Rebuild a document's extracted text from its chunks. Chunk offsets index into
// the original extraction; each chunk is written at its own char_start, padding
// gaps from skipped whitespace-only slices and letting overlapping chunks (the
// structure-aware chunker carries a trailing overlap) write the same region
// twice without duplicating it. Anchoring on char_start keeps every stored
// offset valid against the rebuilt string.
export function reconstructText(db: Database, docPath: string) {
  const rows = db
    .query("SELECT text, char_start FROM chunks WHERE doc_path = ? ORDER BY char_start")
    .all(docPath) as Array<{ text: string; char_start: number }>
  let out = ""
  for (const row of rows) {
    if (row.char_start > out.length) out += " ".repeat(row.char_start - out.length)
    out = out.slice(0, row.char_start) + row.text
  }
  return out
}

// A clause line like "7.2 Termination" or "12. Governing Law": the dotted number
// is the hierarchy key, the trailing text is the heading title.
const CLAUSE_RE = /^\s*(\d+(?:\.\d+)*)[.)]?\s+(.+)$/
// A heading line: short, mostly uppercase, no trailing sentence punctuation.
const HEADING_RE = /^[A-Z0-9][A-Z0-9 ,'\-&/]{2,79}$/

// label is the hierarchy key (the dotted clause number, or the heading text when
// there is no number); title is the human heading line; number is the dotted
// numeric when present, which drives the parent/child path nesting. level is the
// Markdown heading level ("##" is 2) when the extractor marked the heading line
// as Markdown. opening marks the text before the first heading, which has no
// heading of its own; its "Preamble" label and title only name it in paths.
export type Section = {
  label: string
  title: string
  number: string | null
  level?: number
  opening?: boolean
  text: string
  charStart: number
  charEnd: number
}

export function sectionize(text: string): Section[] {
  const sections: Section[] = []
  let current: Section = { label: "Preamble", title: "Preamble", number: null, opening: true, text: "", charStart: 0, charEnd: 0 }
  let offset = 0

  for (const line of text.split("\n")) {
    // PDF extractors can emit Markdown headings; strip hashes so "## 7.2 Termination"
    // sections the same as a plain "7.2 Termination" line.
    const markdown = line.trim().match(/^(#{1,6})\s+/)
    const trimmed = line.trim().slice(markdown?.[0].length ?? 0)
    const isMdHeading = Boolean(markdown)
    const clause = trimmed.match(CLAUSE_RE)
    const isHeading = clause || (isMdHeading ? isMarkdownHeading(trimmed) : trimmed.length > 0 && HEADING_RE.test(trimmed))
    if (isHeading && trimmed.length > 0) {
      if (current.text.trim()) {
        current.charEnd = offset
        sections.push(current)
      }
      const number = clause?.[1] ?? null
      const title = (clause?.[2] ?? trimmed).trim()
      current = { label: number ?? trimmed, title, number, level: markdown?.[1]?.length, text: line + "\n", charStart: offset, charEnd: 0 }
    } else {
      current.text += line + "\n"
    }
    offset += line.length + 1
  }
  if (current.text.trim()) {
    current.charEnd = offset
    sections.push(current)
  }
  return sections
}

function isMarkdownHeading(text: string) {
  if (CLAUSE_RE.test(text)) return true
  if (/^(appendix|annex|exhibit|schedule|article)\b/i.test(text)) return true
  const words = text.match(/[A-Za-z][A-Za-z'’-]*/g) ?? []
  if (words.length < 2) return false
  if (/^[A-Z]{1,4}\s+\d+[A-Z0-9.-]*$/i.test(text)) return false
  return !words.every((word) => ["a", "an", "and", "between", "for", "of", "or", "the"].includes(word.toLowerCase()))
}

// Build a human structural path per section from its heading numbering. Sections
// arrive in document order; a dotted
// number nests under whichever earlier section is its numeric prefix ("7.2"
// under "7"), so the path carries the full ancestry. Unnumbered headings
// (ALL-CAPS titles, the preamble) stand alone — they have no numeric parent to
// hang under and pollute the ancestry if guessed. The result is keyed by the
// Section object so the caller looks each one up while iterating the same list.
export function buildSectionPaths(docName: string, sections: Section[]) {
  const paths = new Map<Section, string>()
  const root = `Doc(${docName})`
  // Ancestors with a number, in document order, used to find a section's parent.
  const numbered: Section[] = []
  for (const section of sections) {
    const crumb = section.number ? `§${section.number} ${section.title}`.trim() : section.title
    if (!section.number) {
      paths.set(section, `${root} › ${crumb}`)
      numbered.length = 0
      continue
    }
    // Parent = the most recent earlier numbered section whose number is a strict
    // dotted prefix ("7" of "7.2", "7.2" of "7.2.1").
    const parent = [...numbered].reverse().find((s) => section.number!.startsWith(s.number + "."))
    const prefix = parent ? paths.get(parent)! : root
    paths.set(section, `${prefix} › ${crumb}`)
    numbered.push(section)
  }
  return paths
}

type SectionRow = {
  ordinal: number
  parentOrdinal: number | null
  depth: number
  label: string | null
  heading: string | null
  charStart: number
  charEnd: number
  structureSource: "sectionize" | "markdown"
}

// One document's rows for the sections table (see db.ts), from its
// reconstructed text. The hierarchy is the one buildSectionPaths gives the
// structural paths (a numbered section nests under the latest earlier numbered
// section whose number is its dotted prefix, and an unnumbered heading ends that
// numbering) plus the Markdown heading levels the extractor emitted: a Markdown
// heading closes every open Markdown heading at its level or deeper, and a
// section with no numbered parent belongs to the innermost Markdown heading
// still open. Nothing else is inferred from the text.
export function documentSections(text: string) {
  // The reconstructed text normally ends with the newline sectionize counts
  // after the last line; sectionize the text before it so the last span ends
  // with the text. Text without that newline still counts one past its end, and
  // leading blank lines sectionize drops belong to no section, so the first span
  // starts at 0 and the last ends at text.length below: the spans tile the text.
  const sections = sectionize(text.endsWith("\n") ? text.slice(0, -1) : text)
  const rows: SectionRow[] = []
  const numbered: Array<{ number: string; row: SectionRow }> = []
  const markdown: Array<{ level: number; row: SectionRow }> = []
  for (const section of sections) {
    const level = section.level
    const numberedParent = section.number
      ? numbered.findLast((open) => section.number!.startsWith(open.number + "."))
      : undefined
    if (level) while (markdown.length && markdown.at(-1)!.level >= level) markdown.pop()
    const parent = numberedParent?.row ?? markdown.at(-1)?.row
    const row: SectionRow = {
      ordinal: rows.length + 1,
      parentOrdinal: parent?.ordinal ?? null,
      depth: (parent?.depth ?? 0) + 1,
      label: section.number,
      heading: section.opening ? null : section.title,
      charStart: rows.length ? section.charStart : 0,
      charEnd: Math.min(section.charEnd, text.length),
      structureSource: level ? "markdown" : "sectionize",
    }
    if (section.number) numbered.push({ number: section.number, row })
    else numbered.length = 0
    if (level) markdown.push({ level, row })
    rows.push(row)
  }
  return rows
}

// Structure-aware chunking. A section is the atomic unit; it is emitted whole
// when it fits the token budget so a coherent component is not split across
// chunks. An over-long section is split on blank-line paragraph boundaries (and
// never inside a table — consecutive pipe/grid rows are held together) with a
// trailing-paragraph overlap so content that straddles a boundary still shows
// up whole in one chunk. Offsets stay anchored to the original extracted text.
const CHUNK_TOKENS = 512
const CHARS_PER_TOKEN = 4
const MAX_CHARS = CHUNK_TOKENS * CHARS_PER_TOKEN
const OVERLAP_RATIO = 0.12

// A table row in extracted/markdown text: a pipe-delimited row or a markdown
// header-separator border ("|---|---|").
function isTableLine(line: string) {
  const t = line.trim()
  return t.startsWith("|") || /^\|?[-+| :]+\|[-+| :]*$/.test(t)
}

export function chunkSection(section: { text: string; charStart: number }) {
  if (section.text.length <= MAX_CHARS) {
    if (!section.text.trim()) return []
    return [{ text: section.text, charStart: section.charStart, charEnd: section.charStart + section.text.length }]
  }

  // Group lines into blocks that must stay whole: a run of table rows is one
  // block, every other line is its own block. Blocks are then packed into
  // chunks up to the budget, keeping each block (and so each table) atomic.
  const lines = section.text.split("\n")
  const blocks: { text: string; start: number }[] = []
  let offset = 0
  let i = 0
  while (i < lines.length) {
    const start = offset
    if (isTableLine(lines[i]!)) {
      let text = ""
      while (i < lines.length && isTableLine(lines[i]!)) {
        text += lines[i] + "\n"
        offset += lines[i]!.length + 1
        i++
      }
      blocks.push({ text, start })
      continue
    }
    blocks.push({ text: lines[i] + "\n", start })
    offset += lines[i]!.length + 1
    i++
  }

  const chunks: { text: string; charStart: number; charEnd: number }[] = []
  let buffer: { text: string; start: number }[] = []
  let bufferChars = 0
  const flush = () => {
    if (!buffer.length) return
    const text = buffer.map((b) => b.text).join("")
    const charStart = section.charStart + buffer[0]!.start
    if (text.trim()) chunks.push({ text, charStart, charEnd: charStart + text.length })
    // Carry the trailing blocks (~OVERLAP_RATIO of the budget) into the next
    // chunk so content spanning the boundary appears whole somewhere.
    const overlap: { text: string; start: number }[] = []
    let overlapChars = 0
    for (let j = buffer.length - 1; j >= 0 && overlapChars < MAX_CHARS * OVERLAP_RATIO; j--) {
      overlap.unshift(buffer[j]!)
      overlapChars += buffer[j]!.text.length
    }
    buffer = overlap
    bufferChars = overlapChars
  }
  for (const block of blocks) {
    // A single block over budget (a giant table) is emitted on its own rather
    // than dropped — atomicity wins over the soft size cap.
    if (bufferChars && bufferChars + block.text.length > MAX_CHARS) flush()
    buffer.push(block)
    bufferChars += block.text.length
  }
  flush()
  return chunks
}

// Backfill matters indexed before structure extraction existed (or before the
// current pattern VERSION). Text comes back out of the chunks themselves, so no
// re-parse of the source file is needed.
export function migrateStructure(db: Database) {
  if (getMeta(db, "structure_version") === VERSION) return 0
  const docs = db.query("SELECT doc_path, name FROM documents WHERE role = 'source'").all() as Array<{
    doc_path: string
    name: string
  }>
  for (const doc of docs) extractStructure(db, doc.doc_path, doc.name, reconstructText(db, doc.doc_path))
  setMeta(db, "structure_version", VERSION)
  return docs.length
}

// Record a source document's sections, each chunk's section ordinal and the
// document revision. All three derive from the chunks already stored, so ingest
// (right after writing them) and the backfill below write the same rows.
export function recordSections(db: Database, docPath: string, docName: string) {
  const text = reconstructText(db, docPath)
  const sections = documentSections(text)
  db.run("DELETE FROM sections WHERE doc_path = ?", [docPath])
  const insert = db.query(
    "INSERT INTO sections (doc_path, doc_name, ordinal, parent_ordinal, depth, label, heading, char_start, char_end, structure_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
  )
  for (const s of sections)
    insert.run(docPath, docName, s.ordinal, s.parentOrdinal, s.depth, s.label, s.heading, s.charStart, s.charEnd, s.structureSource)
  // Each chunk was cut from one section of this same text, so the section that
  // holds a chunk's first character is the one it came from. (chunkSection
  // counts one newline past a section it splits, so that section's last chunk
  // ends one character after the section does.)
  const chunks = db
    .query("SELECT id, char_start FROM chunks WHERE doc_path = ? ORDER BY char_start")
    .all(docPath) as Array<{ id: number; char_start: number }>
  const link = db.query("UPDATE chunks SET section_ordinal = ? WHERE id = ?")
  let at = 0
  for (const chunk of chunks) {
    while (sections[at + 1] && sections[at + 1]!.charStart <= chunk.char_start) at++
    link.run(sections[at]?.ordinal ?? null, chunk.id)
  }
  db.run("UPDATE documents SET rev = ? WHERE doc_path = ?", [revision(text), docPath])
}

// The first 12 hex digits of the sha256 of the canonical text: the untrimmed
// reconstructed text that chunk offsets, section spans and citation checks all
// index. A span is valid only under the rev it was issued with, so any change
// to that text, leading or trailing whitespace included, changes the rev.
function revision(text: string) {
  return createHash("sha256").update(text).digest("hex").slice(0, 12)
}

// Bump when sectionize, its heading patterns or the derivation of the recorded
// rows (documentSections, recordSections, revision) change, so existing matters
// re-record at next boot (server.ts) or ingest (ingest.ts).
const SECTIONS_VERSION = "1"

// Backfill matters indexed before sections were recorded or before the current
// SECTIONS_VERSION, plus any source document still without a rev, and drop
// sections no source document owns. Like the structure backfill it reads the
// chunks, which is what ingest records from too, so a backfilled document
// matches a fresh ingest. One transaction never leaves a matter half-recorded.
// With nothing to do it costs the single query below.
export function migrateSections(db: Database) {
  const state = db
    .query(
      `SELECT (SELECT value FROM meta WHERE key = 'sections_version') AS version,
        EXISTS (SELECT 1 FROM documents WHERE role = 'source' AND rev IS NULL) AS unrecorded,
        EXISTS (SELECT 1 FROM sections WHERE doc_path NOT IN (SELECT doc_path FROM documents WHERE role = 'source')) AS orphaned`,
    )
    .get() as { version: string | null; unrecorded: number; orphaned: number }
  const current = state.version === SECTIONS_VERSION
  if (current && !state.unrecorded && !state.orphaned) return 0
  const docs = db
    .query(`SELECT doc_path, name FROM documents WHERE role = 'source'${current ? " AND rev IS NULL" : ""}`)
    .all() as Array<{ doc_path: string; name: string }>
  db.transaction(() => {
    db.run("DELETE FROM sections WHERE doc_path NOT IN (SELECT doc_path FROM documents WHERE role = 'source')")
    for (const doc of docs) recordSections(db, doc.doc_path, doc.name)
    setMeta(db, "sections_version", SECTIONS_VERSION)
  })()
  return docs.length
}
