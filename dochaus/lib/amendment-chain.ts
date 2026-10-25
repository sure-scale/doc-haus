import { Database } from "bun:sqlite"
import path from "node:path"
import { chunkText, type ChunkTextRow } from "./chunk-text"

// Shared chain semantics for the amendment surfaces (issue #69). Three tools
// (amendment-chain, current-section, and the supersession pass inside
// search-document) consume these helpers so the definition of "which document
// amends which, transitively, and what each amender operates on" lives in one
// place. Read-only against the per-matter legal.db.

export type RelationRow = {
  doc_path: string
  doc_name: string
  relation: string
  target_name: string
  char_start: number
  char_end: number
}

export type SectionOpRow = {
  doc_path: string
  doc_name: string
  section_label: string
  operation: "amend" | "restate" | "delete" | "add"
  char_start: number
  char_end: number
}

export type DocRow = { doc_path: string; name: string; created_at: number }

export type Resolution =
  | { kind: "resolved"; doc: DocRow }
  | { kind: "ambiguous"; names: string[] }
  | { kind: "unmatched" }

export type ChainEntry = {
  rel: RelationRow
  resolution: Resolution
  recital: string
  ops: SectionOpRow[]
}

export type Chain = {
  key: string
  base:
    | { kind: "resolved"; doc: DocRow }
    | { kind: "unresolved"; targetName: string }
  entries: ChainEntry[] // oldest-first, by created_at then insert order
}

// Structural normalisation for comparing a stated target with a file name:
// case and runs of whitespace, underscores and hyphens. Nothing else is folded.
const normalizeTitle = (value: string) => value.toLowerCase().replace(/[\s_-]+/g, " ").trim()

// Document types ingest accepts (MATTER_DOCUMENT_EXTENSIONS in
// services/ingest/src/ingest.ts). Only one of these is ever treated as a file
// name's extension, so a dotted word in a name ("Lease.final", "v1.2") stays.
const DOCUMENT_EXTENSIONS = new Set([".docx", ".pdf", ".xlsx", ".eml", ".msg", ".zip"])
const documentExtension = (name: string) => {
  const ext = path.extname(name).toLowerCase()
  return DOCUMENT_EXTENSIONS.has(ext) ? ext : ""
}

// A document's file name without its real extension: the one its indexed path
// records (the name's own when the path records none), when that is a document
// type. Null when the name carries no such extension.
function bareFileTitle(doc: DocRow) {
  const ext = documentExtension(path.extname(doc.doc_path) ? doc.doc_path : doc.name)
  return ext && doc.name.toLowerCase().endsWith(ext) ? normalizeTitle(doc.name.slice(0, -ext.length)) : null
}

// A document resolves a stated target only on an exact match: its file name
// equals the target after structural normalisation, or it records the target
// as a defined term (NOCASE). A target may omit the file's extension; one that
// states a document extension names the whole file. Partial overlap never
// matches. The amender itself is excluded. File-name and defined-term matches
// are not ranked against each other; non-amending documents are preferred over
// amending ones, except that this preference never drops every file-name match
// in favour of a defined-term match (the target is then ambiguous). More than
// one remaining candidate is ambiguous and none is unmatched; both are
// reported, never forced.
function resolveTarget(db: Database, target: RelationRow, docs: DocRow[], relations: RelationRow[]): Resolution {
  const stated = target.target_name.trim()
  const title = normalizeTitle(stated)
  const statesExtension = documentExtension(stated) !== ""
  const definerPaths = new Set(
    (
      db
        .query("SELECT DISTINCT doc_path FROM defined_terms WHERE term = ? COLLATE NOCASE")
        .all(stated) as { doc_path: string }[]
    ).map((r) => r.doc_path),
  )
  const others = docs.filter((d) => d.doc_path !== target.doc_path)
  const byName = others.filter(
    (d) => title !== "" && (normalizeTitle(d.name) === title || (!statesExtension && bareFileTitle(d) === title)),
  )
  const candidates = others.filter((d) => byName.includes(d) || definerPaths.has(d.doc_path))
  const amenderPaths = new Set(relations.map((r) => r.doc_path))
  const nonAmending = candidates.filter((d) => !amenderPaths.has(d.doc_path))
  const pool = nonAmending.length > 0 ? nonAmending : candidates
  if (byName.length > 0 && !pool.some((d) => byName.includes(d)))
    return { kind: "ambiguous", names: candidates.map((d) => d.name) }
  if (pool.length === 1) return { kind: "resolved", doc: pool[0]! }
  if (pool.length > 1) return { kind: "ambiguous", names: pool.map((d) => d.name) }
  return { kind: "unmatched" }
}

// A document's extracted text rebuilt from its stored chunks (chunkText), with
// the [start, end) ranges stored chunk text fills; padding is not covered. When
// `before` is given, only chunks starting before it are read.
export type IndexedText = { text: string; covered: [number, number][] }
function indexedText(db: Database, docPath: string, before?: number): IndexedText {
  const rows = (
    before === undefined
      ? db.query("SELECT text, char_start FROM chunks WHERE doc_path = ? ORDER BY char_start, rowid").all(docPath)
      : db.query("SELECT text, char_start FROM chunks WHERE doc_path = ? AND char_start < ? ORDER BY char_start, rowid").all(docPath, before)
  ) as ChunkTextRow[]
  return chunkText(rows)
}

// End of the contiguous stored-text run that contains `from`, or `from` itself
// when that position is not covered by any chunk.
function coveredUntil(covered: [number, number][], from: number) {
  return covered.find(([start, end]) => start <= from && from < end)?.[1] ?? from
}

// Unicode code points in text[from, to). Callers count only stored text: a
// padded gap holds one space per UTF-16 unit, so a count across it could be
// wrong. Document readers (read-document) count offsets in code points, while
// stored chunk and op offsets are UTF-16 indexes.
function codePoints(text: string, from: number, to: number) {
  return Array.from(text.slice(from, to)).length
}

// The verbatim text at start..end of a document's extracted text, rebuilt from
// its stored chunks. Only stored chunk text is returned — never anything
// synthesized — and null when stored chunks do not cover the whole span.
// Exported so tool surfaces can quote the operative sentence at a section_op's
// offsets without re-implementing the reconstruction.
export function verbatimSpan(db: Database, docPath: string, start: number, end: number) {
  if (start >= end) return null
  const { text, covered } = indexedText(db, docPath, end)
  if (coveredUntil(covered, start) < end) return null
  return text.slice(start, end).trim() || null
}

// A document reader (read-document) returns at most this many code points.
const READ_LIMIT = 4000
// How far past an operative statement the following text may run when no
// later operative statement in the same document bounds it.
const FOLLOWING_CAP = 3000

export type OperativePassage = {
  // Verbatim statement at the op's span (the amendment's own wording); null
  // when stored chunks do not cover the span.
  operativeText: string | null
  // Verbatim text after the statement, up to the next operative statement in
  // the same amending document or FOLLOWING_CAP characters. Position-based only:
  // it may hold the restated wording or unrelated clauses, and is never
  // verified as the resulting section text.
  followingText: string | null
  // Position of the statement and the text after it in the amending document,
  // in code points (offset, and limit = length). Both are null when stored
  // text is not contiguous from the document start through that range, since
  // the count could then be wrong.
  read: { document: string; offset: number | null; limit: number | null }
}

// The amending document's own words around one op, located purely by stored
// offsets. `docOps` are that document's ops; the next one bounds the passage.
// `texts` caches each document's rebuilt text across calls (one search call
// may hit several passages superseded by the same amendment).
export function operativePassage(
  db: Database,
  op: SectionOpRow,
  docOps: SectionOpRow[],
  texts = new Map<string, IndexedText>(),
): OperativePassage {
  const next = Math.min(...docOps.filter((o) => o.char_start >= op.char_end).map((o) => o.char_start))
  const boundary = Math.min(next, op.char_end + FOLLOWING_CAP)
  const indexed = texts.get(op.doc_path) ?? indexedText(db, op.doc_path)
  texts.set(op.doc_path, indexed)
  const { text, covered } = indexed
  const anchored = op.char_start < op.char_end && coveredUntil(covered, op.char_start) >= op.char_end
  const readEnd = Math.max(op.char_end, Math.min(boundary, text.length))
  const followingEnd = anchored ? Math.min(readEnd, coveredUntil(covered, op.char_end)) : op.char_end
  const positioned = coveredUntil(covered, 0) >= readEnd
  return {
    operativeText: anchored ? text.slice(op.char_start, op.char_end).trim() || null : null,
    followingText: text.slice(op.char_end, followingEnd).trim() || null,
    read: {
      document: op.doc_name,
      offset: positioned ? codePoints(text, 0, op.char_start) : null,
      limit: positioned ? Math.min(READ_LIMIT, Math.max(1, codePoints(text, op.char_start, readEnd))) : null,
    },
  }
}

export function buildChains(db: Database): Chain[] {
  const relations = db
    .query("SELECT doc_path, doc_name, relation, target_name, char_start, char_end FROM doc_relations")
    .all() as RelationRow[]
  if (relations.length === 0) return []

  const docs = db.query("SELECT doc_path, name, created_at FROM documents").all() as DocRow[]
  const opsByDoc = new Map<string, SectionOpRow[]>()
  for (const op of db
    // Document order: ingest stores ops grouped by verb, not by position.
    .query("SELECT doc_path, doc_name, section_label, operation, char_start, char_end FROM section_ops ORDER BY char_start, id")
    .all() as SectionOpRow[]) {
    const list = opsByDoc.get(op.doc_path) ?? []
    list.push(op)
    opsByDoc.set(op.doc_path, list)
  }

  const resolved = relations.map((rel) => ({
    rel,
    resolution: resolveTarget(db, rel, docs, relations),
    recital: verbatimSpan(db, rel.doc_path, rel.char_start, rel.char_end) ?? "",
    ops: opsByDoc.get(rel.doc_path) ?? [],
  }))

  // Group by resolved target (or the trimmed stated name when unresolved), then
  // merge transitively from root targets: a target that is itself an amender
  // continues another document's chain. The seen set terminates a cycle.
  const groups = new Map<string, typeof resolved>()
  for (const entry of resolved) {
    const key =
      entry.resolution.kind === "resolved"
        ? entry.resolution.doc.doc_path
        : `unresolved:${entry.rel.target_name.trim().toLowerCase()}`
    groups.set(key, [...(groups.get(key) ?? []), entry])
  }
  const amenderPaths = new Set(resolved.map((entry) => entry.rel.doc_path))
  const flatten = (key: string, seen: Set<string>): typeof resolved => {
    if (seen.has(key)) return []
    seen.add(key)
    return (groups.get(key) ?? []).flatMap((entry) => [entry, ...flatten(entry.rel.doc_path, seen)])
  }
  const createdAt = new Map(docs.map((d) => [d.doc_path, d.created_at]))
  return [...groups.keys()]
    .filter((key) => !amenderPaths.has(key))
    .map((key) => {
      const entries = flatten(key, new Set()).sort(
        (a, b) => (createdAt.get(a.rel.doc_path) ?? 0) - (createdAt.get(b.rel.doc_path) ?? 0),
      )
      const base: Chain["base"] = key.startsWith("unresolved:")
        ? { kind: "unresolved", targetName: (groups.get(key) ?? [])[0]?.rel.target_name.trim() ?? key }
        : { kind: "resolved", doc: docs.find((d) => d.doc_path === key)! }
      return { key, base, entries }
    })
}

// For each resolved base, map every section_label any amender touches to the
// latest amender's op (newest-first walk), with all of that amender's ops so a
// caller can bound the op's passage (operativePassage). Used by search-document
// to flag and down-rank chunks whose section has been superseded.
type Supersession = { amender: string; op: SectionOpRow; amenderOps: SectionOpRow[] }
export function supersededSections(db: Database, chains: Chain[]): Map<string, Map<string, Supersession>> {
  const result = new Map<string, Map<string, Supersession>>()
  for (const chain of chains) {
    if (chain.base.kind !== "resolved") continue
    const basePath = chain.base.doc.doc_path
    const baseMap = result.get(basePath) ?? new Map<string, Supersession>()
    // Walk amenders newest-first. Within one amender, the LAST op in document
    // order wins per label (matches resolveCurrentSection's contract — a
    // document that says "§4 amended" then later "§4 deleted" reports deleted).
    for (let i = chain.entries.length - 1; i >= 0; i--) {
      const entry = chain.entries[i]!
      const latestPerLabel = new Map<string, SectionOpRow>()
      for (const op of entry.ops) latestPerLabel.set(op.section_label, op)
      for (const [label, op] of latestPerLabel) {
        if (!baseMap.has(label)) baseMap.set(label, { amender: entry.rel.doc_name, op, amenderOps: entry.ops })
      }
    }
    if (baseMap.size > 0) result.set(basePath, baseMap)
  }
  return result
}

// Walk a chain newest-first for the given section label. Returns the first
// matching amender's op (the operative one) with its operativePassage: the
// verbatim statement at that op's offsets in the amending document, the text
// that follows it there, and their code-point position. That is the
// amendment's own wording ("Section 4 ... is hereby amended and restated ..."),
// not the resulting text of the base section. The amender's chunks are never
// looked up by label: an amendment's own §4 is a different clause from the
// base document's §4. Returns null when no amender touches the label.
export function resolveCurrentSection(
  db: Database,
  chain: Chain,
  label: string,
): ({ operation: SectionOpRow["operation"]; amender: ChainEntry; op: SectionOpRow } & OperativePassage) | null {
  for (let i = chain.entries.length - 1; i >= 0; i--) {
    const entry = chain.entries[i]!
    const matching = entry.ops.filter((op) => op.section_label === label)
    if (matching.length === 0) continue
    // Within one amender, the last operative sentence in document order wins.
    const op = matching[matching.length - 1]!
    return { operation: op.operation, amender: entry, op, ...operativePassage(db, op, entry.ops) }
  }
  return null
}

// Every indexed document a rendered chain names, for its provenance footer.
export function chainNames(chain: Chain) {
  return [
    ...(chain.base.kind === "resolved" ? [chain.base.doc.name] : []),
    ...chain.entries.flatMap((entry) => [
      entry.rel.doc_name,
      ...(entry.resolution.kind === "resolved" ? [entry.resolution.doc.name] : entry.resolution.kind === "ambiguous" ? entry.resolution.names : []),
    ]),
  ]
}
