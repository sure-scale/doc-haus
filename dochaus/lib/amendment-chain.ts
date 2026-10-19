import { Database } from "bun:sqlite"

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

// A document resolves a stated target when its indexed name contains every
// token of the target (NOCASE), or when it defines the target as a term. The
// amender itself is excluded, and amenders are dropped whenever a non-amending
// candidate exists. Anything still unclear is reported, never forced. Verbatim
// from the original inline resolver in dochaus/tool/amendment-chain.ts.
function resolveTarget(db: Database, target: RelationRow, docs: DocRow[], relations: RelationRow[]): Resolution {
  const tokens = target.target_name
    .toLowerCase()
    .split(/\s+/)
    .filter((t) => /[\p{L}\p{N}]/u.test(t))
  const definerPaths = new Set(
    (
      db
        .query("SELECT DISTINCT doc_path FROM defined_terms WHERE term = ? COLLATE NOCASE")
        .all(target.target_name) as { doc_path: string }[]
    ).map((r) => r.doc_path),
  )
  const candidates = docs.filter(
    (d) =>
      d.doc_path !== target.doc_path &&
      (tokens.every((t) => d.name.toLowerCase().includes(t)) || definerPaths.has(d.doc_path)),
  )
  const amenderPaths = new Set(relations.map((r) => r.doc_path))
  const nonAmending = candidates.filter((d) => !amenderPaths.has(d.doc_path))
  const pool = nonAmending.length > 0 ? nonAmending : candidates
  if (pool.length === 1) return { kind: "resolved", doc: pool[0]! }
  if (pool.length > 1) return { kind: "ambiguous", names: pool.map((d) => d.name) }
  return { kind: "unmatched" }
}

// The verbatim text at start..end of a document's extracted text, rebuilt from
// its stored chunks the same way services/ingest/src/structure.ts
// reconstructText does. Only stored chunk text is returned — never anything
// synthesized. Exported so tool surfaces can quote the operative sentence at a
// section_op's offsets without re-implementing the reconstruction.
export function verbatimSpan(db: Database, docPath: string, start: number, end: number) {
  const rows = db
    .query("SELECT text, char_start FROM chunks WHERE doc_path = ? AND char_start < ? ORDER BY char_start")
    .all(docPath, end) as { text: string; char_start: number }[]
  let rebuilt = ""
  for (const row of rows) {
    if (row.char_start > rebuilt.length) rebuilt += " ".repeat(row.char_start - rebuilt.length)
    rebuilt += row.text
  }
  return rebuilt.slice(start, end).trim()
}

export function buildChains(db: Database): Chain[] {
  const relations = db
    .query("SELECT doc_path, doc_name, relation, target_name, char_start, char_end FROM doc_relations")
    .all() as RelationRow[]
  if (relations.length === 0) return []

  const docs = db.query("SELECT doc_path, name, created_at FROM documents").all() as DocRow[]
  const opsByDoc = new Map<string, SectionOpRow[]>()
  for (const op of db
    .query("SELECT doc_path, doc_name, section_label, operation, char_start, char_end FROM section_ops")
    .all() as SectionOpRow[]) {
    const list = opsByDoc.get(op.doc_path) ?? []
    list.push(op)
    opsByDoc.set(op.doc_path, list)
  }

  const resolved = relations.map((rel) => ({
    rel,
    resolution: resolveTarget(db, rel, docs, relations),
    recital: verbatimSpan(db, rel.doc_path, rel.char_start, rel.char_end),
    ops: opsByDoc.get(rel.doc_path) ?? [],
  }))

  // Group by resolved target (or the stated name when unresolved), then merge
  // transitively from root targets: a target that is itself an amender
  // continues another document's chain. The seen set terminates a cycle.
  const groups = new Map<string, typeof resolved>()
  for (const entry of resolved) {
    const key =
      entry.resolution.kind === "resolved"
        ? entry.resolution.doc.doc_path
        : `unresolved:${entry.rel.target_name.toLowerCase()}`
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
        ? { kind: "unresolved", targetName: (groups.get(key) ?? [])[0]?.rel.target_name ?? key }
        : { kind: "resolved", doc: docs.find((d) => d.doc_path === key)! }
      return { key, base, entries }
    })
}

// For each resolved base, map every section_label any amender touches to the
// latest amender's op (newest-first walk). Used by search-document to flag and
// down-rank chunks whose section has been superseded.
export function supersededSections(
  db: Database,
  chains: Chain[],
): Map<string, Map<string, { amender: string; op: SectionOpRow }>> {
  const result = new Map<string, Map<string, { amender: string; op: SectionOpRow }>>()
  for (const chain of chains) {
    if (chain.base.kind !== "resolved") continue
    const basePath = chain.base.doc.doc_path
    const baseMap = result.get(basePath) ?? new Map<string, { amender: string; op: SectionOpRow }>()
    // Walk amenders newest-first. Within one amender, the LAST op in document
    // order wins per label (matches resolveCurrentSection's contract — a
    // document that says "§4 amended" then later "§4 deleted" reports deleted).
    for (let i = chain.entries.length - 1; i >= 0; i--) {
      const entry = chain.entries[i]!
      const latestPerLabel = new Map<string, SectionOpRow>()
      for (const op of entry.ops) latestPerLabel.set(op.section_label, op)
      for (const [label, op] of latestPerLabel) {
        if (!baseMap.has(label)) baseMap.set(label, { amender: entry.rel.doc_name, op })
      }
    }
    if (baseMap.size > 0) result.set(basePath, baseMap)
  }
  return result
}

// Walk a chain newest-first for the given section label. Returns the first
// matching amender's op (the operative one) and operativeText: the verbatim
// statement at that op's offsets in the amending document. That is the
// amendment's own wording ("Section 4 ... is hereby amended and restated ..."),
// not the resulting text of the base section, and it is null when the index
// cannot rebuild the span. The amender's chunks are never looked up by label:
// an amendment's own §4 is a different clause from the base document's §4.
// Returns null when no amender touches the label.
export function resolveCurrentSection(
  db: Database,
  chain: Chain,
  label: string,
): { operation: SectionOpRow["operation"]; amender: ChainEntry; op: SectionOpRow; operativeText: string | null } | null {
  for (let i = chain.entries.length - 1; i >= 0; i--) {
    const entry = chain.entries[i]!
    const matching = entry.ops.filter((op) => op.section_label === label)
    if (matching.length === 0) continue
    // Within one amender, the last operative sentence in document order wins.
    const op = matching[matching.length - 1]!
    const operativeText = verbatimSpan(db, op.doc_path, op.char_start, op.char_end)
    return { operation: op.operation, amender: entry, op, operativeText: operativeText || null }
  }
  return null
}
