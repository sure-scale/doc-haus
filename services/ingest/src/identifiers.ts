import nlp from "compromise"

export type IdentifierType =
  | "person" | "organization" | "place" | "money" | "date"
  | "phone" | "email" | "url" | "acronym"
  | "case_reference" | "statute_citation"

export type IdentifierRow = {
  type: IdentifierType
  value: string
  source: string
  chunk_id: string
  line: number
  context: string
}

// The version of matter-identifiers.json's extraction and shape; the one
// value the extractor writes and the reader accepts. Bump it when either
// changes: a stored file of another version is stale to the startup check
// (source-artifacts.ts), so each matter extracts again at its next start and
// matches from then on. 2: values are read from each chunk's unwrapped source
// slice, so a soft-wrap line break no longer splits a name.
export const IDENTIFIERS_SCHEMA_VERSION = 2

export type MatterIdentifiers = {
  schema_version: typeof IDENTIFIERS_SCHEMA_VERSION
  generated_at: string
  documents: number
  rows: IdentifierRow[]
}

export type IdentifierChunk = {
  id: string
  document: string
  // Values are read from this text: the chunk's source slice.
  text: string
  // The chunk file a reader opens, when it is soft-wrapped; `line` and
  // `context` point into it. The wrap turned spaces into line breaks one for
  // one and inserted a line break before each offset of `text` in `inserted`
  // (ascending).
  file?: { text: string; inserted: number[] }
}

// The offset in a wrapped chunk file of the character at `offset` in its
// source text.
export function fileOffset(offset: number, inserted: number[]): number {
  let shift = 0
  while (shift < inserted.length && inserted[shift]! <= offset) shift++
  return offset + shift
}

const CONTEXT_MAX_CHARS = 200

// Compromise selectors that work reliably in v14. `.dates()` is gone (use
// #Date+ tag match); `.money()` returns noise; `.organizations()` fragments
// on "of" — all three are handled as regex below.
type CompromiseSelector = {
  match: string
  type: IdentifierType
  filter?: (value: string) => boolean
}

const compromiseSelectors: CompromiseSelector[] = [
  { match: "#Person+", type: "person" },
  { match: "#Place+", type: "place" },
  { match: "#Date+", type: "date" },
  { match: "#PhoneNumber+", type: "phone" },
  { match: "#Email+", type: "email" },
  { match: "#Url+", type: "url" },
  {
    match: "#Acronym+",
    type: "acronym",
    filter: (value) => {
      const cleaned = value.replace(/\.$/, "")
      return cleaned.length >= 2 && /[A-Z]{2,}/.test(cleaned)
    },
  },
]

type RegexExtractor = {
  type: IdentifierType
  pattern: RegExp
  captureGroup?: number
  keep?: (value: string) => boolean
}

const moneyPattern = /\$\d[\d,]*(?:\.\d{2})?\b/g

const organizationSequencePattern =
  /\b[A-Z][a-zA-Z']+(?:(?:\s+(?:of|the|and|for|&)?\s*[A-Z][a-zA-Z']+)){1,15}/g

const organizationSuffixes = new Set([
  "department", "services", "service", "institute", "institution", "foundation",
  "association", "authority", "bureau", "office", "commission", "committee",
  "council", "board", "agency", "center", "centre", "hospital", "school",
  "university", "college", "court", "tribunal", "llc", "inc", "corp",
  "corporation", "llp", "lp", "ltd", "co", "solutions", "group", "holdings",
  "partners", "associates", "systems", "technologies", "industries",
])

function looksLikeOrganization(value: string): boolean {
  const words = value.toLowerCase().split(/\s+/)
  return words.some((w) => organizationSuffixes.has(w.replace(/\.$/, "")))
}

const caseReferencePattern =
  /\b(?:Case|Docket)\s*(?:No\.?|#)\s*([0-9]{2,4}-[A-Z]{1,4}-[0-9]{3,8}|[0-9]{4}-[A-Z]{2,}-[0-9]+)\b/g

const statuteCitationPattern =
  /\b(?:\d{1,2}\s+U\.S\.C\.\s*§\s*\d+(?:-\d+[A-Za-z]?)*|[A-Z]{2}\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?\s+§\s*\d+(?:-[A-Z0-9]+)*|§\s*\d+(?:-[A-Z0-9]+)*|(?:[A-Z][a-zA-Z']+\.?\s*(?:and\s+)?){1,5}(?:Article|Code|Law|Procedure)\s*,?\s*(?:Title|Subtitle|Chapter|Division)\s+\d+(?:\s*,?\s*(?:Title|Subtitle|Chapter|Division)\s+\d+)*)\b/g

const regexExtractors: RegexExtractor[] = [
  { type: "money", pattern: moneyPattern },
  { type: "organization", pattern: organizationSequencePattern, keep: looksLikeOrganization },
  { type: "case_reference", pattern: caseReferencePattern, captureGroup: 1 },
  { type: "statute_citation", pattern: statuteCitationPattern },
]

export function extractIdentifiers(chunks: IdentifierChunk[]): MatterIdentifiers {
  const documents = new Set(chunks.map((chunk) => chunk.document)).size
  const rows = chunks.flatMap(scanChunk)
  return {
    schema_version: IDENTIFIERS_SCHEMA_VERSION,
    generated_at: new Date().toISOString(),
    documents,
    rows: collapseInitialMiddleVariants(rows),
  }
}

// Canonical-form normalization for person-name variants. When the same
// entity appears with both an initial-form middle name and a full-form
// middle name (same first + last word, single-initial middle matching the
// first letter of a multi-letter middle), every variant row's Value column
// is normalized to the longest form. The original surface form survives in
// the Context column (source text is never rewritten); only the Value is
// normalized so downstream stages reading the Value for atom pinning
// always carry the full legal name.
//
// Conservative on purpose:
//   - Two-word names (no middle) are never collapsed — there is no initial
//     to compare, and a third word added by another source may be a title
//     ("Sarah Chen" vs "Sarah Chen Esq" already differ in last word, so
//     they sit in different buckets anyway).
//   - Two multi-letter middles that share a stem but are not the same name
//     ("Alex" vs "Alexis") are NOT collapsed — could be different people.
//   - The collapse key is (first, last) lowercased; punctuation stripped.
//     A single initial middle ("R." or "R") matches a multi-letter middle
//     starting with the same letter ("Renee").
function collapseInitialMiddleVariants(rows: IdentifierRow[]): IdentifierRow[] {
  // Collect person rows with ≥3 alpha words; group by (first, last).
  type PersonPick = { row: IdentifierRow; first: string; last: string; middleKey: string; middleInitial?: string }
  const groups = new Map<string, PersonPick[]>()
  for (const row of rows) {
    if (row.type !== "person") continue
    const cleaned = row.value.replace(/[^a-zA-Z\s]/g, "")
    const words = cleaned.split(/\s+/).filter(Boolean)
    if (words.length < 3) continue
    const first = words[0]!.toLowerCase()
    const last = words[words.length - 1]!.toLowerCase()
    const middle = words.slice(1, -1).map((w) => w.toLowerCase())
    const middleKey = middle.join(" ")
    // A "single initial" middle is one word that is a single letter.
    const middleInitial = middle.length === 1 && middle[0]!.length === 1 ? middle[0] : undefined
    const key = `${first}|${last}`
    if (!groups.has(key)) groups.set(key, [])
    groups.get(key)!.push({ row, first, last, middleKey, middleInitial })
  }
  // For each group, find the longest middle (most words; ties broken by character length).
  // If a longest middle exists and any other row has a single-initial middle matching
  // its first letter, normalize that row's Value to the row carrying the longest middle.
  const replacements = new Map<IdentifierRow, string>()
  for (const group of groups.values()) {
    if (group.length < 2) continue
    const longest = group.reduce((best, p) => {
      const bestLen = best.middleKey.split(" ").length
      const pLen = p.middleKey.split(" ").length
      if (pLen > bestLen) return p
      if (pLen === bestLen && p.middleKey.length > best.middleKey.length) return p
      return best
    })
    if (!longest) continue
    const longestFirstLetter = longest.middleKey.charAt(0)
    // The longest form itself must have a multi-letter middle (>=2 chars) — otherwise
    // we'd be normalizing TOWARD an initial form, which is the wrong direction.
    if (longest.middleKey.replace(/[^a-z]/gi, "").length < 2) continue
    for (const p of group) {
      if (p.row === longest.row) continue
      if (p.middleInitial && p.middleInitial === longestFirstLetter) {
        replacements.set(p.row, longest.row.value)
      }
    }
  }
  if (!replacements.size) return rows
  return rows.map((row) => (replacements.has(row) ? { ...row, value: replacements.get(row)! } : row))
}

function scanChunk(chunk: IdentifierChunk): IdentifierRow[] {
  const doc = nlp(chunk.text)
  const fromRegex = regexExtractors.flatMap((extractor) => extractFromRegex(chunk, extractor))
  const fromCompromise = compromiseSelectors.flatMap((selector) =>
    extractViaMatch(doc, selector, chunk),
  )
  return [...fromRegex, ...fromCompromise]
}

function extractViaMatch(
  doc: ReturnType<typeof nlp>,
  selector: CompromiseSelector,
  chunk: IdentifierChunk,
): IdentifierRow[] {
  const view = doc.match(selector.match)
  const out: IdentifierRow[] = []
  for (const match of view.json({ offset: true })) {
    const value = cleanValue(match.text ?? "")
    if (!value) continue
    if (selector.filter && !selector.filter(value)) continue
    const offset = match.start ?? match.offset?.start ?? 0
    out.push({
      type: selector.type,
      value,
      source: chunk.document,
      chunk_id: chunk.id,
      ...locate(chunk, offset),
    })
  }
  return out
}

function extractFromRegex(chunk: IdentifierChunk, extractor: RegexExtractor): IdentifierRow[] {
  const out: IdentifierRow[] = []
  const pattern = new RegExp(extractor.pattern.source, extractor.pattern.flags)
  for (const match of chunk.text.matchAll(pattern)) {
    const raw = extractor.captureGroup ? match[extractor.captureGroup] : match[0]
    const value = cleanValue(raw ?? "")
    if (!value) continue
    if (extractor.keep && !extractor.keep(value)) continue
    const offset = match.index ?? 0
    out.push({
      type: extractor.type,
      value,
      source: chunk.document,
      chunk_id: chunk.id,
      ...locate(chunk, offset),
    })
  }
  return out
}

function cleanValue(value: string): string {
  return value.replace(/[.\s]+$/, "").trim()
}

function locate(chunk: IdentifierChunk, offset: number) {
  const text = chunk.file?.text ?? chunk.text
  const at = chunk.file ? fileOffset(offset, chunk.file.inserted) : offset
  return { line: lineAt(text, at), context: contextAt(text, at) }
}

function lineAt(text: string, offset: number): number {
  if (offset <= 0) return 1
  return text.slice(0, offset).split("\n").length
}

function contextAt(text: string, offset: number): string {
  if (!Number.isFinite(offset) || offset < 0) {
    return text.slice(0, CONTEXT_MAX_CHARS).replaceAll("\n", " ").trim()
  }
  const lineStart = text.lastIndexOf("\n", offset) + 1
  const lineEnd = text.indexOf("\n", offset)
  const line = text.slice(lineStart, lineEnd === -1 ? text.length : lineEnd)
  return line.length > CONTEXT_MAX_CHARS
    ? `${line.slice(0, CONTEXT_MAX_CHARS - 1)}…`
    : line.trim()
}

export function renderIdentifiersMarkdown(data: MatterIdentifiers): string {
  const lines: string[] = [
    "# Matter Identifiers",
    "",
    `Generated ${data.generated_at} from ${data.documents} document${data.documents === 1 ? "" : "s"}, ${data.rows.length} identifier${data.rows.length === 1 ? "" : "s"}.`,
    "",
  ]
  if (!data.rows.length) {
    lines.push("_No identifiers extracted._")
    return lines.join("\n") + "\n"
  }
  const byType = new Map<IdentifierType, IdentifierRow[]>()
  for (const row of data.rows) {
    const bucket = byType.get(row.type) ?? []
    bucket.push(row)
    byType.set(row.type, bucket)
  }
  for (const type of typeDisplayOrder) {
    const bucket = byType.get(type)
    if (!bucket || !bucket.length) continue
    lines.push(`## ${type} (${bucket.length})`, "")
    lines.push("| Value | Source | Chunk | Line | Context |")
    lines.push("| --- | --- | --- | ---: | --- |")
    for (const row of bucket) {
      lines.push(
        `| ${mdCell(row.value)} | ${mdCell(row.source)} | ${mdCell(row.chunk_id)} | ${row.line} | ${mdCell(row.context)} |`,
      )
    }
    lines.push("")
  }
  return lines.join("\n")
}

const typeDisplayOrder: IdentifierType[] = [
  "person", "organization", "place", "case_reference", "statute_citation",
  "date", "money", "phone", "email", "url", "acronym",
]

function mdCell(value: string): string {
  return value.replaceAll("|", "\\|").replaceAll("\n", " ").trim()
}
