import { createHash } from "node:crypto"
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
  // The value as the source wrote it, present when the matter-level collapse
  // (collapseInitialMiddleVariants) replaced `value` with a longer form.
  surface?: string
}

// The version of matter-identifiers.json's extraction and shape; the one
// value the extractor writes and the reader accepts. Bump it when either
// changes: a stored file of another version is stale to the startup check
// (source-artifacts.ts), so each matter extracts again at its next start and
// matches from then on. 2: values are read from each chunk's unwrapped source
// slice, so a soft-wrap line break no longer splits a name. 3: `sources` keys
// each document's rows, which a later extraction reuses while the document is
// unchanged, and a collapsed row keeps its `surface`.
export const IDENTIFIERS_SCHEMA_VERSION = 3

export type MatterIdentifiers = {
  schema_version: typeof IDENTIFIERS_SCHEMA_VERSION
  generated_at: string
  documents: number
  // Each document, in extraction order, with the scan key (see scanKey) of
  // the chunks its rows were read from.
  sources: { document: string; sha256: string }[]
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

// The matter's identifiers, read from its chunks document by document (in the
// order each document's first chunk appears). Parsing chunks with compromise
// is nearly all of the time, so the rows of a document whose chunks are
// unchanged since `previous` (the matter's stored extraction) are taken from
// it: a document's rows depend only on what scanKey hashes, so they are the
// rows parsing it again would give. The collapse then runs over every
// document's rows, as in a full extraction, so the result is the same. The
// extraction yields to the event loop after each document and each parsed
// chunk, so a long one does not stall the service's other requests.
export async function extractIdentifiers(chunks: IdentifierChunk[], previous?: MatterIdentifiers): Promise<MatterIdentifiers> {
  const reusable = reusableRows(previous)
  const sources: MatterIdentifiers["sources"] = []
  const rows: IdentifierRow[] = []
  for (const [document, documentChunks] of chunksByDocument(chunks)) {
    const sha256 = scanKey(documentChunks)
    sources.push({ document, sha256 })
    const reused = reusable.get(sha256)
    if (reused) for (const row of reused) rows.push(row)
    else for (const chunk of documentChunks) {
      for (const row of scanChunk(chunk)) rows.push(row)
      await yieldToEventLoop()
    }
    await yieldToEventLoop()
  }
  return {
    schema_version: IDENTIFIERS_SCHEMA_VERSION,
    generated_at: new Date().toISOString(),
    documents: sources.length,
    sources,
    rows: collapseInitialMiddleVariants(rows),
  }
}

function yieldToEventLoop() {
  return new Promise<void>((resolve) => setImmediate(resolve))
}

function chunksByDocument(chunks: IdentifierChunk[]) {
  const documents = new Map<string, IdentifierChunk[]>()
  for (const chunk of chunks) {
    const documentChunks = documents.get(chunk.document)
    if (documentChunks) documentChunks.push(chunk)
    else documents.set(chunk.document, [chunk])
  }
  return documents
}

// The sha256 of everything a document's rows are read from: the extraction
// and compromise versions, and each chunk's id, document, source slice and
// wrapped file (where line and context point).
function scanKey(chunks: IdentifierChunk[]) {
  const hash = createHash("sha256").update(JSON.stringify([IDENTIFIERS_SCHEMA_VERSION, nlp.version]))
  for (const chunk of chunks) hash.update(JSON.stringify([chunk.id, chunk.document, chunk.text, chunk.file ?? null]))
  return hash.digest("hex")
}

// Each document's rows in `previous` as scanning yielded them (before the
// collapse), keyed by its scan key. A file whose rows are not exactly the rows
// of its listed documents is not reused at all.
function reusableRows(previous: MatterIdentifiers | undefined) {
  const reusable = new Map<string, IdentifierRow[]>()
  const sources: unknown = previous?.sources
  if (!previous || !Array.isArray(previous.rows) || !Array.isArray(sources) || !sources.every(isSource)) return reusable
  const rowsByDocument = new Map(sources.map((source) => [source.document, [] as IdentifierRow[]]))
  if (rowsByDocument.size !== sources.length) return reusable
  for (const row of previous.rows) {
    const documentRows = isStoredRow(row) ? rowsByDocument.get(row.source) : undefined
    if (!documentRows) return reusable
    const { surface, ...scanned } = row
    documentRows.push(surface === undefined ? scanned : { ...scanned, value: surface })
  }
  for (const source of sources) reusable.set(source.sha256, rowsByDocument.get(source.document)!)
  return reusable
}

function isSource(value: unknown): value is MatterIdentifiers["sources"][number] {
  if (!value || typeof value !== "object") return false
  const source = value as Record<string, unknown>
  return Object.keys(source).length === 2 && typeof source.document === "string" && typeof source.sha256 === "string"
}

function isStoredRow(value: unknown): value is IdentifierRow {
  if (!value || typeof value !== "object") return false
  const row = value as Record<string, unknown>
  return Object.keys(row).length === ("surface" in row ? 7 : 6) &&
    typeof row.type === "string" && typeof row.value === "string" && typeof row.source === "string" &&
    typeof row.chunk_id === "string" && typeof row.line === "number" && typeof row.context === "string" &&
    (!("surface" in row) || typeof row.surface === "string")
}

// Canonical-form normalization for person-name variants. When the same
// entity appears with both an initial-form middle name and a full-form
// middle name (same first + last word, single-initial middle matching the
// first letter of a multi-letter middle), every variant row's Value column
// is normalized to the longest form. The original surface form survives in
// the Context column (source text is never rewritten) and the row's
// `surface`; only the Value is normalized so downstream stages reading the
// Value for atom pinning always carry the full legal name.
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
  return rows.map((row) => (replacements.has(row) ? { ...row, value: replacements.get(row)!, surface: row.value } : row))
}

// One chunk's rows, before the matter-level collapse.
export function scanChunk(chunk: IdentifierChunk): IdentifierRow[] {
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
