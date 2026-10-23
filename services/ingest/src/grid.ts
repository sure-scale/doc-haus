import { Database } from "bun:sqlite"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { verifyCitation } from "../../../dochaus/plugin/legal/citations"
import { formatCitations } from "../../../dochaus/lib/citations"
import { indexIsContained, lexicalChannel, phraseChannel, type ChunkRow } from "../../../dochaus/lib/search/channels"

// The tabular-review grid for a matter. Persisted as a single JSON file next to
// the matter's retrieval index (`<matter>/.dochaus/grid.json`) so the review
// work product lives with the documents it summarizes and survives across
// devices and sessions. The web app owns the merge; this service just persists.
//
// Rows are NOT stored here — they are the matter's documents, enumerated live
// from the index. We store the columns (the questions) and the computed cells.

export type GridColumn = { id: string; question: string }

export type GridCell = {
  answer: string
  citation?: {
    documentName: string
    docPath: string
    section: string
    // The clause hierarchy the search hit carried, when the index had one.
    structuralPath?: string
    excerpt: string
    charStart: number
    charEnd: number
  }
  // `reviewed` cells are locked: a lawyer has signed off, so they are never
  // recomputed. `questionHash` is the hash of the column's question at compute
  // time, so editing a column marks its cells stale without changing them.
  status: "filled" | "reviewed"
  questionHash: string
  // Reviewer comment thread on the cell; survives recomputes.
  comments?: { id: string; text: string; at: number }[]
  // A cell computed without a citation gets one re-check of its own document
  // (see recheckPassages); this records its outcome and what it searched.
  recheck?: GridRecheck
}

export type GridRecheck = { outcome: "found" | "not_found" } & RecheckScope

// What one re-check searched: the document's indexed passages, how many the
// first extraction's search had already returned, how many passages were shown
// to the model and how many of those were already read (only an uncited
// answer's re-check re-reads them), the question terms and the channels.
export type RecheckScope = { passages: number; read: number; candidates: number; reread: number; terms: string[]; channels: string[] }

// Cells are keyed `<documentName>::<columnId>`.
export type Grid = { columns: GridColumn[]; cells: Record<string, GridCell> }

function gridFile(matterDir: string) {
  return path.join(matterDir, ".dochaus", "grid.json")
}

export function readGrid(matterDir: string): Grid {
  const file = gridFile(matterDir)
  if (!existsSync(file)) return { columns: [], cells: {} }
  return JSON.parse(readFileSync(file, "utf8")) as Grid
}

export function writeGrid(matterDir: string, grid: Grid) {
  mkdirSync(path.join(matterDir, ".dochaus"), { recursive: true })
  writeFileSync(gridFile(matterDir), JSON.stringify(grid, null, 2))
}

// --- Source check of cited cells -------------------------------------------
//
// Whether each cell's recorded excerpt is still the document's text at its
// offsets, through the same exact -> re-anchor -> reject ladder that verifies
// search citations (dochaus/plugin/legal/citations.ts). Read-only: it reports
// what the run recorded against the current file and never rewrites a cell.

export type SourceCheckRequest = { cells: Array<{ key: string; documentName: string; excerpt: string; charStart: number; charEnd: number }> }
// verified: the excerpt is at its offsets. reanchored: found elsewhere in the
// document. not_found: the document no longer contains it. missing: the
// document is no longer in the matter.
export type SourceCheck = "verified" | "reanchored" | "not_found" | "missing"

// Each well-formed cell is checked; a malformed one (an older grid's offsets,
// say) is skipped and reads as unchecked, never failing the others.
export function sourceCheckRequest(value: unknown): SourceCheckRequest | undefined {
  const cells = (value as { cells?: unknown } | undefined)?.cells
  if (!Array.isArray(cells)) return
  return {
    cells: cells.flatMap((cell) => {
      const { key, documentName, excerpt, charStart, charEnd } = (cell ?? {}) as Record<string, unknown>
      return typeof key === "string" && typeof documentName === "string" && typeof excerpt === "string" && Number.isInteger(charStart) && Number.isInteger(charEnd)
        ? [{ key, documentName, excerpt, charStart: charStart as number, charEnd: charEnd as number }]
        : []
    }),
  }
}

// `resolve` maps a document name to the matter's registered source file; the
// request's own paths are never read. A cell whose document cannot be read is
// left out and reads as unchecked.
export async function checkCellSources(request: SourceCheckRequest, resolve: (name: string) => { name: string; docPath: string } | undefined) {
  const checks: Record<string, SourceCheck> = {}
  for (const cell of request.cells) {
    const target = resolve(cell.documentName)
    if (!target || !existsSync(target.docPath)) {
      checks[cell.key] = "missing"
      continue
    }
    const verified = await verifyCitation({ ...cell, documentName: target.name, docPath: target.docPath, section: "" }).catch(() => null)
    if (verified === null) continue
    checks[cell.key] = !verified ? "not_found" : verified.reanchored ? "reanchored" : "verified"
  }
  return checks
}

// --- Re-check of a negative cell -------------------------------------------
//
// A cell with no citation gets one deterministic pass over its own document:
// the question's terms through the lexical and phrase channels. For "Not
// addressed" the passages the first extraction already saw are skipped, since
// the miss was retrieval; for an answer whose source label did not resolve
// (`includeSeen`) they lead the candidates, since the answer most likely came
// from one of them. With no candidates the cell stays negative with this scope
// and no model call; otherwise the web app shows the numbered candidates to one
// extract call and adopts its answer only if it cites one of them by number.

export type RecheckRequest = {
  document: string
  question: string
  seen: Array<{ docPath: string; charStart: number; charEnd: number }>
  includeSeen?: boolean
}

export type RecheckCitation = {
  documentName: string
  docPath: string
  section: string
  structuralPath?: string
  party?: string
  excerpt: string
  charStart: number
  charEnd: number
  score: number
  flagged: boolean
}

// Candidates shown to the one re-check call: a small, focused set.
const RECHECK_PASSAGES = 8
const RECHECK_CHANNELS = ["lexical", "phrase"]
const RRF_K = 60

// Words that carry no subject on their own: question words, articles,
// auxiliaries, pronouns and prepositions. Everything else in the question is a
// term the document would have to contain to address it.
const STOPWORDS = new Set(
  ("a an the and or but nor not no if then than so as of in on at to for from by with without within about into onto upon " +
    "over under between among after before during until per via is are was were be been being am do does did doing have has had " +
    "having can could shall should will would may might must what which who whom whose when where why how whether there here " +
    "this that these those it its they them their theirs he him his she her hers we us our ours you your yours i me my any all " +
    "each every some such other own same also only just very more most less least much many both either neither one ones " +
    "anything something nothing everything document documents").split(" "),
)

export function recheckRequest(value: unknown): RecheckRequest | undefined {
  if (!value || typeof value !== "object") return
  const { document, question, seen, includeSeen } = value as Record<string, unknown>
  if (typeof document !== "string" || !document.trim() || typeof question !== "string" || !question.trim() || !Array.isArray(seen)) return
  if (includeSeen !== undefined && typeof includeSeen !== "boolean") return
  const spans = seen.flatMap((span) => {
    const { docPath, charStart, charEnd } = (span ?? {}) as Record<string, unknown>
    return typeof docPath === "string" && Number.isInteger(charStart) && Number.isInteger(charEnd)
      ? [{ docPath, charStart: charStart as number, charEnd: charEnd as number }]
      : []
  })
  if (spans.length !== seen.length) return
  return includeSeen === undefined ? { document, question, seen: spans } : { document, question, seen: spans, includeSeen }
}

/** The question's searchable terms: every content word and exact anchor (a
 * number, amount, date, identifier or section reference), plus as phrases the
 * quoted strings and runs of adjacent content words. Punctuation ends a run. */
export function questionTerms(question: string) {
  const words: string[] = []
  const phrases: string[] = []
  const add = (list: string[], term: string) => {
    if (!list.some((known) => known.toLowerCase() === term.toLowerCase())) list.push(term)
  }
  for (const [, quoted] of question.matchAll(/["\u201C\u201D]([^"\u201C\u201D]+)["\u201C\u201D]/g)) {
    const term = quoted!.trim()
    if (/[\p{L}\p{N}]/u.test(term)) add(term.split(/\s+/).length > 1 ? phrases : words, term)
  }
  let run: string[] = []
  const close = () => {
    if (run.length > 1) add(phrases, run.join(" "))
    run = []
  }
  for (const raw of question.split(/\s+/)) {
    const token = raw.replace(/^[^\p{L}\p{N}$\u00A3\u20AC\u00A7]+/u, "").replace(/[^\p{L}\p{N}%]+$/u, "")
    const term = /\p{N}/u.test(token) || (token.length > 1 && /\p{L}/u.test(token) && !STOPWORDS.has(token.toLowerCase()))
    if (term) {
      add(words, token)
      run.push(token)
    } else close()
    if (!term || !raw.endsWith(token)) close()
  }
  close()
  return { words, phrases }
}

/** The one deterministic re-check pass for a negative cell: passages of the
 * cell's own document that contain the question's terms, best first, either
 * skipping or (`includeSeen`) leading with the passages the first extraction
 * saw, rendered with the shared untrusted-excerpt framing for the re-check call.
 * A candidate's number in that list is its label for this re-check. */
export function recheckPassages(matterDir: string, request: RecheckRequest) {
  const { words, phrases } = questionTerms(request.question)
  const terms = [...phrases, ...words]
  const none = { citations: [] as RecheckCitation[], passages: "", scope: { passages: 0, read: 0, candidates: 0, reread: 0, terms, channels: RECHECK_CHANNELS } }
  const dbPath = path.join(matterDir, ".dochaus", "legal.db")
  if (!existsSync(dbPath) || !indexIsContained(matterDir, dbPath)) return none
  const db = new Database(dbPath, { readonly: true })
  try {
    const scope = { document: request.document }
    const identity = (span: { docPath: string; charStart: number; charEnd: number }) => JSON.stringify([span.docPath, span.charStart, span.charEnd])
    const span = (row: ChunkRow) => identity({ docPath: row.doc_path, charStart: row.char_start, charEnd: row.char_end })
    // First-search order of each passage already read.
    const seen = new Map<string, number>()
    for (const passage of request.seen) if (!seen.has(identity(passage))) seen.set(identity(passage), seen.size)
    const rows = db.query(
      "SELECT id, doc_name, doc_path, section, structural_path, party, text, char_start, char_end, flagged FROM chunks WHERE doc_name = ?",
    ).all(request.document) as ChunkRow[]
    const read = rows.filter((row) => seen.has(span(row)))
    // Reciprocal-rank fusion, as search-document fuses its channels.
    const fused = new Map<number, { row: ChunkRow; score: number }>()
    const channels = [
      words.length ? lexicalChannel(db, words.join(" "), scope) : [],
      ...phrases.map((phrase) => phraseChannel(db, phrase, scope)),
    ]
    for (const channel of channels)
      channel.forEach((row, rank) => {
        const entry = fused.get(row.id) ?? { row, score: 0 }
        entry.score += 1 / (RRF_K + 1 + rank)
        fused.set(row.id, entry)
      })
    const score = (row: ChunkRow) => fused.get(row.id)?.score ?? 0
    const reread = request.includeSeen
      ? [...read].sort((a, b) => score(b) - score(a) || seen.get(span(a))! - seen.get(span(b))!)
      : []
    const further = [...fused.values()].filter(({ row }) => !seen.has(span(row))).sort((a, b) => b.score - a.score).map(({ row }) => row)
    const picked = [...reread, ...further].slice(0, RECHECK_PASSAGES)
    const citations = picked
      .map((row): RecheckCitation => ({
        documentName: row.doc_name,
        docPath: row.doc_path,
        section: row.section,
        structuralPath: row.structural_path || undefined,
        party: row.party || undefined,
        excerpt: row.text,
        charStart: row.char_start,
        charEnd: row.char_end,
        score: score(row),
        flagged: row.flagged === 1,
      }))
    return {
      citations,
      passages: formatCitations(citations, { full: true }),
      scope: {
        passages: rows.length,
        read: read.length,
        candidates: citations.length,
        reread: picked.filter((row) => seen.has(span(row))).length,
        terms,
        channels: RECHECK_CHANNELS,
      },
    }
  } finally {
    db.close()
  }
}
