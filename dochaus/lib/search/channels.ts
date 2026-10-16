import type { Database } from "bun:sqlite"
import { realpathSync } from "node:fs"
import path from "node:path"
import { DIM, dot } from "./query-embedding"

// Retrieval channels for the search-document tool: embedding cosine, FTS5/BM25
// tokens and the literal phrase, each returning its best candidates over the
// matter's chunks. The tool fuses them with reciprocal-rank fusion.

// Candidates each channel contributes to fusion.
const CANDIDATES = 20

export type ChunkRow = {
  id: number
  doc_name: string
  doc_path: string
  section: string
  structural_path: string
  party: string
  text: string
  char_start: number
  char_end: number
  flagged: number
}

// Columns every channel reads back, in one place so the three SELECTs stay in
// lockstep. structural_path/party may be empty on chunks ingested before R2.
const CHUNK_COLS = "id, doc_name, doc_path, section, structural_path, party, text, char_start, char_end, flagged"

export type SearchScope = { document?: string; structural_path?: string }

export function chunkScope(scope: SearchScope, prefix = "") {
  const conditions = [
    ...(scope.document ? [`${prefix}doc_name = ?`] : []),
    ...(scope.structural_path === undefined
      ? []
      : [`(${prefix}structural_path = ? OR instr(${prefix}structural_path, ?) = 1)`]),
  ]
  return {
    sql: conditions.join(" AND "),
    params: [
      ...(scope.document ? [scope.document] : []),
      ...(scope.structural_path === undefined ? [] : [scope.structural_path, `${scope.structural_path} › `]),
    ],
  }
}

export function indexIsContained(directory: string, dbPath: string) {
  const root = realpathSync.native(directory)
  return [path.dirname(dbPath), dbPath].every((file) => {
    const relative = path.relative(root, realpathSync.native(file))
    return relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative)
  })
}

export function vectorChannel(db: Database, queryVec: Float32Array, scope: SearchScope): ChunkRow[] {
  const filter = chunkScope(scope)
  const sql =
    "SELECT id, doc_name, doc_path, section, structural_path, party, text, char_start, char_end, embedding, flagged FROM chunks" +
    (filter.sql ? ` WHERE ${filter.sql}` : "")
  const rows = db.query(sql).all(...filter.params) as Array<
    ChunkRow & { embedding: Uint8Array }
  >
  return rows
    .map((row) => ({
      row,
      score: dot(queryVec, new Float32Array(row.embedding.buffer, row.embedding.byteOffset, DIM)),
    }))
    .sort((a, b) => b.score - a.score)
    .slice(0, CANDIDATES)
    .map((scored) => scored.row)
}

// FTS5 MATCH has its own query syntax that throws on raw punctuation, so every
// whitespace token is wrapped in double quotes (a phrase). Quoting also makes
// dotted identifiers work: unicode61 splits "8.3" into adjacent tokens, and the
// quoted phrase matches exactly that sequence. Tokens are OR'd — BM25's IDF
// weighting lets a rare token (a section number, a party name) dominate the
// ranking while near-stopwords contribute almost nothing.
export function lexicalChannel(db: Database, query: string, scope: SearchScope): ChunkRow[] {
  const match = query
    .split(/\s+/)
    .filter((token) => /[\p{L}\p{N}]/u.test(token))
    .map((token) => `"${token.replaceAll('"', '""')}"`)
    .join(" OR ")
  if (!match) return []
  const filter = chunkScope(scope, "c.")
  const sql =
    `SELECT ${CHUNK_COLS.split(", ").map((c) => "c." + c).join(", ")}` +
    " FROM chunks_fts f JOIN chunks c ON c.id = f.rowid WHERE chunks_fts MATCH ?" +
    (filter.sql ? ` AND ${filter.sql}` : "") +
    // bm25() is best-first ascending; weight the section label and document name
    // above body text so a query naming a clause or a document ranks that
    // clause's or document's own chunks before passing mentions.
    " ORDER BY bm25(chunks_fts, 1.0, 2.0, 2.0) LIMIT ?"
  const params = [match, ...filter.params, CANDIDATES]
  return db.query(sql).all(...params) as ChunkRow[]
}

// Third channel: the whole query as one FTS5 phrase. It only matches chunks
// containing the query's tokens as a literal sequence — an exact section
// reference, defined term, or party name — and is empty for paraphrased
// semantic queries. Without it, a chunk that uniquely contains the full literal
// can be outscored in fusion by chunks the fuzzy channels both like; the extra
// rank contribution here keeps the literal hit on top, which is the point of
// hybrid retrieval. Single-token queries are already covered by the OR channel,
// and a phrase needs two tokens to add ordering signal.
export function phraseChannel(db: Database, query: string, scope: SearchScope): ChunkRow[] {
  const tokens = query.split(/\s+/).filter((token) => /[\p{L}\p{N}]/u.test(token))
  if (tokens.length < 2) return []
  const phrase = `"${tokens.map((token) => token.replaceAll('"', '""')).join(" ")}"`
  const filter = chunkScope(scope, "c.")
  const sql =
    `SELECT ${CHUNK_COLS.split(", ").map((c) => "c." + c).join(", ")}` +
    " FROM chunks_fts f JOIN chunks c ON c.id = f.rowid WHERE chunks_fts MATCH ?" +
    (filter.sql ? ` AND ${filter.sql}` : "") +
    " ORDER BY bm25(chunks_fts, 1.0, 2.0, 2.0) LIMIT ?"
  const params = [phrase, ...filter.params, CANDIDATES]
  return db.query(sql).all(...params) as ChunkRow[]
}
