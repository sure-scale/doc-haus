import { tool } from "@opencode-ai/plugin"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { formatCitations } from "../lib/citations"
import { pendingRedlinesForDoc } from "../lib/redlines"

// doc.haus retrieval tool. Reads the per-matter legal.db that
// `services/ingest` populates and runs the query through two channels — the
// same local MiniLM embedding used at ingest time, and the BM25-ranked FTS5
// index ingest maintains over the same chunks — fused into one citation list.
//
// The database lives inside the matter directory (`<matter>/.dochaus/legal.db`)
// so retrieval is naturally scoped to the active matter and never crosses into
// another matter's privileged material. This tool is read-only; all writes
// happen in the ingest service.

const MODEL = "Xenova/all-MiniLM-L6-v2"
const DIM = 384

let extractor: any
async function embed(text: string) {
  if (!extractor) {
    const { pipeline } = await import("@xenova/transformers")
    extractor = await pipeline("feature-extraction", MODEL)
  }
  const output = await extractor(text, { pooling: "mean", normalize: true })
  return output.data as Float32Array
}

function dot(a: Float32Array, b: Float32Array) {
  let sum = 0
  for (let i = 0; i < DIM; i++) sum += a[i] * b[i]
  return sum
}

// Hybrid retrieval (issue #67): two channels over the same chunks — embedding
// cosine for meaning, FTS5/BM25 for exact tokens (section numbers, defined
// terms, party names) — fused with reciprocal-rank fusion. RRF works on ranks
// alone, so the channels' incomparable score scales never need normalizing.
const CANDIDATES = 20
const RRF_K = 60

type ChunkRow = {
  id: number
  doc_name: string
  doc_path: string
  section: string
  text: string
  char_start: number
  char_end: number
  flagged: number
}

function vectorChannel(db: Database, queryVec: Float32Array, document?: string): ChunkRow[] {
  const sql =
    "SELECT id, doc_name, doc_path, section, text, char_start, char_end, embedding, flagged FROM chunks" +
    (document ? " WHERE doc_name = ?" : "")
  const rows = (document ? db.query(sql).all(document) : db.query(sql).all()) as Array<
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
function lexicalChannel(db: Database, query: string, document?: string): ChunkRow[] {
  const match = query
    .split(/\s+/)
    .filter((token) => /[\p{L}\p{N}]/u.test(token))
    .map((token) => `"${token.replaceAll('"', '""')}"`)
    .join(" OR ")
  if (!match) return []
  const sql =
    "SELECT c.id, c.doc_name, c.doc_path, c.section, c.text, c.char_start, c.char_end, c.flagged" +
    " FROM chunks_fts f JOIN chunks c ON c.id = f.rowid WHERE chunks_fts MATCH ?" +
    (document ? " AND c.doc_name = ?" : "") +
    // bm25() is best-first ascending; weight the section label above body text so
    // a query naming a clause ranks the clause's own chunks before passing mentions.
    " ORDER BY bm25(chunks_fts, 1.0, 2.0) LIMIT ?"
  const params = document ? [match, document, CANDIDATES] : [match, CANDIDATES]
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
function phraseChannel(db: Database, query: string, document?: string): ChunkRow[] {
  const tokens = query.split(/\s+/).filter((token) => /[\p{L}\p{N}]/u.test(token))
  if (tokens.length < 2) return []
  const phrase = `"${tokens.map((token) => token.replaceAll('"', '""')).join(" ")}"`
  const sql =
    "SELECT c.id, c.doc_name, c.doc_path, c.section, c.text, c.char_start, c.char_end, c.flagged" +
    " FROM chunks_fts f JOIN chunks c ON c.id = f.rowid WHERE chunks_fts MATCH ?" +
    (document ? " AND c.doc_name = ?" : "") +
    " ORDER BY bm25(chunks_fts, 1.0, 2.0) LIMIT ?"
  const params = document ? [phrase, document, CANDIDATES] : [phrase, CANDIDATES]
  return db.query(sql).all(...params) as ChunkRow[]
}

// Per-session memory of the most recent result-sets. A model in a search loop
// keeps rephrasing the query but gets back the *same* passages every time (e.g.
// hunting a section that does not exist). When an incoming result-set matches
// one already returned in the last few calls, we short-circuit and steer the
// model to answer from what it has rather than searching again. Upstream's V2
// runner does not yet bound repeated identical tool calls (see runner/llm.ts),
// so the bound lives here at the tool boundary.
const RECENT_LIMIT = 5
// Cap how many sessions we keep loop-detection state for, so a long-lived
// server process does not accumulate one entry per session forever.
const SESSION_LIMIT = 256
const recentBySession = new Map<string, string[]>()

export default tool({
  description:
    "Search the current matter's documents for passages relevant to a query and return them as citations. Use this before answering any question about a document.",
  args: {
    query: tool.schema.string().describe("What to search for, in natural language"),
    k: tool.schema.number().int().min(1).max(20).optional().describe("Number of passages to return (default 5)"),
    document: tool.schema
      .string()
      .optional()
      .describe("Restrict the search to a single document by its exact name. Omit to search the whole matter."),
  },
  async execute(args, ctx) {
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) {
      return "No documents have been indexed for this matter yet."
    }

    const queryVec = await embed(args.query)
    const k = args.k ?? 5

    const db = new Database(dbPath, { readonly: true })
    const channels = [
      vectorChannel(db, queryVec, args.document),
      lexicalChannel(db, args.query, args.document),
      phraseChannel(db, args.query, args.document),
    ]
    db.close()

    // score = Σ 1/(RRF_K + rank) over the channels a chunk appears in, divided
    // by the best possible sum (rank 1 in every channel) to cap it at 1. A hit
    // from a single channel therefore tops out near 1/3 — the scale ranks
    // results against each other, it is not a calibrated relevance probability.
    const fused = new Map<number, { row: ChunkRow; score: number }>()
    for (const channel of channels)
      channel.forEach((row, i) => {
        const entry = fused.get(row.id) ?? { row, score: 0 }
        entry.score += 1 / (RRF_K + 1 + i)
        fused.set(row.id, entry)
      })
    const ranked = [...fused.values()]
      .sort((a, b) => b.score - a.score)
      .slice(0, k)
      .map(({ row, score }) => ({ row, score: score / (channels.length / (RRF_K + 1)) }))

    const signature = ranked.map(({ row }) => `${row.doc_name}§${row.section}`).join("|")
    const recent = recentBySession.get(ctx.sessionID) ?? []
    if (recent.includes(signature)) {
      console.warn(
        `[search-document] repeated result-set in session ${ctx.sessionID} for query "${args.query}" — short-circuiting to break a search loop`,
      )
      return "These passages were already returned by an earlier search this turn — the same results matched again, so searching further will not surface anything new. Answer from the passages you have already retrieved; if they do not address the question, say the documents do not cover it. Do not repeat this search."
    }
    if (!recentBySession.has(ctx.sessionID) && recentBySession.size >= SESSION_LIMIT) {
      const oldest = recentBySession.keys().next().value
      if (oldest !== undefined) recentBySession.delete(oldest)
    }
    recentBySession.set(ctx.sessionID, [signature, ...recent].slice(0, RECENT_LIMIT))

    const citations = ranked.map(({ row, score }) => ({
      documentName: row.doc_name,
      docPath: row.doc_path,
      section: row.section,
      excerpt: row.text,
      charStart: row.char_start,
      charEnd: row.char_end,
      score,
      // Carried through to formatCitations, which warns the model before the
      // excerpt that ingest flagged this passage as instruction-like (issue #17).
      flagged: row.flagged === 1,
    }))

    // Surface proposals already pending on the cited documents. The retrieved
    // passages reflect the clean (accepted) document on disk — they do NOT include
    // edits the assistant proposed earlier this negotiation but that are not yet
    // accepted. Listing them keeps the model from re-proposing or contradicting a
    // change it already made, and lets it compose new edits against the running state.
    const pending = [...new Set(citations.map((c) => c.docPath))].flatMap((docPath) =>
      pendingRedlinesForDoc(ctx.directory, docPath).map((r) => ({ docName: path.basename(docPath), ...r })),
    )
    const pendingNote = pending.length
      ? `\n\nPending redlines on these documents (proposed but not yet accepted — not reflected in the passages above):\n` +
        pending.map((r) => `- #${r.id} (${r.author}) in ${r.docName}: proposes "${r.new_text}"`).join("\n")
      : ""

    return {
      title: `${citations.length} passage(s) for "${args.query}"`,
      output: (formatCitations(citations) || "No relevant passages found.") + pendingNote,
      metadata: { citations, pending },
    }
  },
})
