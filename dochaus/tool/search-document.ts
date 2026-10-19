import { searchCurrentDocument } from "../lib/current-document-retrieval"
import { tool } from "./legacy-definition"
import { Database } from "bun:sqlite"
import { existsSync } from "node:fs"
import path from "node:path"
import { formatCitations } from "../lib/citations"
import { buildChains, operativePassage, supersededSections, type IndexedText } from "../lib/amendment-chain"
import { pendingRedlinesForDoc } from "../lib/redlines"
import { draftingTraceReference } from "../lib/drafting-trace-identity"
import { indexIsContained, lexicalChannel, phraseChannel, vectorChannel, type ChunkRow } from "../lib/search/channels"
import { attachCrossRefs, attachDefinitions } from "../lib/search/excerpt-context"
import { embed, MODEL } from "../lib/search/query-embedding"

// doc.haus retrieval tool. Reads the per-matter legal.db that
// `services/ingest` populates and runs the query through two channels — the
// same local embedding model used at ingest time, and the BM25-ranked FTS5
// index ingest maintains over the same chunks — fused into one citation list.
//
// Resolve the metadata directory and database against the physical matter
// directory before loading the model or opening the index: a lexical path
// inside a matter can be a symbolic-link alias into a sibling matter.
// This tool is read-only; all writes happen in the ingest service.

// Query embedding, the three retrieval channels and the excerpt context live in
// lib/search/; this file fuses the channels and renders the result.

// Hybrid retrieval (issue #67): two channels over the same chunks — embedding
// cosine for meaning, FTS5/BM25 for exact tokens (section numbers, defined
// terms, party names) — fused with reciprocal-rank fusion. RRF works on ranks
// alone, so the channels' incomparable score scales never need normalizing.
const RRF_K = 60

// Issue #69: a chunk whose (doc, section) is touched by a later amender's
// section_op is potentially superseded. Apply a modest penalty so the
// amender's operative version outranks it when both match a query, and flag
// the citation so the renderer can warn the model. The chunk still surfaces
// for historical-research queries.
const SUPERSESSION_PENALTY = 0.5

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
    "view=proposed_body reads/searches the pending generated revision from current initial paragraphs plus amendments, with initial-PID/part coordinates; it is NOT actual Word or source authority. Search an explicit evidence corpus: view=source (default) searches indexed uploads; current_body, comments and deleted_text search exact current Word artifacts lexically with native provenance. Current views require an exact plan ID or filename and do not search uploaded sources. Unavailable is distinct from searched with no hits; neither empty lexical matches nor a different view proves absence. Read relevant consecutive current entries and related provisions before deciding.",
  args: {
    view: tool.schema.enum(["source", "current_body", "comments", "deleted_text", "proposed_body"]).optional().describe("Select source, a distinct actual Word view, or proposed_body for the pending generated revision."),
    query: tool.schema.string().describe("What to search for, in natural language"),
    k: tool.schema.number().int().min(1).max(20).optional().describe("Number of passages to return (default 5)"),
    document: tool.schema
      .string()
      .optional()
      .describe("Restrict the search to a single document by its exact name. Omit to search the whole matter."),
    structural_path: tool.schema
      .string()
      .min(1)
      .optional()
      .describe("Select passages at an exact stored structural path and its descendants separated by ' › '. Copy the full path from a citation; matching is literal and case-sensitive. Intersects the document filter. Selects passages and supported supplemental context; indexed amendment warnings may refer to other paths."),
  },
  async execute(args, ctx) {
    if (args.view && args.view !== "source") {
      if (!args.document) return "Current view unavailable: provide one exact planned document ID or filename; no search was performed."
      if (args.structural_path) return "Current view unavailable: structural_path belongs to uploaded-source citations. Use current entry offsets and PIDs instead; no search was performed."
      return searchCurrentDocument(ctx.directory, { ...args, document: args.document, view: args.view })
    }
    const dbPath = path.join(ctx.directory, ".dochaus", "legal.db")
    if (!existsSync(dbPath)) {
      return "No documents have been indexed for this matter yet."
    }
    if (!indexIsContained(ctx.directory, dbPath)) {
      return "This matter's search index is outside its matter directory. Search was not performed."
    }

    if (args.document) {
      const sourceIndex = new Database(dbPath, { readonly: true })
      try {
        if (!sourceIndex.query("SELECT 1 FROM chunks WHERE doc_name = ? LIMIT 1").get(args.document))
          return `Source view unavailable: ${JSON.stringify(args.document)} has no indexed uploaded-source passages. No source search was performed. For a planned generated document, select view=current_body, comments or deleted_text with its exact plan ID or filename.`
      } finally { sourceIndex.close() }
    }
    const queryVec = await embed(args.query)
    const k = typeof args.k === "number" && Number.isFinite(args.k) && args.k > 0 ? Math.floor(args.k) : 5

    const db = new Database(dbPath, { readonly: true })
    // Vectors from a different embedding model are not comparable with this
    // tool's query vectors — rankings would be silently wrong. Refuse instead.
    const hasMeta = db.query("SELECT 1 FROM sqlite_master WHERE name = 'meta'").get()
    const indexModel = hasMeta
      ? (db.query("SELECT value FROM meta WHERE key = 'embedding_model'").get() as { value: string } | null)?.value
      : null
    if (indexModel !== MODEL) {
      db.close()
      return "This matter's search index was built by an older embedding model. Restart the ingest service to migrate it, then search again."
    }
    // Preserve matter-wide indexed-chain warnings and penalties for scoped passages.
    const chains = buildChains(db)
    const superseded = supersededSections(db, chains)
    const channels = [
      vectorChannel(db, queryVec, args),
      lexicalChannel(db, args.query, args),
      phraseChannel(db, args.query, args),
    ]

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
      .map(({ row, score }) => {
        const normalized = score / (channels.length / (RRF_K + 1))
        const op = superseded.get(row.doc_path)?.get(row.section)
        return op ? { row, score: normalized * SUPERSESSION_PENALTY, supersededBy: op } : { row, score: normalized, supersededBy: undefined }
      })
      .sort((a, b) => b.score - a.score)
      .slice(0, k)

    // Code-point position of each superseding statement, so the warning says
    // where to read it. Each amending document is rebuilt once per call.
    const amenderTexts = new Map<string, IndexedText>()
    const supersessionReads = new Map(
      ranked.flatMap(({ row, supersededBy }) =>
        supersededBy
          ? [[row.id, operativePassage(db, supersededBy.op, supersededBy.amenderOps, amenderTexts).read] as const]
          : [],
      ),
    )
    const definitionsNote = attachDefinitions(db, ranked, args)
    const crossRefsNote = attachCrossRefs(db, ranked, args)
    db.close()

    // Compare the exact chunks, not their section labels: one section may span
    // several independently useful chunks. Empty searches are not loops because
    // no passage was returned, and a differently phrased query may still find
    // something useful.
    const signature = ranked
      .map(({ row }) => row.id)
      .toSorted((a, b) => a - b)
      .join("|")
    const recent = recentBySession.get(ctx.sessionID) ?? []
    if (signature && recent.includes(signature)) {
      console.warn(
        "[search-document] repeated result-set; search loop short-circuited",
        { session_ref: draftingTraceReference(ctx.sessionID), query_ref: draftingTraceReference(args.query), result_count: ranked.length },
      )
      return "These passages were already returned by an earlier search this turn — the same results matched again, so searching further will not surface anything new. Answer from the passages you have already retrieved; if they do not address the question, say the documents do not cover it. Do not repeat this search."
    }
    if (!recentBySession.has(ctx.sessionID) && recentBySession.size >= SESSION_LIMIT) {
      const oldest = recentBySession.keys().next().value
      if (oldest !== undefined) recentBySession.delete(oldest)
    }
    if (signature) recentBySession.set(ctx.sessionID, [signature, ...recent].slice(0, RECENT_LIMIT))

    const citations = ranked.map(({ row, score, supersededBy }) => ({
      documentName: row.doc_name,
      docPath: row.doc_path,
      section: row.section,
      // Empty on chunks ingested before structure-aware chunking; formatCitations
      // falls back to "doc § section" when the path is absent.
      structuralPath: row.structural_path || undefined,
      party: row.party || undefined,
      excerpt: row.text,
      charStart: row.char_start,
      charEnd: row.char_end,
      score,
      // Carried through to formatCitations, which warns the model before the
      // excerpt that ingest flagged this passage as instruction-like (issue #17).
      flagged: row.flagged === 1,
      superseded_by: supersededBy
        ? {
            amenderName: supersededBy.amender,
            operation: supersededBy.op.operation,
            charStart: supersededBy.op.char_start,
            charEnd: supersededBy.op.char_end,
            read: supersessionReads.get(row.id)!,
          }
        : undefined,
    }))

    // Surface proposals already pending on the cited documents. The retrieved
    // passages reflect the clean (accepted) document on disk — they do NOT include
    // edits the assistant proposed earlier in this review or revision but that are not yet
    // accepted. Listing them keeps the model from re-proposing or contradicting a
    // change it already made, and lets it compose new edits against the running state.
    // Proposals carry document/anchor identity, but no stored structural path
    // or span. Their text cannot be safely attached to a strict path scope.
    const pending = args.structural_path === undefined
      ? [...new Set(citations.map((c) => c.docPath))].flatMap((docPath) =>
          pendingRedlinesForDoc(ctx.directory, docPath).map((r) => ({ docName: path.basename(docPath), ...r })),
        )
      : []
    const pendingNote = args.structural_path !== undefined
      ? "\n\nPending redline context omitted for this structural scope; pending proposals have not been assessed."
      : pending.length
      ? `\n\nPending redlines on these documents (proposed but not yet accepted — not reflected in the passages above):\n` +
        pending.map((r) => `- #${r.id} (${r.author}) in ${r.docName}: proposes "${r.new_text}"`).join("\n")
      : ""

    return {
      title: `${citations.length} passage(s) for "${args.query}"`,
      output: (formatCitations(citations) || "No relevant passages found.") + definitionsNote + crossRefsNote + pendingNote,
      metadata: {
        citations,
        pending,
        ...(args.structural_path === undefined
          ? {}
          : { pending_context: { status: "omitted", reason: "unassessed_for_structural_scope" } }),
      },
    }
  },
})
