// Shared citation shape and rendering, used by the tools that produce citations
// (search-document's vector hits, cite's anchored quotations) and the legal
// plugin (which re-renders the surviving set after span verification). One
// formatter keeps every output identical — including the untrusted-data framing,
// so no path can hand the model a document excerpt without it.

import type { Database } from "bun:sqlite"
import { FLAGGED_PASSAGE_NOTE, UNTRUSTED_EXCERPT_NOTICE } from "./untrusted"

const SUPERSEDED_PASSAGE_NOTE =
  "[⚠ SUPERSEDED] A later amendment restates, amends, or deletes this passage's section, so it may not be the current text."

// Ingest's extraction quality gate (services/ingest/src/extraction-quality.ts)
// marked the cited document degraded: its text was indexed but may be
// incomplete or garbled.
const DEGRADED_EXTRACTION_NOTE =
  "[⚠ DEGRADED EXTRACTION] Ingest found this document's extracted text may be incomplete or garbled"

// Each degraded source's tripped signal details, by doc_path, from the
// documents.extraction_quality column ingest writes. Sources that are ok, have
// no readable text (and so no passages) or predate the gate are absent. A
// matter database ingest has not opened since the gate (its startup migration
// failed) lacks the column and has no degraded sources yet.
export function degradedExtractions(db: Database): Map<string, string[]> {
  const columns = db.query("SELECT name FROM pragma_table_info('documents')").all() as { name: string }[]
  if (!columns.some((column) => column.name === "extraction_quality")) return new Map()
  const rows = db
    .query("SELECT doc_path, extraction_quality FROM documents WHERE extraction_quality IS NOT NULL")
    .all() as { doc_path: string; extraction_quality: string }[]
  return new Map(
    rows.flatMap((row) => {
      const quality = JSON.parse(row.extraction_quality) as { status: string; signals?: { detail: string }[] }
      return quality.status === "degraded" ? [[row.doc_path, (quality.signals ?? []).map((signal) => signal.detail)] as const] : []
    }),
  )
}

const DEFAULT_EXCERPT_DISPLAY_CHARS = 1000

export type DocumentCitation = {
  documentName: string
  docPath: string
  section: string
  // Full clause hierarchy ("Doc › §7 Indemnification › §7.2 Cap") and the
  // document's parties, attached by ingest's structure-aware chunking. Both are
  // optional — citations from sources without them (older indexes, the cite
  // tool) simply omit the line.
  structuralPath?: string
  party?: string
  excerpt: string
  charStart: number
  charEnd: number
  // Vector hits from search-document carry a similarity score; quotations the
  // cite tool anchored do not, so score is optional.
  score?: number
  reason?: string
  confidence?: number
  context?: string
  verified?: boolean
  reanchored?: boolean
  // Set when ingest flagged the chunk as containing instruction-like text
  // (see services/ingest/src/sanitize.ts), so the rendering can warn the model
  // before it reads the excerpt.
  flagged?: boolean
  // The signals ingest's extraction quality gate tripped when the cited
  // document's text is degraded; absent when it is ok or was never assessed.
  degradedExtraction?: string[]
  // Set when this chunk belongs to a section of an amended base that a later
  // amender restates, amends, or deletes (issue #69). The renderer warns the
  // model above the excerpt; retrieval also down-ranks the chunk so the
  // amender's operative version outranks it.
  superseded_by?: {
    amenderName: string
    operation: "amend" | "restate" | "delete" | "add"
    charStart: number
    charEnd: number
    // Position (Unicode code points) of the amender's operative statement and
    // the text after it; offset and limit are null when the indexed text is
    // incomplete and the position cannot be counted safely.
    read?: { document: string; offset: number | null; limit: number | null }
  }
}

// `full` shows every excerpt unclipped, for a job that cannot read further
// (the review grid's re-check sees only the passages it is given).
export function formatCitations(citations: DocumentCitation[], options: { full?: boolean } = {}) {
  if (!citations.length) return ""
  const body = citations
    .map((c, i) => {
      const score = c.score === undefined ? "" : ` (score ${c.score.toFixed(3)})`
      const flag = c.flagged ? `\n${FLAGGED_PASSAGE_NOTE}` : ""
      const degraded = c.degradedExtraction
        ? `\n${DEGRADED_EXTRACTION_NOTE} (${c.degradedExtraction.join("; ")}). Check it against the original before relying on it, and do not treat a term missing from it as absent from the document.`
        : ""
      const superseded = c.superseded_by
        ? `\n${SUPERSEDED_PASSAGE_NOTE}\n   superseded by: ${c.superseded_by.amenderName} (${c.superseded_by.operation}, chars ${c.superseded_by.charStart}..${c.superseded_by.charEnd})` +
          (c.superseded_by.read ? `\n   ${supersededReadHint(c.superseded_by.read)}` : "")
        : ""
      const reason =
        c.reason === undefined
          ? ""
          : `\n   reason: ${c.reason}${c.confidence === undefined ? "" : ` (confidence ${c.confidence}/5)`}`
      // Prefer the full structural path over the bare section label when ingest
      // supplied one; append the parties so the passage carries its matter context.
      const heading = c.structuralPath || `${c.documentName} § ${c.section}`
      const party = c.party ? `\n   parties: ${c.party}` : ""
      return `${i + 1}. [${heading}]${score}${flag}${degraded}${superseded}${party}\n${options.full ? c.excerpt : displayExcerpt(c.excerpt)}${reason}`
    })
    .join("\n\n")
  return `${UNTRUSTED_EXCERPT_NOTICE}\n\n${body}`
}

// Tool-neutral: names the amending document and where its operative wording
// sits, without assuming which reading tool the agent has.
function supersededReadHint(read: { document: string; offset: number | null; limit: number | null }) {
  const at =
    read.offset === null || read.limit === null
      ? `; its position is unavailable because the indexed text is incomplete.`
      : ` at code-point offset ${read.offset} (${read.limit} characters).`
  return `before quoting this passage as current, read the amendment's operative wording in ${read.document}${at}`
}

function displayExcerpt(excerpt: string) {
  const max = citationExcerptDisplayChars()
  if (excerpt.length <= max) return excerpt
  return (
    excerpt.slice(0, max).trimEnd() +
    `\n[passage preview clipped at ${max} characters for context budget; use cite or read-document for the full verified passage before quoting beyond this preview.]`
  )
}

function citationExcerptDisplayChars() {
  const value = Number(process.env.DOCHAUS_CITATION_EXCERPT_CHARS ?? DEFAULT_EXCERPT_DISPLAY_CHARS)
  if (!Number.isFinite(value) || value < 400) return DEFAULT_EXCERPT_DISPLAY_CHARS
  return Math.floor(value)
}
